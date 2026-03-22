"""
Entity Classifier
=================
Maps parsed DXF entities to MEP disciplines and BOQ element categories
using configurable rules from config.json.
"""

import logging
from dataclasses import dataclass, field

from .parser import ParsedDrawing, BlockInstance, LineEntity

logger = logging.getLogger(__name__)

# The four MEP disciplines
DISCIPLINES = ["ELECTRICAL", "PLUMBING", "HVAC", "FIRE_FIGHTING"]


@dataclass
class BOQItem:
    """A single line item in the Bill of Quantities."""
    description: str
    unit: str
    quantity: float
    element: str  # e.g., "Lighting Fixtures", "Sanitary Fixtures"
    discipline: str  # e.g., "ELECTRICAL", "PLUMBING"
    source_layer: str = ""
    source_blocks: list[str] = field(default_factory=list)


@dataclass
class ClassifiedResult:
    """All BOQ items organized by discipline and element."""
    items: list[BOQItem] = field(default_factory=list)
    unclassified_blocks: dict[str, int] = field(default_factory=dict)
    unclassified_layers: set[str] = field(default_factory=set)

    def by_discipline(self) -> dict[str, list[BOQItem]]:
        """Group items by discipline."""
        result: dict[str, list[BOQItem]] = {d: [] for d in DISCIPLINES}
        for item in self.items:
            result.setdefault(item.discipline, []).append(item)
        return result

    def by_discipline_and_element(self) -> dict[str, dict[str, list[BOQItem]]]:
        """Group items by discipline then element."""
        result: dict[str, dict[str, list[BOQItem]]] = {}
        for item in self.items:
            disc = result.setdefault(item.discipline, {})
            elem = disc.setdefault(item.element, [])
            elem.append(item)
        return result


    # Architectural/structural layers and blocks to exclude from BOQ
EXCLUDE_LAYER_KEYWORDS = [
    "WALL", "DOOR", "WINDOW", "SLAB", "BEAM", "COLUMN", "STAIR",
    "ROOF", "RAMP", "FURNITURE", "MEUBLES", "ALUMINUM", "GLASS",
    "TILE", "WOODEN", "CEILING", "STRUCTURE", "STEEL", "FRAME",
    "AXES", "GRID", "TITLE", "BORDER", "ANNO", "TEXT", "DIM",
    "DEFPOINTS", "HATCH", "VIEWPORT", "PROJ", "VOITURE", "CAR",
    "S-GRID", "G-ANNO", "A-WALL", "A-DOOR", "A-GLAZ", "A-BLDG",
    "A-DETL", "A-ANNO", "V-CTRL", "V-NODE", "V-ROAD", "V-SITE",
    "C-ROAD", "C-SSWR", "C-STRM", "C-TOPO", "C-WATR", "C-PROP",
    "BMED-MEDICAL EQUIPMENT", "BMED-MEDICAL FURNITURE", "BMED-NON",
    "INTERIOR", "EXTERNAL WALLS", "Reflective", "False Ceiling",
]

EXCLUDE_BLOCK_KEYWORDS = [
    "armchair", "bar stool", "cabinet", "office chair", "desk",
    "sofa", "table", "chair", "couch", "bed", "wardrobe", "shelf",
    "bookcase", "drawer", "lamp_decor", "plant", "rug", "curtain",
    "beam", "column", "slab", "wall_section", "ramp", "stair",
    "door_", "window_", "mullion", "panel_wall", "grid",
    "Seat", "SIEGE", "FAUT", "PORTE", "lav-manta", "WC manta",
    "voiture", "car_", "tree", "person", "people", "north_arrow",
    "title_block", "revision_", "section_head", "elev_marker",
]


class EntityClassifier:
    """Classifies parsed DXF entities into MEP BOQ categories."""

    # Drawing units are typically millimeters in these files
    # Convert to meters for linear BOQ items
    DRAWING_UNIT_TO_METER = 0.001  # mm -> m

    def __init__(self, config: dict):
        self.layer_map = config.get("layer_discipline_map", {})
        self.block_catalog = config.get("block_catalog", {})

    def classify(self, drawings: list[ParsedDrawing]) -> ClassifiedResult:
        """Classify entities from one or more parsed drawings."""
        result = ClassifiedResult()

        # Aggregate blocks across all drawings, filtering out architectural entities
        all_blocks: dict[str, list[BlockInstance]] = {}
        all_lines_by_layer: dict[str, float] = {}

        for drawing in drawings:
            for block in drawing.blocks:
                if self._is_excluded_block(block.name, block.layer):
                    continue
                if self._is_junk_block(block.name):
                    continue
                bl = all_blocks.setdefault(block.name, [])
                bl.append(block)
            for line in drawing.lines:
                if self._is_excluded_layer(line.layer):
                    continue
                all_lines_by_layer[line.layer] = (
                    all_lines_by_layer.get(line.layer, 0) + line.length
                )

        # Step 1: Classify blocks using the block catalog
        # Each block can only be assigned to ONE element (best match wins)
        classified_block_names: set[str] = set()
        # Collect all candidate matches, pick the best one per block
        block_candidates: dict[str, tuple[str, str, str, str, list]] = {}
        # key=block_name, value=(discipline, element, unit, desc_prefix, catalog_entry)

        for discipline, elements in self.block_catalog.items():
            for element_name, element_def in elements.items():
                catalog_blocks = element_def.get("blocks", [])
                unit = element_def.get("unit", "NR")
                desc_prefix = element_def.get("description_prefix", "Supply and install")

                for block_name in all_blocks:
                    match_score = self._block_match_score(block_name, catalog_blocks)
                    if match_score > 0:
                        existing = block_candidates.get(block_name)
                        if existing is None or match_score > existing[5]:
                            block_candidates[block_name] = (
                                discipline, element_name, unit, desc_prefix,
                                all_blocks[block_name], match_score
                            )

        for block_name, (discipline, element_name, unit, desc_prefix, instances, _score) in block_candidates.items():
            classified_block_names.add(block_name)
            count = len(instances)
            desc = self._build_description(desc_prefix, block_name, instances)
            layer = instances[0].layer if instances else ""

            result.items.append(BOQItem(
                description=desc,
                unit=unit,
                quantity=count,
                element=element_name,
                discipline=discipline,
                source_layer=layer,
                source_blocks=[block_name],
            ))

        # Step 2: Classify remaining blocks by layer name
        for block_name, instances in all_blocks.items():
            if block_name in classified_block_names:
                continue

            layers = {inst.layer for inst in instances}
            discipline = None
            for layer in layers:
                discipline = self._classify_layer(layer)
                if discipline:
                    break

            if discipline:
                classified_block_names.add(block_name)
                count = len(instances)
                # Use block name to infer element more precisely
                element = self._infer_element_from_block_and_layer(block_name, layers)

                result.items.append(BOQItem(
                    description=block_name,
                    unit="NR",
                    quantity=count,
                    element=element,
                    discipline=discipline,
                    source_layer=next(iter(layers)),
                    source_blocks=[block_name],
                ))
            else:
                # Truly unclassified
                result.unclassified_blocks[block_name] = len(instances)
                result.unclassified_layers.update(layers)

        # Step 3: Classify linear entities (pipes, ducts, cable trays) by layer
        # Convert drawing units (typically mm) to meters
        for layer, total_length in all_lines_by_layer.items():
            discipline = self._classify_layer(layer)
            if not discipline:
                continue

            # Determine if this is pipe, duct, cable tray, etc.
            item_desc, unit, element = self._classify_linear_by_layer(layer)
            if item_desc:
                length_meters = total_length * self.DRAWING_UNIT_TO_METER
                result.items.append(BOQItem(
                    description=f"{item_desc} (on layer: {layer})",
                    unit=unit,
                    quantity=round(length_meters, 2),
                    element=element,
                    discipline=discipline,
                    source_layer=layer,
                ))

        # Merge duplicate items (same description + element + discipline)
        result.items = self._merge_duplicates(result.items)

        logger.info(
            f"Classified: {len(result.items)} BOQ items, "
            f"{len(result.unclassified_blocks)} unclassified block types"
        )
        return result

    def _is_excluded_layer(self, layer_name: str) -> bool:
        """Check if a layer is architectural/structural and should be excluded."""
        layer_upper = layer_name.upper()
        for kw in EXCLUDE_LAYER_KEYWORDS:
            if kw.upper() in layer_upper:
                return True
        return False

    def _is_excluded_block(self, block_name: str, layer_name: str) -> bool:
        """Check if a block should be excluded (architectural/structural)."""
        if self._is_excluded_layer(layer_name):
            # Only exclude if the block is NOT in the MEP catalog
            bn_lower = block_name.lower()
            for discipline, elements in self.block_catalog.items():
                for element_def in elements.values():
                    for entry in element_def.get("blocks", []):
                        if entry.lower() in bn_lower or bn_lower in entry.lower():
                            return False  # Keep MEP blocks even on architectural layers
            return True
        bn_lower = block_name.lower()
        for kw in EXCLUDE_BLOCK_KEYWORDS:
            if kw.lower() in bn_lower:
                return True
        return False

    # Known MEP abbreviations that should NOT be filtered as junk
    KNOWN_MEP_NAMES = {
        "pump", "tank", "fan", "ahu", "fcu", "ats", "mdb", "smdb", "sdb",
        "led", "ups", "dvr", "nvr", "rccb", "mccb", "mcb", "pipe", "vrf",
        "db", "fhr", "wc", "pvc", "ppr", "cpvc", "ahu", "fau", "eow",
        "eos", "lv", "hv", "mv", "ct", "ac", "tv", "pa",
    }

    @staticmethod
    def _is_junk_block(name: str) -> bool:
        """Filter out junk/noise block names that aren't real MEP equipment."""
        import re
        # Skip AutoCAD anonymous/internal blocks
        if name.startswith("A$C") or name.startswith("*") or name.startswith("_"):
            return True
        # Skip very short names (1-2 chars) - too ambiguous
        if len(name) <= 2:
            return True
        # Skip Revit internal blocks
        if name.startswith("Aecb_") or name.startswith("Rpt"):
            return True
        # Skip Revit DVM/system internal markers
        if name in ("INOUT_MARK", "AC_RF"):
            return True

        # Detect random/gibberish block names (both uppercase and lowercase)
        clean = name.replace("_", "").replace("-", "").replace(" ", "")
        clean_lower = clean.lower()

        # If it's a known MEP name, keep it
        if clean_lower in EntityClassifier.KNOWN_MEP_NAMES:
            return False

        # Random alphanumeric strings with no meaningful structure
        if re.match(r'^[A-Z]{4,}$', clean) and len(clean) <= 12:
            # All-caps random string (FEFEFE, GTJYUKJU, SFEFEV, etc.)
            # Allow known patterns: CCTV, HVAC, etc.
            if clean not in ("CCTV", "HVAC", "RCCB", "MCCB", "SMDB", "SMATV",
                             "PUMP", "TANK", "FIRE", "DATA", "SPEAKER", "CABLE",
                             "LIGHT", "POWER", "SHOWER", "BASIN", "DRAIN"):
                # Check for repeating chars or lack of vowels (gibberish indicator)
                vowel_count = sum(1 for c in clean_lower if c in 'aeiou')
                unique_ratio = len(set(clean_lower)) / len(clean_lower)
                vowel_ratio = vowel_count / len(clean_lower)
                if vowel_ratio < 0.3 or unique_ratio < 0.75:
                    return True
        # Lowercase gibberish
        if re.match(r'^[a-z]{4,}$', clean) and len(clean) <= 12:
            vowel_count = sum(1 for c in clean if c in 'aeiou')
            unique_ratio = len(set(clean)) / len(clean)
            vowel_ratio = vowel_count / len(clean)
            if vowel_ratio < 0.3 or unique_ratio < 0.75:
                return True
        # Mixed random alphanumeric (like "4y55r8u68", "CDCFR45667")
        if re.match(r'^[A-Za-z0-9]{4,12}$', clean):
            digit_count = sum(1 for c in clean if c.isdigit())
            alpha_count = sum(1 for c in clean if c.isalpha())
            if digit_count >= 2 and alpha_count >= 2:
                # Could be random. Check if it looks like a model number (has structure)
                if not re.search(r'[A-Z]{2,}\d+|[A-Z]+[-_ ]\d+', name):
                    # No model number pattern - likely junk
                    if len(clean) <= 10:
                        return True
        return False

    def _block_match_score(self, block_name: str, catalog_entries: list[str]) -> int:
        """
        Score how well a block name matches catalog entries.
        Returns 0 for no match, higher = better match.
        Requires minimum 3-character overlap to avoid false positives.
        """
        bn_lower = block_name.lower()
        best_score = 0
        for entry in catalog_entries:
            entry_lower = entry.lower()
            # Exact match (highest priority)
            if bn_lower == entry_lower:
                return 1000
            # Catalog entry contained in block name
            if entry_lower in bn_lower and len(entry_lower) >= 3:
                score = len(entry_lower) * 2
                best_score = max(best_score, score)
            # Block name contained in catalog entry (only if block name is meaningful)
            elif bn_lower in entry_lower and len(bn_lower) >= 4:
                score = len(bn_lower)
                best_score = max(best_score, score)
        return best_score

    def _classify_layer(self, layer_name: str) -> str | None:
        """Determine which discipline a layer belongs to."""
        layer_upper = layer_name.upper()
        for discipline, keywords in self.layer_map.items():
            for kw in keywords:
                if kw.upper() in layer_upper:
                    return discipline
        return None

    def _infer_element_from_block_and_layer(self, block_name: str, layers: set[str]) -> str:
        """Infer element from both block name content and layer name."""
        bn_upper = block_name.upper()

        # Pipe fittings (Elbow, Tee, Flange, Reducer, Cap, Nipple, etc.)
        fitting_keywords = ["ELBOW", "TEE ", "TEE-", "FLANGE", "REDUCER",
                            "CAP -", "NIPPLE", "COUPLING", "UNION", "VALVE",
                            "STRAINER", "SOCKET REDUCING", "BRANCH TAKEOFF"]
        if any(kw in bn_upper for kw in fitting_keywords):
            return "Pipes & Fittings"

        # Specific block name patterns
        if "SPRINKLER" in bn_upper:
            return "Sprinklers"
        if "ALARM" in bn_upper and "FIRE" not in bn_upper:
            return "Fire Alarm System"
        if any(kw in bn_upper for kw in ["PUMP", "BOOSTER"]):
            return "Equipment"
        if any(kw in bn_upper for kw in ["GUN SPRAY", "NOZZLE"]):
            return "Fire Fighting Equipment"
        if "LEVEL" in bn_upper and "HEAD" in bn_upper:
            return "Equipment"

        return self._infer_element_from_layer(layers)

    def _infer_element_from_layer(self, layers: set[str]) -> str:
        """Infer a BOQ element name from layer name(s)."""
        combined = " ".join(layers).upper()

        if any(kw in combined for kw in ["LIGHT", "LAMP", "LED"]):
            return "Lighting Fixtures"
        if any(kw in combined for kw in ["POWER", "SOCKET", "PWR"]):
            return "Power Outlets"
        if any(kw in combined for kw in ["CABLE TRAY", "CT-"]):
            return "Cable Trays"
        if any(kw in combined for kw in ["DATA", "PHONE", "TEL"]):
            return "Data Communication"
        if any(kw in combined for kw in ["FIRE ALARM", "SMOKE", "E-FIRE"]):
            return "Fire Alarm System"
        if any(kw in combined for kw in ["CCTV", "CAMERA"]):
            return "CCTV System"
        if any(kw in combined for kw in ["PA", "SPEAKER", "ADDRESS"]):
            return "Public Address"
        if any(kw in combined for kw in ["DUCT", "H-SUPPLY", "M-DUCT"]):
            return "Ductwork"
        if any(kw in combined for kw in ["GRILLE", "DIFFUSER", "M-TERM"]):
            return "Diffusers & Grilles"
        if any(kw in combined for kw in ["DVM", "FCU", "YORK", "SPLIT"]):
            return "AC Units"
        if any(kw in combined for kw in ["VENT", "FAN", "EXHAUST"]):
            return "Ventilation Equipment"
        if any(kw in combined for kw in ["WATER", "PPR", "CPVC"]):
            return "Pipes & Fittings"
        if any(kw in combined for kw in ["SEWER", "DRAIN", "SANIT"]):
            return "Sanitary Fixtures"
        if any(kw in combined for kw in ["HOSE", "FHR"]):
            return "Fire Hose Reels"
        if any(kw in combined for kw in ["SPRINKLER"]):
            return "Sprinklers"
        if any(kw in combined for kw in ["LANDING VALVE"]):
            return "Landing Valves"
        if any(kw in combined for kw in ["FIRE FIGHT", "FF-"]):
            return "Fire Fighting Equipment"

        return "Other Items"

    def _classify_linear_by_layer(self, layer: str) -> tuple[str | None, str, str]:
        """
        Classify a linear entity layer into description, unit, and element.
        Returns (description, unit, element) or (None, '', '') if not relevant.
        """
        layer_upper = layer.upper()

        if any(kw in layer_upper for kw in ["DUCT", "H-SUPPLY", "H-FLEXIBLE", "M-DUCT", "M-FLEX"]):
            return "Ductwork", "M", "Ductwork"
        if any(kw in layer_upper for kw in ["DVM_PIPE"]):
            return "Refrigerant piping", "M", "Refrigerant Piping"
        if any(kw in layer_upper for kw in ["CABLE TRAY"]):
            return "Cable tray", "M", "Cable Trays"
        if any(kw in layer_upper for kw in ["WATER SUPPLY"]):
            return "Water supply piping", "M", "Pipes & Fittings"
        if any(kw in layer_upper for kw in ["SEWERAGE", "DRAINAGE"]):
            return "Drainage/sewage piping", "M", "Soil & Vent System"
        if any(kw in layer_upper for kw in ["RAIN WATER"]):
            return "Rain water piping", "M", "Rain Water System"
        if any(kw in layer_upper for kw in ["FIRE FIGHT", "FIRE HOSE", "FF-"]):
            return "Fire fighting piping", "M", "Fire Fighting Pipes"

        return None, "", ""

    def _build_description(
        self, prefix: str, block_name: str, instances: list[BlockInstance]
    ) -> str:
        """Build a human-readable description for a BOQ item."""
        # Clean up block name
        clean_name = block_name.replace("_", " ").replace("-", " ").strip()

        # Check if instances have useful attributes
        attr_info = ""
        if instances:
            sample_attrs = instances[0].attributes
            useful_keys = [k for k in sample_attrs if k not in ("", "0")]
            if useful_keys:
                parts = [f"{k}: {sample_attrs[k]}" for k in useful_keys[:3]]
                attr_info = f" ({', '.join(parts)})"

        return f"{prefix} {clean_name}{attr_info}"

    def _merge_duplicates(self, items: list[BOQItem]) -> list[BOQItem]:
        """Merge items with the same description, element, and discipline."""
        merged: dict[str, BOQItem] = {}
        for item in items:
            key = (item.description, item.element, item.discipline, item.unit)
            if key in merged:
                merged[key].quantity += item.quantity
                merged[key].source_blocks.extend(item.source_blocks)
            else:
                merged[key] = BOQItem(
                    description=item.description,
                    unit=item.unit,
                    quantity=item.quantity,
                    element=item.element,
                    discipline=item.discipline,
                    source_layer=item.source_layer,
                    source_blocks=list(item.source_blocks),
                )
        return list(merged.values())
