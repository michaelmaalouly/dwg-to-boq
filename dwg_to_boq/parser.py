"""
DXF Parser
==========
Parses DXF files using ezdxf to extract MEP entities: blocks, text,
attributes, lines, and polylines organized by layer.
"""

import math
import logging
from dataclasses import dataclass, field

import ezdxf

logger = logging.getLogger(__name__)


@dataclass
class BlockInstance:
    """A block reference (INSERT) found in the drawing."""
    name: str
    layer: str
    x: float
    y: float
    attributes: dict[str, str] = field(default_factory=dict)
    count: int = 1


@dataclass
class TextEntity:
    """A text or mtext entity found in the drawing."""
    content: str
    layer: str
    x: float
    y: float
    height: float = 0.0


@dataclass
class LineEntity:
    """A line, polyline, or arc entity with a measurable length."""
    entity_type: str
    layer: str
    length: float  # in drawing units


@dataclass
class ParsedDrawing:
    """All extracted data from a single DXF file."""
    source_file: str
    layers: dict[str, dict] = field(default_factory=dict)
    blocks: list[BlockInstance] = field(default_factory=list)
    texts: list[TextEntity] = field(default_factory=list)
    lines: list[LineEntity] = field(default_factory=list)

    def block_counts_by_layer(self) -> dict[str, dict[str, int]]:
        """Group block counts by layer then block name."""
        result: dict[str, dict[str, int]] = {}
        for b in self.blocks:
            layer_blocks = result.setdefault(b.layer, {})
            layer_blocks[b.name] = layer_blocks.get(b.name, 0) + b.count
        return result

    def total_block_counts(self) -> dict[str, int]:
        """Total count of each block name across all layers."""
        counts: dict[str, int] = {}
        for b in self.blocks:
            counts[b.name] = counts.get(b.name, 0) + b.count
        return counts

    def line_lengths_by_layer(self) -> dict[str, float]:
        """Total line length per layer."""
        lengths: dict[str, float] = {}
        for ln in self.lines:
            lengths[ln.layer] = lengths.get(ln.layer, 0) + ln.length
        return lengths


class DXFParser:
    """Parses DXF files and extracts MEP-relevant entities."""

    # Layers to skip (architectural/structural background)
    SKIP_LAYER_KEYWORDS = [
        "Defpoints", "0",
    ]

    def parse(self, dxf_path: str) -> ParsedDrawing:
        """Parse a DXF file and return structured data."""
        logger.info(f"Parsing: {dxf_path}")
        try:
            doc = ezdxf.readfile(dxf_path)
        except Exception as e:
            if "sort handle" in str(e).lower() or "331" in str(e):
                # Strip problematic SORTENTSTABLE objects and retry
                logger.warning(f"Patching sort handle issue in: {dxf_path}")
                patched_path = self._patch_sort_handles(dxf_path)
                try:
                    doc = ezdxf.readfile(patched_path)
                except Exception:
                    from ezdxf import recover
                    doc, auditor = recover.readfile(patched_path)
                    if auditor.has_errors:
                        logger.warning(f"Recovered with {len(auditor.errors)} errors")
            else:
                from ezdxf import recover
                doc, auditor = recover.readfile(dxf_path)
                if auditor.has_errors:
                    logger.warning(f"Recovered with {len(auditor.errors)} errors: {dxf_path}")
        msp = doc.modelspace()

        drawing = ParsedDrawing(source_file=dxf_path)

        # Extract layer info
        for layer in doc.layers:
            drawing.layers[layer.dxf.name] = {
                "color": layer.dxf.color,
                "linetype": layer.dxf.linetype if layer.dxf.hasattr("linetype") else "Continuous",
                "is_off": layer.is_off(),
                "is_frozen": layer.is_frozen(),
            }

        # Extract entities from modelspace
        for entity in msp:
            try:
                self._process_entity(entity, drawing)
            except Exception as e:
                logger.debug(f"Skipping entity {entity.dxftype()}: {e}")

        logger.info(
            f"Parsed: {len(drawing.blocks)} blocks, "
            f"{len(drawing.texts)} texts, "
            f"{len(drawing.lines)} line entities"
        )
        return drawing

    @staticmethod
    def _patch_sort_handles(dxf_path: str) -> str:
        """Remove SORTENTSTABLE objects that cause parsing errors."""
        import re
        import tempfile

        with open(dxf_path, "r", errors="replace") as f:
            content = f.read()

        # Remove SORTENTSTABLE objects (from "0\nSORTENTSTABLE" to the next "0\n<TYPE>")
        # These objects contain code 331 entries that ezdxf can't handle in some versions
        patched = re.sub(
            r'  0\nSORTENTSTABLE\n.*?(?=  0\n[A-Z])',
            '',
            content,
            flags=re.DOTALL,
        )

        patched_path = dxf_path.replace(".dxf", "_patched.dxf")
        with open(patched_path, "w") as f:
            f.write(patched)

        return patched_path

    def _process_entity(self, entity, drawing: ParsedDrawing):
        """Route entity to the appropriate handler."""
        etype = entity.dxftype()
        layer = entity.dxf.layer if entity.dxf.hasattr("layer") else "0"

        if etype == "INSERT":
            self._process_insert(entity, layer, drawing)
        elif etype in ("MTEXT", "TEXT"):
            self._process_text(entity, layer, drawing)
        elif etype == "LINE":
            self._process_line(entity, layer, drawing)
        elif etype == "LWPOLYLINE":
            self._process_lwpolyline(entity, layer, drawing)
        elif etype == "POLYLINE":
            self._process_polyline(entity, layer, drawing)
        elif etype == "ARC":
            self._process_arc(entity, layer, drawing)
        elif etype == "CIRCLE":
            self._process_circle(entity, layer, drawing)

    def _process_insert(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract block reference and its attributes."""
        block_name = entity.dxf.name
        # Skip anonymous blocks (start with *)
        if block_name.startswith("*"):
            return

        pos = entity.dxf.insert
        attrs = {}
        if entity.attribs:
            for attrib in entity.attribs:
                tag = attrib.dxf.tag
                value = attrib.dxf.text
                if tag and value:
                    attrs[tag] = value

        drawing.blocks.append(BlockInstance(
            name=block_name,
            layer=layer,
            x=pos.x,
            y=pos.y,
            attributes=attrs,
        ))

    def _process_text(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract text content."""
        if entity.dxftype() == "MTEXT":
            content = entity.plain_text()
        else:
            content = entity.dxf.text

        if not content or not content.strip():
            return

        pos = entity.dxf.insert if entity.dxf.hasattr("insert") else (0, 0, 0)
        height = entity.dxf.char_height if entity.dxf.hasattr("char_height") else 0

        drawing.texts.append(TextEntity(
            content=content.strip(),
            layer=layer,
            x=pos[0] if hasattr(pos, '__getitem__') else getattr(pos, 'x', 0),
            y=pos[1] if hasattr(pos, '__getitem__') else getattr(pos, 'y', 0),
            height=height,
        ))

    def _process_line(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract a LINE entity with its length."""
        start = entity.dxf.start
        end = entity.dxf.end
        length = math.sqrt(
            (end.x - start.x) ** 2 +
            (end.y - start.y) ** 2 +
            (end.z - start.z) ** 2
        )
        if length > 0:
            drawing.lines.append(LineEntity(
                entity_type="LINE",
                layer=layer,
                length=length,
            ))

    def _process_lwpolyline(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract a lightweight polyline with total length."""
        try:
            length = 0.0
            points = list(entity.get_points(format="xy"))
            for i in range(len(points) - 1):
                dx = points[i + 1][0] - points[i][0]
                dy = points[i + 1][1] - points[i][1]
                length += math.sqrt(dx * dx + dy * dy)
            if entity.closed and len(points) > 1:
                dx = points[0][0] - points[-1][0]
                dy = points[0][1] - points[-1][1]
                length += math.sqrt(dx * dx + dy * dy)
            if length > 0:
                drawing.lines.append(LineEntity(
                    entity_type="LWPOLYLINE",
                    layer=layer,
                    length=length,
                ))
        except Exception:
            pass

    def _process_polyline(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract a 2D/3D polyline with total length."""
        try:
            points = [(v.dxf.location.x, v.dxf.location.y) for v in entity.vertices]
            length = 0.0
            for i in range(len(points) - 1):
                dx = points[i + 1][0] - points[i][0]
                dy = points[i + 1][1] - points[i][1]
                length += math.sqrt(dx * dx + dy * dy)
            if entity.is_closed and len(points) > 1:
                dx = points[0][0] - points[-1][0]
                dy = points[0][1] - points[-1][1]
                length += math.sqrt(dx * dx + dy * dy)
            if length > 0:
                drawing.lines.append(LineEntity(
                    entity_type="POLYLINE",
                    layer=layer,
                    length=length,
                ))
        except Exception:
            pass

    def _process_arc(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract an ARC entity with its arc length."""
        radius = entity.dxf.radius
        start_angle = math.radians(entity.dxf.start_angle)
        end_angle = math.radians(entity.dxf.end_angle)
        angle = end_angle - start_angle
        if angle < 0:
            angle += 2 * math.pi
        length = abs(radius * angle)
        if length > 0:
            drawing.lines.append(LineEntity(
                entity_type="ARC",
                layer=layer,
                length=length,
            ))

    def _process_circle(self, entity, layer: str, drawing: ParsedDrawing):
        """Extract a CIRCLE (stored as a line entity with circumference)."""
        radius = entity.dxf.radius
        circumference = 2 * math.pi * radius
        if circumference > 0:
            drawing.lines.append(LineEntity(
                entity_type="CIRCLE",
                layer=layer,
                length=circumference,
            ))
