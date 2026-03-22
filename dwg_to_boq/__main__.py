"""
CLI Entry Point
================
Usage:
    python -m dwg_to_boq <input_dwg_or_folder> [--output <output.xlsx>] [--project <name>] [--config <config.json>]

Examples:
    # Single DWG file
    python -m dwg_to_boq "drawing.dwg" --output boq.xlsx

    # Folder of DWG files
    python -m dwg_to_boq "./MEP Drawings/" --output "MEP BOQ.xlsx" --project "Hospital Project"

    # With custom config
    python -m dwg_to_boq drawings/ --config my_config.json --output boq.xlsx
"""

import argparse
import json
import logging
import os
import sys
import glob

from .converter import DWGConverter
from .parser import DXFParser
from .classifier import EntityClassifier
from .boq_generator import BOQGenerator

DEFAULT_CONFIG = os.path.join(os.path.dirname(__file__), "config.json")


def find_dwg_files(path: str) -> list[str]:
    """Find all .dwg files in a path (file or directory)."""
    if os.path.isfile(path) and path.lower().endswith(".dwg"):
        return [path]

    if os.path.isdir(path):
        dwg_files = []
        for root, dirs, files in os.walk(path):
            for f in files:
                if f.lower().endswith(".dwg"):
                    dwg_files.append(os.path.join(root, f))
        return sorted(dwg_files)

    # Try glob pattern
    matches = glob.glob(path)
    return [m for m in matches if m.lower().endswith(".dwg")]


def main():
    parser = argparse.ArgumentParser(
        description="Convert DWG MEP drawings to Excel Bill of Quantities (BOQ)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m dwg_to_boq "drawing.dwg"
  python -m dwg_to_boq "./MEP Drawings/" --output "BOQ.xlsx" --project "Hospital"
  python -m dwg_to_boq drawings/ --config custom_config.json
        """,
    )
    parser.add_argument(
        "input",
        help="Path to a DWG file or folder containing DWG files",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output Excel file path (default: <input_name>_BOQ.xlsx)",
    )
    parser.add_argument(
        "--project", "-p",
        default="",
        help="Project name to include in the BOQ header",
    )
    parser.add_argument(
        "--config", "-c",
        default=DEFAULT_CONFIG,
        help="Path to config.json (default: built-in config)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()

    # Set up logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    # Load config
    with open(args.config, "r") as f:
        config = json.load(f)

    # Find DWG files
    dwg_files = find_dwg_files(args.input)
    if not dwg_files:
        print(f"Error: No DWG files found at '{args.input}'", file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(dwg_files)} DWG file(s):")
    for f in dwg_files:
        print(f"  - {os.path.basename(f)}")
    print()

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        base = os.path.basename(args.input.rstrip("/\\"))
        base = os.path.splitext(base)[0] if base.endswith(".dwg") else base
        output_path = f"{base}_BOQ.xlsx"

    # Step 1: Convert DWG -> DXF
    print("=" * 60)
    print("STEP 1: Converting DWG to DXF...")
    print("=" * 60)
    converter = DWGConverter(config.get("dwg2dxf_path", "/tmp/libredwg/programs/dwg2dxf"))
    dxf_files = converter.convert_batch(dwg_files)

    if not dxf_files:
        print("Error: No files were successfully converted.", file=sys.stderr)
        sys.exit(1)
    print(f"  Converted {len(dxf_files)}/{len(dwg_files)} files successfully.\n")

    # Step 2: Parse DXF files
    print("=" * 60)
    print("STEP 2: Parsing DXF entities...")
    print("=" * 60)
    parser_obj = DXFParser()
    drawings = []
    for dxf in dxf_files:
        try:
            drawing = parser_obj.parse(dxf)
            drawings.append(drawing)
            print(f"  {os.path.basename(dxf)}: "
                  f"{len(drawing.blocks)} blocks, "
                  f"{len(drawing.texts)} texts, "
                  f"{len(drawing.lines)} lines")
        except Exception as e:
            print(f"  Error parsing {dxf}: {e}", file=sys.stderr)
    print()

    if not drawings:
        print("Error: No files were successfully parsed.", file=sys.stderr)
        sys.exit(1)

    # Step 3: Classify entities
    print("=" * 60)
    print("STEP 3: Classifying entities into BOQ categories...")
    print("=" * 60)
    classifier = EntityClassifier(config)
    result = classifier.classify(drawings)

    by_disc = result.by_discipline()
    for disc, items in by_disc.items():
        if items:
            total_qty = sum(i.quantity for i in items)
            print(f"  {disc}: {len(items)} line items, {total_qty:.0f} total quantity")
    if result.unclassified_blocks:
        print(f"  UNCLASSIFIED: {len(result.unclassified_blocks)} block types")
    print()

    # Step 4: Generate Excel
    print("=" * 60)
    print("STEP 4: Generating Excel BOQ...")
    print("=" * 60)
    generator = BOQGenerator(config)
    generator.generate(result, output_path, project_name=args.project)
    print(f"  Output: {os.path.abspath(output_path)}")
    print()
    print("Done! Open the Excel file and fill in UNIT/RATE column to complete the BOQ.")


if __name__ == "__main__":
    main()
