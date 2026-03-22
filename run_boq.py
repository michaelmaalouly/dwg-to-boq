#!/usr/bin/env python3
"""
DWG to BOQ - Convenience Runner
================================
Converts AutoCAD DWG files (MEP designs) into Excel Bill of Quantities.

Usage:
    python run_boq.py <input_dwg_or_folder> [--output output.xlsx] [--project "Project Name"]

Examples:
    python run_boq.py "25 08 2024 ZIA Terminal Block - MEP DESIGN (GENERAL FILE).dwg"
    python run_boq.py "2025 06 12 MEP FINAL SUBMITTED FILE/" --output "Maitama_BOQ.xlsx" --project "Maitama Hospital"
    python run_boq.py "Imo State Radiology Center - MEP Design/" --output "Imo_BOQ.xlsx"
"""

import sys
import os

# Ensure the parent directory is in the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dwg_to_boq.__main__ import main

if __name__ == "__main__":
    main()
