"""
DWG to BOQ Tool
===============
Converts AutoCAD DWG files (MEP designs) into Excel Bill of Quantities (BOQ).

Pipeline: DWG → DXF (via LibreDWG) → Parse (ezdxf) → Classify → Excel BOQ
"""

__version__ = "1.0.0"
