"""
DWG to DXF Converter
====================
Converts DWG files to DXF format using LibreDWG's dwg2dxf utility.
"""

import os
import shutil
import subprocess
import tempfile
import logging

logger = logging.getLogger(__name__)


class DWGConverter:
    """Converts DWG files to DXF using LibreDWG's dwg2dxf."""

    def __init__(self, dwg2dxf_path: str = "/tmp/libredwg/programs/dwg2dxf"):
        self.dwg2dxf_path = dwg2dxf_path
        # Try common install locations if configured path doesn't exist
        if not os.path.isfile(self.dwg2dxf_path):
            for fallback in [
                "/usr/local/bin/dwg2dxf",
                "/usr/bin/dwg2dxf",
                shutil.which("dwg2dxf"),
            ]:
                if fallback and os.path.isfile(fallback):
                    self.dwg2dxf_path = fallback
                    break
            else:
                raise FileNotFoundError(
                    f"dwg2dxf not found at {self.dwg2dxf_path}. "
                    "Install LibreDWG or update the path in config.json."
                )

    def convert(self, dwg_path: str, output_dir: str | None = None) -> str:
        """
        Convert a DWG file to DXF.

        Args:
            dwg_path: Path to the input .dwg file.
            output_dir: Directory for the output .dxf file.
                        If None, uses a temporary directory.

        Returns:
            Path to the generated .dxf file.
        """
        if not os.path.isfile(dwg_path):
            raise FileNotFoundError(f"DWG file not found: {dwg_path}")

        if output_dir is None:
            output_dir = tempfile.mkdtemp(prefix="dwg_to_boq_")
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(dwg_path))[0]
        dxf_path = os.path.join(output_dir, f"{base_name}.dxf")

        logger.info(f"Converting: {os.path.basename(dwg_path)} -> DXF")

        result = subprocess.run(
            [self.dwg2dxf_path, "-o", dxf_path, dwg_path],
            capture_output=True,
            text=True,
            timeout=120,
        )

        if not os.path.isfile(dxf_path):
            raise RuntimeError(
                f"dwg2dxf failed for {dwg_path}.\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        logger.info(f"Converted successfully: {dxf_path}")
        return dxf_path

    def convert_batch(self, dwg_paths: list[str], output_dir: str | None = None) -> list[str]:
        """Convert multiple DWG files. Returns list of DXF paths."""
        dxf_paths = []
        for path in dwg_paths:
            try:
                dxf_path = self.convert(path, output_dir)
                dxf_paths.append(dxf_path)
            except Exception as e:
                logger.error(f"Failed to convert {path}: {e}")
        return dxf_paths
