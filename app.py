"""
DWG to BOQ - Web Interface
===========================
Flask web application for converting DWG files to Excel BOQ.
"""

import json
import logging
import os
import shutil
import time
import uuid
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file

from dwg_to_boq.converter import DWGConverter
from dwg_to_boq.parser import DXFParser
from dwg_to_boq.classifier import EntityClassifier
from dwg_to_boq.boq_generator import BOQGenerator

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB max upload

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
CONFIG_PATH = BASE_DIR / "dwg_to_boq" / "config.json"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# In-memory job tracking
jobs: dict[str, dict] = {}


def load_config():
    with open(CONFIG_PATH) as f:
        return json.load(f)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload_files():
    """Handle DWG file uploads and start processing."""
    if "files" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files")
    dwg_files = [f for f in files if f.filename and f.filename.lower().endswith(".dwg")]

    if not dwg_files:
        return jsonify({"error": "No .dwg files found in upload"}), 400

    project_name = request.form.get("project_name", "").strip()

    # Create a unique job
    job_id = str(uuid.uuid4())[:8]
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    saved_paths = []
    filenames = []
    for f in dwg_files:
        safe_name = f.filename.replace("/", "_").replace("\\", "_")
        save_path = job_dir / safe_name
        f.save(str(save_path))
        saved_paths.append(str(save_path))
        filenames.append(safe_name)

    jobs[job_id] = {
        "status": "uploaded",
        "progress": 0,
        "step": "Files uploaded",
        "files": filenames,
        "file_count": len(saved_paths),
        "dwg_paths": saved_paths,
        "project_name": project_name,
        "output_path": None,
        "error": None,
        "result_summary": None,
    }

    return jsonify({"job_id": job_id, "file_count": len(saved_paths), "files": filenames})


@app.route("/api/process/<job_id>", methods=["POST"])
def process_job(job_id):
    """Process uploaded DWG files into BOQ Excel."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404

    job = jobs[job_id]
    if job["status"] == "processing":
        return jsonify({"error": "Job already processing"}), 409

    job["status"] = "processing"
    job["progress"] = 0

    try:
        config = load_config()
        dwg_paths = job["dwg_paths"]
        project_name = job["project_name"]
        total_steps = len(dwg_paths) * 2 + 2  # convert + parse + classify + generate

        # Step 1: Convert DWG to DXF
        job["step"] = "Converting DWG to DXF..."
        converter = DWGConverter(config.get("dwg2dxf_path"))
        dxf_paths = []
        for i, dwg_path in enumerate(dwg_paths):
            try:
                dxf_path = converter.convert(dwg_path)
                dxf_paths.append(dxf_path)
            except Exception as e:
                logger.error(f"Convert failed: {dwg_path}: {e}")
            job["progress"] = int(((i + 1) / total_steps) * 100)

        if not dxf_paths:
            raise RuntimeError("No files were successfully converted to DXF")

        # Step 2: Parse DXF files
        job["step"] = "Parsing drawing entities..."
        parser = DXFParser()
        drawings = []
        for i, dxf_path in enumerate(dxf_paths):
            try:
                drawing = parser.parse(dxf_path)
                drawings.append(drawing)
            except Exception as e:
                logger.error(f"Parse failed: {dxf_path}: {e}")
            job["progress"] = int(((len(dwg_paths) + i + 1) / total_steps) * 100)

        if not drawings:
            raise RuntimeError("No files were successfully parsed")

        # Step 3: Classify
        job["step"] = "Classifying MEP entities..."
        job["progress"] = int(((total_steps - 1) / total_steps) * 100)
        classifier = EntityClassifier(config)
        result = classifier.classify(drawings)

        # Step 4: Generate Excel
        job["step"] = "Generating Excel BOQ..."
        output_name = f"BOQ_{project_name.replace(' ', '_') or job_id}_{int(time.time())}.xlsx"
        output_path = OUTPUT_DIR / output_name
        generator = BOQGenerator(config)
        generator.generate(result, str(output_path), project_name=project_name)

        # Build summary
        by_disc = result.by_discipline()
        summary = {}
        for disc, items in by_disc.items():
            if items:
                summary[disc] = {
                    "line_items": len(items),
                    "total_quantity": sum(i.quantity for i in items),
                }
        summary["unclassified"] = len(result.unclassified_blocks)

        job["status"] = "completed"
        job["progress"] = 100
        job["step"] = "Done!"
        job["output_path"] = str(output_path)
        job["output_name"] = output_name
        job["result_summary"] = summary

        return jsonify({
            "status": "completed",
            "output_name": output_name,
            "summary": summary,
        })

    except Exception as e:
        logger.exception(f"Job {job_id} failed")
        job["status"] = "failed"
        job["error"] = str(e)
        job["step"] = f"Error: {e}"
        return jsonify({"error": str(e)}), 500


@app.route("/api/status/<job_id>")
def job_status(job_id):
    """Get current job status."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    return jsonify({
        "status": job["status"],
        "progress": job["progress"],
        "step": job["step"],
        "files": job["files"],
        "error": job["error"],
        "summary": job.get("result_summary"),
        "output_name": job.get("output_name"),
    })


@app.route("/api/download/<job_id>")
def download_result(job_id):
    """Download the generated Excel file."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "completed" or not job["output_path"]:
        return jsonify({"error": "No output file available"}), 404

    return send_file(
        job["output_path"],
        as_attachment=True,
        download_name=job["output_name"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/cleanup/<job_id>", methods=["DELETE"])
def cleanup_job(job_id):
    """Clean up uploaded and output files for a job."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404

    job_dir = UPLOAD_DIR / job_id
    if job_dir.exists():
        shutil.rmtree(job_dir)

    if job_id in jobs:
        output_path = jobs[job_id].get("output_path")
        if output_path and os.path.exists(output_path):
            os.remove(output_path)
        del jobs[job_id]

    return jsonify({"status": "cleaned"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(debug=debug, host="0.0.0.0", port=port)
