#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
photo_web_app.py
================
Opens a web page in your browser (http://127.0.0.1:5050) where you can upload
the Excel file and the ZIP file. The page runs insert_profile_photos.py, shows
the summary, and gives you download links for the Excel with photos and the
missing-photo report.

Keep this file in the SAME folder as insert_profile_photos.py.

Setup (once):   pip install flask openpyxl pillow
Run:            python photo_web_app.py
"""

import contextlib
import io
import shutil
import tempfile
import threading
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, abort, render_template_string, request, send_file

import insert_profile_photos as core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024 * 1024   # up to 2 GB uploads

WORK_ROOT = Path(tempfile.gettempdir()) / "photo_inserter_jobs"
WORK_ROOT.mkdir(exist_ok=True)

PAGE = """
<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Student Photo Inserter</title>
<style>
 body{font-family:Segoe UI,system-ui,sans-serif;background:#f4f6f9;margin:0;color:#1b1f24}
 main{max-width:720px;margin:30px auto;padding:0 16px}
 .card{background:#fff;border:1px solid #d9dee5;border-radius:12px;padding:18px;margin-bottom:14px}
 h1{margin:0 0 6px} label{font-weight:600;display:block;margin-bottom:6px}
 input[type=file]{width:100%;padding:10px;border:1px dashed #aab;border-radius:8px;box-sizing:border-box}
 button,.btn{background:#1a6ed8;color:#fff;border:0;border-radius:8px;padding:11px 18px;font-size:1rem;
   cursor:pointer;text-decoration:none;display:inline-block;margin:6px 8px 0 0}
 .btn.g{background:#1a7f45} pre{background:#101418;color:#d7e0ea;padding:14px;border-radius:8px;
   overflow:auto;font-size:.85rem} .err{color:#c0392b;font-weight:600}
</style></head><body><main>
<h1>Student Photo Inserter <small style="font-size:.5em;color:#888">v3</small></h1>
<p>Upload the Excel and the ZIP. Photos are matched by <b>App.No</b> and embedded in the
<b>Profile Picture</b> column.</p>
{% if error %}<div class="card err">{{ error }}</div>{% endif %}
<form class="card" method="post" action="/" enctype="multipart/form-data"
      onsubmit="document.getElementById('b').disabled=true;document.getElementById('b').textContent='Processing... please wait';">
 <label>1. Excel file (.xlsx)</label>
 <input type="file" name="excel" accept=".xlsx,.xlsm" required><br><br>
 <label>2. ZIP file with student folders</label>
 <input type="file" name="zip" accept=".zip" required><br>
 <button id="b" type="submit">Insert photos</button>
</form>
{% if job %}
<div class="card"><h3>Done</h3>
 <a class="btn g" href="/download/{{ job }}/xlsx">Download Excel with photos</a>
 <a class="btn g" href="/download/{{ job }}/full">Download FULL report (all details + photo)</a>
 <a class="btn" href="/download/{{ job }}/csv">Download missing-photo report (CSV)</a>
 </div>
{% endif %}
</main></body></html>
"""


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(PAGE)

    xl, zp = request.files.get("excel"), request.files.get("zip")
    if not xl or not xl.filename or not zp or not zp.filename:
        return render_template_string(PAGE, error="Please choose both files.")

    job = uuid.uuid4().hex[:12]
    folder = WORK_ROOT / job
    folder.mkdir(parents=True)
    excel_path = folder / Path(xl.filename).name
    zip_path = folder / Path(zp.filename).name
    xl.save(excel_path)
    zp.save(zip_path)

    buffer = io.StringIO()
    try:
        core.check_inputs(excel_path, zip_path)
        with contextlib.redirect_stdout(buffer):
            core.process(excel_path, zip_path)
    except Exception as exc:                                   # show, never crash
        shutil.rmtree(folder, ignore_errors=True)
        return render_template_string(PAGE, error="ERROR: %s" % exc)

    return render_template_string(PAGE, job=job, log=buffer.getvalue().strip())


@app.route("/download/<job>/<kind>")
def download(job, kind):
    folder = WORK_ROOT / Path(job).name
    if not folder.is_dir():
        abort(404)
    pattern = {"xlsx": "*_with_photos*.xls*", "csv": "*_photo_report*.csv",
               "full": "*_full_report*.xlsx"}.get(kind)
    if not pattern:
        abort(404)
    found = sorted(folder.glob(pattern))
    if not found:
        abort(404)
    return send_file(found[-1], as_attachment=True, download_name=found[-1].name)


if __name__ == "__main__":
    url = "http://127.0.0.1:5050"
    print("Opening %s  (press Ctrl+C in this window to stop)" % url)
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=5050, debug=False)
