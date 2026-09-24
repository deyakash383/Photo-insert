#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
insert_profile_photos_by_serial.py
===================================
Alternate matching mode for photo insertion.

insert_profile_photos.py matches a ZIP folder/photo to an Excel row mainly by
the WHOLE App.No text (exact name, then separators ignored, then "contains
the App.No"), and only falls back to the serial number as a last resort.

This script does the opposite: it matches PURELY by the serial number - the
digits after the LAST "/" in the App.No - and ignores everything else in the
App.No or in the folder/file name.

    App.No format : MSU-WI / 2025-26 / 15432
                     prefix   year      serial number  <- unique matching code

    Excel App.No : MSU-WI/2025-26/15432
    ZIP folder   : <any name that contains 15432 as a standalone number>
                    e.g. "15432", "15432 - Rahul Sharma", "MSU_WI_2025-26_15432"
    ZIP photo    : <any image whose file name contains 15432>
                    e.g. "15432.jpg", "Photo_15432.png"

"Standalone number" means 15432 must not be part of a longer number (it will
NOT match "115432" or "154320"), so short serials do not accidentally match
the wrong folder.

Everything else - reading the Excel file, resizing/embedding photos, the
CSV report, the full Excel report, the file-picker window, etc. - is re-used
from insert_profile_photos.py, so keep both files in the SAME folder.

Quick start (Windows):
    1. pip install openpyxl pillow
    2. python insert_profile_photos_by_serial.py
    3. A window opens: choose the Excel file, then choose the ZIP file.
       (You can also pass the two paths on the command line.)

Output files are named "..._with_photos_by_serial.xlsx" etc. so running this
script never overwrites the output of insert_profile_photos.py (or vice
versa), even on the same input files.

Requires: Python 3.8+, openpyxl, pillow
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from collections import Counter
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Set, Tuple

try:
    import insert_profile_photos as core
except ImportError:
    print("\nERROR: insert_profile_photos.py must be in the same folder as this script.")
    sys.exit(1)

try:
    from PIL import Image as PILImage
except ImportError as exc:  # pragma: no cover
    print("\nERROR: A required package is missing (%s)." % exc)
    print("Install the requirements with:\n\n    pip install openpyxl pillow\n")
    sys.exit(1)


# =============================================================================
# SETTINGS - only the naming differs from insert_profile_photos.py; every
# other setting (photo size, matching tolerances, etc.) is re-used from
# core.<SETTING> so the two scripts always stay in sync.
# =============================================================================
OUTPUT_SUFFIX = "_with_photos_by_serial"
REPORT_SUFFIX = "_photo_report_by_serial.csv"
FULL_REPORT_SUFFIX = "_full_report_by_serial.xlsx"

DEFAULT_EXCEL_NAME = core.DEFAULT_EXCEL_NAME
DEFAULT_ZIP_NAME = core.DEFAULT_ZIP_NAME
MIN_SERIAL_DIGITS = core.MIN_SERIAL_DIGITS

# Outcome codes (re-exported for readability below)
INSERTED = core.INSERTED
MISSING_FOLDER = core.MISSING_FOLDER
MISSING_PHOTO = core.MISSING_PHOTO
INVALID_APP_NO = core.INVALID_APP_NO
AMBIGUOUS_DUPLICATE = core.AMBIGUOUS_DUPLICATE
UNREADABLE_IMAGE = core.UNREADABLE_IMAGE


# -----------------------------------------------------------------------------
# Serial-number extraction
# -----------------------------------------------------------------------------
def serial_of(text: str) -> str:
    """The unique matching code: the digits after the LAST '/' in the text,
    e.g. 'MSU-WI/2025-26/15432' -> '15432'.
    If there is no '/', or nothing numeric after it, falls back to the last
    group of digits found anywhere in the text (so folder/file names, which
    have no '/', are handled the same way)."""
    cleaned = core.clean_text(text)
    if "/" in cleaned:
        tail = cleaned.rsplit("/", 1)[-1]
        digits = re.findall(r"\d+", tail)
        if digits and len(digits[-1]) >= MIN_SERIAL_DIGITS:
            return digits[-1]
    return core.last_serial(cleaned)


def _serial_pattern(serial: str) -> "re.Pattern":
    # digit-boundary lookaround so "15432" never matches inside "115432"
    return re.compile(r"(?<!\d)%s(?!\d)" % re.escape(serial))


# -----------------------------------------------------------------------------
# Matching - the only real difference from insert_profile_photos.py
# -----------------------------------------------------------------------------
def find_student_folders(rec: "core.Record", index: "core.ZipIndex",
                         excel_serials: Set[str]) -> Tuple[List[str], str]:
    """Match a ZIP folder purely by the serial number. The rest of the
    folder name can say anything at all."""
    serial = rec.alnum                    # this script stores the SERIAL here
    if not serial:
        return [], ""
    others = excel_serials - {serial}
    pattern = _serial_pattern(serial)

    hits = []
    for d, (name, _na) in index.dir_info.items():
        own = core.last_serial(name)
        if own and own != serial and own in others:
            continue                      # folder's own number belongs to someone else
        if pattern.search(name):
            hits.append(d)
    if hits:
        return core.top_level(hits), "folder name contains serial number %s" % serial
    return [], ""


def find_photo_files(rec: "core.Record", index: "core.ZipIndex",
                     excel_serials: Set[str]) -> Tuple[List["core.FileEntry"], str]:
    """For ZIPs where photos sit loose (no per-student folder): match a
    photo FILE purely by the serial number in its file name."""
    serial = rec.alnum
    if not serial:
        return [], ""
    others = excel_serials - {serial}
    pattern = _serial_pattern(serial)

    def file_serial(f: "core.FileEntry") -> str:
        groups = re.findall(r"\d+", Path(f.filename).stem)
        return groups[-1] if groups else ""

    hits = []
    for f in index.all_images:
        own = file_serial(f)
        if own and own != serial and own in others:
            continue                      # file's own number belongs to someone else
        if pattern.search(Path(f.filename).stem):
            hits.append(f)
    if hits:
        return hits, "file name contains serial number %s" % serial
    return [], ""


# -----------------------------------------------------------------------------
# Main workflow (mirrors core.process(), matching section replaced)
# -----------------------------------------------------------------------------
def process(excel_path: Path, zip_path: Path) -> Tuple[int, str]:
    started = datetime.now()

    print("Loading Excel : %s" % excel_path.name)
    keep_vba = excel_path.suffix.lower() == ".xlsm"
    wb = core.load_workbook(excel_path, keep_vba=keep_vba)
    ws, header_row, app_col, pic_col = core.locate_sheet_and_headers(wb)
    print("Worksheet     : '%s'  (header on row %d)" % (ws.title, header_row))

    if pic_col is None:
        pic_col = ws.max_column + 1
        header_cell = ws.cell(header_row, pic_col, core.PICTURE_HEADER)
        prev = ws.cell(header_row, pic_col - 1)
        if prev.has_style:
            header_cell._style = copy(prev._style)
        print("Note: no '%s' column existed, so one was added (column %d)."
              % (core.PICTURE_HEADER, pic_col))
    print("Columns       : App.No = %d, %s = %d  (matching by serial number)"
          % (app_col, core.PICTURE_HEADER, pic_col))

    print("Reading ZIP   : %s" % zip_path.name)
    zf = zipfile.ZipFile(zip_path, "r")
    try:
        index = core.build_zip_index(zf)
        print("ZIP contains  : %d folders holding files, %d files in total"
              % (len(index.leaf_dirs), len({e for d in index.leaf_dirs for e in index.files_under[d]})))

        # ---- Collect Excel records (rec.alnum holds the SERIAL, not alnum text) --
        records: List["core.Record"] = []
        for r in range(header_row + 1, ws.max_row + 1):
            if core.row_is_blank(ws, r):
                continue
            raw = core.clean_text(ws.cell(r, app_col).value)
            rec = core.Record(row=r, app_no=raw, key=core.make_key(raw), alnum=serial_of(raw))
            records.append(rec)

        key_counts = Counter(rec.key for rec in records if rec.key)
        flat_dirs: Set[str] = set()
        used_files: Set[str] = set()
        excel_serials = {rec.alnum for rec in records if rec.alnum}

        cell_w = core.MAX_PHOTO_WIDTH_PX + 2 * core.CELL_PADDING_PX
        cell_h = core.MAX_PHOTO_HEIGHT_PX + 2 * core.CELL_PADDING_PX
        if core.row_height_for_pixels(cell_h) > 409:
            raise ValueError("MAX_PHOTO_HEIGHT_PX is too large: Excel rows "
                             "cannot be taller than 409 points.")
        ws.column_dimensions[
            ws.cell(1, pic_col).column_letter].width = core.col_width_for_pixels(cell_w)

        total = len(records)
        print("\nProcessing %d record(s)..." % total)
        for n, rec in enumerate(records, start=1):
            if n % 100 == 0:
                print("  ...%d / %d" % (n, total))

            if not rec.app_no or not core.VALID_APP_NO_PATTERN.match(rec.app_no):
                rec.outcome = INVALID_APP_NO
                rec.issues.append((INVALID_APP_NO,
                                   "App.No is blank or contains invalid characters"))
                continue
            if not rec.alnum:
                rec.outcome = INVALID_APP_NO
                rec.issues.append((INVALID_APP_NO,
                                   "Could not find a serial number (%d+ digits) after "
                                   "the last '/' in the App.No" % MIN_SERIAL_DIGITS))
                continue

            if key_counts[rec.key] > 1:
                rec.dup_excel = True
                rec.issues.append(("DUPLICATE_APP_NO_IN_EXCEL",
                                   "App.No appears %d times in the Excel sheet"
                                   % key_counts[rec.key]))

            folder_paths, how = find_student_folders(rec, index, excel_serials)
            files: Dict[str, "core.FileEntry"] = {}
            if not folder_paths:
                flat, fhow = find_photo_files(rec, index, excel_serials)
                if not flat:
                    rec.outcome = MISSING_FOLDER
                    rec.issues.append((MISSING_FOLDER,
                                       "No ZIP folder or photo file contains serial number %s"
                                       % rec.alnum))
                    continue
                files = {f.zip_path: f for f in flat}
                flat_dirs.update(f.dir_path for f in flat)
                rec.matched_folders = sorted({f.dir_path for f in flat if f.dir_path})
                rec.partial_folder = True
                rec.issues.append(("MATCHED_BY_SERIAL",
                                   "Matched because the %s -> %s (please verify)"
                                   % (fhow, " | ".join(f.zip_path for f in flat))))
            else:
                rec.matched_folders = folder_paths
                rec.partial_folder = True
                rec.issues.append(("MATCHED_BY_SERIAL",
                                   "Matched by %s -> %s (please verify)"
                                   % (how, " | ".join(folder_paths))))

                if len(folder_paths) > 1:
                    rec.dup_zip = True
                    rec.issues.append(("DUPLICATE_FOLDER_IN_ZIP",
                                       "Folder found %d times: %s"
                                       % (len(folder_paths), " | ".join(folder_paths))))

                for fp in folder_paths:
                    files.update(index.files_under.get(fp, {}))

            candidates, tier = core.pick_photo_candidates(list(files.values()))

            if not candidates:
                rec.outcome = MISSING_PHOTO
                detail = ("Folder '%s' has only document scans (Aadhar/ID/etc.), no photograph"
                          if tier == "only-documents"
                          else "Folder '%s' has no .jpg/.jpeg/.png image")
                rec.issues.append((MISSING_PHOTO, detail % (
                    folder_paths[0] if folder_paths else "matched files")))
                continue

            if len(candidates) > 1:
                rec.dup_zip = True
                names = " | ".join(c.zip_path for c in candidates)
                if len({c.crc for c in candidates}) > 1:
                    rec.outcome = AMBIGUOUS_DUPLICATE
                    if tier == "photo-name":
                        msg = "%d different 'Photo' files" % len(candidates)
                    else:
                        msg = "%d images and none is named 'Photo'" % len(candidates)
                    rec.issues.append(("MULTIPLE_DIFFERENT_PHOTOS",
                                       "%s - none inserted (check manually): %s"
                                       % (msg, names)))
                    continue
                rec.issues.append(("DUPLICATE_IDENTICAL_PHOTOS",
                                   "%d identical copies, first one used: %s"
                                   % (len(candidates), names)))

            chosen = candidates[0]
            if tier != "photo-name":
                rec.fallback_photo = True
                rec.issues.append(("PHOTO_FILE_NAME_FALLBACK",
                                   "No file starting with 'Photo'; used '%s' (%s) - please verify"
                                   % (chosen.zip_path,
                                      "name contains photo/profile/pic" if tier == "photo-keyword"
                                      else "only non-document image in the folder")))
            try:
                buf, w, h = core.prepare_photo(zf.read(chosen.zip_path))
                used_files.add(chosen.zip_path)
                rec.photo_file, rec.photo_bytes, rec.photo_size = \
                    chosen.zip_path, buf.getvalue(), (w, h)
                core.add_centered_image(ws, buf, w, h, rec.row, pic_col, cell_w, cell_h)
                rec.outcome = INSERTED
            except Exception as exc:
                rec.outcome = UNREADABLE_IMAGE
                reason = ("file is corrupt or not a real image"
                          if isinstance(exc, PILImage.UnidentifiedImageError)
                          else "%s: %s" % (type(exc).__name__, exc))
                rec.issues.append((UNREADABLE_IMAGE,
                                   "Could not read '%s' (%s)" % (chosen.zip_path, reason)))

        for rec in records:
            ws.row_dimensions[rec.row].height = core.row_height_for_pixels(cell_h)
            if core.CENTER_ROW_CONTENT_VERTICALLY:
                for cell in ws[rec.row]:
                    al = copy(cell.alignment)
                    al.vertical = "center"
                    cell.alignment = al

        matched = {fp for rec in records for fp in rec.matched_folders}
        ancestors = {"/".join(m.split("/")[:i]) for m in matched
                     for i in range(1, len(m.split("/")))}
        orphan_folders = []
        for d in sorted(index.leaf_dirs):
            parts = d.split("/")
            inside_matched = any("/".join(parts[:i]) in matched
                                 for i in range(1, len(parts) + 1))
            if not inside_matched and d not in ancestors:
                orphan_folders.append(d)
        for e in index.all_images:
            if e.dir_path in flat_dirs and e.zip_path not in used_files \
                    and not core.looks_like_document(e.filename):
                orphan_folders.append(e.zip_path)

        out_path = excel_path.with_name(excel_path.stem + OUTPUT_SUFFIX + excel_path.suffix)
        if out_path.resolve() == excel_path.resolve():
            raise RuntimeError("Refusing to overwrite the original file.")
        print("\nSaving workbook (this can take a moment for large files)...")
        try:
            wb.save(out_path)
        except PermissionError:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = out_path.with_name("%s_%s%s" % (out_path.stem, stamp, out_path.suffix))
            print("Output file is open in Excel; saving as %s instead." % out_path.name)
            wb.save(out_path)
    finally:
        zf.close()

    report_path = excel_path.with_name(excel_path.stem + REPORT_SUFFIX)
    rows_written = core.write_report(report_path, records, orphan_folders)
    full_path = core.write_full_report(excel_path, ws, header_row, pic_col, records,
                                       orphan_folders, cell_w, cell_h)
    # write_full_report names the file from FULL_REPORT_SUFFIX defined in core; move/rename
    # it to this script's own suffix so it never collides with insert_profile_photos.py's output.
    wanted_full_path = excel_path.with_name(excel_path.stem + FULL_REPORT_SUFFIX)
    if full_path != wanted_full_path:
        try:
            if wanted_full_path.exists():
                wanted_full_path.unlink()
            full_path.rename(wanted_full_path)
            full_path = wanted_full_path
        except OSError:
            pass                              # keep core's name if a rename isn't possible
    print("\nFull report (all details + photo): %s" % full_path)

    return core.print_summary(records, orphan_folders, out_path, report_path, rows_written, started)


# -----------------------------------------------------------------------------
# Command-line entry point (mirrors core.main())
# -----------------------------------------------------------------------------
def parse_args(argv):
    import argparse
    p = argparse.ArgumentParser(
        description="Embed learner photographs from a ZIP into an Excel sheet, "
                    "matching purely by the serial number in the App.No "
                    "(the digits after the last '/').")
    p.add_argument("excel", nargs="?",
                   help="Input .xlsx (leave out to choose it in a window)")
    p.add_argument("zip", nargs="?",
                   help="Input .zip (leave out to choose it in a window)")
    return p.parse_args(argv)


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass

    args = parse_args(argv)
    script_dir = Path(__file__).resolve().parent
    exit_code = 1
    interactive = not (args.excel or args.zip)
    try:
        if interactive:
            excel_path = core.ask_for_file("Excel", (("Excel files", "*.xlsx *.xlsm"),
                                                      ("All files", "*.*")),
                                           DEFAULT_EXCEL_NAME, ("*.xlsx", "*.xlsm"), script_dir)
            zip_path = core.ask_for_file("ZIP", (("ZIP files", "*.zip"), ("All files", "*.*")),
                                         DEFAULT_ZIP_NAME, ("*.zip",), script_dir)
        else:
            excel_path = core.resolve_input(args.excel, DEFAULT_EXCEL_NAME,
                                            ("*.xlsx", "*.xlsm"), "Excel file", script_dir)
            zip_path = core.resolve_input(args.zip, DEFAULT_ZIP_NAME,
                                          ("*.zip",), "ZIP file", script_dir)
        core.check_inputs(excel_path, zip_path)
        exit_code, popup = process(excel_path, zip_path)
        if interactive:
            core.show_popup("Photo insertion finished (serial-number matching)", popup)
        try:
            import os
            os.startfile(str(excel_path.parent))
        except Exception:
            pass
    except core.UserCancelled as exc:
        print("\nCancelled: %s" % exc)
    except FileNotFoundError as exc:
        print("\nERROR: %s" % exc)
    except zipfile.BadZipFile:
        print("\nERROR: The ZIP file is corrupt or is not a valid ZIP archive.")
    except PermissionError as exc:
        print("\nERROR: Permission denied (%s). Close the file in Excel and try again." % exc)
    except ValueError as exc:
        print("\nERROR: %s" % exc)
    except Exception as exc:
        print("\nUNEXPECTED ERROR: %s: %s" % (type(exc).__name__, exc))
    finally:
        if len(sys.argv) == 1 and sys.stdin and sys.stdin.isatty():
            try:
                input("\nPress Enter to close...")
            except EOFError:
                pass
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
