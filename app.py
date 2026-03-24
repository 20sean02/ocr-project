#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import csv
import uuid
import json
import base64
import hashlib
import tempfile
from datetime import datetime

from flask import Flask, request, render_template, send_file, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import anthropic

import extract_subjects_batch_full_withEN as ocrmod

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

# ── LLM fallback for partial OCR results ──────────────────
_ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
if not _ANTHROPIC_KEY:
    _key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "API_key.env")
    if os.path.exists(_key_path):
        with open(_key_path, "r") as _f:
            _ANTHROPIC_KEY = _f.read().strip()
_llm_client = anthropic.Anthropic(api_key=_ANTHROPIC_KEY) if _ANTHROPIC_KEY else None

_LLM_PROMPT = """你是公文 OCR 資料擷取助手。這張圖片是一份台灣政府公文。
請從圖片中擷取以下欄位，只回傳 JSON，不要多餘文字：

{
  "來文機關": "發文機關名稱（如：消防局、澎湖縣政府、消防署）",
  "來文字號": "來文的科室（如：人事室、搶救科、預防科、督企科、火調科、救護科、訓練科、府建商）",
  "收文文號": "發文字號中的數字編號",
  "收文日期": "發文日期，格式為 民國年.月.日（如 115.1.13）",
  "事由": "主旨欄的完整內容"
}

規則：
- 來文機關：若文件頂部寫「澎湖縣政府消防局 函」則為「消防局」；若寫「澎湖縣政府 函」則為「澎湖縣政府」；若寫「內政部消防署」則為「消防署」
- 來文字號：根據發文字號中「澎消X字」的X判斷科室，例如「澎消護字」→「救護科」，「澎消搶字」→「搶救科」，「府建商字」→「府建商」
- 收文文號：發文字號中「第XXXXXXX號」的數字部分
- 收文日期：「中華民國115年1月13日」→「115.1.13」
- 事由：主旨欄的完整內容，到「請查照」為止（含「請查照」），結尾加句號
- 如果某個欄位確實無法從圖片中辨識，該欄位留空字串 ""
- 只回傳 JSON，不要 markdown 代碼塊"""


def _llm_extract(image_path: str, missing_fields: list) -> dict:
    """Use Claude vision to extract missing fields from a document image."""
    if not _llm_client:
        return {}
    try:
        with open(image_path, "rb") as f:
            img_data = base64.standard_b64encode(f.read()).decode("utf-8")

        ext = os.path.splitext(image_path)[1].lower()
        media_type = {
            ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".png": "image/png", ".webp": "image/webp",
            ".gif": "image/gif",
        }.get(ext, "image/jpeg")

        resp = _llm_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_data}},
                    {"type": "text", "text": _LLM_PROMPT + "\n\n只需要擷取以下欄位：" + "、".join(missing_fields)},
                ],
            }],
        )

        text = resp.content[0].text.strip()
        # Strip markdown code block if present
        if text.startswith("```"):
            text = text.split("\n", 1)[1] if "\n" in text else text[3:]
            if text.endswith("```"):
                text = text[:-3]
            text = text.strip()
        return json.loads(text)
    except Exception as e:
        print(f"[LLM fallback error] {e}")
        return {}

# In-memory storage
RESULTS = {}          # token -> list of row dicts (editable)
IMAGE_CACHE = {}      # file content hash -> row dict (skip duplicates)
EXPORT_LOG = []       # list of {"filename", "time", "rows", "format"}

# Server-side export directory
EXPORT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "csv_exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

HEADERS = ["工作項目編號", "原始檔名", "來文機關", "來文字號", "收文文號", "收文日期", "事由", "status"]
FIELD_MAP = [
    ("工作項目編號", "工作項目編號"),
    ("original_filename", "原始檔名"),
    ("來文機關", "來文機關"),
    ("來文字號", "來文字號"),
    ("收文文號", "收文文號"),
    ("收文日期", "收文日期"),
    ("事由", "事由"),
    ("status", "status"),
]


def _hash_file(file_obj) -> str:
    data = file_obj.read()
    file_obj.seek(0)
    return hashlib.sha256(data).hexdigest()


def _rows_to_csv(rows):
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(HEADERS)
    for r in rows:
        writer.writerow([r.get(key, "") or "" for key, _ in FIELD_MAP])
    return buf.getvalue()


XLSX_HEADERS = ["工作項目編號", "來文機關", "來文字號", "收文文號", "收文日期", "事由"]
XLSX_FIELD_MAP = [
    ("工作項目編號", "工作項目編號"),
    ("來文機關", "來文機關"),
    ("來文字號", "來文字號"),
    ("收文文號", "收文文號"),
    ("收文日期", "收文日期"),
    ("事由", "事由"),
]


def _rows_to_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "辨識結果"

    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    hdr_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(bottom=Side(style="thin", color="D9D9D9"))

    for col, name in enumerate(XLSX_HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align

    for row_idx, r in enumerate(rows, 2):
        for col_idx, (key, _) in enumerate(XLSX_FIELD_MAP, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=r.get(key, "") or "")
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=(key == "事由"))

    col_widths = [16, 16, 16, 14, 12, 50]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Export helpers ──────────────────────────────────────────────

def _save_export(rows, token, fmt):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ext = {"csv": ".csv", "xlsx": ".xlsx"}.get(fmt, ".csv")
    filename = f"export_{ts}_{token}{ext}"
    path = os.path.join(EXPORT_DIR, filename)

    if fmt == "xlsx":
        buf = _rows_to_xlsx(rows)
        with open(path, "wb") as f:
            f.write(buf.getvalue())
    else:
        csv_data = _rows_to_csv(rows)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(csv_data)

    EXPORT_LOG.append({
        "filename": filename,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rows": len(rows),
        "format": fmt.upper(),
    })
    if len(EXPORT_LOG) > 100:
        EXPORT_LOG[:] = EXPORT_LOG[-100:]

    return filename


def _sort_and_renumber(rows):
    """Sort rows by 收文日期 ascending and re-assign 工作項目編號."""
    rows.sort(key=lambda r: (
        r.get("_sort_key") or ocrmod.parse_roc_date_to_sort_key(r.get("收文日期", ""))
    ))
    _renumber(rows)


def _renumber(rows):
    """Re-assign 工作項目編號 without changing order."""
    for idx, r in enumerate(rows, start=1):
        r["工作項目編號"] = ocrmod.format_work_item_id(idx, r.get("收文日期", ""))


def _read_csv_file(path):
    """Read a CSV export file and return list of row dicts."""
    rows = []
    with open(path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append({
                "工作項目編號": r.get("工作項目編號", ""),
                "original_filename": r.get("原始檔名", ""),
                "來文機關": r.get("來文機關", "N/A"),
                "來文字號": r.get("來文字號", "N/A"),
                "收文文號": r.get("收文文號", ""),
                "收文日期": r.get("收文日期", ""),
                "事由": r.get("事由", ""),
                "status": r.get("status", ""),
            })
    return rows


# ── Routes ─────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    uploaded = request.files.getlist("files")
    if not uploaded or all(f.filename == "" for f in uploaded):
        return jsonify({"error": "請選擇至少一張圖片"}), 400

    token = request.form.get("token", "").strip()
    existing_rows = []
    if token and token in RESULTS:
        existing_rows = RESULTS[token]
    else:
        token = uuid.uuid4().hex[:12]

    new_rows = []
    skipped = 0
    existing_hashes = {r.get("_file_hash") for r in existing_rows if r.get("_file_hash")}

    with tempfile.TemporaryDirectory() as tmpdir:
        saved = []
        for i, f in enumerate(uploaded):
            orig = os.path.basename(f.filename) or f"upload_{i}.jpg"
            file_hash = _hash_file(f)

            if file_hash in existing_hashes:
                skipped += 1
                continue

            if file_hash in IMAGE_CACHE:
                cached = IMAGE_CACHE[file_hash].copy()
                cached["original_filename"] = orig
                cached["_file_hash"] = file_hash
                new_rows.append(cached)
                existing_hashes.add(file_hash)
                skipped += 1
                continue

            safe = f"{i:04d}_{orig}"
            path = os.path.join(tmpdir, safe)
            f.save(path)
            saved.append((path, orig, file_hash))
            existing_hashes.add(file_hash)

        for path, orig_name, file_hash in saved:
            try:
                import time as _time
                _t0 = _time.time()
                subject, issue_date, laiwen_dept, doc_no, from_agency = ocrmod.process_one(path)
                _t_ocr = _time.time() - _t0
                status = "ok" if (subject and from_agency != "N/A" and doc_no and issue_date) else "partial"
                print(f"[OCR] {orig_name} -> {status} ({_t_ocr:.1f}s)", flush=True)

                # LLM fallback for partial results
                if status == "partial":
                    missing = []
                    if not subject:
                        missing.append("事由")
                    if not from_agency or from_agency == "N/A":
                        missing.append("來文機關")
                    if not laiwen_dept or laiwen_dept == "N/A":
                        missing.append("來文字號")
                    if not doc_no:
                        missing.append("收文文號")
                    if not issue_date:
                        missing.append("收文日期")

                    if missing:
                        print(f"[LLM] {orig_name} missing: {missing}", flush=True)
                        _t1 = _time.time()
                        llm_data = _llm_extract(path, missing)
                        print(f"[LLM] {orig_name} done ({_time.time() - _t1:.1f}s)", flush=True)
                        if llm_data:
                            if not subject and llm_data.get("事由"):
                                subject = llm_data["事由"]
                            if (not from_agency or from_agency == "N/A") and llm_data.get("來文機關"):
                                from_agency = llm_data["來文機關"]
                            if (not laiwen_dept or laiwen_dept == "N/A") and llm_data.get("來文字號"):
                                laiwen_dept = llm_data["來文字號"]
                            if not doc_no and llm_data.get("收文文號"):
                                doc_no = llm_data["收文文號"]
                            if not issue_date and llm_data.get("收文日期"):
                                issue_date = llm_data["收文日期"]
                            # Re-evaluate status
                            status = "ok" if (subject and from_agency != "N/A" and doc_no and issue_date) else "partial"

                row = {
                    "original_filename": orig_name,
                    "來文機關": from_agency or "N/A",
                    "來文字號": laiwen_dept or "N/A",
                    "收文文號": doc_no or "",
                    "收文日期": issue_date or "",
                    "事由": subject or "",
                    "status": status,
                    "_sort_key": ocrmod.parse_roc_date_to_sort_key(issue_date),
                    "_file_hash": file_hash,
                }
                new_rows.append(row)
                IMAGE_CACHE[file_hash] = row.copy()
            except Exception as e:
                new_rows.append({
                    "original_filename": orig_name,
                    "來文機關": "N/A",
                    "來文字號": "N/A",
                    "收文文號": "",
                    "收文日期": "",
                    "事由": "",
                    "status": f"error: {e}",
                    "_sort_key": (1, 0, 0, 0),
                    "_file_hash": file_hash,
                })

    all_rows = existing_rows + new_rows
    _sort_and_renumber(all_rows)
    RESULTS[token] = all_rows

    if len(RESULTS) > 20:
        oldest = list(RESULTS.keys())[0]
        del RESULTS[oldest]
    if len(IMAGE_CACHE) > 500:
        keys = list(IMAGE_CACHE.keys())
        for k in keys[:250]:
            del IMAGE_CACHE[k]

    ok_count = sum(1 for r in all_rows if r.get("status") == "ok")
    partial_count = sum(1 for r in all_rows if r.get("status") == "partial")
    error_count = sum(1 for r in all_rows if r.get("status", "").startswith("error"))

    safe_rows = []
    for r in all_rows:
        safe_rows.append({
            "工作項目編號": r.get("工作項目編號", ""),
            "original_filename": r.get("original_filename", ""),
            "來文機關": r.get("來文機關", "N/A"),
            "來文字號": r.get("來文字號", "N/A"),
            "收文文號": r.get("收文文號", ""),
            "收文日期": r.get("收文日期", ""),
            "事由": r.get("事由", ""),
            "status": r.get("status", ""),
            "_file_hash": r.get("_file_hash", ""),
        })

    return jsonify({
        "token": token,
        "rows": safe_rows,
        "ok_count": ok_count,
        "partial_count": partial_count,
        "error_count": error_count,
        "skipped": skipped,
        "new_count": len(new_rows),
    })


def _get_edited_rows(token):
    """Extract edited rows from request, update caches."""
    stored_rows = RESULTS.get(token)
    if stored_rows is None:
        return None

    edited = request.get_json(silent=True)
    if edited and isinstance(edited, list):
        rows = edited
        for edited_row, stored_row in zip(edited, stored_rows):
            file_hash = edited_row.get("_file_hash") or stored_row.get("_file_hash")
            if file_hash and file_hash in IMAGE_CACHE:
                for field in ("來文機關", "來文字號", "收文文號", "收文日期", "事由"):
                    if field in edited_row:
                        IMAGE_CACHE[file_hash][field] = edited_row[field]
        RESULTS[token] = edited
    else:
        rows = stored_rows
    return rows


@app.route("/download/<token>/csv", methods=["POST"])
def download_csv(token):
    rows = _get_edited_rows(token)
    if rows is None:
        return "結果已過期，請重新上傳辨識", 404

    csv_data = _rows_to_csv(rows)
    _save_export(rows, token, "csv")

    buf = io.BytesIO(csv_data.encode("utf-8-sig"))
    return send_file(buf, as_attachment=True, download_name="subject_output.csv", mimetype="text/csv")


@app.route("/download/<token>/xlsx", methods=["POST"])
def download_xlsx(token):
    rows = _get_edited_rows(token)
    if rows is None:
        return "結果已過期，請重新上傳辨識", 404

    _save_export(rows, token, "xlsx")
    buf = _rows_to_xlsx(rows)

    return send_file(
        buf, as_attachment=True,
        download_name="subject_output.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )



@app.route("/download/<token>/append", methods=["POST"])
def download_append(token):
    """Append current data to an existing CSV export file."""
    stored_rows = RESULTS.get(token)
    if stored_rows is None:
        return "結果已過期，請重新上傳辨識", 404

    data = request.get_json(silent=True)
    if not data or not data.get("target"):
        return "Missing target filename", 400

    target_name = data["target"]
    edited_rows = data.get("rows", stored_rows)

    if ".." in target_name or "/" in target_name:
        return "Invalid filename", 400
    target_path = os.path.join(EXPORT_DIR, target_name)
    if not os.path.isfile(target_path):
        return "目標檔案不存在", 404

    # Read existing CSV
    existing = _read_csv_file(target_path)

    # Merge: append new rows, skip duplicates by 收文文號
    existing_doc_nos = {r.get("收文文號") for r in existing if r.get("收文文號")}
    duplicates = []
    added = 0
    for r in edited_rows:
        doc_no = r.get("收文文號", "")
        if doc_no and doc_no in existing_doc_nos:
            duplicates.append(doc_no)
            continue
        existing.append(r)
        added += 1
        if doc_no:
            existing_doc_nos.add(doc_no)

    # If all rows were duplicates, don't write and inform user
    if added == 0:
        return jsonify({
            "ok": False,
            "all_duplicates": True,
            "duplicates": duplicates,
            "total_rows": len(existing),
            "filename": target_name,
        })

    # Sort by date and re-number
    _sort_and_renumber(existing)

    # Write back
    csv_data = _rows_to_csv(existing)
    with open(target_path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(csv_data)

    return jsonify({
        "ok": True,
        "all_duplicates": False,
        "duplicates": duplicates,
        "total_rows": len(existing),
        "filename": target_name,
    })


@app.route("/exports", methods=["GET"])
def list_exports():
    files = []
    for name in sorted(os.listdir(EXPORT_DIR), reverse=True):
        if name.endswith((".csv", ".xlsx")):
            path = os.path.join(EXPORT_DIR, name)
            stat = os.stat(path)
            fmt = "XLSX" if name.endswith(".xlsx") else "CSV"
            files.append({
                "filename": name,
                "size": stat.st_size,
                "time": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                "format": fmt,
            })
    return jsonify(files)


@app.route("/exports/<filename>", methods=["GET"])
def download_export(filename):
    if ".." in filename or "/" in filename:
        return "Invalid filename", 400
    path = os.path.join(EXPORT_DIR, filename)
    if not os.path.isfile(path):
        return "檔案不存在", 404
    mimes = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".csv": "text/csv",
    }
    ext = os.path.splitext(filename)[1]
    mime = mimes.get(ext, "application/octet-stream")
    return send_file(path, as_attachment=True, download_name=filename, mimetype=mime)


@app.route("/exports/<filename>/preview", methods=["GET"])
def preview_export(filename):
    if ".." in filename or "/" in filename:
        return "Invalid filename", 400
    if not filename.endswith(".csv"):
        return "只支援預覽 CSV 檔案", 400
    path = os.path.join(EXPORT_DIR, filename)
    if not os.path.isfile(path):
        return "檔案不存在", 404
    rows = _read_csv_file(path)
    return jsonify({"headers": HEADERS, "rows": rows})


@app.route("/exports/<filename>/save", methods=["POST"])
def save_export(filename):
    """Save edited rows back to a CSV file."""
    if ".." in filename or "/" in filename:
        return "Invalid filename", 400
    if not filename.endswith(".csv"):
        return "只支援編輯 CSV 檔案", 400
    path = os.path.join(EXPORT_DIR, filename)
    if not os.path.isfile(path):
        return "檔案不存在", 404

    rows = request.get_json(silent=True)
    if not rows or not isinstance(rows, list):
        return "無效的資料", 400

    _renumber(rows)
    csv_data = _rows_to_csv(rows)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(csv_data)

    return jsonify({"ok": True, "rows": len(rows)})


@app.route("/exports/<filename>/download/<fmt>", methods=["GET"])
def download_export_as(filename, fmt):
    """Download a history CSV file as CSV or XLSX."""
    if ".." in filename or "/" in filename:
        return "Invalid filename", 400
    if fmt not in ("csv", "xlsx"):
        return "不支援的格式", 400
    path = os.path.join(EXPORT_DIR, filename)
    if not os.path.isfile(path):
        return "檔案不存在", 404

    if filename.endswith(".csv") and fmt == "xlsx":
        rows = _read_csv_file(path)
        buf = _rows_to_xlsx(rows)
        dl_name = os.path.splitext(filename)[0] + ".xlsx"
        return send_file(
            buf, as_attachment=True, download_name=dl_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Default: serve the file as-is
    mimes = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".csv": "text/csv",
    }
    ext = os.path.splitext(filename)[1]
    mime = mimes.get(ext, "application/octet-stream")
    return send_file(path, as_attachment=True, download_name=filename, mimetype=mime)


@app.route("/exports/<filename>", methods=["DELETE"])
def delete_export(filename):
    if ".." in filename or "/" in filename:
        return "Invalid filename", 400
    path = os.path.join(EXPORT_DIR, filename)
    if not os.path.isfile(path):
        return "檔案不存在", 404
    os.remove(path)
    return jsonify({"ok": True})


@app.route("/exports/<filename>/rename", methods=["POST"])
def rename_export(filename):
    if ".." in filename or "/" in filename:
        return "Invalid filename", 400
    data = request.get_json(silent=True)
    if not data or not data.get("new_name"):
        return "Missing new_name", 400

    new_name = data["new_name"].strip()
    if ".." in new_name or "/" in new_name or "\\" in new_name:
        return "Invalid new name", 400

    old_ext = os.path.splitext(filename)[1]
    if not new_name.endswith((".csv", ".xlsx")):
        new_name += old_ext

    old_path = os.path.join(EXPORT_DIR, filename)
    new_path = os.path.join(EXPORT_DIR, new_name)
    if not os.path.isfile(old_path):
        return "檔案不存在", 404
    if os.path.exists(new_path) and old_path != new_path:
        return "檔名已存在", 409

    os.rename(old_path, new_path)
    return jsonify({"ok": True, "new_name": new_name})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)
