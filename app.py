"""
Daily Sales Register dashboard
Run start.bat or:  python app.py
Then open http://127.0.0.1:5050
"""

from __future__ import annotations

import json
import os
import re
import threading
import traceback
import uuid
from datetime import datetime

from flask import Flask, Response, jsonify, render_template, request

from analytics import (
    compute_charts,
    compute_from_rows,
    compute_report_analytics,
    compute_team_performance,
    iter_export_rows,
    list_filter_options,
)
from dates import parse_date_from_name
from ingest import cache_path, ingest_file, is_ready, query_rows, read_summary
from schema import COLUMNS

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data")
CACHE_DIR = os.path.join(DATA_DIR, "cache")
DOWNLOAD_DIR = os.path.join(DATA_DIR, "downloads")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
CONFIG_PATH = os.path.join(ROOT, "config.json")

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 400 * 1024 * 1024
app.config["TEMPLATES_AUTO_RELOAD"] = True

_lock = threading.Lock()
_files: dict[str, dict] = {}
_jobs: dict[str, dict] = {}
_active: dict = {"folderId": "", "localPath": "", "latestDate": ""}


def _ensure_dirs() -> None:
    for path in (DATA_DIR, CACHE_DIR, DOWNLOAD_DIR, UPLOAD_DIR):
        os.makedirs(path, exist_ok=True)


def _load_config() -> dict:
    if not os.path.isfile(CONFIG_PATH):
        return {"defaultFolderId": "", "localFolder": ""}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return {
            "defaultFolderId": str(data.get("defaultFolderId") or data.get("default_folder_id") or ""),
            "localFolder": str(data.get("localFolder") or data.get("local_folder") or ""),
        }
    except (OSError, json.JSONDecodeError):
        return {"defaultFolderId": "", "localFolder": ""}


def _sanitize_folder_id(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if not re.fullmatch(r"[a-zA-Z0-9_-]{10,80}", value):
        raise ValueError("That folder ID looks invalid. Paste only the ID from the Drive folder URL.")
    return value


def _safe_local_path(value: str) -> str:
    path = os.path.abspath(os.path.expanduser((value or "").strip().strip('"')))
    if not path or not os.path.isdir(path):
        raise ValueError("That local folder does not exist.")
    return path


def _file_payload(meta: dict) -> dict:
    return {
        "dateKey": meta["dateKey"],
        "name": meta.get("name") or "",
        "size": int(meta.get("size") or 0),
        "source": meta.get("source") or "",
        "fingerprint": meta["fingerprint"],
        "cached": is_ready(cache_path(CACHE_DIR, meta["fingerprint"])),
    }


def _cache_ready(meta: dict) -> bool:
    fp = meta.get("fingerprint") or ""
    return bool(fp) and is_ready(cache_path(CACHE_DIR, fp))


def _register_file(meta: dict) -> None:
    key = meta["dateKey"]
    existing = _files.get(key)
    if not existing:
        _files[key] = meta
        return

    def rank(item: dict) -> tuple:
        return (
            1 if _cache_ready(item) else 0,
            str(item.get("modifiedTime") or ""),
            int(item.get("size") or 0),
        )

    if rank(meta) > rank(existing):
        _files[key] = meta


def _scan_local(folder: str) -> int:
    count = 0
    for dirpath, dirnames, filenames in os.walk(folder):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")][:50]
        for name in filenames:
            if not name.lower().endswith((".xlsx", ".xls", ".xlsm")):
                continue
            if name.startswith("~$"):
                continue
            date_key = parse_date_from_name(name)
            if not date_key:
                continue
            full = os.path.join(dirpath, name)
            try:
                st = os.stat(full)
            except OSError:
                continue
            _register_file(
                {
                    "dateKey": date_key,
                    "name": name,
                    "path": full,
                    "size": st.st_size,
                    "modifiedTime": datetime.fromtimestamp(st.st_mtime).isoformat(),
                    "source": "local",
                    "fingerprint": f"local_{st.st_size}_{int(st.st_mtime)}_{name}",
                }
            )
            count += 1
        if dirpath.count(os.sep) - folder.count(os.sep) >= 2:
            dirnames[:] = []
    return count


def _scan_stored_uploads() -> int:
    """Re-attach previously uploaded Excel files so a restart still has data."""
    if not os.path.isdir(UPLOAD_DIR):
        return 0
    count = 0
    for name in os.listdir(UPLOAD_DIR):
        if not name.lower().endswith((".xlsx", ".xls", ".xlsm")) or name.startswith("~$"):
            continue
        full = os.path.join(UPLOAD_DIR, name)
        if not os.path.isfile(full):
            continue
        date_key = parse_date_from_name(name)
        if not date_key:
            continue
        try:
            st = os.stat(full)
        except OSError:
            continue
        display = name.split("_", 2)[-1] if name.count("_") >= 2 else name
        _register_file(
            {
                "dateKey": date_key,
                "name": display,
                "path": full,
                "size": st.st_size,
                "modifiedTime": datetime.fromtimestamp(st.st_mtime).isoformat(),
                "source": "upload",
                "fingerprint": f"upload_{st.st_size}_{int(st.st_mtime)}_{display}",
            }
        )
        count += 1
    return count


def _restore_known_files() -> int:
    cfg = _load_config()
    found = _scan_stored_uploads()
    local = (cfg.get("localFolder") or "").strip()
    if local and os.path.isdir(local):
        try:
            found += _scan_local(_safe_local_path(local))
            _active["localPath"] = os.path.abspath(local)
        except Exception:
            pass
    folder_id = (cfg.get("defaultFolderId") or "").strip()
    if folder_id:
        _active["folderId"] = folder_id
    return found


def _save_config(folder_id: str = "", local_path: str = "") -> None:
    cfg = _load_config()
    if folder_id:
        cfg["defaultFolderId"] = folder_id
    if local_path:
        cfg["localFolder"] = local_path
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2)
    except OSError:
        pass


def _list_payload() -> dict:
    files = sorted(_files.values(), key=lambda f: f["dateKey"], reverse=True)
    dates = [f["dateKey"] for f in files]
    return {
        "ok": True,
        "dates": dates,
        "latestDate": dates[0] if dates else "",
        "files": [_file_payload(f) for f in files],
        "folderId": _active.get("folderId") or "",
        "localPath": _active.get("localPath") or "",
    }


def _job_view(job: dict) -> dict:
    return {
        "ok": job.get("status") != "error",
        "jobId": job["id"],
        "status": job["status"],
        "progress": job.get("progress") or 0,
        "message": job.get("message") or "",
        "error": job.get("error") or "",
        "dateKey": job.get("dateKey") or "",
        "fileName": job.get("fileName") or "",
        "fileSize": job.get("fileSize") or 0,
        "kpis": job.get("kpis"),
        "rowCount": job.get("rowCount") or 0,
        "cached": bool(job.get("cached")),
        "columns": job.get("columns") or COLUMNS,
    }


def _set_progress(job_id: str, pct: int, message: str) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        job["progress"] = pct
        job["message"] = message


def _prepare_source(meta: dict, job_id: str) -> str:
    if meta.get("path") and os.path.isfile(meta["path"]):
        return meta["path"]
    if meta.get("source") == "drive":
        import drive_io

        dest = os.path.join(DOWNLOAD_DIR, f"{meta['id']}.xlsx")
        stale = True
        if os.path.isfile(dest) and os.path.getsize(dest) == int(meta.get("size") or 0) and meta.get("size"):
            stale = False
        if stale:
            _set_progress(job_id, 3, "Downloading from Google Drive…")
            drive_io.download_file(ROOT, meta, dest, lambda p, m: _set_progress(job_id, p, m))
        return dest
    raise FileNotFoundError("The source file is no longer available.")


def _run_job(job_id: str, date_key: str, bypass: bool) -> None:
    try:
        with _lock:
            meta = dict(_files[date_key])
        db = cache_path(CACHE_DIR, meta["fingerprint"])
        if not bypass and is_ready(db):
            summary = read_summary(db)
            with _lock:
                job = _jobs[job_id]
                job.update(
                    {
                        "status": "ready",
                        "progress": 100,
                        "message": "Loaded from cache",
                        "kpis": summary["kpis"],
                        "rowCount": summary["rowCount"],
                        "cached": True,
                        "fileName": meta.get("name"),
                        "fileSize": meta.get("size") or 0,
                        "columns": summary.get("columns") or COLUMNS,
                    }
                )
            return

        src = _prepare_source(meta, job_id)
        _set_progress(job_id, 10, "Parsing Excel and computing KPIs…")
        stats = ingest_file(src, db, lambda p, m: _set_progress(job_id, p, m))
        with _lock:
            job = _jobs[job_id]
            job.update(
                {
                    "status": "ready",
                    "progress": 100,
                    "message": f"Indexed {stats['invoiceCount']:,} invoices",
                    "kpis": stats,
                    "rowCount": stats["invoiceCount"],
                    "cached": False,
                    "fileName": meta.get("name"),
                    "fileSize": meta.get("size") or 0,
                    "columns": stats.get("columns") or COLUMNS,
                }
            )
    except Exception as exc:
        traceback.print_exc()
        with _lock:
            job = _jobs[job_id]
            job["status"] = "error"
            job["error"] = str(exc) or "Failed to load this workbook."
            job["message"] = job["error"]


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/config")
def api_config():
    cfg = _load_config()
    import drive_io

    return jsonify(
        {
            "ok": True,
            "defaultFolderId": cfg["defaultFolderId"],
            "localFolder": cfg["localFolder"],
            "driveReady": drive_io.has_oauth_files(ROOT),
            "columns": COLUMNS,
        }
    )


@app.post("/api/connect")
def api_connect():
    body = request.get_json(silent=True) or {}
    folder_id = (body.get("folderId") or "").strip()
    local_path = (body.get("localPath") or "").strip()
    cfg = _load_config()
    if not folder_id:
        folder_id = cfg["defaultFolderId"]
    if not local_path:
        local_path = cfg["localFolder"]

    with _lock:
        _files.clear()
        _active["folderId"] = ""
        _active["localPath"] = ""
        _scan_stored_uploads()

    errors = []
    found = 0

    if folder_id:
        try:
            folder_id = _sanitize_folder_id(folder_id)
            import drive_io

            remote = drive_io.list_folder_files(ROOT, folder_id)
            with _lock:
                for item in remote:
                    item["fingerprint"] = (
                        f"drive_{item['id']}_{item.get('modifiedTime')}_{item.get('size')}"
                    )
                    _register_file(item)
                _active["folderId"] = folder_id
            found += len(remote)
        except Exception as exc:
            errors.append(str(exc))

    if local_path:
        try:
            folder = _safe_local_path(local_path)
            added = _scan_local(folder)
            with _lock:
                _active["localPath"] = folder
            found += added
        except Exception as exc:
            errors.append(str(exc))

    with _lock:
        payload = _list_payload()
    if not payload["dates"]:
        msg = errors[0] if errors else (
            "No dated Excel files were found. Use names like "
            "'Daily Sales Register on 18-Aug-2026.xlsx'."
        )
        return jsonify({"ok": False, "error": msg}), 400
    if folder_id or local_path:
        _save_config(folder_id=folder_id, local_path=local_path)
    payload["warnings"] = errors
    return jsonify(payload)


@app.get("/api/reports")
def api_reports():
    with _lock:
        if not _files:
            _restore_known_files()
        payload = _list_payload()
    return jsonify(payload)


@app.post("/api/load")
def api_load():
    body = request.get_json(silent=True) or {}
    date_key = (body.get("dateKey") or "").strip()
    bypass = bool(body.get("bypassCache"))
    with _lock:
        if not _files:
            _restore_known_files()
        if not _files:
            return jsonify({"ok": False, "error": "Connect a Drive or local folder first, or upload an Excel file."}), 400
        if not date_key:
            date_key = sorted(_files.keys(), reverse=True)[0]
        if date_key not in _files:
            return jsonify({"ok": False, "error": f"No file matches {date_key}."}), 404
        meta = _files[date_key]
        job_id = uuid.uuid4().hex[:12]
        _jobs[job_id] = {
            "id": job_id,
            "status": "running",
            "progress": 1,
            "message": "Starting…",
            "dateKey": date_key,
            "fileName": meta.get("name"),
            "fileSize": meta.get("size") or 0,
            "error": "",
            "kpis": None,
            "rowCount": 0,
            "cached": False,
        }
    thread = threading.Thread(target=_run_job, args=(job_id, date_key, bypass), daemon=True)
    thread.start()
    with _lock:
        return jsonify(_job_view(_jobs[job_id]))


@app.get("/api/job/<job_id>")
def api_job(job_id: str):
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return jsonify({"ok": False, "error": "Unknown job."}), 404
        return jsonify(_job_view(job))


@app.route("/api/rows", methods=["GET", "POST"])
def api_rows():
    src = request.get_json(silent=True) if request.method == "POST" else request.args
    src = src or {}
    date_key = str(src.get("dateKey") or "").strip()
    search = str(src.get("search") or "").strip()
    sort_col = str(src.get("sortCol") or "").strip()
    sort_dir = str(src.get("sortDir") or "asc").strip()
    try:
        start = int(src.get("start") if src.get("start") is not None else src.get("startRow") or 0)
        end = int(src.get("end") if src.get("end") is not None else src.get("endRow") or 100)
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "Invalid page range."}), 400
    with _lock:
        if not _files:
            _restore_known_files()
        meta = _files.get(date_key)
    if not meta:
        return jsonify({"ok": False, "error": "Load a report date first."}), 400
    db = cache_path(CACHE_DIR, meta["fingerprint"])
    if not is_ready(db):
        return jsonify({"ok": False, "error": "This file is still being indexed."}), 409
    try:
        payload = query_rows(
            db,
            start,
            end,
            search,
            sort_col,
            sort_dir,
            parse_filters(src),
            parse_grid_filters(src),
        )
        payload["ok"] = True
        return jsonify(payload)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


def parse_filters(src) -> dict:
    def getlist(key):
        if hasattr(src, "getlist"):
            vals = [str(v).strip() for v in src.getlist(key) if str(v).strip()]
            if len(vals) == 1 and "," in vals[0]:
                return [x.strip() for x in vals[0].split(",") if x.strip()][:40]
            return vals[:40]
        raw = src.get(key) if hasattr(src, "get") else None
        if isinstance(raw, str):
            return [x.strip() for x in raw.split(",") if x.strip()][:40]
        if isinstance(raw, list):
            return [str(x).strip() for x in raw if str(x).strip()][:40]
        return []

    field = str(src.get("dateField") or "Invoice Date").strip()
    if field not in ("Invoice Date", "Service Period"):
        field = "Invoice Date"
    date_from = str(src.get("dateFrom") or "").strip()[:10]
    date_to = str(src.get("dateTo") or "").strip()[:10]
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_from):
        date_from = ""
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_to):
        date_to = ""
    return {
        "dateFrom": date_from,
        "dateTo": date_to,
        "dateField": field,
        "currency": str(src.get("currency") or "").strip()[:20],
        "brands": getlist("brands"),
        "teams": getlist("teams"),
    }


def parse_grid_filters(src) -> dict:
    raw = src.get("filterModel") if hasattr(src, "get") else None
    if raw is None and hasattr(src, "get"):
        raw = src.get("filter_model")
    if isinstance(raw, str) and raw.strip():
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return {}
    if not isinstance(raw, dict):
        return {}
    out = {}
    for i, (key, val) in enumerate(raw.items()):
        if i >= 40:
            break
        label = str(key).strip()[:120]
        if label and isinstance(val, dict):
            out[label] = val
    return out


def _db_for_date(date_key: str):
    with _lock:
        if not _files:
            _restore_known_files()
        meta = _files.get(date_key)
    if not meta:
        return None, "Load a report date first."
    db = cache_path(CACHE_DIR, meta["fingerprint"])
    if not is_ready(db):
        return None, "This file is still being indexed."
    return db, None


@app.get("/api/stats")
def api_stats():
    date_key = (request.args.get("dateKey") or "").strip()
    search = (request.args.get("search") or "").strip()
    db, err = _db_for_date(date_key)
    if err:
        return jsonify({"ok": False, "error": err}), 400 if "Load" in err else 409
    try:
        stats = compute_report_analytics(db, search=search, filters=parse_filters(request.args))
        try:
            charts = compute_charts(db, search=search, filters=parse_filters(request.args))
        except Exception:
            traceback.print_exc()
            charts = None
        return jsonify({"ok": True, "all": stats, "charts": charts})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.post("/api/stats/selected")
def api_stats_selected():
    body = request.get_json(silent=True) or {}
    date_key = str(body.get("dateKey") or "").strip()
    search = str(body.get("search") or "").strip()
    invoices = body.get("invoices") or []
    rows = body.get("rows") or []
    if not isinstance(invoices, list):
        invoices = []
    if not isinstance(rows, list):
        rows = []
    db, err = _db_for_date(date_key)
    if err:
        return jsonify({"ok": False, "error": err}), 400 if "Load" in err else 409
    try:
        if invoices:
            selected = compute_report_analytics(db, search=search, invoice_numbers=invoices)
        else:
            selected = compute_from_rows(rows)
        return jsonify({"ok": True, "selected": selected})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.get("/api/team")
def api_team():
    date_key = (request.args.get("dateKey") or "").strip()
    search = (request.args.get("search") or "").strip()
    db, err = _db_for_date(date_key)
    if err:
        return jsonify({"ok": False, "error": err}), 400 if "Load" in err else 409
    try:
        payload = compute_team_performance(
            db,
            search=search,
            filters=parse_filters(request.args),
            month=str(request.args.get("month") or "all"),
            segment=str(request.args.get("segment") or "all"),
            aging=str(request.args.get("aging") or "all"),
            group_by=str(request.args.get("groupBy") or "manager"),
        )
        payload["ok"] = True
        return jsonify(payload)
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.get("/api/filters")
def api_filters():
    date_key = (request.args.get("dateKey") or "").strip()
    db, err = _db_for_date(date_key)
    if err:
        return jsonify({"ok": False, "error": err}), 400 if "Load" in err else 409
    try:
        return jsonify({"ok": True, **list_filter_options(db)})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/charts", methods=["GET"])
def api_charts():
    date_key = (request.args.get("dateKey") or "").strip()
    search = (request.args.get("search") or "").strip()
    db, err = _db_for_date(date_key)
    if err:
        return jsonify({"ok": False, "error": err}), 400 if "Load" in err else 409
    try:
        charts = compute_charts(db, search=search, filters=parse_filters(request.args))
        return jsonify({"ok": True, "charts": charts})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.get("/api/export")
def api_export():
    import csv
    import io

    date_key = (request.args.get("dateKey") or "").strip()
    search = (request.args.get("search") or "").strip()
    fmt = str(request.args.get("format") or "csv").lower()
    db, err = _db_for_date(date_key)
    if err:
        return jsonify({"ok": False, "error": err}), 400 if "Load" in err else 409
    filters = parse_filters(request.args)
    rows_iter = iter_export_rows(db, search=search, filters=filters)
    labels = next(rows_iter)
    stamp = date_key or "export"
    if fmt == "xlsx":
        from openpyxl import Workbook

        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Daily Sales Register")
        ws.append(labels)
        count = 0
        for rec in rows_iter:
            ws.append(list(rec))
            count += 1
            if count >= 40000:
                break
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return Response(
            buf.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=DSR_{stamp}.xlsx"},
        )
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(labels)
    for rec in rows_iter:
        writer.writerow(rec)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=DSR_{stamp}.csv"},
    )


@app.post("/api/upload")
def api_upload():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"ok": False, "error": "Choose an Excel file to upload."}), 400
    name = os.path.basename(uploaded.filename)
    if not name.lower().endswith((".xlsx", ".xls", ".xlsm")):
        return jsonify({"ok": False, "error": "Only Excel files (.xlsx) are supported."}), 400
    date_key = parse_date_from_name(name)
    if not date_key:
        date_key = datetime.now().strftime("%Y-%m-%d")
    _ensure_dirs()
    dest = os.path.join(UPLOAD_DIR, f"{date_key}_{uuid.uuid4().hex[:8]}_{name}")
    uploaded.save(dest)
    st = os.stat(dest)
    meta = {
        "dateKey": date_key,
        "name": name,
        "path": dest,
        "size": st.st_size,
        "modifiedTime": datetime.fromtimestamp(st.st_mtime).isoformat(),
        "source": "upload",
        "fingerprint": f"upload_{st.st_size}_{int(st.st_mtime)}_{name}",
    }
    with _lock:
        _files[date_key] = meta
    return jsonify({"ok": True, "dateKey": date_key, "file": _file_payload(meta)})


@app.get("/favicon.ico")
def favicon():
    return ("", 204)


def main() -> None:
    _ensure_dirs()
    restored = _restore_known_files()
    print("Daily Sales Register")
    print("Open http://127.0.0.1:5050")
    if restored:
        print(f"Restored {restored} saved report file(s) from previous uploads.")
    print("150+ MB Excel files are streamed, cached, and shown via virtual scrolling.")
    app.run(host="127.0.0.1", port=5050, debug=False, threaded=True)


if __name__ == "__main__":
    main()
