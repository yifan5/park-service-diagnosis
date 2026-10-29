"""Small asynchronous HTTP API for the park review analysis demo."""

from __future__ import annotations

import cgi
import json
import mimetypes
import os
import shutil
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import parse_qs, quote, urlparse

from .analysis import (
    ASPECTS,
    DEFAULT_SETTINGS,
    InputValidationError,
    build_overview,
    build_screening_rows,
    create_attributions,
    enrich_frame,
    export_xlsx,
    group_evidence,
    make_adapter,
    normalize_rows,
    read_input_file,
    safe_settings,
    split_batches,
    _compact_ds_guide,
    _load_ds_scopes,
    load_initial_result,
)
import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
FRONTEND_DIR = BASE_DIR / "frontend"
JOBS_DIR = BASE_DIR / "var" / "jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)
MAX_UPLOAD_BYTES = int(os.environ.get("PARK_MAX_UPLOAD_BYTES", str(100 * 1024 * 1024)))
MAX_ROWS = int(os.environ.get("PARK_MAX_ROWS", "500000"))
EXECUTOR = ThreadPoolExecutor(max_workers=4)


class JobStore:
    def __init__(self) -> None:
        self.lock = threading.RLock()
        self.jobs: Dict[str, Dict[str, Any]] = {}

    def create(self, upload_path: Path, filename: str, settings: Dict[str, Any], mapping: Dict[str, str], job_id: Optional[str] = None) -> Dict[str, Any]:
        job_id = job_id or uuid.uuid4().hex
        job_dir = upload_path.parent
        job = {
            "id": job_id,
            "filename": filename,
            "job_dir": str(job_dir),
            "upload_path": str(upload_path),
            "settings": settings,
            "mapping": mapping,
            "status": "queued",
            "phase": "排队中",
            "created_at": time.time(),
            "updated_at": time.time(),
            "processed_rows": 0,
            "total_rows": 0,
            "completed_batches": 0,
            "total_batches": 0,
            "failed_batches": 0,
            "attribution_done": 0,
            "attribution_total": 0,
            "attribution_stage": "等待归因",
            "attribution_stage_index": 0,
            "attribution_stage_total": 4,
            "attribution_stage_done": 0,
            "attribution_stage_items": 0,
            "errors": [],
            "overview": None,
            "attributions": [],
        }
        with self.lock:
            self.jobs[job_id] = job
        return job

    def update(self, job_id: str, **changes: Any) -> None:
        with self.lock:
            job = self.jobs[job_id]
            job.update(changes)
            job["updated_at"] = time.time()

    def get(self, job_id: str) -> Optional[Dict[str, Any]]:
        with self.lock:
            job = self.jobs.get(job_id)
            return dict(job) if job else None


STORE = JobStore()


def _public_job(job: Dict[str, Any]) -> Dict[str, Any]:
    result = {key: value for key, value in job.items() if key not in ("job_dir", "upload_path", "overview", "attributions")}
    result["overview_ready"] = bool(job.get("overview"))
    return result


def _load_persisted_job(job_id: str) -> Optional[Dict[str, Any]]:
    """Rehydrate a completed job from its result workbooks after a restart."""
    job_dir = JOBS_DIR / job_id
    initial_path = job_dir / "初步分析结果.xlsx"
    summary_path = job_dir / "问题归因总结.xlsx"
    if not initial_path.exists():
        return None
    try:
        frame = pd.read_excel(initial_path)
        park_col = next((c for c in ("公园名", "公园", "park") if c in frame.columns), None)
        comment_col = next((c for c in ("评论", "评论内容", "comment", "review") if c in frame.columns), None)
        if not park_col or not comment_col:
            return None
        row_col = "原始行号" if "原始行号" in frame.columns else None
        rows = []
        scores = {}
        for index, record in frame.iterrows():
            source_row = int(record[row_col]) if row_col and pd.notna(record[row_col]) else index + 2
            park = "" if pd.isna(record[park_col]) else str(record[park_col]).strip()
            comment = "" if pd.isna(record[comment_col]) else str(record[comment_col]).strip()
            if not park or not comment:
                continue
            rows.append({"source_row": source_row, "park": park, "comment": comment})
            scores[source_row] = {aspect: int(pd.to_numeric(pd.Series([record[aspect]]), errors="coerce").fillna(0).iloc[0]) if aspect in frame.columns else 0 for aspect in ASPECTS}
        attributions = []
        if summary_path.exists():
            workbook_sheets = pd.ExcelFile(summary_path).sheet_names
            summary = pd.read_excel(summary_path, sheet_name="归因汇总")
            evidence = pd.read_excel(summary_path, sheet_name="证据明细") if "证据明细" in workbook_sheets else pd.DataFrame()
            excluded_frame = pd.read_excel(summary_path, sheet_name="范围排除") if "范围排除" in workbook_sheets else pd.DataFrame()
            for _, item in summary.iterrows():
                park = str(item.get("公园名", "")); aspect = str(item.get("方面", ""))
                negatives = [x for x in group_evidence(rows, scores).get((park, aspect), [])]
                subset = evidence[(evidence.get("公园名", pd.Series(dtype=str)).astype(str) == park) & (evidence.get("方面", pd.Series(dtype=str)).astype(str) == aspect)] if not evidence.empty else pd.DataFrame()
                cards = []
                if not subset.empty and "归因短句" in subset.columns:
                    for text, group in subset.groupby("归因短句", sort=False):
                        cards.append({"text": str(text), "source_rows": sorted({int(value) for value in group["原始行号"].tolist()})})
                elif str(item.get("问题列表", "")).strip():
                    cards = [{"text": text.strip(), "source_rows": [x["source_row"] for x in negatives]} for text in str(item.get("问题列表", "")).split("；") if text.strip()]
                excluded = []
                if not excluded_frame.empty:
                    selected_excluded = excluded_frame[(excluded_frame.get("公园名", pd.Series(dtype=str)).astype(str) == park) & (excluded_frame.get("方面", pd.Series(dtype=str)).astype(str) == aspect)]
                    excluded = [{"source_row": int(value.get("原始行号", 0)), "reason": str(value.get("排除原因", ""))} for _, value in selected_excluded.iterrows()]
                summary_text = item.get("问题概要", item.get("问题摘要", ""))
                attributions.append({"公园名": park, "方面": aspect, "负面评论数": int(item.get("负面评论数", len(negatives)) or len(negatives)), "归因状态": str(item.get("归因状态", "完成")), "归因错误": str(item.get("归因错误", "")) if pd.notna(item.get("归因错误", "")) else "", "问题摘要": str(summary_text), "问题列表": "；".join(card["text"] for card in cards), "问题卡片": cards, "负面证据行号": ",".join(str(x["source_row"]) for x in negatives), "负面证据": negatives, "范围排除": excluded})
        settings = safe_settings({"provider": "persisted"})
        overview = build_overview(rows, scores, attributions, settings, [])
        overview["adapter"] = "persisted"
        overview["selected_columns"] = {"park": park_col, "comment": comment_col}
        return {"id": job_id, "filename": "已保存分析结果", "job_dir": str(job_dir), "upload_path": str(job_dir / "input.xlsx"), "settings": settings, "mapping": {}, "status": "completed", "phase": "完成", "created_at": job_dir.stat().st_ctime, "updated_at": job_dir.stat().st_mtime, "processed_rows": len(rows), "total_rows": len(rows), "completed_batches": 0, "total_batches": 0, "failed_batches": 0, "attribution_done": len(attributions), "attribution_total": len(attributions), "errors": [], "overview": overview, "attributions": attributions}
    except Exception:
        return None


def _get_job(job_id: str) -> Optional[Dict[str, Any]]:
    job = STORE.get(job_id)
    if job:
        return job
    persisted = _load_persisted_job(job_id)
    if persisted:
        with STORE.lock:
            STORE.jobs[job_id] = persisted
        return dict(persisted)
    return None


def _run_job(job: Dict[str, Any]) -> None:
    job_id = job["id"]
    job_dir = Path(job["job_dir"])
    settings = job["settings"]
    try:
        STORE.update(job_id, status="running", phase="读取并校验文件")
        frame, selected = read_input_file(Path(job["upload_path"]), job.get("mapping"))
        if len(frame.index) > MAX_ROWS:
            raise InputValidationError("文件超过 {} 行的 Demo 上限。".format(MAX_ROWS))
        rows, validation_errors = normalize_rows(frame, selected)
        if not rows:
            raise InputValidationError("没有可分析的有效行，请检查公园和评论列。")
        STORE.update(job_id, total_rows=len(rows), errors=validation_errors, phase="分批调用模型")
        ds_guide = _compact_ds_guide(BASE_DIR / "DS prompt.md")
        ds_scopes = _load_ds_scopes(BASE_DIR / "DS prompt.md")
        adapter = make_adapter(settings["provider"])
        batches = split_batches(rows, settings["batch_size"], settings["max_chars"])
        STORE.update(job_id, total_batches=len(batches))
        score_map: Dict[int, Dict[str, int]] = {}
        failed_rows = set()
        def analyze_one(batch_index: int, batch: Any) -> Any:
            last_error = ""
            for attempt in range(settings["retries"] + 1):
                try:
                    return batch_index, batch, adapter.analyze_batch(batch, settings, ds_guide), ""
                except Exception as exc:  # model adapters intentionally normalize failures here
                    last_error = str(exc)
                    if attempt < settings["retries"]:
                        time.sleep(min(2 ** attempt, 4))
            return batch_index, batch, None, last_error

        completed_rows = 0
        with ThreadPoolExecutor(max_workers=settings["concurrency"]) as batch_pool:
            futures = [batch_pool.submit(analyze_one, index, batch) for index, batch in enumerate(batches, 1)]
            for completed_index, future in enumerate(as_completed(futures), 1):
                batch_index, batch, batch_scores, last_error = future.result()
                completed_rows += len(batch)
                if batch_scores is None:
                    failed_rows.update(row["source_row"] for row in batch)
                    validation_errors.append({"批次": batch_index, "原始行号": ",".join(str(row["source_row"]) for row in batch), "错误": last_error or "模型调用失败"})
                else:
                    score_map.update(batch_scores)
                STORE.update(job_id, completed_batches=completed_index, processed_rows=completed_rows, failed_batches=sum(1 for item in validation_errors if "批次" in item), errors=validation_errors)

        # Failed rows receive an explicit analysis status in the workbook; their
        # zero scores are never used for screening or attribution.
        frame_with_scores = enrich_frame(frame, rows, score_map)
        frame_with_scores.insert(len(frame_with_scores.columns) - len(ASPECTS), "分析状态", ["完成" if int(index) + 2 in score_map else "未分析" for index in frame.index])
        screening = build_screening_rows(rows, score_map)
        export_xlsx(job_dir, frame_with_scores, screening, [], validation_errors)

        STORE.update(job_id, phase="生成问题归因", attribution_total=0)
        # Only groups containing at least one 1—2 score are sent to attribution.
        grouped = group_evidence(rows, score_map)
        STORE.update(job_id, attribution_total=len(grouped))
        attributions = create_attributions(
            rows, score_map, adapter, settings, ds_scopes,
            progress=lambda done, total: STORE.update(job_id, attribution_done=done, attribution_total=total),
        )
        export_xlsx(job_dir, frame_with_scores, screening, attributions, validation_errors)
        overview = build_overview(rows, score_map, attributions, settings, validation_errors)
        overview["selected_columns"] = selected
        overview["failed_row_count"] = len(failed_rows)
        overview["adapter"] = adapter.name
        failed_attributions = sum(1 for item in attributions if item.get("归因状态") == "失败")
        STORE.update(
            job_id,
            status="partial" if failed_rows or failed_attributions else "completed",
            phase="归因部分失败" if failed_attributions else ("完成" if not failed_rows else "部分完成"),
            errors=validation_errors,
            overview=overview,
            attributions=attributions,
        )
    except InputValidationError as exc:
        errors = [{"错误": str(exc)}]
        STORE.update(job_id, status="failed", phase="输入错误", errors=errors)
    except Exception as exc:
        STORE.update(job_id, status="failed", phase="任务失败", errors=[{"错误": str(exc)}])


def _import_initial_result(job_id: str, upload_path: Path, filename: str, requested_settings: Optional[Dict[str, Any]] = None) -> None:
    job_dir = upload_path.parent
    try:
        frame, rows, scores, errors = load_initial_result(upload_path)
        if not rows:
            raise InputValidationError("初步结果中没有可展示的有效公园和评论记录。")
        grouped = group_evidence(rows, scores)
        screening = build_screening_rows(rows, scores)
        # Persist the scored overview immediately. Attribution runs separately
        # and never invokes the 22-aspect scoring model.
        export_xlsx(job_dir, frame, screening, [], errors)
        settings = safe_settings({**(requested_settings or {}), "provider": "codex"})
        STORE.update(job_id, settings=settings)
        overview = build_overview(rows, scores, [], settings, errors)
        overview["adapter"] = settings["provider"]
        overview["selected_columns"] = {"park": "自动识别", "comment": "自动识别"}
        STORE.update(job_id, status="running", phase="语义梳理", total_rows=len(rows), processed_rows=len(rows), completed_batches=0, total_batches=0, attribution_done=0, attribution_total=len(grouped), attribution_stage="语义梳理", attribution_stage_index=1, attribution_stage_total=4, attribution_stage_done=0, attribution_stage_items=len(grouped), errors=errors, overview=overview, attributions=[])
        adapter = make_adapter(settings["provider"])
        ds_guide = _load_ds_scopes(BASE_DIR / "DS prompt.md")
        attributions = create_attributions(
            rows, scores, adapter, settings, ds_guide,
            progress=lambda done, total: STORE.update(job_id, attribution_done=done, attribution_total=total, attribution_stage_done=done, attribution_stage_items=total),
            stage_progress=lambda stage, index, done, items: STORE.update(job_id, attribution_stage=stage, attribution_stage_index=index, attribution_stage_done=done, attribution_stage_items=items),
        )
        export_xlsx(job_dir, frame, screening, attributions, errors)
        overview = build_overview(rows, scores, attributions, settings, errors)
        overview["adapter"] = adapter.name
        overview["selected_columns"] = {"park": "自动识别", "comment": "自动识别"}
        failed = sum(1 for item in attributions if item.get("归因状态") == "失败")
        STORE.update(job_id, status="partial" if failed else "completed", phase="归因部分失败" if failed else "完成", attribution_stage="结果校验", attribution_stage_index=4, attribution_stage_done=len(grouped), attribution_stage_items=len(grouped), total_rows=len(rows), processed_rows=len(rows), completed_batches=0, total_batches=0, attribution_done=len(attributions), attribution_total=len(attributions), errors=errors, overview=overview, attributions=attributions)
    except InputValidationError as exc:
        STORE.update(job_id, status="failed", phase="初步结果格式错误", errors=[{"错误": str(exc), "columns": exc.columns, "missing": exc.missing}])
    except Exception as exc:
        STORE.update(job_id, status="failed", phase="导入失败", errors=[{"错误": str(exc)}])


def _retry_attributions(job_id: str) -> None:
    job = _get_job(job_id)
    if not job:
        return
    job_dir = Path(job["job_dir"])
    initial_path = job_dir / "初步分析结果.xlsx"
    try:
        frame, rows, scores, errors = load_initial_result(initial_path)
        grouped = group_evidence(rows, scores)
        existing = {(item.get("公园名"), item.get("方面")): item for item in job.get("attributions", [])}
        retry_keys = [key for key in grouped if key not in existing or existing[key].get("归因状态") == "失败"]
        if not retry_keys:
            STORE.update(job_id, phase="没有需要重试的归因单元")
            return
        settings = safe_settings(job.get("settings"))
        adapter = make_adapter(settings["provider"])
        scopes = _load_ds_scopes(BASE_DIR / "DS prompt.md")
        STORE.update(job_id, status="running", phase="语义梳理", attribution_stage="语义梳理", attribution_stage_index=1, attribution_stage_total=4, attribution_stage_done=0, attribution_stage_items=len(retry_keys), attribution_done=0, attribution_total=len(retry_keys))
        retried = create_attributions(
            rows, scores, adapter, settings, scopes,
            progress=lambda done, total: STORE.update(job_id, attribution_done=done, attribution_total=total),
            stage_progress=lambda stage, index, done, items: STORE.update(job_id, attribution_stage=stage, attribution_stage_index=index, attribution_stage_done=done, attribution_stage_items=items),
            only=retry_keys,
        )
        merged = {key: value for key, value in existing.items() if key not in set(retry_keys)}
        merged.update({(item["公园名"], item["方面"]): item for item in retried})
        attributions = [merged[key] for key in sorted(merged)]
        screening = build_screening_rows(rows, scores)
        export_xlsx(job_dir, frame, screening, attributions, errors)
        overview = build_overview(rows, scores, attributions, settings, errors)
        overview["adapter"] = adapter.name
        overview["selected_columns"] = job.get("overview", {}).get("selected_columns", {"park": "自动识别", "comment": "自动识别"})
        failed = sum(1 for item in attributions if item.get("归因状态") == "失败")
        STORE.update(job_id, status="partial" if failed else "completed", phase="归因部分失败" if failed else "完成", attribution_stage="结果校验", attribution_stage_index=4, attribution_stage_done=len(retry_keys), attribution_stage_items=len(retry_keys), attribution_done=len(retry_keys), attribution_total=len(retry_keys), overview=overview, attributions=attributions)
    except Exception as exc:
        STORE.update(job_id, status="partial", phase="归因重试失败", errors=list(job.get("errors", [])) + [{"错误": str(exc)}])


def cleanup_old_jobs() -> None:
    cutoff = time.time() - 24 * 60 * 60
    for path in JOBS_DIR.iterdir():
        try:
            if path.is_dir() and path.stat().st_mtime < cutoff:
                shutil.rmtree(path)
        except OSError:
            continue


class Handler(BaseHTTPRequestHandler):
    server_version = "ParkAnalysisDemo/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        # Keep model prompts and comment text out of the default server log.
        print("{} - {}".format(self.address_string(), fmt % args))

    def _send_json(self, payload: Any, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: Path, download_name: str) -> None:
        if not path.exists() or not path.is_file():
            self._send_json({"error": "文件尚未生成"}, 404)
            return
        content = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Content-Disposition", "attachment; filename*=UTF-8''{}".format(quote(download_name)))
        self.end_headers()
        self.wfile.write(content)

    def _serve_static(self, path: str) -> None:
        relative = "index.html" if path in ("", "/") else path.lstrip("/")
        file_path = (FRONTEND_DIR / relative).resolve()
        if FRONTEND_DIR.resolve() not in file_path.parents or not file_path.exists():
            self._send_json({"error": "未找到页面"}, 404)
            return
        content = file_path.read_bytes()
        content_type = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type + ("; charset=utf-8" if content_type.startswith("text/") or content_type == "application/javascript" else ""))
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"
        if path == "/api/schema":
            self._send_json({"aspects": ASPECTS, "defaults": DEFAULT_SETTINGS, "max_upload_bytes": MAX_UPLOAD_BYTES, "max_rows": MAX_ROWS, "aliases": {"park": ["公园名", "公园"], "comment": ["评论"]}})
            return
        parts = [part for part in path.split("/") if part]
        if len(parts) >= 3 and parts[0] == "api" and parts[1] == "jobs":
            job = _get_job(parts[2])
            if not job:
                self._send_json({"error": "任务不存在"}, 404)
                return
            if len(parts) == 3:
                payload = _public_job(job)
                if job.get("overview"):
                    payload["overview"] = job["overview"]
                    payload["attributions"] = job.get("attributions", [])
                self._send_json(payload)
                return
            if len(parts) == 4 and parts[3] == "overview":
                if not job.get("overview"):
                    self._send_json({"error": "概况尚未生成"}, 409)
                else:
                    self._send_json({"overview": job["overview"], "attributions": job.get("attributions", [])})
                return
            if len(parts) == 5 and parts[3] == "download":
                names = {"initial": "初步分析结果.xlsx", "screening": "筛选明细.xlsx", "attribution": "问题归因总结.xlsx"}
                filename = names.get(parts[4])
                if not filename:
                    self._send_json({"error": "未知下载类型"}, 404)
                else:
                    self._send_file(Path(job["job_dir"]) / filename, filename)
                return
        self._serve_static(path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        endpoint = parsed.path.rstrip("/")
        parts = [part for part in endpoint.split("/") if part]
        if len(parts) == 5 and parts[:2] == ["api", "jobs"] and parts[3:] == ["attribution", "retry"]:
            job = _get_job(parts[2])
            if not job:
                self._send_json({"error": "任务不存在"}, 404)
                return
            if job.get("status") == "running":
                self._send_json({"error": "任务正在运行"}, 409)
                return
            EXECUTOR.submit(_retry_attributions, parts[2])
            self._send_json({"job_id": parts[2], "status": "queued", "mode": "attribution_retry"}, 202)
            return
        if endpoint not in ("/api/jobs", "/api/import-results"):
            self._send_json({"error": "不支持的接口"}, 404)
            return
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0 or content_length > MAX_UPLOAD_BYTES:
            self._send_json({"error": "文件为空或超过上传大小限制"}, 413)
            return
        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data"):
            self._send_json({"error": "请使用 multipart/form-data 上传文件"}, 415)
            return
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type, "CONTENT_LENGTH": str(content_length)})
        file_item = form["file"] if "file" in form else None
        if file_item is None or not getattr(file_item, "filename", None):
            self._send_json({"error": "缺少 file 文件字段"}, 400)
            return
        original_name = os.path.basename(file_item.filename)
        suffix = Path(original_name).suffix.lower()
        if suffix not in (".xlsx", ".xls", ".csv"):
            self._send_json({"error": "仅支持 .xlsx、.xls 和 .csv 文件"}, 400)
            return
        try:
            raw_settings = json.loads(form.getfirst("config", "{}"))
        except json.JSONDecodeError:
            raw_settings = {}
        try:
            mapping = json.loads(form.getfirst("mapping", "{}"))
        except json.JSONDecodeError:
            mapping = {}
        cleanup_old_jobs()
        job_id = uuid.uuid4().hex
        job_dir = JOBS_DIR / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        upload_path = job_dir / ("input" + suffix)
        with upload_path.open("wb") as stream:
            shutil.copyfileobj(file_item.file, stream)
        if endpoint == "/api/import-results":
            try:
                _frame, _rows, _scores, _errors = load_initial_result(upload_path)
            except InputValidationError as exc:
                shutil.rmtree(job_dir, ignore_errors=True)
                self._send_json({"error": str(exc), "code": "invalid_initial_result", "columns": exc.columns, "missing": exc.missing}, 422)
                return
            settings = safe_settings({**raw_settings, "provider": "codex"})
            job = STORE.create(upload_path, original_name, settings, {}, job_id=job_id)
            EXECUTOR.submit(_import_initial_result, job_id, upload_path, original_name, settings)
            self._send_json({"job_id": job["id"], "status": "queued", "mode": "imported", "settings": settings}, 202)
            return
        try:
            _frame, _selected = read_input_file(upload_path, mapping)
        except InputValidationError as exc:
            shutil.rmtree(job_dir, ignore_errors=True)
            self._send_json({"error": str(exc), "code": "mapping_required" if exc.columns else "invalid_input", "columns": exc.columns, "missing": exc.missing}, 422)
            return
        settings = safe_settings(raw_settings)
        job = STORE.create(upload_path, original_name, settings, mapping, job_id=job_id)
        EXECUTOR.submit(_run_job, job)
        self._send_json({"job_id": job["id"], "status": "queued", "settings": settings}, 202)


def serve(host: str = "127.0.0.1", port: int = 8765) -> None:
    server = ThreadingHTTPServer((host, port), Handler)
    print("Park analysis demo: http://{}:{}/".format(host, port), flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        EXECUTOR.shutdown(wait=False)


if __name__ == "__main__":
    serve(os.environ.get("PARK_HOST", "127.0.0.1"), int(os.environ.get("PARK_PORT", "8765")))
