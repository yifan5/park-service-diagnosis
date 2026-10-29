"""Data normalization, model adapters, aggregation, and export helpers.

The web server deliberately keeps this module independent from a web framework so
the demo can run with the Python standard library plus pandas/openpyxl.
"""

from __future__ import annotations

import json
import os
import re
import shlex
import subprocess
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd


ASPECTS = [
    "共享服务", "导览服务", "交通配套", "门票收费", "休闲饮食", "特色正餐",
    "静态展览", "动态演艺", "参与体验", "游乐设施", "游览观光", "社交互动",
    "运动租赁", "赛事竞技", "基础科普", "课程培训", "动手实践", "康体养生",
    "自然体验", "共享农田", "基础零售", "主题零售",
]

DS_ASPECT_NAMES = {
    "共享服务": "便民生活",
    "导览服务": "信息导览",
    "交通配套": "交通配套",
    "门票收费": "门票收费",
    "休闲饮食": "休闲饮食",
    "特色正餐": "特色正餐",
    "静态展览": "静态展览",
    "动态演艺": "动态演艺",
    "参与体验": "参与体验",
    "游乐设施": "游乐设施",
    "游览观光": "游览观光",
    "社交互动": "社交互动",
    "运动租赁": "场地设施租用",
    "赛事竞技": "赛事竞技",
    "基础科普": "基础科普",
    "课程培训": "课程培训",
    "动手实践": "动手实践",
    "康体养生": "康体养生",
    "自然体验": "自然体验",
    "共享农田": "共享田园",
    "基础零售": "基础零售",
    "主题零售": "主题零售",
}

PARK_ALIASES = ["公园名", "公园", "park", "park_name", "name"]
COMMENT_ALIASES = ["评论", "评论内容", "comment", "review", "text"]

DEFAULT_SETTINGS: Dict[str, Any] = {
    "provider": "codex",
    "model": "",
    "batch_size": 20,
    "max_chars": 24000,
    "temperature": 0.1,
    "max_output_tokens": 4096,
    "concurrency": 1,
    "retries": 2,
    "timeout_seconds": 180,
}


class InputValidationError(ValueError):
    def __init__(self, message: str, columns: Optional[Sequence[str]] = None,
                 missing: Optional[Sequence[str]] = None):
        super().__init__(message)
        self.columns = list(columns or [])
        self.missing = list(missing or [])


class ModelError(RuntimeError):
    pass


def _clean_column(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").strip())


def _choose_column(columns: Sequence[str], aliases: Sequence[str], requested: Optional[str]) -> Optional[str]:
    if requested and requested in columns:
        return requested
    normalized = {_clean_column(column): column for column in columns}
    for alias in aliases:
        if _clean_column(alias) in normalized:
            return normalized[_clean_column(alias)]
    return None


def _read_csv(path: Path) -> pd.DataFrame:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return pd.read_csv(path, encoding=encoding)
        except (UnicodeDecodeError, pd.errors.ParserError) as exc:
            last_error = exc
    raise InputValidationError("CSV 无法解析，请检查编码或分隔符。") from last_error


def read_input_file(path: Path, mapping: Optional[Dict[str, str]] = None) -> Tuple[pd.DataFrame, Dict[str, str]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        frame = _read_csv(path)
    elif suffix in (".xlsx", ".xls"):
        try:
            frame = pd.read_excel(path)
        except ImportError as exc:
            raise InputValidationError("读取 .xls 需要安装 xlrd；.xlsx 文件可直接使用。") from exc
        except Exception as exc:
            raise InputValidationError("Excel 文件无法读取，请确认文件未损坏。") from exc
    else:
        raise InputValidationError("仅支持 .xlsx、.xls 和 .csv 文件。")

    frame.columns = [str(column).strip() for column in frame.columns]
    columns = list(frame.columns)
    mapping = mapping or {}
    selected = {
        "park": _choose_column(columns, PARK_ALIASES, mapping.get("park")),
        "comment": _choose_column(columns, COMMENT_ALIASES, mapping.get("comment")),
    }
    missing = [name for name, value in selected.items() if not value]
    if missing:
        labels = {"park": "公园", "comment": "评论"}
        missing_text = "、".join(labels[name] for name in missing)
        raise InputValidationError(
            "无法识别必需列：{}。可从上传文件列名中手动映射。".format(missing_text),
            columns=columns,
            missing=missing,
        )
    return frame, selected  # type: ignore[return-value]


def normalize_rows(frame: pd.DataFrame, selected: Dict[str, str]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Return valid analysis rows and row-level validation errors.

    The original frame is never reduced here; the caller uses it for pass-through
    exports. Invalid rows are reported and skipped from model calls.
    """
    valid: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []
    for index, row in frame.iterrows():
        source_row = int(index) + 2  # Excel physical row: header is row 1.
        park = "" if pd.isna(row[selected["park"]]) else str(row[selected["park"]]).strip()
        comment = "" if pd.isna(row[selected["comment"]]) else str(row[selected["comment"]]).strip()
        row_errors: List[str] = []
        if not park:
            row_errors.append("公园为空")
        if not comment:
            row_errors.append("评论为空")
        if row_errors:
            errors.append({"原始行号": source_row, "错误": "；".join(row_errors)})
            continue
        valid.append({
            "source_row": source_row,
            "park": park,
            "comment": comment,
        })
    return valid, errors


def _load_ds_scopes(prompt_path: Path) -> Dict[str, str]:
    """Parse one complete DS prompt section per result-table aspect."""
    if not prompt_path.exists():
        return {aspect: "未找到 DS prompt 对应段落。" for aspect in ASPECTS}
    try:
        lines = prompt_path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return {aspect: "无法读取 DS prompt 对应段落。" for aspect in ASPECTS}
    headings = {name for name in DS_ASPECT_NAMES.values()}
    sections: Dict[str, List[str]] = {}
    current: Optional[str] = None
    for line in lines:
        heading = line.strip().strip("#").strip()
        if heading in headings and line.strip() == heading:
            current = heading
            sections.setdefault(current, []).append(line)
        elif current:
            sections[current].append(line)
    result: Dict[str, str] = {}
    for aspect, ds_name in DS_ASPECT_NAMES.items():
        text = "\n".join(sections.get(ds_name, []))
        result[aspect] = re.sub(r"\s+", " ", text).strip() or "未找到 DS prompt 对应段落。"
    return result


def _compact_ds_guide(prompt_path: Path) -> str:
    if not prompt_path.exists():
        return "DS prompt 文件未找到；只使用下方方面标签和评分规则。"
    try:
        raw = prompt_path.read_text(encoding="utf-8")
    except OSError:
        return "DS prompt 文件无法读取；只使用下方方面标签和评分规则。"
    # Keep the source prompt influential without sending the entire 80KB document
    # on every batch. Each aspect gets its first compact scope paragraph.
    chunks: List[str] = []
    scopes = _load_ds_scopes(prompt_path)
    for aspect in ASPECTS:
        chunks.append("{}: {}".format(aspect, scopes.get(aspect, "" )[:900]))
    return "\n".join(chunks)[:18000]


def _json_text(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _extract_json(text: str) -> Any:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        candidates = re.findall(r"(\{.*\}|\[.*\])", text, flags=re.S)
        for candidate in reversed(candidates):
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue
    raise ModelError("模型没有返回可解析的 JSON。")


def _score(value: Any) -> int:
    try:
        number = int(round(float(value)))
    except (TypeError, ValueError):
        return 0
    return max(0, min(5, number))


def _normalize_batch_output(payload: Any, rows: Sequence[Dict[str, Any]]) -> Dict[int, Dict[str, int]]:
    if isinstance(payload, dict):
        output_rows = payload.get("rows", payload.get("results", []))
    else:
        output_rows = payload
    if not isinstance(output_rows, list):
        raise ModelError("模型输出缺少 rows 数组。")
    by_row: Dict[int, Dict[str, int]] = {}
    for item in output_rows:
        if not isinstance(item, dict):
            continue
        raw_row = item.get("source_row", item.get("row", item.get("原始行号")))
        try:
            source_row = int(raw_row)
        except (TypeError, ValueError):
            continue
        scores = item.get("scores", item.get("aspects", item))
        if not isinstance(scores, dict):
            scores = {}
        by_row[source_row] = {aspect: _score(scores.get(aspect, 0)) for aspect in ASPECTS}
    missing = [row["source_row"] for row in rows if row["source_row"] not in by_row]
    if missing:
        raise ModelError("模型输出缺少 {} 条记录。".format(len(missing)))
    return by_row


class BaseAdapter:
    name = "base"

    def analyze_batch(self, rows: Sequence[Dict[str, Any]], settings: Dict[str, Any], ds_guide: str) -> Dict[int, Dict[str, int]]:
        raise NotImplementedError

    def summarize_group(self, park: str, aspect: str, negatives: Sequence[Dict[str, Any]], settings: Dict[str, Any], ds_guide: str) -> Dict[str, Any]:
        raise NotImplementedError


class MockAdapter(BaseAdapter):
    name = "mock"

    KEYWORDS = {
        "交通配套": "停车 车位 停车场 电瓶车 接驳 充电 缆车 代步",
        "门票收费": "门票 票价 门票费 联票 年卡 收费 免费",
        "休闲饮食": "咖啡 奶茶 饮料 小吃 冰淇淋 面包 饮品",
        "特色正餐": "餐厅 吃饭 正餐 火锅 烧烤 菜馆 聚餐",
        "基础零售": "小卖部 商店 自动售货 水 零食 便利店",
        "主题零售": "文创 纪念品 玩具 风筝 特许 商店",
        "游乐设施": "游乐 滑梯 摩天轮 碰碰船 儿童乐园",
        "自然体验": "花 草 树 湖 景色 风景 自然 动物",
        "基础科普": "科普 博物馆 展览 标识 介绍 知识",
        "导览服务": "导览 讲解员 讲解器 导游",
        "社交互动": "人多 排队 拥挤 朋友 家人 亲子",
        "康体养生": "散步 健身 养生 按摩 康体",
        "运动租赁": "自行车 滑板车 球 租赁 运动",
        "共享服务": "充电宝 储物柜 婴儿车 轮椅 WiFi",
        "静态展览": "展览 展馆 博物馆 陈列",
        "动态演艺": "演出 表演 演艺 音乐 剧场",
        "参与体验": "体验 手工 互动 活动 采摘",
        "游览观光": "观光 游船 观景 玻璃栈道 热气球",
        "赛事竞技": "比赛 赛事 竞技 球赛",
        "课程培训": "课程 培训 研学 讲座 夏令营",
        "动手实践": "种植 手作 DIY 实践",
        "共享农田": "农田 菜园 认领 农场",
    }
    NEGATIVE = ["不满意", "很差", "太贵", "贵了", "失望", "糟糕", "坑人", "难找", "找不到", "排队很久", "脏乱", "损坏", "态度不好", "体验差"]
    POSITIVE = ["很好", "满意", "推荐", "方便", "划算", "便宜", "精彩", "喜欢", "不错", "舒服", "丰富"]

    def analyze_batch(self, rows: Sequence[Dict[str, Any]], settings: Dict[str, Any], ds_guide: str) -> Dict[int, Dict[str, int]]:
        result: Dict[int, Dict[str, int]] = {}
        for row in rows:
            text = row["comment"]
            result[row["source_row"]] = {}
            for aspect in ASPECTS:
                words = self.KEYWORDS.get(aspect, "").split()
                mentioned = any(word in text for word in words)
                if not mentioned:
                    value = 0
                elif any(word in text for word in self.NEGATIVE):
                    value = 2
                elif any(word in text for word in self.POSITIVE):
                    value = 4
                else:
                    value = 3
                result[row["source_row"]][aspect] = value
        return result

    def summarize_group(self, park: str, aspect: str, negatives: Sequence[Dict[str, Any]], settings: Dict[str, Any], ds_guide: str) -> Dict[str, Any]:
        groups: List[Dict[str, Any]] = []
        for evidence in negatives:
            text = re.sub(r"\s+", " ", str(evidence.get("comment", ""))).strip()
            text = re.sub(r"^(感觉|觉得|认为|整体|就是)\s*", "", text)
            text = re.split(r"[。！？!?；;]", text)[0].strip()[:48].rstrip("，,、 ") or "该方面体验不佳"
            chars = set(re.sub(r"[^0-9A-Za-z一-龥]", "", text))
            match = next((item for item in groups if chars and len(chars & item["chars"]) / max(1, len(chars | item["chars"])) >= 0.45), None)
            if match:
                match["source_rows"].append(int(evidence["source_row"]))
            elif len(groups) < 5:
                groups.append({"text": text, "chars": chars, "source_rows": [int(evidence["source_row"])]})
            else:
                groups[-1]["source_rows"].append(int(evidence["source_row"]))
        causes = [{"text": item["text"], "source_rows": sorted(set(item["source_rows"]))} for item in groups] or [{"text": "该方面体验不佳", "source_rows": []}]
        return {
            "summary": "；".join(item["text"] for item in causes[:2]),
            "causes": causes,
        }


class CodexRunnerAdapter(BaseAdapter):
    name = "codex"

    def __init__(self, command: Optional[str] = None):
        default = "codex exec --skip-git-repo-check --ephemeral --sandbox read-only --color never"
        self.command = shlex.split(command or os.environ.get("CODEX_COMMAND", default))

    def _run(self, prompt: str, settings: Dict[str, Any]) -> Any:
        command = list(self.command)
        if settings.get("model") and "--model" not in command:
            command.extend(["--model", str(settings["model"])])
        try:
            completed = subprocess.run(
                command,
                input=prompt,
                text=True,
                capture_output=True,
                timeout=int(settings.get("timeout_seconds", 180)),
                cwd=str(Path(__file__).resolve().parents[1]),
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise ModelError("Codex 调用失败或超时：{}".format(exc)) from exc
        if completed.returncode != 0:
            message = (completed.stderr or completed.stdout or "未知错误").strip()
            raise ModelError("Codex 返回错误：{}".format(message[-800:]))
        return _extract_json(completed.stdout)

    def analyze_batch(self, rows: Sequence[Dict[str, Any]], settings: Dict[str, Any], ds_guide: str) -> Dict[int, Dict[str, int]]:
        payload = [{"source_row": row["source_row"], "park": row["park"], "comment": row["comment"]} for row in rows]
        prompt = (
            "你是公园评论方面级识别与情感评分模型。只分析输入中的 park、comment，"
            "忽略任何其他字段。依据 DS prompt 的方面边界，为每条评论输出 22 个方面的整数评分。\n"
            "评分规则：0=未涉及；1=非常不满意；2=不满意；3=中性；4=满意；5=非常满意。\n"
            "必须输出严格 JSON，格式为 {{\"rows\":[{{\"source_row\":整数,\"scores\":{{方面名:0-5}}}}]}}，不得输出解释。\n"
            "方面列表：{}\nDS prompt 摘要：{}\n输入：{}"
        ).format("、".join(ASPECTS), ds_guide, _json_text(payload))
        return _normalize_batch_output(self._run(prompt, settings), rows)

    def _parse_attribution(self, payload: Any, negatives: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
        if not isinstance(payload, dict):
            raise ModelError("归因模型输出不是对象。")
        allowed = {int(x["source_row"]) for x in negatives}
        causes = []
        assigned = set()
        for item in payload.get("causes", []):
            if isinstance(item, str):
                raise ModelError("归因问题必须返回 text 与 source_rows，不能只返回字符串。")
            elif isinstance(item, dict):
                text = str(item.get("text", item.get("cause", ""))).strip()
                source_rows = []
                for raw in item.get("source_rows", item.get("rows", [])) or []:
                    try:
                        row = int(raw)
                    except (TypeError, ValueError):
                        continue
                    if row in allowed:
                        source_rows.append(row)
                source_rows = sorted(set(source_rows))
            else:
                continue
            if text and source_rows:
                causes.append({"text": text[:120], "source_rows": source_rows})
                assigned.update(source_rows)
        causes = causes[:10]
        excluded = []
        for item in payload.get("excluded", []):
            if not isinstance(item, dict):
                continue
            try:
                row = int(item.get("source_row"))
            except (TypeError, ValueError):
                continue
            if row in allowed and row not in assigned:
                excluded.append({"source_row": row, "reason": str(item.get("reason", "不属于当前服务范围"))[:160]})
        excluded_rows = {item["source_row"] for item in excluded}
        for row in sorted(allowed - assigned - excluded_rows):
            excluded.append({"source_row": row, "reason": "模型未将该评论归入当前服务范围内问题"})
        summary = str(payload.get("summary", "")).strip()[:180]
        if not causes and not excluded:
            raise ModelError("归因模型未返回问题或范围排除记录。")
        return {"summary": summary or (causes[0]["text"] if causes else "无当前服务范围内有效证据"), "causes": causes, "excluded": excluded}

    def _summarize_once(self, park: str, aspect: str, evidence: Sequence[Dict[str, Any]], scope: str, settings: Dict[str, Any], merge_input: Optional[Sequence[Dict[str, Any]]] = None) -> Dict[str, Any]:
        payload = {
            "park": park,
            "aspect": aspect,
            "evidence": [{"source_row": x["source_row"], "score": x["score"], "comment": x["comment"][:1400]} for x in evidence],
        }
        if merge_input is not None:
            payload["batch_candidates"] = list(merge_input)
        mode = "跨批合并候选问题" if merge_input is not None else "提炼本批评论中的独立问题"
        prompt = (
            "你是公园评论问题归因模型。当前方面是【{aspect}】，只能使用下方 DS prompt 对应服务范围。\n"
            "DS 服务规则：{scope}\n"
            "任务：{mode}。只接收评分 1—2 的评论；排除与当前服务无关的内容并记录行号和原因。"
            "问题短句必须是语义提炼和概括，不能照搬原评论，需指出具体对象、问题和影响。"
            "相近语义必须合并；一条评论可以支持多个确实相关的问题。最多输出 10 个问题。"
            "每个 causes 项必须有 text 和 source_rows，source_rows 只能使用输入行号且至少一个。"
            "严格输出 JSON：{{\"summary\":字符串,\"causes\":[{{\"text\":字符串,\"source_rows\":[整数]}}],\"excluded\":[{{\"source_row\":整数,\"reason\":字符串}}]}}，不得输出其他内容。\n输入：{payload}"
        ).format(aspect=aspect, scope=scope, mode=mode, payload=_json_text(payload))
        return self._parse_attribution(self._run(prompt, settings), evidence)

    def summarize_group(self, park: str, aspect: str, negatives: Sequence[Dict[str, Any]], settings: Dict[str, Any], ds_guide: str) -> Dict[str, Any]:
        batches = split_batches(negatives, int(settings.get("batch_size", 20)), int(settings.get("max_chars", 24000)))
        candidates: List[Dict[str, Any]] = []
        excluded: List[Dict[str, Any]] = []
        for batch in batches:
            result = self._summarize_once(park, aspect, batch, ds_guide, settings)
            candidates.extend(result.get("causes", []))
            excluded.extend(result.get("excluded", []))
        if len(batches) > 1 or len(candidates) > 10:
            merged = self._summarize_once(park, aspect, negatives, ds_guide, settings, candidates)
            merged["excluded"] = merged.get("excluded", [])
            return merged
        result = {"summary": "；".join(x["text"] for x in candidates[:2]) or "无当前服务范围内有效证据", "causes": candidates[:10], "excluded": excluded}
        return self._parse_attribution(result, negatives)


def make_adapter(provider: str) -> BaseAdapter:
    if provider == "mock":
        return MockAdapter()
    return CodexRunnerAdapter()


def split_batches(rows: Sequence[Dict[str, Any]], batch_size: int, max_chars: int) -> List[List[Dict[str, Any]]]:
    batches: List[List[Dict[str, Any]]] = []
    current: List[Dict[str, Any]] = []
    current_chars = 0
    for row in rows:
        row_chars = len(_json_text(row))
        if current and (len(current) >= batch_size or current_chars + row_chars > max_chars):
            batches.append(current)
            current = []
            current_chars = 0
        current.append(row)
        current_chars += row_chars
    if current:
        batches.append(current)
    return batches


def _sentiment(score: int) -> str:
    if score == 0:
        return "未涉及"
    if score <= 2:
        return "负面"
    if score == 3:
        return "中性"
    return "正面"


def enrich_frame(frame: pd.DataFrame, rows: Sequence[Dict[str, Any]], scores: Dict[int, Dict[str, int]]) -> pd.DataFrame:
    output = frame.copy()
    output.insert(0, "原始行号", [int(index) + 2 for index in frame.index])
    by_row = {row["source_row"]: scores.get(row["source_row"], {aspect: 0 for aspect in ASPECTS}) for row in rows}
    for aspect in ASPECTS:
        output[aspect] = [by_row.get(int(index) + 2, {}).get(aspect, 0) for index in frame.index]
    return output


def build_screening_rows(rows: Sequence[Dict[str, Any]], scores: Dict[int, Dict[str, int]]) -> List[Dict[str, Any]]:
    output: List[Dict[str, Any]] = []
    for row in rows:
        for aspect in ASPECTS:
            value = int(scores.get(row["source_row"], {}).get(aspect, 0))
            if 1 <= value <= 2:
                output.append({
                    "原始行号": row["source_row"],
                    "公园名": row["park"],
                    "评论": row["comment"],
                    "方面": aspect,
                    "方面评分": value,
                    "情感": _sentiment(value),
                })
    return output


def group_evidence(rows: Sequence[Dict[str, Any]], scores: Dict[int, Dict[str, int]]) -> Dict[Tuple[str, str], List[Dict[str, Any]]]:
    grouped: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
    by_row = {row["source_row"]: row for row in rows}
    for source_row, aspect_scores in scores.items():
        row = by_row.get(source_row)
        if not row:
            continue
        for aspect, value in aspect_scores.items():
            if 1 <= value <= 2:
                grouped[(row["park"], aspect)].append({"source_row": source_row, "score": value, "comment": row["comment"]})
    return dict(grouped)


def build_overview(rows: Sequence[Dict[str, Any]], scores: Dict[int, Dict[str, int]], attributions: Sequence[Dict[str, Any]], settings: Dict[str, Any], errors: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    parks = sorted({row["park"] for row in rows})
    rows_by_park: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        rows_by_park[row["park"]].append(row)
    def metrics(items: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
        values = [scores.get(row["source_row"], {}) for row in items]
        result: Dict[str, Any] = {}
        for aspect in ASPECTS:
            seen = [int(value.get(aspect, 0)) for value in values if int(value.get(aspect, 0)) > 0]
            negative = [value for value in seen if value <= 2]
            result[aspect] = {
                "mention_count": len(seen),
                "mention_rate": len(seen) / len(items) if items else 0,
                "average_score": sum(seen) / len(seen) if seen else 0,
                "negative_count": len(negative),
                "positive_count": sum(1 for value in seen if value >= 4),
                "neutral_count": sum(1 for value in seen if value == 3),
            }
        return result
    park_data = [{"name": park, "total": len(rows_by_park[park]), "values": metrics(rows_by_park[park])} for park in parks]
    return {
        "aspects": ASPECTS,
        "park_count": len(parks),
        "total_rows": len(rows),
        "attribution_count": len(attributions),
        "invalid_row_count": sum(1 for error in errors if "批次" not in error),
        "parks": park_data,
        "overall": metrics(rows),
        "scatter": {
            "by_park": [{"park": park["name"], "aspect": aspect, **metric} for park in park_data for aspect, metric in park["values"].items()],
            "by_aspect": [{"park": park["name"], "aspect": aspect, **metric} for aspect in ASPECTS for park in park_data for aspect_name, metric in park["values"].items() if aspect_name == aspect],
        },
        "settings": settings,
        "provider": settings.get("provider"),
        "score_rules": {"0": "未涉及", "1-2": "负面", "3": "中性", "4-5": "正面"},
    }


def create_attributions(rows: Sequence[Dict[str, Any]], scores: Dict[int, Dict[str, int]], adapter: BaseAdapter, settings: Dict[str, Any], ds_guide: Any, progress: Optional[Callable[[int, int], None]] = None, only: Optional[Sequence[Tuple[str, str]]] = None, completed_item: Optional[Callable[[Dict[str, Any]], None]] = None, stage_progress: Optional[Callable[[str, int, int, int], None]] = None) -> List[Dict[str, Any]]:
    grouped = group_evidence(rows, scores)
    if only is not None:
        allowed_groups = set(only)
        grouped = {key: value for key, value in grouped.items() if key in allowed_groups}
    total = len(grouped)
    if stage_progress:
        stage_progress("语义梳理", 1, 0, total)

    def process_group(key: Tuple[str, str], negatives: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
        park, aspect = key
        summary = None
        last_error: Optional[Exception] = None
        for attempt in range(int(settings.get("retries", 0)) + 1):
            try:
                scope = ds_guide.get(aspect, "未找到 DS prompt 对应服务规则。") if isinstance(ds_guide, dict) else str(ds_guide)
                summary = adapter.summarize_group(park, aspect, negatives, settings, scope)
                break
            except Exception as exc:
                last_error = exc
                if attempt < int(settings.get("retries", 0)):
                    time.sleep(min(2 ** attempt, 4))
        if summary is None:
            return {
                "公园名": park, "方面": aspect, "负面评论数": len(negatives),
                "归因状态": "失败", "归因错误": str(last_error or "模型调用失败"),
                "问题摘要": "归因生成失败，请重试", "问题列表": "", "问题卡片": [],
                "负面证据行号": ",".join(str(x["source_row"]) for x in negatives),
                "负面证据": negatives, "范围排除": [],
            }
        normalized_causes = []
        for cause in summary.get("causes", []):
            if isinstance(cause, str):
                continue
            elif isinstance(cause, dict) and str(cause.get("text", "")).strip():
                allowed = {int(x["source_row"]) for x in negatives}
                source_rows_set = set()
                for raw in cause.get("source_rows", []):
                    try:
                        source_row = int(raw)
                    except (TypeError, ValueError):
                        continue
                    if source_row in allowed:
                        source_rows_set.add(source_row)
                source_rows = sorted(source_rows_set)
                if source_rows:
                    normalized_causes.append({"text": str(cause["text"]).strip()[:120], "source_rows": source_rows})
        excluded = []
        for item in summary.get("excluded", []):
            if isinstance(item, dict):
                try:
                    row = int(item.get("source_row"))
                except (TypeError, ValueError):
                    continue
                excluded.append({"source_row": row, "reason": str(item.get("reason", "不属于当前服务范围"))[:160]})
        assigned = {row for cause in normalized_causes for row in cause["source_rows"]}
        excluded_rows = {item["source_row"] for item in excluded}
        for row in negatives:
            if row["source_row"] not in assigned and row["source_row"] not in excluded_rows:
                excluded.append({"source_row": row["source_row"], "reason": "模型未将该评论归入当前服务范围内问题"})
        return {
            "公园名": park,
            "方面": aspect,
            "负面评论数": len(negatives),
            "归因状态": "完成" if normalized_causes else "无有效范围证据",
            "问题摘要": summary.get("summary", ""),
            "问题列表": "；".join(item["text"] for item in normalized_causes),
            "问题卡片": normalized_causes,
            "负面证据行号": ",".join(str(x["source_row"]) for x in negatives),
            "负面证据": negatives,
            "范围排除": excluded,
        }

    items = sorted(grouped.items())
    if not items:
        return []
    results: List[Dict[str, Any]] = []
    workers = max(1, int(settings.get("concurrency", 1)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(process_group, key, negatives) for key, negatives in items]
        for completed, future in enumerate(as_completed(futures), 1):
            item = future.result()
            results.append(item)
            if completed_item:
                completed_item(item)
            if progress:
                progress(completed, total)
            if stage_progress:
                stage_progress("语义梳理", 1, completed, total)
    if stage_progress:
        stage_progress("范围筛选", 2, total, total)
        stage_progress("相似合并", 3, total, total)
        stage_progress("结果校验", 4, total, total)
    return sorted(results, key=lambda item: (item["公园名"], item["方面"]))


def export_xlsx(path: Path, frame: pd.DataFrame, screening: Sequence[Dict[str, Any]], attributions: Sequence[Dict[str, Any]], errors: Sequence[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(path / "初步分析结果.xlsx", index=False, engine="openpyxl")
    screening_columns = ["原始行号", "公园名", "评论", "方面", "方面评分", "情感"]
    pd.DataFrame(list(screening), columns=screening_columns).to_excel(path / "筛选明细.xlsx", index=False, engine="openpyxl")
    summary_rows = []
    evidence_rows = []
    excluded_rows = []
    for item in attributions:
        causes = item.get("问题卡片", [])
        summary_rows.append({
            "公园名": item.get("公园名", ""), "方面": item.get("方面", ""),
            "负面评论数": item.get("负面评论数", 0), "归因数量": len(causes),
            "归因状态": item.get("归因状态", "完成"), "问题概要": item.get("问题摘要", ""),
            "问题列表": "；".join(str(x.get("text", "")) for x in causes),
            "支持行号": "；".join("、".join(str(row) for row in x.get("source_rows", [])) for x in causes),
            "排除数量": len(item.get("范围排除", [])),
            "排除行号": "、".join(str(x.get("source_row", "")) for x in item.get("范围排除", [])),
            "归因错误": item.get("归因错误", ""),
        })
        evidence_by_row = {int(evidence["source_row"]): evidence for evidence in item.get("负面证据", [])}
        for cause in causes:
            for source_row in cause.get("source_rows", []):
                evidence = evidence_by_row.get(int(source_row))
                if evidence:
                    evidence_rows.append({
                        "公园名": item["公园名"], "方面": item["方面"], "归因短句": cause.get("text", ""),
                        "原始行号": evidence["source_row"], "方面评分": evidence["score"], "评论": evidence["comment"],
                    })
        for excluded in item.get("范围排除", []):
            evidence = evidence_by_row.get(int(excluded.get("source_row", 0)))
            excluded_rows.append({
                "公园名": item.get("公园名", ""), "方面": item.get("方面", ""),
                "原始行号": excluded.get("source_row", ""), "评论": evidence.get("comment", "") if evidence else "",
                "排除原因": excluded.get("reason", ""),
            })
    with pd.ExcelWriter(path / "问题归因总结.xlsx", engine="openpyxl") as writer:
        summary_columns = ["公园名", "方面", "负面评论数", "归因数量", "归因状态", "问题概要", "问题列表", "支持行号", "排除数量", "排除行号", "归因错误"]
        evidence_columns = ["公园名", "方面", "归因短句", "原始行号", "方面评分", "评论"]
        excluded_columns = ["公园名", "方面", "原始行号", "评论", "排除原因"]
        pd.DataFrame(summary_rows, columns=summary_columns).to_excel(writer, sheet_name="归因汇总", index=False)
        pd.DataFrame(evidence_rows, columns=evidence_columns).to_excel(writer, sheet_name="证据明细", index=False)
        pd.DataFrame(excluded_rows, columns=excluded_columns).to_excel(writer, sheet_name="范围排除", index=False)
        if errors:
            pd.DataFrame(list(errors)).to_excel(writer, sheet_name="错误明细", index=False)


def rebuild_existing_exports(path: Path) -> Dict[str, int]:
    """Rewrite an existing result directory to the global negative-only format.

    This consumes workbooks already produced by an earlier run and never calls a
    model. Existing aspect scores and summaries are retained; only positive and
    conflict fields are removed from the attribution export.
    """
    initial_path = path / "初步分析结果.xlsx"
    screening_path = path / "筛选明细.xlsx"
    summary_path = path / "问题归因总结.xlsx"
    frame = pd.read_excel(initial_path)
    screening_frame = pd.read_excel(screening_path) if screening_path.exists() else pd.DataFrame()
    screening = screening_frame.to_dict("records")
    summary = pd.read_excel(summary_path, sheet_name="归因汇总") if summary_path.exists() else pd.DataFrame()
    sheets = pd.ExcelFile(summary_path).sheet_names if summary_path.exists() else []
    evidence = pd.read_excel(summary_path, sheet_name="证据明细") if "证据明细" in sheets else pd.DataFrame()
    attributions: List[Dict[str, Any]] = []
    for _, item in summary.iterrows():
        park = str(item.get("公园名", "")); aspect = str(item.get("方面", ""))
        negatives: List[Dict[str, Any]] = []
        if not evidence.empty:
            subset = evidence[(evidence.get("公园名", pd.Series(dtype=str)).astype(str) == park) & (evidence.get("方面", pd.Series(dtype=str)).astype(str) == aspect)]
            if "证据类型" in subset.columns:
                subset = subset[subset["证据类型"].astype(str) == "负面"]
            for _, ev in subset.iterrows():
                negatives.append({"source_row": int(ev.get("原始行号", 0)), "score": int(ev.get("方面评分", 0)), "comment": str(ev.get("评论", ""))})
        if not negatives:
            continue
        attributions.append({
            "公园名": park, "方面": aspect, "负面评论数": len(negatives),
            "问题摘要": str(item.get("问题摘要", "")), "问题列表": str(item.get("问题列表", "")),
            "负面证据行号": ",".join(str(x["source_row"]) for x in negatives), "负面证据": negatives,
        })
    export_xlsx(path, frame, screening, attributions, [])
    return {"rows": len(frame), "screening": len(screening), "attributions": len(attributions)}


def load_initial_result(path: Path) -> Tuple[pd.DataFrame, List[Dict[str, Any]], Dict[int, Dict[str, int]], List[Dict[str, Any]]]:
    """Load an already-scored workbook without invoking a model."""
    suffix = path.suffix.lower()
    frame = _read_csv(path) if suffix == ".csv" else pd.read_excel(path)
    frame.columns = [str(column).strip() for column in frame.columns]
    park_col = _choose_column(list(frame.columns), PARK_ALIASES, None)
    comment_col = _choose_column(list(frame.columns), COMMENT_ALIASES, None)
    if not park_col or not comment_col:
        raise InputValidationError("初步结果必须包含公园和评论列。", columns=list(frame.columns), missing=[x for x, value in (("park", park_col), ("comment", comment_col)) if not value])
    missing_aspects = [aspect for aspect in ASPECTS if aspect not in frame.columns]
    if missing_aspects:
        raise InputValidationError("初步结果缺少 22 类方面评分列：{}。".format("、".join(missing_aspects[:8])), columns=list(frame.columns), missing=missing_aspects)
    rows: List[Dict[str, Any]] = []
    scores: Dict[int, Dict[str, int]] = {}
    errors: List[Dict[str, Any]] = []
    row_col = "原始行号" if "原始行号" in frame.columns else None
    for index, record in frame.iterrows():
        source_row = int(record[row_col]) if row_col and pd.notna(record[row_col]) else index + 2
        park = "" if pd.isna(record[park_col]) else str(record[park_col]).strip()
        comment = "" if pd.isna(record[comment_col]) else str(record[comment_col]).strip()
        if not park or not comment:
            errors.append({"原始行号": source_row, "错误": "公园或评论为空，已忽略"})
            continue
        rows.append({"source_row": source_row, "park": park, "comment": comment})
        scores[source_row] = {aspect: _score(record[aspect]) for aspect in ASPECTS}
    return frame, rows, scores, errors


def safe_settings(raw: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    settings = dict(DEFAULT_SETTINGS)
    if raw:
        settings.update({key: value for key, value in raw.items() if key in settings})
    for key, low, high in (("batch_size", 1, 500), ("max_chars", 2000, 100000), ("max_output_tokens", 512, 16000), ("concurrency", 1, 4), ("retries", 0, 5), ("timeout_seconds", 20, 900)):
        try:
            settings[key] = max(low, min(high, int(settings[key])))
        except (TypeError, ValueError):
            settings[key] = DEFAULT_SETTINGS[key]
    try:
        settings["temperature"] = max(0.0, min(2.0, float(settings["temperature"])))
    except (TypeError, ValueError):
        settings["temperature"] = DEFAULT_SETTINGS["temperature"]
    settings["provider"] = str(settings.get("provider") or "codex").lower()
    if settings["provider"] not in ("codex", "mock"):
        settings["provider"] = "codex"
    return settings
