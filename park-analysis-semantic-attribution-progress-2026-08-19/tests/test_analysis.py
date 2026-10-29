import tempfile
import unittest
from pathlib import Path

import pandas as pd

from backend.analysis import (
    ASPECTS,
    CodexRunnerAdapter,
    MockAdapter,
    build_screening_rows,
    build_overview,
    create_attributions,
    enrich_frame,
    export_xlsx,
    group_evidence,
    load_initial_result,
    normalize_rows,
    read_input_file,
    safe_settings,
    split_batches,
    _load_ds_scopes,
)


ROOT = Path(__file__).resolve().parents[1]


class AnalysisPipelineTests(unittest.TestCase):
    def test_codex_prompts_render_and_parse_strict_json(self):
        class FakeCodex(CodexRunnerAdapter):
            def _run(self, prompt, settings):
                if "问题归因模型" in prompt:
                    return {"summary": "停车费用偏高", "causes": [{"text": "停车费用偏高，降低到访意愿", "source_rows": [2]}], "excluded": []}
                return {"rows": [{"source_row": 2, "scores": {"交通配套": 2}}]}

        adapter = FakeCodex("codex exec")
        rows = [{"source_row": 2, "park": "A", "comment": "停车太贵", "review_score": 2}]
        scores = adapter.analyze_batch(rows, safe_settings({"provider": "codex"}), "DS规则")
        self.assertEqual(scores[2]["交通配套"], 2)
        self.assertEqual(scores[2]["共享服务"], 0)
        summary = adapter.summarize_group("A", "交通配套", [{"source_row": 2, "score": 2, "comment": "停车太贵"}], safe_settings({"provider": "codex"}), "DS规则")
        self.assertEqual(summary["causes"][0]["source_rows"], [2])
        self.assertEqual(summary["excluded"], [])

    def test_ds_scopes_are_mapped_and_isolated(self):
        scopes = _load_ds_scopes(ROOT / "DS prompt.md")
        self.assertIn("便民生活", scopes["共享服务"])
        self.assertNotIn("信息导览", scopes["共享服务"])
        self.assertIn("场地设施租用", scopes["运动租赁"])
        self.assertIn("共享田园", scopes["共享农田"])

    def test_attribution_batches_then_merges_with_exact_rows(self):
        class FakeCodex(CodexRunnerAdapter):
            def __init__(self):
                super().__init__("codex exec")
                self.prompts = []

            def _run(self, prompt, settings):
                self.prompts.append(prompt)
                if "跨批合并候选问题" in prompt:
                    return {"summary": "停车费用和车位问题突出", "causes": [{"text": "停车费用偏高", "source_rows": [2]}, {"text": "停车位紧张", "source_rows": [3]}], "excluded": []}
                row = 2 if "停车收费贵" in prompt else 3
                return {"summary": "候选问题", "causes": [{"text": "停车问题影响体验", "source_rows": [row]}], "excluded": []}

        adapter = FakeCodex()
        negatives = [{"source_row": 2, "score": 2, "comment": "停车收费贵"}, {"source_row": 3, "score": 1, "comment": "停车位太少"}]
        result = adapter.summarize_group("A", "交通配套", negatives, safe_settings({"provider": "codex", "batch_size": 1}), "仅交通配套规则")
        self.assertEqual(len(adapter.prompts), 3)
        self.assertEqual([item["source_rows"] for item in result["causes"]], [[2], [3]])
        self.assertTrue(all("仅交通配套规则" in prompt for prompt in adapter.prompts))

    def test_existing_raw_workbook_schema(self):
        frame, selected = read_input_file(ROOT / "随机抽取_5个公园_10000条_原始.xlsx")
        self.assertEqual(frame.shape, (10000, 13))
        self.assertEqual(selected, {"park": "公园名", "comment": "评论"})
        rows, errors = normalize_rows(frame.head(5), selected)
        self.assertEqual(len(rows), 5)
        self.assertEqual(errors, [])

    def test_rating_is_optional(self):
        frame, selected = read_input_file(ROOT / "tests" / "fixtures" / "no_rating.csv")
        self.assertEqual(selected, {"park": "公园名", "comment": "评论"})
        rows, errors = normalize_rows(frame, selected)
        self.assertEqual(len(rows), 2)
        self.assertEqual(errors, [])

    def test_existing_initial_result_can_be_loaded_without_model(self):
        path = ROOT / "var" / "jobs" / "da5d25b5f25b49c78329ab34b919872a" / "初步分析结果.xlsx"
        if not path.exists():
            self.skipTest("current persisted job is not present")
        frame, rows, scores, errors = load_initial_result(path)
        self.assertEqual(len(frame), 1958)
        self.assertEqual(len(rows), 1958)
        self.assertEqual(len(scores), 1958)
        self.assertEqual(len(errors), 0)
        self.assertEqual(len(next(iter(scores.values()))), 22)

    def test_batching_uses_count_and_character_limits(self):
        rows = [{"source_row": index + 2, "park": "A", "comment": "评论" * 50} for index in range(5)]
        count_batches = split_batches(rows, 2, 100000)
        self.assertEqual([len(batch) for batch in count_batches], [2, 2, 1])
        char_batches = split_batches(rows, 20, 250)
        self.assertTrue(all(len(batch) == 1 for batch in char_batches))

    def test_zero_is_not_negative_and_single_negative_is_attributed(self):
        rows = [
            {"source_row": 2, "park": "A", "comment": "停车太贵"},
            {"source_row": 3, "park": "A", "comment": "停车方便"},
            {"source_row": 4, "park": "B", "comment": "普通评论"},
        ]
        scores = {
            2: {aspect: (2 if aspect == "交通配套" else 0) for aspect in ASPECTS},
            3: {aspect: (4 if aspect == "交通配套" else 0) for aspect in ASPECTS},
            4: {aspect: 0 for aspect in ASPECTS},
        }
        screening = build_screening_rows(rows, scores)
        self.assertEqual(len(screening), 1)
        self.assertEqual(screening[0]["方面评分"], 2)
        grouped = group_evidence(rows, scores)
        self.assertEqual(list(grouped), [("A", "交通配套")])
        self.assertEqual(len(grouped[("A", "交通配套")]), 1)
        attributions = create_attributions(rows, scores, MockAdapter(), safe_settings({"provider": "mock"}), "")
        self.assertEqual(len(attributions), 1)
        self.assertNotIn("矛盾标记", attributions[0])
        self.assertNotIn("正面证据", attributions[0])

    def test_model_failure_is_not_replaced_by_comment_text(self):
        class FailingAdapter(MockAdapter):
            def summarize_group(self, park, aspect, negatives, settings, ds_guide):
                raise RuntimeError("runner unavailable")

        rows = [{"source_row": 2, "park": "A", "comment": "停车太贵"}]
        scores = {2: {aspect: (2 if aspect == "交通配套" else 0) for aspect in ASPECTS}}
        result = create_attributions(rows, scores, FailingAdapter(), safe_settings({"provider": "mock", "retries": 0}), {"交通配套": "规则"})[0]
        self.assertEqual(result["归因状态"], "失败")
        self.assertEqual(result["问题卡片"], [])
        self.assertNotIn("停车太贵", result["问题摘要"])

    def test_out_of_scope_evidence_is_recorded_without_fake_cause(self):
        class ExcludingAdapter(MockAdapter):
            def summarize_group(self, park, aspect, negatives, settings, ds_guide):
                return {"summary": "无当前服务范围内有效证据", "causes": [], "excluded": [{"source_row": 2, "reason": "普通自行车属于交通代步"}]}

        rows = [{"source_row": 2, "park": "A", "comment": "自行车票太贵"}]
        scores = {2: {aspect: (2 if aspect == "运动租赁" else 0) for aspect in ASPECTS}}
        result = create_attributions(rows, scores, ExcludingAdapter(), safe_settings({"provider": "mock"}), {"运动租赁": "场地设施租用规则"})[0]
        self.assertEqual(result["归因状态"], "无有效范围证据")
        self.assertEqual(result["问题卡片"], [])
        self.assertEqual(result["范围排除"][0]["source_row"], 2)

    def test_overview_exposes_both_scatter_dimensions(self):
        rows = [{"source_row": 2, "park": "A", "comment": "停车太贵"}, {"source_row": 3, "park": "B", "comment": "停车方便"}]
        scores = {2: {aspect: (2 if aspect == "交通配套" else 0) for aspect in ASPECTS}, 3: {aspect: (4 if aspect == "交通配套" else 0) for aspect in ASPECTS}}
        overview = build_overview(rows, scores, [], safe_settings({"provider": "mock"}), [])
        self.assertEqual(len(overview["scatter"]["by_park"]), 44)
        self.assertEqual(len(overview["scatter"]["by_aspect"]), 44)
        point = next(item for item in overview["scatter"]["by_park"] if item["park"] == "A" and item["aspect"] == "交通配套")
        self.assertEqual(point["mention_count"], 1)
        self.assertEqual(point["average_score"], 2)

    def test_exports_keep_original_columns_and_create_three_files(self):
        frame = pd.DataFrame([{"公园名": "A", "评分": 2, "评论": "停车太贵", "发布时间": "2024-01-01"}])
        rows, errors = normalize_rows(frame, {"park": "公园名", "comment": "评论"})
        scores = {2: {aspect: (2 if aspect == "交通配套" else 0) for aspect in ASPECTS}}
        enriched = enrich_frame(frame, rows, scores)
        enriched.insert(len(enriched.columns) - len(ASPECTS), "分析状态", ["完成"])
        screening = build_screening_rows(rows, scores)
        attributions = create_attributions(rows, scores, MockAdapter(), safe_settings({"provider": "mock"}), "")
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)
            export_xlsx(target, enriched, screening, attributions, errors)
            self.assertTrue((target / "初步分析结果.xlsx").exists())
            self.assertTrue((target / "筛选明细.xlsx").exists())
            self.assertTrue((target / "问题归因总结.xlsx").exists())
            initial = pd.read_excel(target / "初步分析结果.xlsx")
            self.assertIn("发布时间", initial.columns)
            self.assertEqual(list(initial.columns[-22:]), ASPECTS)
            self.assertEqual(pd.read_excel(target / "筛选明细.xlsx").shape[0], 1)
            book = pd.ExcelFile(target / "问题归因总结.xlsx")
            self.assertEqual(book.sheet_names[:3], ["归因汇总", "证据明细", "范围排除"])
            self.assertIn("归因数量", pd.read_excel(target / "问题归因总结.xlsx", sheet_name="归因汇总").columns)
            self.assertIn("归因短句", pd.read_excel(target / "问题归因总结.xlsx", sheet_name="证据明细").columns)
            self.assertNotIn("正面评论数", pd.read_excel(target / "问题归因总结.xlsx", sheet_name="归因汇总").columns)
            self.assertNotIn("矛盾标记", pd.read_excel(target / "问题归因总结.xlsx", sheet_name="归因汇总").columns)


if __name__ == "__main__":
    unittest.main()
