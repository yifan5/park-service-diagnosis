# 当前版本

- 版本标识：`semantic-attribution-progress-2026-08-19`
- 归因方式：DS 服务范围约束 + 分批语义提炼 + 跨批相似问题合并
- 归因上限：每个公园 × 方面最多 10 个问题短句
- 初步结果模式：概况先展示，后台强制 Codex 归因
- 进度阶段：语义梳理、范围筛选、相似合并、结果校验
- 散点颜色：低于 3 分红橙色，3 分橘黄色，高分逐步过渡到蓝色
- 评分模型：不重新处理已有初步结果中的 22 类评分

源码包：`releases/park-analysis-semantic-attribution-progress-2026-08-19.tar.gz`
