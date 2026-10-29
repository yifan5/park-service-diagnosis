# 模型 API 替换指南

## 当前入口

模型适配层位于：

```text
backend/analysis.py
```

主要接口：

- `BaseAdapter.analyze_batch(...)`：原始评论模式的 22 类方面评分。
- `BaseAdapter.summarize_group(...)`：负面归因，接收一个公园 × 方面的负面评论组。
- `make_adapter(provider)`：根据提供方创建适配器。

当前实现：

- `CodexRunnerAdapter`：调用本机 `codex exec`。
- `MockAdapter`：仅用于离线链路测试；初步结果上传模式不会使用它。

## 接入正式 API 的位置

新增一个继承 `BaseAdapter` 的适配器，例如：

```python
class ApiAdapter(BaseAdapter):
    name = "api"

    def analyze_batch(self, batch, settings, ds_guide):
        # 调用正式 API，返回严格 JSON 解析后的评分结果
        ...

    def summarize_group(self, park, aspect, negatives, settings, ds_scope):
        # 调用正式 API，返回严格 JSON 解析后的归因结果
        ...
```

然后在 `make_adapter(provider)` 增加 `api` 分支，并在 `safe_settings` 的允许提供方中加入 `api`。上传、异步任务、进度、页面、Excel 导出接口不需要改动。

## API 适配器必须遵守的契约

归因返回：

```json
{
  "summary": "一句问题概要",
  "causes": [
    {
      "text": "语义提炼后的问题短句",
      "source_rows": [12, 35]
    }
  ],
  "excluded": [
    {
      "source_row": 52,
      "reason": "不属于当前 DS 服务范围"
    }
  ]
}
```

服务端会再次校验：

- 行号必须属于当前输入评论；
- 每个问题必须有支持行号；
- 最多保留 10 个问题；
- 未分配评论会记录为范围排除；
- API 调用失败不得用原评论截取兜底，应返回失败状态以便重试。

## 配置与凭证

建议由环境变量或后端配置读取：

- API 密钥
- API 地址
- 模型名
- 请求超时
- 最大重试次数
- 并发限制

凭证只放后端运行环境，不写入 `frontend/`、Excel、任务 JSON 或日志。

## Prompt 来源

- 归因 Prompt 模板：`docs/attribution_prompt.md`
- 服务范围规则：`DS prompt.md`
- 运行时服务范围解析：`backend/analysis.py` 的 `_load_ds_scopes(...)`

如果修改 Prompt，建议同步更新 Prompt 版本号，并清理或重新生成旧归因缓存，避免不同版本结果混用。
