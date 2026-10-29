# 项目目录说明

```text
frontend/                         网页页面、交互和样式
backend/                          HTTP 服务、数据处理、模型适配器、导出
tests/                            单元测试和测试样例
docs/                             Prompt、使用说明、API 替换指南
DS prompt.md                      22 类服务范围和评分规则
run_demo.py                       本地启动入口
requirements.txt                  Python 依赖
var/jobs/                         运行时任务和结果，不纳入源码版本
releases/                         可复用源码版本压缩包
```

根目录中的 Excel、图片和旧单页 HTML 是参考素材，不是运行时依赖。新网页从 `frontend/` 加载，数据通过 `backend/` 接口返回。

## 复用步骤

1. 复制或解压一个 `releases/` 版本包。
2. 安装 `requirements.txt`。
3. 确认本机已安装并认证 Codex CLI。
4. 执行 `python3 run_demo.py`。
5. 原始评论选择完整分析模式；已有 22 类评分的文件选择初步结果归因模式。
6. 正式 API 接入时只替换 `backend/analysis.py` 的模型适配器。
