# 制度标尺审计平台

面向生成式 AI 平台的「个人信息保护 / 内容合规」审查工具。以「制度标尺」为判断基准，对平台规则文本或用户提交内容逐条对标，输出「合规 / 存疑 / 不合规 / 未提及」判定及证据说明，供平台自查与监管审查使用。

本仓库当前收录项目早期（大创阶段）的批处理脚本与制度标尺数据，作为可复用的分析工具链开源。

## 制度标尺

覆盖 7 个审查维度、58 条核心对标条款、449 条义务细则：信息收集与授权、用户权利、存储与安全、共享与转移、自动化决策、内设监督机构、格式与显著标识。

## 仓库内容

- 制度标尺提取：`extract_baseline_v7.py`、`extract_baseline_v7_details.py`
- 隐私政策抓取：`fetch_policy.py`、`batch_extract.py`、`extract_brokerage_clauses.py`、`extract_gb39770.py`
- 标尺自检与分片：`review_ruler.py`、`export_shards.py`、`merge_parts.py`、`fix_quotes.py`
- 报告生成：`make_report.py`（核心模块）与 `gen_*.py`、`run_all_reports.py`、`generate_summary_report.py`
- 标尺数据：`baseline_v7.json`、`baseline_v7_details.json`、`focus_details.json`、`clause_prevalence.json`
- 统一配置与依赖：`config.py`、`requirements.txt`

## 运行说明

- 路径统一到 `config.py`，项目根目录默认为本目录，可用环境变量 `PRIVACY_DETECTIVE_BASE` 覆盖。
- 安装依赖：`pip install -r requirements.txt`
- 原始输入文件（制度标尺 Excel、国标 PDF、经纪约 word 等）未随代码提供，需自行放入对应子目录；审查结果、平台规则文本、问卷数据等原始数据也未包含，避免体积与版权问题。

## 项目进度

### 已完成（大创阶段，即本仓库）

- 制度标尺：完成 7 维度、58 条款、449 细则的结构化。
- 平台审查：完成 64 份平台规则文本（模型层 / 应用层 / 分发层，To-C / To-B）的逐条对标与横向对比。
- 报告产出：横向对比、三层样本核验、图表、新旧版本 diff 等报告脚本。
- 用户调研：问卷揭示同意形式化（82.9% 受访者不认为勾选同意出于自愿）。

### 进行中（新工程，另行开发）

- 微信小程序前端：首页 / 提交 / 报告 / 标尺 / 历史 / 我的，跑通「提交 → 审查 → 报告」闭环。
- 后端服务：FastAPI，内容按「文本 / 图片 / 视频 / 音频」多模态建模，可插拔 LLM，未配置 key 时走规则兜底。
- 多 Agent：抓取对标 + 独立审计的实时审查链路。

### 下一步

- 接入真实标尺数据与 LLM，打通实时审查。
- 补齐图片 OCR、视频 / 音频 ASR 等多模态入口。
- 历史记录持久化、部署上线，并将小程序 / 后端整理开源。

## 开源协议

协议待定，推荐 MIT 或 Apache-2.0（需补 `LICENSE` 文件）。
