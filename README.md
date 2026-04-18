# A04 多模态AI互动式教学智能体 Demo（高数版）

[![zread](https://img.shields.io/badge/Ask_Zread-_.svg?style=flat&color=00b0aa&labelColor=000000&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTQuOTYxNTYgMS42MDAxSDIuMjQxNTZDMS44ODgxIDEuNjAwMSAxLjYwMTU2IDEuODg2NjQgMS42MDE1NiAyLjI0MDFWNC45NjAxQzEuNjAxNTYgNS4zMTM1NiAxLjg4ODEgNS42MDAxIDIuMjQxNTYgNS42MDAxSDQuOTYxNTZDNS4zMTUwMiA1LjYwMDEgNS42MDE1NiA1LjMxMzU2IDUuNjAxNTYgNC45NjAxVjIuMjQwMUM1LjYwMTU2IDEuODg2NjQgNS4zMTUwMiAxLjYwMDEgNC45NjE1NiAxLjYwMDFaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik00Ljk2MTU2IDEwLjM5OTlIMi4yNDE1NkMxLjg4ODEgMTAuMzk5OSAxLjYwMTU2IDEwLjY4NjQgMS42MDE1NiAxMS4wMzk5VjEzLjc1OTlDMS42MDE1NiAxNC4xMTM0IDEuODg4MSAxNC4zOTk5IDIuMjQxNTYgMTQuMzk5OUg0Ljk2MTU2QzUuMzE1MDIgMTQuMzk5OSA1LjYwMTU2IDE0LjExMzQgNS42MDE1NiAxMy43NTk5VjExLjAzOTlDNS42MDE1NiAxMC42ODY0IDUuMzE1MDIgMTAuMzk5OSA0Ljk2MTU2IDEwLjM5OTlaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik0xMy43NTg0IDEuNjAwMUgxMS4wMzg0QzEwLjY4NSAxLjYwMDEgMTAuMzk4NCAxLjg4NjY0IDEwLjM5ODQgMi4yNDAxVjQuOTYwMUMxMC4zOTg0IDUuMzEzNTYgMTAuNjg1IDUuNjAwMSAxMS4wMzg0IDUuNjAwMUgxMy43NTg0QzE0LjExMTkgNS42MDAxIDE0LjM5ODQgNS4zMTM1NiAxNC4zOTg0IDQuOTYwMVYyLjI0MDFDMTQuMzk4NCAxLjg4NjY0IDE0LjExMTkgMS42MDAxIDEzLjc1ODQgMS42MDAxWiIgZmlsbD0iI2ZmZiIvPgo8cGF0aCBkPSJNNCAxMkwxMiA0TDQgMTJaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik00IDEyTDEyIDQiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIxLjUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPgo8L3N2Zz4K&logoColor=ffffff)](https://zread.ai/Amadeus2027/a04-multimodal-teaching-agent)                   

版本：`v1.15.0`

这是一个可独立运行的 Flask Web 项目，支持：

- 多轮文字对话，自动澄清教学需求
- 上传 PDF / PPT / 图片 / Word / 视频并提取文本
- 使用 `all-MiniLM-L6-v2` + `FAISS` 构建本地知识库 RAG
- 融合教师意图、会话资料与本地专业知识库生成 `PPTX` 课件和 `DOCX` 教案
- 接收修改意见并重新优化输出
- 页面内预览生成内容，并提供下载入口
- 对视频执行“关键帧 OCR + 离线音频转文字”联合解析
- 在需求完成后自动推荐 PPT 模板，并支持教师手动选择
- 支持项目内 `RAG/` 目录自动扫描入库，作为全局专业知识库
- 支持“知识演示 HTML → 自动录制 MP4 → 嵌入 PPT 演示页（含可点击开源演示链接）”

## 技术栈

- 后端：Python + Flask
- 前端：原生 HTML + CSS + JavaScript
- LLM：兼容 OpenAI 风格的在线 API，当前示例配置为 DeepSeek
- Embedding：`sentence-transformers/all-MiniLM-L6-v2`
- 向量库：`faiss-cpu`
- 文件处理：`PyPDF2`、`PyMuPDF`、`easyocr`、`python-pptx`、`python-docx`、`opencv-python-headless`
- 离线音频转写：`vosk` + `imageio-ffmpeg`
- 知识演示视频生成：`Phaser.js` + `Playwright`（Python）

## 运行步骤

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
Copy-Item .env.example .env
python app.py
```

浏览器访问：

```text
http://127.0.0.1:5000
```

## 使用说明

1. 点击“创建新会话”
2. 在对话区输入教学需求，系统会自动补全课程主题、知识点、重难点、课时和风格
3. 可上传多份 PDF / PPT / 图片 / Word / 视频资料，系统会解析后写入当前会话知识库
4. 系统会结合会话资料与本地 `RAG/` 专业知识库检索增强
5. 点击“生成 PPT”或“生成教案”
6. 如需继续优化，可输入修改意见重新生成

## 更新说明（已整理）

> 维护规则：`README` 仅记录重要功能/体验迭代；小修复不逐条单列，按阶段定期归并到“近期版本聚合”，避免日志噪声。

### v1.15.0（当前）

- 方向升级：将原“互动小游戏 MP4”升级为“知识演示型 MP4”，演示内容以概念渐进、过程推演、参数图像变化为主，更贴合课堂讲授。
- 意图识别扩展：新增“演示视频/知识演示/动态演示/可视化”等触发词，自动补入“知识演示”页面并生成对应视频资产。
- PPT 外链增强：在演示页新增可点击开源链接区，内置 `GeoGebra`、`Desmos`、`Math3D`、`Wolfram Demonstrations`，便于课上延伸展示。

### v1.14.0

- 生成流程解锁修复：修正“首次生成后按钮永久锁住”问题；当前会话在生成完成后可继续对话并再次生成 `PPT/教案`（仍保留“生成中互斥”保护）。
- 互动页兜底增强：当教师意图中包含“动画/互动/小游戏”等需求但 LLM 未产出 `interactive` 页时，系统会自动补入一页标准互动页，确保互动链路可执行。
- 视频嵌入稳定性修复：`PPT` 互动页改为先生成海报帧再 `add_movie()` 嵌入 `MP4`，显著降低嵌入失败回退为纯文本卡片的概率。
- 互动演示链路完善：维持“HTML 互动游戏 + Playwright 录制 + MP4 嵌入”的端到端容错，单步失败不影响整份课件导出。

### v1.13.6（当前）

- 输入区布局优化：将“生成阶段提示文案”与“发送/生成按钮”整合到同一行，减少视线跳转，提升操作连贯性。
- PPT 生成进度提示：新增输入区 `PPT` 生成 loading 条（动画进度轨 + 文案），生成期间可直观看到任务进行中状态。
- 响应式适配补充：桌面端保持同排紧凑布局，移动端自动折行为纵向结构，确保提示文案与 loading 条在小屏可读、可操作。

### v1.13.5

- 生成前置校验：新增“关键信息补全后才能生成”约束，未完成 `课程主题/知识点/重难点/课时安排/课件风格` 时，前后端都会阻止生成并提示缺失项。
- 生成互斥控制：生成 PPT/教案期间，禁止继续对话、上传资料和触发另一项生成；后端接口增加并发保护，避免并行请求导致状态冲突。
- UI 适配增强：输入区新增“生成门禁提示”与“生成中锁定态”（按钮/输入框禁用与文案提示），让不可操作原因可见、可理解。

### v1.13.4

- 排版一致性修复：统一 PPT 正文段落样式（同级统一 `level=0`、左对齐、统一行距与段后距），修复同一层级内容出现“首行/其余行缩进不一致”的问题。
- 内容充实增强：在导出归一化阶段增加要点补强策略（content/case/summary 默认扩展到 6-7 条，interactive 到 4 条），自动补入“方法提示/易错提醒/反馈”信息，减少页面单薄。
- 生成策略增强：LLM 提示词新增页内密度约束（agenda/content/case/summary 5-7 条、interactive 3-4 条），提升首次生成的内容完整度。

### v1.13.3

- 运行态修复：在 `services/document_service.py` 增加统一 `bullet` 提纯函数，程序化页面与模板导出链路统一走纯文本写入，彻底规避 `{'text':..., 'section_hint':...}` 字典字面量入页。
- 模板降级门禁：模板写入前新增稳定性检测（旋转文本框/窄高文本框/可写容量），命中高风险模板时自动回退程序化稳定版式，优先保证“清晰可读、不东倒西歪”。
- 导出兼容闭环：修复运行模块中结构化 bullet 解析分支，保证字典字符串与对象 bullet 均可落为纯文本。

### v1.13.2

- 文本清洗修复：针对 `{'text': '...', 'section_hint': '...'}` 被直接写进 PPT 的问题，新增 bullet 兼容清洗逻辑，支持对象/字典字符串两类输入并统一提纯为纯文本。
- 导出兼容增强：当检测到旧版导出器不支持结构化 bullet 时，自动降级为“纯文本 bullets”导出，避免模板页出现字典字面量脏格式。

### v1.13.1

- 发布后流程固化：每次代码更新后默认执行三步收尾检查——①全局代码检查并评估无用接口清理；②更新 `README` 版本记录；③执行 `start_app.bat --check` 启动自检。
- 维护约定补充：若未发现可安全删除的接口，明确记录“本轮无可删除项”，避免误删外部依赖入口。

### v1.13.0

- 模板排版引擎升级：优化“标题区 + 正文区”识别与文本分配逻辑，降低错位写入导致的文本重叠。
- 投屏可读性增强：正文文本框启用自动缩放、边距约束、容量估算与逐框截断，减少溢出与遮挡。
- 画面填充率提升：正文内容改为“先均匀铺满后按容量补充”，并补充知识点/重难点等结构化信息，减少大面积留白。
- 占位文案清理增强：对模板中 `Add a main point` 等占位文案做自动识别与替换/清理，降低模板残留英文观感。

### 近期版本聚合（v1.12.x）

- `.bat` 启动修复：修复 `--no-browser` 分支中的批处理语法问题（括号解析导致的“此时不应有 .”错误）。
- 启动稳定性增强：`app.py` 改为环境变量控制调试模式（`APP_DEBUG`），并默认 `use_reloader=False`，避免 Windows 下重载子进程导致“脚本看似未正常运行”。
- 启动性能兜底：`start_app.bat` 默认禁用模板 COM 缩略图启动期导出（`ENABLE_TEMPLATE_COM_THUMBNAIL=0`），避免初始化阻塞影响启动成功率。
- 启动脚本修复：`start_app.bat` 增强为多 Python 路径兜底（`.venv` / `py -3` / `python`），修复路径含空格时的启动兼容问题。
- 启动脚本增强：新增 `start_app.bat --check` 自检模式，可在不拉起服务的情况下快速验证启动前置条件。
- 代码清理：移除 `services/document_service.py` 中已无引用的辅助函数与未使用变量，减少冗余代码。
- 模板文本替换优化：改为“标题占位 + 正文占位”智能映射，并过滤竖排/装饰性窄文本框，显著降低模板被硬替换后版式崩坏问题。
- 生成预览链路增强：PPT 预览缩略图导出新增路径兼容处理（中文文件名/临时英文副本/绝对路径），并提升导出稳定性（清理异常不再中断流程）。
- 模板卡片缩略图增强：模板小图优先使用模板首页真实导出图，失败时再回退内置缩略图与 SVG。
- Windows 依赖补充：新增 `pywin32`（仅 Windows），用于 PowerPoint 自动化导出优先链路。
- UI 重构：主界面升级为 Chat 风格双视图（对话页/结果页），新增左侧历史会话列表与会话恢复。
- 对话输入区改版：回形针上传、语音输入、发送按钮整合为同一输入组件；“生成 PPT / 生成教案”入口下沉到输入区右下侧。
- 预览体验优化：课件预览区左侧缩略图与右侧放大图在桌面端统一等高，并保持缩略图滚动与键盘切页联动。
- 模板保真生成升级：当选择 `ppt_format/` 外部模板时，优先保留模板原始设计版式，仅替换文本框内容生成课件（无模板时回退到程序化样式生成）。
- 模板策略延续：默认仅展示和推荐 `ppt_format` 外部模板，避免内置模板干扰演示效果。

### 近期版本聚合（v1.11.8 及之前）

- 启动稳定性修复：模板缩略图的 PowerPoint COM 导出改为**默认关闭**（可通过环境变量 `ENABLE_TEMPLATE_COM_THUMBNAIL=1` 手动开启），避免无 PowerPoint/权限受限环境下启动阻塞。
- 模板交互优化：点击模板卡后弹窗自动收起；外部模板推荐权重提升（尤其数学/高数相关模板）。
- 缩略图可见性增强：无内置缩略图时自动生成 SVG 预览缩略图，模板卡片不再出现无图状态。
- 页面稳定性优化：资料融入说明改为固定高度 + 内部滚动，避免内容突增导致页面拉长。
- 请求容错优化：前端 API 解析增加非 JSON 响应保护，避免出现 `Unexpected token '<'` 直出报错。

- 模板能力：支持 `ppt_format/` 外部模板接入、模板风格推荐增强、模板基底生成、模板卡片真实缩略图（内置缩略图优先，缺失时可选 COM 回退）。
- 资料解析：新增 `PPT/PPTX` 参考资料解析（文本框、表格、讲者备注）。
- 教案质量：`DOCX` 教案补全教学目标、教学方法、课堂活动设计、课后作业。
- 导出体验：`PPTX/DOCX` 文件名改为内容感知命名并完善 URL 下载兼容。
- RAG链路：升级为向量 + 词法回退 + 轻量重排，并优化全局知识库入库与索引版本管理。

## 历史版本摘录

### v1.10.5

- 内置 RAG 入库性能优化：启动自动扫描 `RAG/` 时，PDF 采用快速解析模式，优先提取文本，避免超大教材导致长时间阻塞。

### v1.10.4

- 内置知识库入库策略调整：启动时对 `RAG/` 目录默认采用本地 fallback 检索入库，确保“开箱即用”。
- 检索链路增强：向量检索不可用时自动启用关键词检索回退，仍可为大模型提供本地知识库上下文。

### v1.10.3

- 内置 RAG 启动优化：`RAG/` 自动入库改为后台异步执行，应用可先启动、知识库后加载，避免启动卡顿。

### v1.10.2

- 新增内置 RAG 自动入库：程序启动时自动扫描项目目录下 `RAG/` 中的 PDF / JPG / PNG / BMP / DOCX / DOC 文件并写入全局知识库。

### v1.10.0

- 新增“本地专业知识库 RAG”入口，支持单独导入专业资料并参与检索增强。

## 说明

- 首次加载 `all-MiniLM-L6-v2` 和 `EasyOCR` 可能较慢
- `EasyOCR` 首次安装后可能自动下载识别模型
- `vector_store/` 会持久化保存本地索引
- 当前会话状态保存在内存中，适合单机演示和本地部署
- 若要启用视频音轨离线转写，请下载 Vosk 中文模型并在 `.env` 中设置 `VOSK_MODEL_PATH`
- 若要稳定生成“真实模板缩略图”，建议在 Windows 主机安装 Microsoft PowerPoint（用于无内置缩略图模板的首张幻灯片导出）。
