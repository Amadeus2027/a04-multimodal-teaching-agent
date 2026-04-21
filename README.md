# A04 多模态AI互动式教学智能体

[![zread](https://img.shields.io/badge/Ask_Zread-_.svg?style=flat&color=00b0aa&labelColor=000000&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTQuOTYxNTYgMS42MDAxSDIuMjQxNTZDMS44ODgxIDEuNjAwMSAxLjYwMTU2IDEuODg2NjQgMS42MDE1NiAyLjI0MDFWNC45NjAxQzEuNjAxNTYgNS4zMTM1NiAxLjg4ODEgNS42MDAxIDIuMjQxNTYgNS42MDAxSDQuOTYxNTZDNS4zMTUwMiA1LjYwMDEgNS42MDE1NiA1LjMxMzU2IDUuNjAxNTYgNC45NjAxVjIuMjQwMUM1LjYwMTU2IDEuODg2NjQgNS4zMTUwMiAxLjYwMDEgNC45NjE1NiAxLjYwMDFaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik00Ljk2MTU2IDEwLjM5OTlIMi4yNDE1NkMxLjg4ODEgMTAuMzk5OSAxLjYwMTU2IDEwLjY4NjQgMS42MDE1NiAxMS4wMzk5VjEzLjc1OTlDMS42MDE1NiAxNC4xMTM0IDEuODg4MSAxNC4zOTk5IDIuMjQxNTYgMTQuMzk5OUg0Ljk2MTU2QzUuMzE1MDIgMTQuMzk5OSA1LjYwMTU2IDE0LjExMzQgNS42MDE1NiAxMy43NTk5VjExLjAzOTlDNS42MDE1NiAxMC42ODY0IDUuMzE1MDIgMTAuMzk5OSA0Ljk2MTU2IDEwLjM5OTlaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik0xMy43NTg0IDEuNjAwMUgxMS4wMzg0QzEwLjY4NSAxLjYwMDEgMTAuMzk4NCAxLjg4NjY0IDEwLjM5ODQgMi4yNDAxVjQuOTYwMUMxMC4zOTg0IDUuMzEzNTYgMTAuNjg1IDUuNjAwMSAxMS4wMzg0IDUuNjAwMUgxMy43NTg0QzE0LjExMTkgNS42MDAxIDE0LjM5ODQgNS4zMTM1NiAxNC4zOTg0IDQuOTYwMVYyLjI0MDFDMTQuMzk4NCAxLjg4NjY0IDE0LjExMTkgMS42MDAxIDEzLjc1ODQgMS42MDAxWiIgZmlsbD0iI2ZmZiIvPgo8cGF0aCBkPSJNNCAxMkwxMiA0TDQgMTJaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik00IDEyTDEyIDQiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIxLjUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPgo8L3N2Zz4K&logoColor=ffffff)](https://zread.ai/Amadeus2027/a04-multimodal-teaching-agent)

面向高等数学备课场景的多模态 AI 教学智能体，支持多轮对话澄清需求、多格式资料解析、本地 RAG 检索增强，一键生成 PPTX 课件与 DOCX 教案。

## ✨ 核心功能

| 能力 | 说明 |
|------|------|
| 🗣️ 多轮对话 | 自动澄清课程主题、知识点、重难点、课时与风格，补全教学需求 |
| 📎 多格式解析 | 上传 PDF / PPT / 图片 / Word / 视频，提取文本写入会话知识库 |
| 🔍 本地 RAG | `all-MiniLM-L6-v2` + `FAISS` 向量检索，融合会话资料与 `RAG/` 全局知识库 |
| 📊 课件生成 | 一键生成 PPTX 课件与 DOCX 教案，支持 20+ 主题风格与外部模板保真 |
| 🔄 修改优化 | 输入修改意见后重新生成，迭代优化课件内容 |
| 🎬 知识演示 | 自动生成 HTML 演示动画 → Playwright 录制 MP4 → 嵌入 PPT 演示页 |
| 🔗 开源资源导航 | PPT 内嵌 GeoGebra / Desmos / Math3D / Wolfram 可点击链接，含简要介绍 |
| 🎙️ 语音输入 | 浏览器端语音识别，支持口述教学需求 |
| 📹 视频解析 | 关键帧 OCR + 离线音频转写联合解析 |

## 🛠️ 技术栈

| 层级 | 技术 |
|------|------|
| 后端 | Python · Flask |
| 前端 | 原生 HTML · CSS · JavaScript |
| LLM | 兼容 OpenAI 风格 API（默认 DeepSeek） |
| Embedding | `sentence-transformers/all-MiniLM-L6-v2` |
| 向量库 | `faiss-cpu` |
| 文件处理 | `PyPDF2` · `PyMuPDF` · `easyocr` · `python-pptx` · `python-docx` · `opencv-python-headless` |
| 音频转写 | `vosk` · `imageio-ffmpeg` |
| 视频生成 | `Phaser.js` · `Playwright` |

## 🚀 快速开始

```powershell
# 1. 创建虚拟环境
python -m venv .venv
.venv\Scripts\Activate.ps1

# 2. 安装依赖
pip install -r requirements.txt
playwright install chromium

# 3. 配置环境变量
Copy-Item .env.example .env
# 编辑 .env，填入 LLM_API_KEY（DeepSeek 或其他兼容 API）

# 4. 启动
python app.py
```

浏览器访问 **http://127.0.0.1:5000**

## 📖 使用流程

```
创建会话 → 输入教学需求 → 上传参考资料 → 选择模板 → 生成 PPT / 教案 → 查看预览 → 修改优化
```

1. 点击 **"新对话"** 创建会话
2. 在对话区描述教学需求，系统自动补全课程主题、知识点、重难点、课时和风格
3. 可上传 PDF / PPT / 图片 / Word / 视频资料，系统解析后写入会话知识库
4. 需求补全后选择 PPT 模板（也可跳过使用推荐模板）
5. 点击 **"生成 PPT"** 或 **"生成教案"**
6. 在结果页预览课件、下载文件
7. 返回对话后可点击 **"查看生成结果"** 回到结果页
8. 如需优化，输入修改意见重新生成

## 📁 项目结构

```
a04-multimodal-teaching-agent/
├── app.py                          # Flask 入口与路由
├── services/
│   ├── content_service.py          # 内容生成核心（LLM 调用、RAG 融合、资源导航）
│   ├── document_service.py         # PPTX / DOCX 导出与模板引擎
│   ├── interactive_service.py      # 知识演示 HTML 生成
│   ├── video_renderer.py           # Playwright 视频录制
│   ├── rag_service.py              # 向量检索 + 词法回退 + 重排
│   ├── llm_client.py               # OpenAI 兼容 LLM 客户端
│   ├── file_parser.py              # 多格式文件解析
│   ├── audio_transcriber.py        # 离线音频转写（Vosk）
│   └── session_manager.py          # 会话状态管理
├── templates/index.html            # 前端页面
├── static/
│   ├── css/styles.css              # 样式
│   └── js/app.js                   # 交互逻辑
├── ppt_format/                     # 外部 PPT 模板目录
├── RAG/                            # 全局专业知识库资料
└── output/                         # 生成文件输出目录
```

## ⚙️ 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `LLM_API_KEY` | — | LLM API 密钥（必填） |
| `LLM_API_BASE` | `https://api.deepseek.com/v1` | API 地址 |
| `LLM_MODEL` | `deepseek-chat` | 模型名称 |
| `LLM_TIMEOUT_SECONDS` | `90` | LLM 请求超时（秒） |
| `VOSK_MODEL_PATH` | — | Vosk 中文模型路径（启用视频音轨转写） |
| `INTERACTIVE_CLIP_DURATION_SECONDS` | `18` | 知识演示视频时长（秒） |
| `INTERACTIVE_MAX_VIDEOS` | `1` | 每次生成最大演示视频数 |
| `ENABLE_TEMPLATE_COM_THUMBNAIL` | `0` | 启用 PowerPoint COM 缩略图导出 |

## 📝 更新日志

> 维护规则：仅记录重要功能/体验迭代；小修复按阶段归并，避免日志噪声。

### v1.16.0（当前）

- 导航增强：对话页新增"查看生成结果"按钮，生成后返回对话可一键回到结果页
- 资源导航页增强：每个开源链接新增简要介绍，卡片布局优化适配多段文本
- 生成性能优化：开源资源爬取与互动视频生成并行执行，站点爬取内部并行化，缩短整体生成时间
- 文本渲染修复：修复资源导航卡片文本颜色未正确设置导致白字不可见的问题

### v1.15.0

- 方向升级：将原"互动小游戏 MP4"升级为"知识演示型 MP4"，演示内容以概念渐进、过程推演、参数图像变化为主
- 意图识别扩展：新增"演示视频/知识演示/动态演示/可视化"等触发词
- PPT 外链增强：演示页新增可点击开源链接区（GeoGebra / Desmos / Math3D / Wolfram Demonstrations）

### v1.14.0

- 生成流程解锁修复：首次生成后可继续对话并再次生成
- 互动页兜底增强：意图含互动需求但 LLM 未产出时自动补入
- 视频嵌入稳定性修复：先生成海报帧再嵌入 MP4

### v1.13.x

- 输入区布局优化：生成提示与按钮整合到同一行
- PPT 生成进度提示：新增 loading 条动画
- 生成前置校验：关键信息补全后才能生成
- 生成互斥控制：生成期间禁止对话/上传/并行生成
- 排版一致性修复：统一正文段落样式
- 内容充实增强：要点补强策略与密度约束
- 文本清洗修复：bullet 兼容清洗，规避字典字面量入页
- 模板降级门禁：高风险模板自动回退程序化版式
- 模板排版引擎升级：标题区 + 正文区识别优化

### v1.12.x

- UI 重构：Chat 风格双视图，左侧历史会话列表
- 对话输入区改版：上传/语音/发送整合为同一组件
- 模板保真生成：外部模板保留原始版式，仅替换文本
- 预览体验优化：缩略图与放大图等高联动
- 启动脚本增强：多 Python 路径兜底、自检模式

### v1.11.x 及之前

- 模板能力：外部模板接入、风格推荐、真实缩略图
- 资料解析：PPT/PPTX 参考资料解析
- 教案质量：补全教学目标、方法、活动、作业
- 导出体验：内容感知文件名、URL 下载兼容
- RAG 链路：向量 + 词法回退 + 轻量重排

### v1.10.x

- 内置 RAG 自动入库：启动扫描 `RAG/` 目录
- RAG 启动优化：后台异步入库，避免启动卡顿
- 入库性能优化：PDF 快速解析模式

## ⚠️ 注意事项

- 首次加载 `all-MiniLM-L6-v2` 和 `EasyOCR` 可能较慢，后续会缓存
- `EasyOCR` 首次安装后自动下载识别模型
- `vector_store/` 持久化保存本地索引
- 会话状态保存在内存中，适合单机演示和本地部署
- 启用视频音轨转写需下载 [Vosk 中文模型](https://alphacephei.com/vosk/models) 并设置 `VOSK_MODEL_PATH`
- 稳定生成模板缩略图建议安装 Microsoft PowerPoint（Windows）
