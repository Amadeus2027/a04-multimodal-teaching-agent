from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from textwrap import dedent


class InteractiveService:
    """Generates HTML knowledge-demo animations for lesson slides."""

    SUPPORTED_DEMO_TYPES = {"concept", "timeline", "graph"}
    LEGACY_GAME_TYPE_MAP = {
        "match": "concept",
        "memory": "timeline",
        "quiz": "graph",
    }

    def __init__(self, llm_client) -> None:
        self.llm_client = llm_client

    def generate_game(self, topic: str, game_type: str, output_dir: Path) -> str:
        """Backward-compatible entry for legacy callers."""
        return self.generate_demo(topic=topic, demo_type=game_type, output_dir=output_dir)

    def generate_demo(self, topic: str, demo_type: str, output_dir: Path) -> str:
        """Generate an HTML knowledge-demo file and return its absolute path."""
        safe_topic = str(topic or "课堂知识点演示").strip() or "课堂知识点演示"
        normalized_type = str(demo_type or "concept").strip().lower()
        normalized_type = self.LEGACY_GAME_TYPE_MAP.get(normalized_type, normalized_type)
        if normalized_type not in self.SUPPORTED_DEMO_TYPES:
            normalized_type = "concept"

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        digest = hashlib.md5(f"{safe_topic}::{normalized_type}".encode("utf-8")).hexdigest()[:12]
        html_path = output_dir / f"demo_{normalized_type}_{digest}.html"

        html = self._generate_html_via_llm(topic=safe_topic, demo_type=normalized_type)
        if not html:
            html = self._fallback_demo_html(topic=safe_topic, demo_type=normalized_type)

        html_path.write_text(html, encoding="utf-8")
        return str(html_path.resolve())

    def _generate_html_via_llm(self, topic: str, demo_type: str) -> str:
        fallback_payload = {
            "title": f"{topic} 知识演示动画",
            "demo_type": demo_type,
            "html": "",
        }

        llm_result = self.llm_client.chat_json(
            system_prompt=(
                "你是前端教学可视化开发专家。"
                "请生成一个单文件 HTML5 知识演示页面。"
                "必须可在 file:// 离线运行，不依赖任何外部 CDN/脚本/样式。"
                "页面需包含：标题、阶段提示、动画演示区、重播按钮。"
                "请在脚本中设置 window.__A04_DEMO_READY=true 和 window.__A04_RECOMMENDED_DURATION_SECONDS=18。"
                "不要设计闯关或计分机制。"
                "严格返回 JSON：title、demo_type、html。"
            ),
            user_prompt=json.dumps(
                {
                    "topic": topic,
                    "demo_type": demo_type,
                    "constraints": {
                        "single_file": True,
                        "language": "zh-CN",
                        "offline_ready": True,
                    },
                },
                ensure_ascii=False,
            ),
            fallback=fallback_payload,
        )

        html = str((llm_result or {}).get("html") or "").strip()
        if not html:
            return ""
        html = self._strip_fences(html)
        lowered = html.lower()
        if "<html" not in lowered:
            return ""
        if re.search(r"<script[^>]+src\s*=\s*['\"]https?://", lowered):
            return ""
        if "__a04_demo_ready" not in lowered:
            return ""
        return html

    def _strip_fences(self, text: str) -> str:
        compact = str(text or "").strip()
        compact = re.sub(r"^```(?:html)?\\s*", "", compact, flags=re.IGNORECASE)
        compact = re.sub(r"\\s*```$", "", compact)
        return compact.strip()

    def _fallback_demo_html(self, topic: str, demo_type: str) -> str:
        title = f"{topic} · 知识演示动画"
        guides = {
            "concept": "演示重点：从定义到要点逐步出现，帮助学生建立概念框架。",
            "timeline": "演示重点：按步骤推进，展示知识形成与推导过程。",
            "graph": "演示重点：通过参数变化观察图像动态规律。",
        }
        guide_text = guides.get(demo_type, guides["concept"])

        return dedent(
            f"""
            <!doctype html>
            <html lang="zh-CN">
            <head>
              <meta charset="utf-8" />
              <meta name="viewport" content="width=device-width, initial-scale=1" />
              <title>{title}</title>
              <style>
                body {{ margin: 0; font-family: 'Microsoft YaHei', sans-serif; background: #f4f7ff; color:#1f2a44; }}
                .hud {{ padding: 12px 16px; background: #fff; border-bottom: 1px solid #dbe3f5; }}
                .hud h1 {{ margin: 0 0 6px; font-size: 20px; }}
                .meta {{ display:flex; gap: 12px; align-items:center; flex-wrap:wrap; font-size:14px; }}
                .pill {{ background:#eaf0ff; padding:4px 10px; border-radius:999px; }}
                .stage {{ background:#fff4dc; color:#7a4b00; }}
                #panel {{ width: 100%; height: calc(100vh - 96px); display:flex; justify-content:center; align-items:center; }}
                #board {{ width: 1180px; height: 560px; background:#fff; border:2px solid #3559d5; border-radius:16px; box-shadow:0 8px 24px rgba(53,89,213,.12); position:relative; overflow:hidden; }}
                #canvas {{ width:100%; height:100%; display:block; }}
                button {{ border: none; border-radius: 8px; padding: 6px 10px; background:#3559d5; color:#fff; cursor:pointer; }}
              </style>
            </head>
            <body>
              <div class="hud">
                <h1>{title}</h1>
                <div class="meta">
                  <span class="pill">{guide_text}</span>
                  <span class="pill stage" id="stage">阶段：准备</span>
                  <button id="replayBtn">重播</button>
                </div>
              </div>
              <div id="panel">
                <div id="board"><canvas id="canvas" width="1180" height="560"></canvas></div>
              </div>

              <script>
                const DEMO_TYPE = {json.dumps(demo_type, ensure_ascii=False)};
                const TOPIC = {json.dumps(topic, ensure_ascii=False)};
                const canvas = document.getElementById('canvas');
                const ctx = canvas.getContext('2d');
                const stageEl = document.getElementById('stage');
                const replayBtn = document.getElementById('replayBtn');
                const setStage = (txt) => stageEl.textContent = `阶段：${{txt}}`;

                let startAt = 0;
                let req = 0;

                function clearBoard() {{
                  ctx.fillStyle = '#ffffff';
                  ctx.fillRect(0, 0, canvas.width, canvas.height);
                  ctx.fillStyle = '#2b3f72';
                  ctx.font = 'bold 34px Microsoft YaHei';
                  ctx.fillText('课堂知识演示区', 42, 58);
                  ctx.fillStyle = '#496199';
                  ctx.font = '24px Microsoft YaHei';
                  ctx.fillText(`主题：${{TOPIC}}`, 42, 98);
                }}

                function drawConcept(t) {{
                  const lines = [
                    `1. ${{TOPIC}} 的核心定义：描述对象、条件与结论。`,
                    '2. 关键特征：可观察、可比较、可迁移。',
                    '3. 易错点：不要混淆前提条件与结论。',
                    '4. 应用提示：先判定条件，再选择方法。',
                  ];
                  const phase = Math.floor(t / 1800);
                  if (phase <= 0) setStage('定义引入');
                  else if (phase == 1) setStage('要点展开');
                  else setStage('迁移应用');

                  lines.forEach((line, idx) => {{
                    const revealMs = idx * 1400;
                    if (t >= revealMs) {{
                      const alpha = Math.min(1, (t - revealMs) / 400);
                      ctx.globalAlpha = alpha;
                      ctx.fillStyle = '#1f2a44';
                      ctx.font = '30px Microsoft YaHei';
                      ctx.fillText(line, 110, 190 + idx * 90);
                      ctx.globalAlpha = 1;
                    }}
                  }});
                }}

                function drawTimeline(t) {{
                  const steps = ['问题提出', '方法建模', '结果验证'];
                  const phase = Math.floor(t / 1700);
                  if (phase <= 0) setStage('步骤 1/3');
                  else if (phase == 1) setStage('步骤 2/3');
                  else if (phase == 2) setStage('步骤 3/3');
                  else setStage('总结回顾');

                  ctx.strokeStyle = '#3559d5';
                  ctx.lineWidth = 4;
                  ctx.beginPath();
                  ctx.moveTo(150, 360);
                  ctx.lineTo(1080, 360);
                  ctx.stroke();

                  steps.forEach((label, idx) => {{
                    const x = 260 + idx * 320;
                    const reveal = t >= idx * 1300;
                    if (!reveal) return;
                    const alpha = Math.min(1, (t - idx * 1300) / 500);
                    ctx.globalAlpha = alpha;
                    ctx.fillStyle = '#9bb7ff';
                    ctx.beginPath();
                    ctx.arc(x, 360, 20, 0, Math.PI * 2);
                    ctx.fill();
                    ctx.strokeStyle = '#3559d5';
                    ctx.strokeRect(x - 140, 210, 280, 118);
                    ctx.fillStyle = '#1f2a44';
                    ctx.font = '30px Microsoft YaHei';
                    ctx.fillText(`${{idx + 1}}. ${{label}}`, x - 114, 280);
                    ctx.globalAlpha = 1;
                  }});
                }}

                function drawGraph(t) {{
                  const sec = (t % 6400) / 1000;
                  const a = 0.4 + (Math.sin(sec) + 1) * 0.6;
                  if (a < 0.8) setStage('开口变缓');
                  else if (a < 1.2) setStage('标准抛物线');
                  else setStage('开口变陡');

                  const ox = 180, oy = 500;
                  ctx.strokeStyle = '#3559d5';
                  ctx.lineWidth = 3;
                  ctx.beginPath();
                  ctx.moveTo(ox, 120); ctx.lineTo(ox, oy); ctx.lineTo(1120, oy); ctx.stroke();

                  ctx.fillStyle = '#1f2a44';
                  ctx.font = '40px Microsoft YaHei';
                  ctx.fillText(`y = ${{a.toFixed(1)}}·x²`, 830, 160);

                  ctx.strokeStyle = '#2f64d6';
                  ctx.lineWidth = 4;
                  ctx.beginPath();
                  for (let i = -12; i <= 12; i++) {{
                    const x = ox + i * 34;
                    const y = oy - a * i * i * 7;
                    if (i === -12) ctx.moveTo(x, y);
                    else ctx.lineTo(x, y);
                  }}
                  ctx.stroke();
                }}

                function tick(now) {{
                  if (!startAt) startAt = now;
                  const t = now - startAt;
                  clearBoard();
                  if (DEMO_TYPE === 'timeline') drawTimeline(t);
                  else if (DEMO_TYPE === 'graph') drawGraph(t);
                  else drawConcept(t);
                  req = requestAnimationFrame(tick);
                }}

                function startDemo() {{
                  cancelAnimationFrame(req);
                  startAt = 0;
                  req = requestAnimationFrame(tick);
                }}

                replayBtn.addEventListener('click', startDemo);
                window.__A04_RECOMMENDED_DURATION_SECONDS = 18;
                window.__A04_DEMO_READY = true;
                startDemo();
              </script>
            </body>
            </html>
            """
        ).strip()
