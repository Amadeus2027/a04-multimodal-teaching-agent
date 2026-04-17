from __future__ import annotations

import json
import re
from pathlib import Path
import hashlib
import zipfile
import subprocess
import os
from urllib.parse import quote

from services.document_service import DocumentExportService
from services.session_manager import REQUIRED_FIELDS


class ContentService:
    """Coordinates clarification, retrieval, instruction fusion, and package generation."""

    CREATIVE_KEYWORDS = ["动画", "动效", "小游戏", "互动游戏", "闯关", "拖拽", "配对", "课堂互动", "创意"]

    TEMPLATE_PROFILES = [
        {"name": "学术风", "scene": "概念讲解、推理证明、考试复习", "tags": ["学术", "严谨", "证明", "推理", "理论", "复习"]},
        {"name": "教育蓝", "scene": "通用课堂、教研公开课、教师培训", "tags": ["通用", "教育", "课堂", "公开课", "培训"]},
        {"name": "极简风", "scene": "信息密度高、结构化讲解", "tags": ["极简", "简洁", "条理", "结构"]},
        {"name": "科技风", "scene": "信息技术、科学、编程、AI", "tags": ["科技", "数字", "信息", "编程", "人工智能", "创新"]},
        {"name": "卡通风", "scene": "低龄课堂、趣味导入、闯关互动", "tags": ["卡通", "趣味", "游戏", "闯关", "低龄"]},
        {"name": "可爱风", "scene": "儿童化表达、启蒙课堂", "tags": ["可爱", "儿童", "启蒙", "亲和", "活泼"]},
        {"name": "插画风", "scene": "故事化课程、绘本式表达", "tags": ["插画", "故事", "绘本", "情境"]},
        {"name": "清新风", "scene": "语文、英语、综合素养课", "tags": ["清新", "轻盈", "自然", "柔和"]},
        {"name": "中国风", "scene": "传统文化、历史、古诗文", "tags": ["中国风", "传统文化", "古典", "历史", "诗词"]},
        {"name": "新中式", "scene": "传统文化、校本课程、高级感展示", "tags": ["新中式", "国风", "传统", "文化", "典雅"]},
        {"name": "商务风", "scene": "汇报课、成果展示、说课", "tags": ["商务", "汇报", "成果", "展示", "说课"]},
        {"name": "杂志风", "scene": "案例展示、作品讲评、项目化学习", "tags": ["杂志", "案例", "作品", "项目", "展示"]},
        {"name": "几何风", "scene": "数学、图形、逻辑结构", "tags": ["几何", "数学", "图形", "结构", "逻辑"]},
        {"name": "莫兰迪风", "scene": "温和叙述、心理健康、审美课程", "tags": ["莫兰迪", "柔和", "审美", "心理", "温和"]},
        {"name": "粉彩风", "scene": "低刺激、温暖陪伴型课堂", "tags": ["粉彩", "温暖", "柔和", "陪伴"]},
        {"name": "暖阳风", "scene": "德育、心理、班会、启发式课程", "tags": ["暖阳", "启发", "心理", "德育", "班会"]},
        {"name": "自然风", "scene": "科学观察、生物、劳动教育", "tags": ["自然", "观察", "科学", "生物", "劳动"]},
        {"name": "森林风", "scene": "生态主题、生命教育", "tags": ["森林", "生态", "植物", "生命", "自然"]},
        {"name": "海洋风", "scene": "地理、海洋、环保主题", "tags": ["海洋", "蓝色", "环保", "地理", "水"]},
        {"name": "复古风", "scene": "历史回顾、经典案例", "tags": ["复古", "历史", "经典", "年代"]},
        {"name": "手账风", "scene": "过程记录、学习单式课堂", "tags": ["手账", "记录", "学习单", "任务"]},
        {"name": "未来感", "scene": "前沿主题、科创展示", "tags": ["未来", "前沿", "创新", "科创"]},
        {"name": "黑金风", "scene": "高完成度展示、成果汇报", "tags": ["高端", "汇报", "展示", "黑金"]},
        {"name": "暗夜风", "scene": "深色屏显、沉浸式展示", "tags": ["暗夜", "深色", "沉浸", "夜间"]},
        {"name": "霓虹风", "scene": "未来科技、活动舞台感", "tags": ["霓虹", "未来", "舞台", "冲击"]},
        {"name": "玻璃拟态", "scene": "现代界面感、科技展示", "tags": ["玻璃", "拟态", "现代", "UI"]},
        {"name": "扁平风", "scene": "清晰易读、适合大部分课堂投屏", "tags": ["扁平", "清晰", "通用", "直观"]},
        {"name": "渐变风", "scene": "品牌感、现代感表达", "tags": ["渐变", "现代", "品牌", "流动"]},
    ]

    def __init__(self, llm_client, rag_service, output_dir: Path):
        self.llm_client = llm_client
        self.rag_service = rag_service
        self.export_service = DocumentExportService(output_dir=output_dir)
        self.workspace_root = output_dir.parent
        self.template_thumb_dir = self.workspace_root / "static" / "template_thumbs"
        self.template_thumb_dir.mkdir(parents=True, exist_ok=True)
        self.enable_com_thumbnail_fallback = str(os.getenv("ENABLE_TEMPLATE_COM_THUMBNAIL", "0")).strip().lower() in {"1", "true", "yes", "on"}
        self.only_ppt_format_templates = str(os.getenv("ONLY_PPT_FORMAT_TEMPLATES", "1")).strip().lower() in {"1", "true", "yes", "on"}
        external_profiles = self._load_external_template_profiles(output_dir.parent / "ppt_format")
        if self.only_ppt_format_templates and external_profiles:
            self.template_profiles = list(external_profiles)
        else:
            self.template_profiles = list(self.TEMPLATE_PROFILES)
            self.template_profiles.extend(external_profiles)
        dedup = {}
        for item in self.template_profiles:
            dedup[item["name"]] = item
        self.template_profiles = list(dedup.values())
        self.template_names = [item["name"] for item in self.template_profiles]
        self.template_scenes = {item["name"]: item["scene"] for item in self.template_profiles}

    def _load_external_template_profiles(self, template_dir: Path):
        if not template_dir.exists():
            return []
        profiles = []
        for path in sorted(template_dir.glob("*.pptx")):
            name = path.stem.strip()
            if not name:
                continue
            thumbnail_url = self._extract_template_thumbnail(path)
            if not thumbnail_url:
                thumbnail_url = self._build_svg_thumbnail(name=name, scene=self._infer_template_scene(name), key_seed=str(path))
            profiles.append(
                {
                    "name": name,
                    "scene": self._infer_template_scene(name),
                    "tags": self._infer_template_tags(name),
                    "thumbnail_url": thumbnail_url,
                    "source": "external",
                }
            )
        return profiles

    def _build_svg_thumbnail(self, name: str, scene: str, key_seed: str):
        seed = hashlib.md5(str(key_seed).encode("utf-8")).hexdigest()
        bg1 = f"#{seed[:6]}"
        bg2 = f"#{seed[6:12]}"
        txt = (name or "模板")[:28]
        sub = (scene or "课堂模板")[:32]
        safe_txt = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        safe_sub = sub.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        content = (
            "<svg xmlns='http://www.w3.org/2000/svg' width='960' height='540' viewBox='0 0 960 540'>"
            "<defs>"
            f"<linearGradient id='g' x1='0' y1='0' x2='1' y2='1'><stop offset='0%' stop-color='{bg1}'/><stop offset='100%' stop-color='{bg2}'/></linearGradient>"
            "</defs>"
            "<rect width='960' height='540' fill='url(#g)'/>"
            "<rect x='44' y='44' width='872' height='452' rx='24' fill='rgba(255,255,255,0.84)'/>"
            f"<text x='74' y='170' font-size='42' font-family='Microsoft YaHei, Segoe UI, sans-serif' fill='#1f2937'>{safe_txt}</text>"
            f"<text x='74' y='228' font-size='24' font-family='Microsoft YaHei, Segoe UI, sans-serif' fill='#475569'>{safe_sub}</text>"
            "<text x='74' y='448' font-size='20' font-family='Microsoft YaHei, Segoe UI, sans-serif' fill='#64748b'>A04 Demo Template Preview</text>"
            "</svg>"
        )
        file_name = f"svg_{seed}.svg"
        out_path = self.template_thumb_dir / file_name
        if not out_path.exists():
            out_path.write_text(content, encoding="utf-8")
        return f"/static/template_thumbs/{quote(file_name)}"

    def _extract_template_thumbnail(self, template_path: Path):
        try:
            stat = template_path.stat()
        except OSError:
            return ""

        key = f"{template_path.resolve()}::{stat.st_mtime_ns}::{stat.st_size}"
        digest = hashlib.md5(key.encode("utf-8")).hexdigest()

        candidates = (
            ("docProps/thumbnail.jpeg", ".jpg"),
            ("docProps/thumbnail.jpg", ".jpg"),
            ("docProps/thumbnail.png", ".png"),
        )
        try:
            with zipfile.ZipFile(template_path, "r") as zf:
                for inner_path, ext in candidates:
                    if inner_path not in zf.namelist():
                        continue
                    target_name = f"{digest}{ext}"
                    target_path = self.template_thumb_dir / target_name
                    if not target_path.exists():
                        target_path.write_bytes(zf.read(inner_path))
                    return f"/static/template_thumbs/{target_name}"
        except Exception:
            return ""
        if not self.enable_com_thumbnail_fallback:
            return ""
        fallback_url = self._export_thumbnail_via_powerpoint(template_path, digest)
        return fallback_url or ""

    def _export_thumbnail_via_powerpoint(self, template_path: Path, digest: str):
        target_name = f"{digest}.png"
        target_path = self.template_thumb_dir / target_name
        if target_path.exists():
            return f"/static/template_thumbs/{target_name}"

        src = str(template_path)
        dst = str(target_path)
        src_ps = src.replace("'", "''")
        dst_ps = dst.replace("'", "''")
        script = (
            "$ErrorActionPreference='Stop';"
            "$ppt=$null; $pres=$null;"
            "try {"
            "$ppt=New-Object -ComObject PowerPoint.Application;"
            "$ppt.Visible=0;"
            f"$pres=$ppt.Presentations.Open('{src_ps}', $false, $true, $false);"
            "if($pres.Slides.Count -gt 0){"
            f"$pres.Slides.Item(1).Export('{dst_ps}','PNG',960,540);"
            "}"
            "} finally {"
            "if($pres -ne $null){$pres.Close()}"
            "if($ppt -ne $null){$ppt.Quit()}"
            "}"
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=12,
            )
            if result.returncode == 0 and target_path.exists():
                return f"/static/template_thumbs/{target_name}"
        except Exception:
            return ""
        return ""

    def _infer_template_tags(self, name: str):
        source = str(name or "").lower()
        tags = ["模板", "ppt", "教学"]
        keyword_map = {
            "math": ["数学", "高数", "函数", "几何"],
            "function": ["函数", "图像", "推导"],
            "ratio": ["数学", "比例", "计算"],
            "angle": ["数学", "几何", "图形"],
            "geometric": ["几何", "结构", "理性"],
            "blackboard": ["课堂", "板书", "学术"],
            "cute": ["可爱", "卡通", "轻松"],
            "playful": ["互动", "活泼", "趣味"],
            "notebook": ["手账", "记录", "学习单"],
            "history": ["历史", "人文", "叙事"],
            "psychology": ["心理", "情绪", "案例"],
            "medical": ["医学", "科学", "知识"],
            "stem": ["科技", "综合", "探究"],
            "matisse": ["艺术", "创意", "插画"],
        }
        for keyword, mapped_tags in keyword_map.items():
            if keyword in source:
                tags.extend(mapped_tags)
        if "slidescarnival" in source:
            tags.extend(["演示", "教育"])
        if not any(tag in tags for tag in ["数学", "高数"]):
            tags.append("通用")
        return list(dict.fromkeys(tags))

    def _infer_template_scene(self, name: str):
        source = str(name or "").lower()
        if any(token in source for token in ["math", "function", "ratio", "angle", "geometric"]):
            return "数学/高数讲解、题型训练、概念建构"
        if any(token in source for token in ["psychology", "mental_health"]):
            return "心理与德育主题、讨论式课堂"
        if any(token in source for token in ["history", "medical", "stem"]):
            return "学科专题课、知识讲解与案例呈现"
        if any(token in source for token in ["cute", "playful", "doodles", "notebook"]):
            return "趣味课堂、互动导入、低龄或轻松表达"
        return "通用教学演示、课件设计"

    def handle_chat(self, session_id: str, session: dict, user_message: str):
        session["messages"].append({"role": "user", "content": user_message})
        self._rule_fill_slots(session, user_message)
        self._collect_creative_requests(session, user_message)

        fallback = self._fallback_clarification(session)
        llm_result = self.llm_client.chat_json(
            system_prompt=(
                "你是教师备课智能助理。"
                "请从用户消息中抽取 slots：course_theme、knowledge_points、key_difficulties、lesson_periods、style。"
                "允许从口语化表达、间接描述和上下文中推断字段。"
                "如果信息不全，只追问最关键的一个缺失字段。"
                "如果全部字段完整，则提示教师现在可以选择推荐模板，也可以继续补充课件/教案要求。"
                "严格返回 JSON：slots、assistant_reply、ready_to_generate。"
            ),
            user_prompt=json.dumps(
                {
                    "known_slots": session["slots"],
                    "user_message": user_message[:1200],
                    "required_fields": REQUIRED_FIELDS,
                    "creative_requests": session.get("creative_requests", []),
                },
                ensure_ascii=False,
            ),
            fallback=fallback,
        )

        for key, value in (llm_result.get("slots") or {}).items():
            if key in session["slots"] and str(value).strip():
                session["slots"][key] = str(value).strip()

        assistant_reply = str(llm_result.get("assistant_reply") or fallback["assistant_reply"]).strip()
        session["messages"].append({"role": "assistant", "content": assistant_reply})

        template_meta = self._template_meta(session)
        ready = self._all_slots_ready(session)
        if ready:
            session["template_picker_seen"] = True

        return {
            "assistant_reply": assistant_reply,
            "slots": session["slots"],
            "ready_to_generate": ready,
            "messages": session["messages"],
            "api_mode": "online" if self.llm_client.configured else "fallback",
            "template_suggestions": template_meta["suggestions"],
            "recommended_template": template_meta["recommended"],
            "selected_template": session.get("selected_template", ""),
            "template_picker_visible": ready,
            "template_catalog": self._template_catalog(template_meta["suggestions"]),
        }

    def generate_ppt(self, session_id: str, session: dict, selected_template: str = ""):
        package, retrievals, instruction_bundle = self._generate_package_core(session_id, session, revision="", selected_template=selected_template)
        exported = self.export_service.build_ppt(package)
        session["last_package"] = {"package": package, "files": exported}
        return self._build_package_response(session, package, retrievals, exported, instruction_bundle)

    def generate_doc(self, session_id: str, session: dict, selected_template: str = ""):
        package, retrievals, instruction_bundle = self._generate_package_core(session_id, session, revision="", selected_template=selected_template)
        exported = self.export_service.build_doc(package)
        session["last_package"] = {"package": package, "files": exported}
        return self._build_package_response(session, package, retrievals, exported, instruction_bundle)

    def revise_teaching_package(self, session_id: str, session: dict, revision: str, selected_template: str = ""):
        self._collect_creative_requests(session, revision)
        package, retrievals, instruction_bundle = self._generate_package_core(session_id, session, revision=revision, selected_template=selected_template)
        package["applied_revision"] = revision
        exported = self.export_service.build_package(package)
        session["last_package"] = {"package": package, "files": exported}
        response = self._build_package_response(session, package, retrievals, exported, instruction_bundle)
        response["revision_applied"] = revision
        return response

    def _generate_package_core(self, session_id: str, session: dict, revision: str, selected_template: str):
        template_meta = self._template_meta(session)
        if selected_template and selected_template in self.template_names:
            session["selected_template"] = selected_template
        elif session.get("selected_template") and session.get("selected_template") not in self.template_names:
            session["selected_template"] = ""
        elif not session.get("selected_template"):
            session["selected_template"] = template_meta["recommended"]

        search_plan = self._compose_search_plan(session, revision)
        retrievals = self.rag_service.search(
            session_id=session_id,
            query=search_plan["query"],
            top_k=8,
            include_global=True,
            query_hints=search_plan["query_hints"],
        )
        instruction_bundle = self._build_instruction_bundle(session, retrievals, revision, session["selected_template"])
        package = self._build_package(session=session, retrievals=retrievals, instruction_bundle=instruction_bundle)
        package["instruction_bundle"] = instruction_bundle
        return package, retrievals, instruction_bundle

    def _compose_search_plan(self, session: dict, revision: str):
        slots = session["slots"]
        query_parts = [slots.get("course_theme", ""), slots.get("knowledge_points", ""), slots.get("key_difficulties", "")]
        if revision:
            query_parts.append(revision)
        query_parts.extend(session.get("creative_requests", [])[:3])
        query_parts = [item.strip() for item in query_parts if str(item).strip()]

        keywords = []
        desired_chunk_types = ["knowledge", "overview"]
        if slots.get("knowledge_points"):
            keywords.extend(re.split(r"[，、；;\s]+", slots["knowledge_points"]))
        if slots.get("key_difficulties"):
            keywords.extend(re.split(r"[，、；;\s]+", slots["key_difficulties"]))
        if any(token in f"{revision}{' '.join(session.get('creative_requests', []))}" for token in ["案例", "情境", "应用", "例题"]):
            desired_chunk_types.append("case")
            keywords.extend(["案例", "情境", "应用"])
        if any(token in f"{revision}{slots.get('style', '')}" for token in ["风格", "版式", "排版", "模板"]):
            desired_chunk_types.append("style")
            keywords.extend(["风格", "版式", "排版"])
        if any(token in f"{revision}{' '.join(session.get('creative_requests', []))}" for token in ["动画", "互动", "游戏", "闯关"]):
            keywords.extend(["互动", "动画", "游戏"])

        return {
            "query": " ".join(query_parts),
            "query_hints": {
                "keywords": [token for token in keywords if token][:8],
                "desired_chunk_types": list(dict.fromkeys(desired_chunk_types)),
            },
        }

    def _build_instruction_bundle(self, session: dict, retrievals: list[dict], revision: str, template_choice: str):
        session_documents = []
        for doc in session.get("documents", [])[-8:]:
            session_documents.append(
                {
                    "name": doc.get("original_name", "未命名资料"),
                    "summary": (doc.get("summary") or "")[:150],
                    "knowledge_structure": doc.get("knowledge_structure", [])[:3],
                    "cases": doc.get("cases", [])[:2],
                    "content_style": doc.get("content_style", ""),
                }
            )

        local_kb_hits = []
        uploaded_hits = []
        for item in retrievals:
            target = local_kb_hits if item.get("scope") == "global" else uploaded_hits
            target.append(
                {
                    "source": item.get("source_name", ""),
                    "title": item.get("title", ""),
                    "chunk_type": item.get("chunk_type", ""),
                    "score": item.get("score", 0),
                    "summary": item.get("summary", ""),
                    "excerpt": item.get("text", "")[:180],
                    "use": self._usage_hint(item),
                }
            )

        directives = [
            "以教师意图为主线组织内容，不要只堆砌资料片段。",
            "优先吸收上传资料中的知识结构、案例、风格信号，再用本地专业知识库补足专业讲解深度。",
            "所有页面文案要适合投屏阅读，每页控制在 3-5 条要点内。",
            "若教师提出动画创意或互动小游戏，必须落到具体可执行的课堂设计。",
        ]
        if revision:
            directives.append(f"必须落实教师本轮修改意见：{revision[:120]}")

        return {
            "teacher_intent": {
                "course_theme": session["slots"].get("course_theme", ""),
                "knowledge_points": session["slots"].get("knowledge_points", ""),
                "key_difficulties": session["slots"].get("key_difficulties", ""),
                "lesson_periods": session["slots"].get("lesson_periods", ""),
                "style": session["slots"].get("style", ""),
                "selected_template": template_choice,
                "creative_requests": session.get("creative_requests", [])[:4],
                "revision": revision[:200],
            },
            "teacher_context_digest": self._conversation_digest(session),
            "uploaded_materials": session_documents,
            "uploaded_material_hits": uploaded_hits[:4],
            "local_knowledge_hits": local_kb_hits[:4],
            "generation_directives": directives,
            "output_requirements": {
                "ppt_structure": ["封面", "目录", "正文", "案例或练习", "小结"],
                "readability": "高可读、短句、适合课堂投屏",
                "template_scene": self.template_scenes.get(template_choice, ""),
            },
        }

    def _usage_hint(self, item: dict):
        chunk_type = item.get("chunk_type", "")
        if chunk_type == "knowledge":
            return "用于补充概念定义、原理推导或知识结构"
        if chunk_type == "case":
            return "用于补充例题、情境案例或课堂练习"
        if chunk_type == "style":
            return "用于吸收表达顺序、排版风格或呈现结构"
        return "用于补充课堂背景信息和表达线索"

    def _build_package(self, session: dict, retrievals: list[dict], instruction_bundle: dict):
        slots = session["slots"]
        references = self._build_reference_notes(instruction_bundle)
        creative_requests = session.get("creative_requests", [])
        fallback = self._fallback_package(
            slots=slots,
            references=references,
            revision=instruction_bundle["teacher_intent"].get("revision", ""),
            creative_requests=creative_requests,
            template_choice=instruction_bundle["teacher_intent"]["selected_template"],
        )

        llm_result = self.llm_client.chat_json(
            system_prompt=(
                "你是资深教研员和课件设计师。"
                "请把教师意图、上传资料、本地专业知识库命中内容进行有效融合，形成完整课件。"
                "优先保持教学逻辑完整：导入-建构-案例/练习-互动-小结。"
                "不得堆砌检索原文，不得直接照抄资料。"
                "严格返回 JSON：title、theme_style、slides、closing_points、references。"
                "slides 每项必须包含 title、layout、bullets，layout 仅可为 cover、agenda、content、case、summary、interactive。"
            ),
            user_prompt=json.dumps({"instruction_bundle": instruction_bundle}, ensure_ascii=False),
            fallback=fallback,
        )

        package = dict(fallback)
        package.update(
            {
                "title": llm_result.get("title") or fallback["title"],
                "theme_style": llm_result.get("theme_style") or fallback["theme_style"],
                "slides": llm_result.get("slides") or fallback["slides"],
                "closing_points": llm_result.get("closing_points") or fallback["closing_points"],
                "references": llm_result.get("references") or fallback["references"],
            }
        )
        package["summary"] = slots.copy()
        package["selected_template"] = instruction_bundle["teacher_intent"]["selected_template"]
        return package

    def _build_package_response(self, session: dict, package: dict, retrievals: list[dict], downloads: dict, instruction_bundle: dict):
        template_meta = self._template_meta(session)
        slide_thumbnail_urls = []
        if downloads.get("pptx"):
            slide_thumbnail_urls = self.export_service.build_ppt_slide_thumbnails(downloads["pptx"])
        return {
            "package": package,
            "downloads": downloads,
            "ppt_preview": self._build_ppt_preview(package, slide_thumbnail_urls),
            "doc_preview": self.export_service.build_doc_preview(package),
            "api_mode": "online" if self.llm_client.configured else "fallback",
            "template_suggestions": template_meta["suggestions"],
            "recommended_template": template_meta["recommended"],
            "selected_template": session.get("selected_template", ""),
            "reference_digest": self._build_reference_digest(retrievals),
            "instruction_bundle_summary": {
                "uploaded_material_count": len(instruction_bundle.get("uploaded_materials", [])),
                "uploaded_hit_count": len(instruction_bundle.get("uploaded_material_hits", [])),
                "local_kb_hit_count": len(instruction_bundle.get("local_knowledge_hits", [])),
                "directives": instruction_bundle.get("generation_directives", [])[:4],
            },
            "template_catalog": self._template_catalog(template_meta["suggestions"]),
        }

    def _build_ppt_preview(self, package: dict, slide_thumbnail_urls: list[str] | None = None):
        thumb_urls = slide_thumbnail_urls or []
        slides = []
        for index, slide in enumerate(package.get("slides", [])[:12]):
            bullets = []
            for bullet in (slide.get("bullets") or [])[:4]:
                text = re.sub(r"\s+", " ", str(bullet or "")).strip()
                bullets.append(text[:78] + ("…" if len(text) > 78 else ""))
            slides.append(
                {
                    "title": str(slide.get("title") or "未命名页面")[:40],
                    "layout": str(slide.get("layout") or "content"),
                    "bullets": bullets,
                    "thumbnail_url": thumb_urls[index] if index < len(thumb_urls) else "",
                }
            )
        return {"slides": slides}

    def _template_meta(self, session: dict):
        style = str(session["slots"].get("style", ""))
        creative = " ".join(session.get("creative_requests", []))
        corpus = f"{style} {creative} {session['slots'].get('course_theme', '')} {session['slots'].get('knowledge_points', '')}"
        scores = []
        for profile in self.template_profiles:
            score = 0
            if profile["name"] in style:
                score += 10
            score += sum(2 for tag in profile["tags"] if tag in corpus)
            if profile.get("source") == "external":
                score += 2
                if any(tag in (" ".join(profile.get("tags", []))) for tag in ["数学", "高数", "几何", "函数"]):
                    score += 2
            if any(keyword in creative for keyword in ["动画", "小游戏", "互动"]) and profile["name"] in {"卡通风", "可爱风", "插画风", "未来感"}:
                score += 3
            scores.append((score, profile["name"]))
        index_map = {name: idx for idx, name in enumerate(self.template_names)}
        scores.sort(key=lambda item: (-item[0], index_map.get(item[1], 9999)))
        suggestions = [name for _, name in scores[:6]]
        recommended = suggestions[0] if suggestions else "教育蓝"
        return {"suggestions": suggestions, "recommended": recommended}

    def _template_catalog(self, priority_templates: list[str]):
        ordered = []
        priority = list(dict.fromkeys(priority_templates))
        remaining = [name for name in self.template_names if name not in priority]
        for name in priority + remaining:
            profile = next((item for item in self.template_profiles if item["name"] == name), {})
            ordered.append(
                {
                    "name": name,
                    "scene": self.template_scenes.get(name, ""),
                    "recommended": name in priority[:3],
                    "thumbnail_url": profile.get("thumbnail_url", ""),
                }
            )
        return ordered

    def _build_reference_notes(self, instruction_bundle: dict):
        notes = []
        for item in instruction_bundle.get("uploaded_materials", [])[:6]:
            note = [item["name"]]
            if item.get("summary"):
                note.append(f"摘要：{item['summary']}")
            if item.get("knowledge_structure"):
                note.append(f"知识结构：{'；'.join(item['knowledge_structure'])}")
            if item.get("cases"):
                note.append(f"案例：{'；'.join(item['cases'])}")
            if item.get("content_style"):
                note.append(f"风格：{item['content_style']}")
            notes.append(" | ".join(note)[:280])

        for group_name in ("uploaded_material_hits", "local_knowledge_hits"):
            for item in instruction_bundle.get(group_name, [])[:4]:
                notes.append(f"{item['source']} | {item['use']} | {item['excerpt'][:120]}")
        return notes or ["暂无参考资料，按教师需求直接生成。"]

    def _build_reference_digest(self, retrievals: list[dict]):
        digest = []
        seen = set()
        for item in retrievals:
            key = (item.get("scope"), item.get("source_name"), item.get("title"))
            if key in seen:
                continue
            seen.add(key)
            title_prefix = "专业知识库" if item.get("scope") == "global" else "上传资料"
            digest.append(
                {
                    "title": f"{title_prefix} · {item.get('source_name', '')}",
                    "score": item.get("score", ""),
                    "summary": item.get("summary", "") or item.get("text", ""),
                    "highlights": [self._usage_hint(item), f"片段标题：{item.get('title', '正文')}"],
                }
            )
        return digest

    def _conversation_digest(self, session: dict):
        digest = []
        for item in session["messages"][-6:]:
            role = "教师" if item["role"] == "user" else "助理"
            digest.append(f"{role}：{item['content'][:120]}")
        return digest

    def _fallback_clarification(self, session: dict):
        missing = self._missing_fields(session)
        if missing:
            assistant_reply = f"我已经记录了部分教学需求。为了继续完善，请再补充“{REQUIRED_FIELDS[missing[0]]}”。"
        else:
            recommended = self._template_meta(session)["recommended"]
            assistant_reply = (
                f"教学需求已经完整。现在可以选择推荐模板“{recommended}”，"
                "也可以继续补充希望加入的动画创意、互动小游戏或教案要求。"
            )
        return {"slots": session["slots"], "assistant_reply": assistant_reply, "ready_to_generate": not missing}

    def _fallback_package(self, slots: dict, references: list[str], revision: str, creative_requests: list[str], template_choice: str):
        theme = slots.get("course_theme") or "课堂主题"
        points = slots.get("knowledge_points") or "核心知识点"
        difficulties = slots.get("key_difficulties") or "重难点"
        periods = slots.get("lesson_periods") or "1课时"
        style = template_choice or slots.get("style") or "教育蓝"
        revision_note = revision or "无"

        slides = [
            {"title": theme, "layout": "cover", "bullets": [f"知识点：{points}", f"课时安排：{periods}", f"版式模板：{style}"]},
            {"title": "学习导航", "layout": "agenda", "bullets": ["情境导入", "概念建构", "例题讲解", "互动练习", "课堂小结"]},
            {"title": "情境导入", "layout": "content", "bullets": ["用真实问题引出本课主题。", f"唤醒与“{points}”相关的前置知识。", "用一个核心问题串起整节课。"]},
            {"title": "知识结构梳理", "layout": "content", "bullets": [f"围绕“{points}”拆分知识层级。", f"重点突破：{difficulties}。", "用板书结构或流程框帮助学生建立联系。"]},
            {"title": "案例与练习", "layout": "case", "bullets": ["引入资料中的典型案例。", "设计基础到提升的练习链。", "提醒学生说明方法选择与易错点。"]},
        ]
        if creative_requests:
            slides.append(
                {
                    "title": "互动创意设计",
                    "layout": "interactive",
                    "bullets": [
                        "知识点动画创意：用逐步显现、遮罩揭示或移动路径突出概念变化。",
                        "互动小游戏：设计配对、闯关、拖拽排序或抢答环节，服务于知识点掌握。",
                        f"教师附加要求：{'；'.join(creative_requests[:3])}",
                    ],
                }
            )
        slides.append({"title": "课堂小结", "layout": "summary", "bullets": ["回顾本节课的核心概念、方法与误区。", "用出口任务检查达成度。", f"本次修改要求：{revision_note}"]})
        return {
            "title": f"{theme} 互动式教学课件",
            "theme_style": style,
            "slides": slides,
            "closing_points": [f"本课围绕“{theme}”建立清晰知识链路。", f"重点突破“{difficulties}”。", "建议导出后结合班级学情继续微调活动页。"],
            "references": references[:8],
        }

    def _rule_fill_slots(self, session: dict, message: str):
        text = str(message or "")
        patterns = {
            "course_theme": [r"(?:课程主题|主题|课题|课程)[：: ]?([^\n；;，,]+)"],
            "knowledge_points": [r"(?:知识点|核心知识|教学内容)[：: ]?([^\n；;]+)"],
            "key_difficulties": [r"(?:重难点|重点|难点)[：: ]?([^\n；;]+)"],
            "lesson_periods": [r"(?:课时安排|课时)[：: ]?([^\n；;]+)", r"(\d+\s*课时)", r"(\d+\s*分钟)"],
            "style": [r"(?:课件风格|风格|模板)[：: ]?([^\n；;]+)"],
        }
        for slot, regexes in patterns.items():
            for pattern in regexes:
                match = re.search(pattern, text)
                if match:
                    session["slots"][slot] = match.group(1).strip("。；;，, ")
                    break

        if not session["slots"].get("course_theme", "").strip():
            match = re.search(r"(?:这节课讲|本节课讲|准备讲|主要讲|我们讲)\s*([^\n，。；;]{2,30})", text)
            if match:
                session["slots"]["course_theme"] = match.group(1).strip("。；;，, ")

        if not session["slots"].get("knowledge_points", "").strip():
            match = re.search(r"(?:重点放在|核心是|主要是|会讲到|包括)\s*([^\n。；;]{3,60})", text)
            if match:
                session["slots"]["knowledge_points"] = match.group(1).strip("。；;，, ")

        if not session["slots"].get("key_difficulties", "").strip():
            match = re.search(r"(?:难点在于|学生容易|学生常错|容易混淆|最难的是)\s*([^\n。；;]{3,60})", text)
            if match:
                session["slots"]["key_difficulties"] = match.group(1).strip("。；;，, ")

        if not session["slots"].get("lesson_periods", "").strip():
            match = re.search(r"(\d+\s*(?:课时|分钟|min|节课))", text, re.IGNORECASE)
            if match:
                session["slots"]["lesson_periods"] = match.group(1).strip("。；;，, ")

        if not session["slots"].get("style", "").strip():
            match = re.search(r"(?:想要|希望|偏好|做成)\s*([^\n。；;]{2,20})(?:风|风格|模板)", text)
            if match:
                session["slots"]["style"] = f"{match.group(1).strip('。；;，, ')}风"

    def _collect_creative_requests(self, session: dict, text: str):
        compact = str(text or "").strip()
        if compact and any(keyword in compact for keyword in self.CREATIVE_KEYWORDS):
            if compact not in session["creative_requests"]:
                session["creative_requests"].append(compact[:200])

    def _missing_fields(self, session: dict):
        return [key for key, value in session["slots"].items() if not str(value).strip()]

    def _all_slots_ready(self, session: dict):
        return not self._missing_fields(session)
