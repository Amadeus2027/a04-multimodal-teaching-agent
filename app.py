from pathlib import Path
import os
import threading
import traceback

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_from_directory

from services.audio_transcriber import OfflineAudioTranscriber
from services.content_service import ContentService
from services.file_parser import FileParserService
from services.llm_client import OnlineLLMClient
from services.rag_service import LocalRAGService
from services.session_manager import SessionManager, REQUIRED_FIELDS

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
VECTOR_DIR = BASE_DIR / "vector_store"
BUILTIN_RAG_DIR = BASE_DIR / "RAG"

for directory in (UPLOAD_DIR, OUTPUT_DIR, VECTOR_DIR):
    directory.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


@app.errorhandler(Exception)
def handle_api_exception(error):
    if request.path.startswith("/api/"):
        app.logger.error("API 未捕获异常：%s\n%s", error, traceback.format_exc())
        return jsonify({"error": f"服务端异常：{str(error)}"}), 500
    raise error

session_manager = SessionManager()
llm_client = OnlineLLMClient()
rag_service = LocalRAGService(vector_dir=VECTOR_DIR)
audio_transcriber = OfflineAudioTranscriber(workspace_dir=UPLOAD_DIR, model_path=os.getenv("VOSK_MODEL_PATH", ""))
file_parser = FileParserService(upload_dir=UPLOAD_DIR, llm_client=llm_client, audio_transcriber=audio_transcriber)
content_service = ContentService(llm_client=llm_client, rag_service=rag_service, output_dir=OUTPUT_DIR)


def _missing_required_slots(session: dict) -> list[dict]:
    slots = session.get("slots") or {}
    missing = []
    for key, label in REQUIRED_FIELDS.items():
        if not str(slots.get(key) or "").strip():
            missing.append({"key": key, "label": label})
    return missing


def _is_generating(session: dict) -> bool:
    state = session.get("state") or {}
    return bool(state.get("generating_ppt") or state.get("generating_doc"))


def _bootstrap_builtin_rag():
    """Fast pass: parse local RAG files and write lexical fallback chunks first."""
    if not BUILTIN_RAG_DIR.exists():
        app.logger.info("内置 RAG 目录不存在，跳过自动入库：%s", BUILTIN_RAG_DIR)
        return

    existing_files = set(rag_service.get_global_stats().get("indexed_files", []))
    allowed = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".docx", ".doc"}
    added = 0
    skipped = 0

    for path in sorted(BUILTIN_RAG_DIR.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in allowed:
            continue
        source_name = str(path.relative_to(BASE_DIR)).replace("\\", "/")
        if source_name in existing_files:
            skipped += 1
            continue
        try:
            parsed = file_parser.parse_existing_file(path)
            text = str(parsed.get("rag_text") or parsed.get("text", "")).strip()
            if not text:
                skipped += 1
                continue
            rag_service.add_global_document_fallback(
                source_name=source_name,
                text=text,
                source_type=parsed.get("file_type", "file"),
                extra_meta={
                    "summary": parsed.get("summary", ""),
                    "title": parsed.get("original_name", source_name),
                },
            )
            added += 1
        except BaseException as error:  # pragma: no cover
            app.logger.warning("内置 RAG 入库失败：%s (%s)", source_name, error)

    stats = rag_service.get_global_stats()
    app.logger.info(
        "内置 RAG 快速入库完成：新增文件 %s，跳过 %s，当前全局切片 %s（向量 %s / fallback %s）",
        added,
        skipped,
        stats.get("indexed_chunks", 0),
        stats.get("vector_chunks", 0),
        stats.get("fallback_chunks", 0),
    )


def _upgrade_builtin_rag_vectors():
    """Second pass: convert global fallback chunks to vectors in background."""
    try:
        upgraded = rag_service.upgrade_scope_to_vector(rag_service.global_session_id)
        if upgraded:
            stats = rag_service.get_global_stats()
            app.logger.info(
                "内置 RAG 向量化升级完成：新增向量切片 %s，当前全局切片 %s（向量 %s / fallback %s）",
                upgraded,
                stats.get("indexed_chunks", 0),
                stats.get("vector_chunks", 0),
                stats.get("fallback_chunks", 0),
            )
    except BaseException as error:  # pragma: no cover
        app.logger.warning("内置 RAG 向量化升级失败：%s", error)


def _bootstrap_builtin_rag_async():
    def runner():
        _bootstrap_builtin_rag()
        _upgrade_builtin_rag_vectors()

    threading.Thread(target=runner, name="builtin-rag-loader", daemon=True).start()


_bootstrap_builtin_rag_async()


@app.route("/")
def index():
    return render_template("index.html")


@app.post("/api/session/new")
def create_session():
    return jsonify(session_manager.create_session())


@app.get("/api/sessions")
def list_sessions():
    return jsonify({"sessions": session_manager.list_sessions()})


@app.get("/api/session/<session_id>")
def get_session_detail(session_id: str):
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在，请刷新页面后重试"}), 404
    return jsonify(content_service.build_session_snapshot(session_id=session_id, session=session))


@app.post("/api/chat")
def chat():
    payload = request.get_json(force=True)
    session_id = payload.get("session_id", "")
    message = payload.get("message", "").strip()
    if not session_id:
        return jsonify({"error": "缺少 session_id"}), 400
    if not message:
        return jsonify({"error": "请输入消息内容"}), 400
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在，请刷新页面后重试"}), 404
    if _is_generating(session):
        return jsonify({"error": "正在生成文件，请稍候再继续对话。"}), 409
    result = content_service.handle_chat(session_id=session_id, session=session, user_message=message)
    session_manager.touch_session(session_id)
    return jsonify(result)


@app.post("/api/upload")
def upload_files():
    session_id = request.form.get("session_id", "")
    if not session_id:
        return jsonify({"error": "缺少 session_id"}), 400
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在，请刷新页面后重试"}), 404
    if _is_generating(session):
        return jsonify({"error": "正在生成文件，暂时不能上传资料。"}), 409

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "请至少选择一个文件"}), 400

    parsed_documents = []
    for uploaded_file in files:
        try:
            parsed = file_parser.save_and_parse(uploaded_file, session_id=session_id)
        except ValueError as error:
            return jsonify({"error": str(error)}), 400
        parsed_documents.append(parsed)
        if parsed["text"].strip():
            rag_service.add_document(
                session_id=session_id,
                source_name=parsed["saved_name"],
                text=parsed.get("rag_text") or parsed["text"],
                source_type=parsed["file_type"],
                extra_meta={"summary": parsed.get("summary", ""), "title": parsed.get("original_name", "")},
            )

    session["documents"].extend(parsed_documents)
    session_manager.touch_session(session_id)
    return jsonify(
        {
            "files": parsed_documents,
            "all_files": session["documents"],
            "rag_stats": rag_service.get_stats(session_id=session_id),
            "parse_mode": "deepseek" if llm_client.configured else "local",
        }
    )


@app.post("/api/generate/ppt")
def generate_ppt():
    payload = request.get_json(force=True)
    session_id = payload.get("session_id", "")
    selected_template = payload.get("selected_template", "").strip()
    if not session_id:
        return jsonify({"error": "缺少 session_id"}), 400
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在，请刷新页面后重试"}), 404

    state = session["state"]
    missing = _missing_required_slots(session)
    if missing:
        labels = "、".join(item["label"] for item in missing)
        return jsonify({"error": f"请先补全关键信息后再生成：{labels}", "missing_fields": missing}), 400
    if state["generating_ppt"]:
        return jsonify({"error": "PPT 正在生成中，请勿重复点击。"}), 409
    if state["generating_doc"]:
        return jsonify({"error": "教案正在生成中，请稍后再生成 PPT。"}), 409
    if state["revising"]:
        return jsonify({"error": "正在根据意见优化，请稍后再试。"}), 409

    state["generating_ppt"] = True
    try:
        result = content_service.generate_ppt(session_id=session_id, session=session, selected_template=selected_template)
        session_manager.touch_session(session_id)
        return jsonify(result)
    finally:
        state["generating_ppt"] = False


@app.post("/api/generate/docx")
def generate_docx():
    payload = request.get_json(force=True)
    session_id = payload.get("session_id", "")
    selected_template = payload.get("selected_template", "").strip()
    if not session_id:
        return jsonify({"error": "缺少 session_id"}), 400
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在，请刷新页面后重试"}), 404

    state = session["state"]
    missing = _missing_required_slots(session)
    if missing:
        labels = "、".join(item["label"] for item in missing)
        return jsonify({"error": f"请先补全关键信息后再生成：{labels}", "missing_fields": missing}), 400
    if state["generating_doc"]:
        return jsonify({"error": "教案正在生成中，请勿重复点击。"}), 409
    if state["generating_ppt"]:
        return jsonify({"error": "PPT 正在生成中，请稍后再生成教案。"}), 409
    if state["revising"]:
        return jsonify({"error": "正在根据意见优化，请稍后再试。"}), 409

    state["generating_doc"] = True
    try:
        result = content_service.generate_doc(session_id=session_id, session=session, selected_template=selected_template)
        session_manager.touch_session(session_id)
        return jsonify(result)
    finally:
        state["generating_doc"] = False


@app.post("/api/revise")
def revise():
    payload = request.get_json(force=True)
    session_id = payload.get("session_id", "")
    revision = payload.get("revision", "").strip()
    selected_template = payload.get("selected_template", "").strip()
    if not session_id:
        return jsonify({"error": "缺少 session_id"}), 400
    if not revision:
        return jsonify({"error": "请输入修改意见"}), 400
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在，请刷新页面后重试"}), 404

    state = session["state"]
    if _is_generating(session):
        return jsonify({"error": "正在生成文件，请稍后再提交优化意见。"}), 409
    if state["revising"]:
        return jsonify({"error": "正在根据修改意见优化，请勿重复提交。"}), 409

    state["revising"] = True
    try:
        result = content_service.revise_teaching_package(
            session_id=session_id,
            session=session,
            revision=revision,
            selected_template=selected_template,
        )
        session_manager.touch_session(session_id)
        return jsonify(
            result
        )
    finally:
        state["revising"] = False


@app.get("/api/progress/<session_id>")
def get_progress(session_id: str):
    session = session_manager.get_session(session_id)
    if session is None:
        return jsonify({"error": "会话不存在"}), 404
    state = session.get("state") or {}
    return jsonify({
        "percent": state.get("progress_percent", 0),
        "label": state.get("progress_label", ""),
        "generating": bool(state.get("generating_ppt") or state.get("generating_doc") or state.get("revising")),
    })


@app.get("/api/download/<path:filename>")
def download_file(filename: str):
    return send_from_directory(OUTPUT_DIR, Path(filename).name, as_attachment=True)


if __name__ == "__main__":
    debug_mode = str(os.getenv("APP_DEBUG", "0")).strip().lower() in {"1", "true", "yes", "on"}
    app.run(host="0.0.0.0", port=5000, debug=debug_mode, use_reloader=False)
