from copy import deepcopy
import uuid
from datetime import datetime


REQUIRED_FIELDS = {
    "course_theme": "课程主题",
    "knowledge_points": "知识点",
    "key_difficulties": "重难点",
    "lesson_periods": "课时安排",
    "style": "课件风格",
}


DEFAULT_SESSION = {
    "created_at": "",
    "updated_at": "",
    "messages": [],
    "slots": {key: "" for key in REQUIRED_FIELDS},
    "documents": [],
    "last_package": None,
    "creative_requests": [],
    "selected_template": "",
    "template_picker_seen": False,
    "state": {
        "generating_ppt": False,
        "generating_doc": False,
        "revising": False,
        "generation_locked": False,
        "progress_percent": 0,
        "progress_label": "",
    },
}


class SessionManager:
    """Tracks lightweight in-memory conversation sessions."""

    def __init__(self):
        self._sessions = {}

    def create_session(self):
        session_id = str(uuid.uuid4())
        now = datetime.utcnow().isoformat(timespec="seconds")
        payload = deepcopy(DEFAULT_SESSION)
        payload["created_at"] = now
        payload["updated_at"] = now
        self._sessions[session_id] = payload
        return {"session_id": session_id, "required_fields": REQUIRED_FIELDS}

    def get_session(self, session_id: str):
        return self._sessions.get(session_id)

    def touch_session(self, session_id: str):
        session = self._sessions.get(session_id)
        if not session:
            return
        session["updated_at"] = datetime.utcnow().isoformat(timespec="seconds")

    def list_sessions(self):
        items = []
        for session_id, session in self._sessions.items():
            title = str(session.get("slots", {}).get("course_theme") or "").strip()
            if not title:
                first_user = next((m.get("content", "") for m in session.get("messages", []) if m.get("role") == "user"), "")
                title = str(first_user).strip()[:26] or "新对话"
            items.append(
                {
                    "session_id": session_id,
                    "title": title,
                    "updated_at": session.get("updated_at", ""),
                    "created_at": session.get("created_at", ""),
                    "message_count": len(session.get("messages", [])),
                    "has_result": bool(session.get("last_package")),
                }
            )
        items.sort(key=lambda item: item.get("updated_at", ""), reverse=True)
        return items
