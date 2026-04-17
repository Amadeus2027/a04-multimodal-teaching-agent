from copy import deepcopy
import uuid


REQUIRED_FIELDS = {
    "course_theme": "课程主题",
    "knowledge_points": "知识点",
    "key_difficulties": "重难点",
    "lesson_periods": "课时安排",
    "style": "课件风格",
}


DEFAULT_SESSION = {
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
    },
}


class SessionManager:
    """Tracks lightweight in-memory conversation sessions."""

    def __init__(self):
        self._sessions = {}

    def create_session(self):
        session_id = str(uuid.uuid4())
        self._sessions[session_id] = deepcopy(DEFAULT_SESSION)
        return {"session_id": session_id, "required_fields": REQUIRED_FIELDS}

    def get_session(self, session_id: str):
        return self._sessions.get(session_id)
