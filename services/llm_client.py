import json
import os
import re

import requests


class OnlineLLMClient:
    """
    Calls a free online OpenAI-compatible chat-completions API.
    By default the environment points at OpenRouter free models, but any
    compatible provider can be used.
    """

    def __init__(self):
        self.api_base = os.getenv("LLM_API_BASE", "https://api.deepseek.com/v1")
        self.api_key = os.getenv("LLM_API_KEY", "")
        self.model = os.getenv("LLM_MODEL", "deepseek-chat")
        self.site_url = os.getenv("LLM_SITE_URL", "http://localhost:5000")
        self.site_name = os.getenv("LLM_SITE_NAME", "Multimodal Teaching Agent")
        self.timeout = int(os.getenv("LLM_TIMEOUT_SECONDS", "90"))

    @property
    def configured(self) -> bool:
        return bool(self.api_key.strip())

    def chat_json(self, system_prompt: str, user_prompt: str, fallback: dict):
        if not self.configured:
            return fallback

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        if "openrouter.ai" in self.api_base:
            headers["HTTP-Referer"] = self.site_url
            headers["X-Title"] = self.site_name
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.4,
            "response_format": {"type": "json_object"},
        }

        try:
            response = requests.post(
                f"{self.api_base}/chat/completions",
                headers=headers,
                json=payload,
                timeout=self.timeout,
            )
            response.raise_for_status()
            content = response.json()["choices"][0]["message"]["content"]
            return self._parse_json(content) or fallback
        except Exception:
            return fallback

    def _parse_json(self, raw_text: str):
        if not raw_text:
            return None
        try:
            return json.loads(raw_text)
        except json.JSONDecodeError:
            pass

        fenced_match = re.search(r"```json\s*(\{.*\})\s*```", raw_text, re.DOTALL)
        if fenced_match:
            try:
                return json.loads(fenced_match.group(1))
            except json.JSONDecodeError:
                return None

        object_match = re.search(r"(\{.*\})", raw_text, re.DOTALL)
        if object_match:
            try:
                return json.loads(object_match.group(1))
            except json.JSONDecodeError:
                return None
        return None
