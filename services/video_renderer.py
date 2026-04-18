from __future__ import annotations

import subprocess
import tempfile
import time
from pathlib import Path

from imageio_ffmpeg import get_ffmpeg_exe


class VideoRenderer:
    """Renders interactive HTML pages into MP4 clips using Playwright."""

    def __init__(self, timeout_seconds: int = 45) -> None:
        self.timeout_seconds = max(8, int(timeout_seconds or 45))

    def render(self, html_path: Path, output_path: Path, duration: float = 18.0) -> str:
        """Render the given HTML file into an MP4 file and return output path.

        Raises RuntimeError when rendering/conversion fails.
        """
        src = Path(html_path).resolve()
        if not src.exists() or src.suffix.lower() != ".html":
            raise RuntimeError(f"invalid html source: {src}")

        output = Path(output_path).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        duration_ms = int(max(6.0, float(duration or 18.0)) * 1000)

        started = time.monotonic()
        with tempfile.TemporaryDirectory(prefix="interactive_video_") as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            webm_path = self._capture_webm(src, temp_dir, duration_ms)
            self._convert_webm_to_mp4(webm_path=webm_path, mp4_path=output)

        elapsed = time.monotonic() - started
        if elapsed > self.timeout_seconds:
            raise RuntimeError(f"interactive video rendering timed out after {elapsed:.1f}s")
        if not output.exists() or output.stat().st_size <= 0:
            raise RuntimeError("rendered MP4 is missing or empty")
        return str(output)

    def _capture_webm(self, html_path: Path, temp_dir: Path, duration_ms: int) -> Path:
        try:
            from playwright.sync_api import sync_playwright  # type: ignore[import-not-found]
        except Exception as exc:  # pragma: no cover - environment dependent
            raise RuntimeError("Playwright is not available in current environment") from exc

        local_url = html_path.resolve().as_uri()
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            context = browser.new_context(
                viewport={"width": 1280, "height": 720},
                record_video_dir=str(temp_dir),
                record_video_size={"width": 1280, "height": 720},
            )
            page = context.new_page()
            page.goto(local_url, wait_until="domcontentloaded", timeout=min(self.timeout_seconds * 1000, 30000))
            try:
                page.wait_for_function(
                    "() => !!(window.__A04_DEMO_READY) || (document.body && document.body.innerText.trim().length > 0)",
                    timeout=4000,
                )
            except Exception:
                pass

            recommended_seconds = 0.0
            try:
                recommended_seconds = float(
                    page.evaluate("() => Number(window.__A04_RECOMMENDED_DURATION_SECONDS || 0)")
                )
            except Exception:
                recommended_seconds = 0.0

            capture_ms = max(duration_ms, int(max(0.0, recommended_seconds) * 1000))
            page.wait_for_timeout(capture_ms)
            page.close()
            context.close()
            browser.close()

        videos = sorted(temp_dir.glob("*.webm"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not videos:
            raise RuntimeError("Playwright did not output video file")
        return videos[0]

    def _convert_webm_to_mp4(self, webm_path: Path, mp4_path: Path) -> None:
        ffmpeg = get_ffmpeg_exe()
        command = [
            ffmpeg,
            "-y",
            "-i",
            str(webm_path),
            "-movflags",
            "+faststart",
            "-pix_fmt",
            "yuv420p",
            "-vcodec",
            "libx264",
            "-acodec",
            "aac",
            str(mp4_path),
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=self.timeout_seconds)
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "ffmpeg conversion failed").strip()
            raise RuntimeError(f"webm->mp4 conversion failed: {err[:220]}")
