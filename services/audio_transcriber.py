from __future__ import annotations

from pathlib import Path
import json
import logging
import subprocess
import wave

try:
    import imageio_ffmpeg
except Exception:  # pragma: no cover
    imageio_ffmpeg = None

try:
    from vosk import KaldiRecognizer, Model
except Exception:  # pragma: no cover
    KaldiRecognizer = None
    Model = None


class OfflineAudioTranscriber:
    """
    Offline audio transcription using Vosk.

    Expected model path examples:
    - ./models/vosk-model-small-cn-0.22
    - value from environment variable VOSK_MODEL_PATH
    """

    def __init__(self, workspace_dir: Path, model_path: str = ""):
        self.workspace_dir = workspace_dir
        self.model_path = model_path
        self._logger = logging.getLogger(__name__)
        self._model = None
        self._load_attempted = False

    def transcribe_video(self, video_path: Path) -> str:
        model = self._ensure_model()
        if model is None:
            return ""

        wav_path = self.workspace_dir / f"{video_path.stem}_audio.wav"
        if not self._extract_audio(video_path=video_path, wav_path=wav_path):
            return ""

        try:
            return self._transcribe_wav(wav_path, model)
        finally:
            try:
                wav_path.unlink(missing_ok=True)
            except Exception:
                pass

    def _ensure_model(self):
        if self._model is not None:
            return self._model
        if self._load_attempted:
            return None
        self._load_attempted = True

        if Model is None or KaldiRecognizer is None:
            self._logger.warning("Vosk 未安装，离线音频转写不可用。")
            return None

        candidate_paths = []
        if self.model_path:
            candidate_paths.append(Path(self.model_path))
        candidate_paths.append(self.workspace_dir.parent / "models" / "vosk-model-small-cn-0.22")
        candidate_paths.append(self.workspace_dir.parent / "models" / "vosk-model-cn-0.22")

        for candidate in candidate_paths:
            if candidate.exists():
                try:
                    self._model = Model(str(candidate))
                    return self._model
                except Exception as error:
                    self._logger.warning("Vosk 模型加载失败：%s", error)
                    return None

        self._logger.warning("未找到 Vosk 中文模型，离线音频转写已跳过。")
        return None

    def _extract_audio(self, video_path: Path, wav_path: Path) -> bool:
        try:
            if imageio_ffmpeg is None:
                self._logger.warning("imageio-ffmpeg 未安装，无法提取视频音频。")
                return False
            ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
        except Exception as error:
            self._logger.warning("FFmpeg 不可用，无法提取音频：%s", error)
            return False

        command = [
            ffmpeg_exe,
            "-y",
            "-i",
            str(video_path),
            "-vn",
            "-acodec",
            "pcm_s16le",
            "-ar",
            "16000",
            "-ac",
            "1",
            str(wav_path),
        ]
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=False)
        except Exception as error:
            self._logger.warning("音频提取失败：%s", error)
            return False

        if result.returncode != 0:
            self._logger.warning("FFmpeg 提取音频失败：%s", result.stderr[-400:])
            return False
        return wav_path.exists()

    def _transcribe_wav(self, wav_path: Path, model) -> str:
        try:
            wav_file = wave.open(str(wav_path), "rb")
        except Exception as error:
            self._logger.warning("无法打开 WAV 文件：%s", error)
            return ""

        try:
            if wav_file.getnchannels() != 1 or wav_file.getsampwidth() != 2 or wav_file.getcomptype() != "NONE":
                self._logger.warning("WAV 格式不符合 Vosk 要求。")
                return ""

            recognizer = KaldiRecognizer(model, wav_file.getframerate())
            recognizer.SetWords(False)
            chunks = []
            while True:
                data = wav_file.readframes(4000)
                if len(data) == 0:
                    break
                if recognizer.AcceptWaveform(data):
                    part = json.loads(recognizer.Result()).get("text", "").strip()
                    if part:
                        chunks.append(part)
            final_part = json.loads(recognizer.FinalResult()).get("text", "").strip()
            if final_part:
                chunks.append(final_part)
            return " ".join(chunks).strip()
        finally:
            wav_file.close()
