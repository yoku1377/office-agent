"""Provider 适配接口：所有外部模型调用唯一入口。
以后切本地部署（FunASR / Qwen3-TTS 等）只新增实现类，skill 代码不动。"""
from abc import ABC, abstractmethod


class LLMProvider(ABC):
    @abstractmethod
    def chat(self, system: str, user: str) -> str:
        """返回模型原始文本输出。"""


class ASRProvider(ABC):
    @abstractmethod
    def transcribe(self, file_url: str, diarization: bool = True, hotwords: list | None = None) -> list:
        """返回 [{begin_ms, end_ms, speaker_id, text}, ...]"""


class TTSProvider(ABC):
    @abstractmethod
    def synthesize(self, text: str, voice: str, out_path: str) -> float:
        """合成到 out_path，返回音频时长（秒）。"""
