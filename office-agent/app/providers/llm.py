"""LLM Provider：百炼 OpenAI 兼容端点 + 离线 Mock（用于无 Key 测试）。"""
import os
from .base import LLMProvider


class DashScopeLLM(LLMProvider):
    def __init__(self, model=None):
        from openai import OpenAI
        self.client = OpenAI(
            api_key=os.environ["DASHSCOPE_API_KEY"],
            base_url=os.environ.get("LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1"),
        )
        self.model = model or os.environ.get("LLM_MODEL", "qwen-max")

    def chat(self, system, user):
        resp = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=0.3,
        )
        return resp.choices[0].message.content


class MockLLM(LLMProvider):
    """离线演示用：从环境变量 MOCK_LLM_FILE 读取固定回复。"""
    def chat(self, system, user):
        with open(os.environ["MOCK_LLM_FILE"], encoding="utf-8") as f:
            return f.read()


def get_llm():
    if os.environ.get("MOCK_LLM_FILE"):
        return MockLLM()
    return DashScopeLLM()
