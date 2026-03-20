# -*- coding: utf-8 -*-
"""LLM 客户端 — 封装 OpenAI API 调用（支持普通对话和 JSON 模式）"""

import json
import os
from typing import Optional

try:
    from openai import OpenAI
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False


_client = None
_config = None


def _load_config() -> dict:
    """加载配置文件"""
    global _config
    if _config is not None:
        return _config

    # 查找 config.json：先当前目录，再项目根目录
    candidates = [
        'config.json',
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'config.json'),
    ]
    for path in candidates:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                _config = json.load(f)
            return _config

    _config = {}
    return _config


def _get_client() -> Optional[object]:
    """获取 OpenAI 客户端实例"""
    global _client
    if _client is not None:
        return _client

    if not HAS_OPENAI:
        print("[WARN] 未安装 openai 库，跳过 AI 分析。请运行: pip install openai")
        return None

    config = _load_config()
    llm_config = config.get('llm', {})

    api_key = llm_config.get('api_key', '')
    if not api_key or api_key.startswith('sk-xxx'):
        print("[WARN] 未配置有效的 API key，跳过 AI 分析。请在 config.json 中设置 llm.api_key")
        return None

    _client = OpenAI(
        api_key=api_key,
        base_url=llm_config.get('base_url', 'https://api.openai.com/v1'),
    )
    return _client


# 公开别名（供 property_cert_parser 等直接调用 Vision API 使用）
def get_client():
    """返回配置好的 OpenAI 客户端实例（供 Vision API 调用）"""
    return _get_client()


def get_model() -> str:
    """返回当前配置的模型名称"""
    config = _load_config()
    return config.get('llm', {}).get('model', 'gpt-4o')


def is_available() -> bool:
    """检测 LLM 是否可用"""
    return _get_client() is not None


def chat(system_prompt: str, user_prompt: str) -> Optional[str]:
    """
    调用大模型进行对话。

    Args:
        system_prompt: 系统提示词（角色设定）
        user_prompt: 用户提示词（具体任务+数据）

    Returns:
        模型回复文本，失败返回 None
    """
    client = _get_client()
    if client is None:
        return None

    config = _load_config()
    llm_config = config.get('llm', {})
    model = llm_config.get('model', 'gpt-5.4')
    temperature = llm_config.get('temperature', 0.3)
    max_tokens = llm_config.get('max_tokens', 4096)

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=temperature,
            max_tokens=max_tokens,
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"[WARN] LLM 调用失败: {e}")
        return None


def chat_json(system_prompt: str, user_prompt: str) -> Optional[str]:
    """
    调用大模型，要求返回 JSON 格式（response_format=json_object）。
    适用于 Brain Agent 获取结构化调用计划。

    Args:
        system_prompt: 系统提示词
        user_prompt:   用户提示词

    Returns:
        JSON 字符串，失败返回 None
    """
    client = _get_client()
    if client is None:
        return None

    config = _load_config()
    llm_config = config.get('llm', {})
    model = llm_config.get('model', 'gpt-4o')
    temperature = llm_config.get('temperature', 0.1)  # JSON模式用低温度更稳定
    max_tokens = llm_config.get('max_tokens', 4096)

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=temperature,
            max_tokens=max_tokens,
            response_format={"type": "json_object"},
        )
        return response.choices[0].message.content
    except Exception as e:
        # 部分旧模型不支持 response_format，fallback 到普通模式
        print(f"[WARN] LLM JSON模式调用失败，尝试普通模式: {e}")
        return chat(system_prompt, user_prompt)


def is_available() -> bool:
    """检查 LLM 是否可用"""
    return _get_client() is not None
