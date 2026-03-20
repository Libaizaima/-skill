# -*- coding: utf-8 -*-
"""
房产证解析器 — 支持扫描图片型 PDF。

由于房产证通常是扫描件（无文字层），使用 LLM Vision 接口
将每页渲染为 Base64 图片后提取关键字段：
  - 权利人（所有者姓名）
  - 坐落（房产地址/位置）
  - 建筑面积（㎡）

返回 List[Dict]，每个元素对应一个PDF文件。
"""

import os
import sys
import json
import base64
from typing import List, Dict, Any, Optional

_src_dir = os.path.dirname(os.path.abspath(__file__))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

import pdfplumber
import llm_client


# ---------------------------------------------------------------------------
# 内部：将 PDF 第一页渲染为 Base64 PNG
# ---------------------------------------------------------------------------
def _pdf_page_to_base64(fpath: str, page_index: int = 0) -> Optional[str]:
    """
    利用 pdfplumber + PIL 将 PDF 页面转为 Base64 PNG。
    若 pdf2image 可用则优先使用（分辨率更高）。
    """
    # 优先尝试 pdf2image（更清晰）
    try:
        from pdf2image import convert_from_path
        pages = convert_from_path(fpath, dpi=200, first_page=page_index + 1, last_page=page_index + 1)
        if pages:
            from io import BytesIO
            buf = BytesIO()
            pages[0].save(buf, format="PNG")
            return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        pass

    # 降级：pdfplumber → page.to_image(PIL)
    try:
        import pdfplumber
        from io import BytesIO
        with pdfplumber.open(fpath) as pdf:
            if page_index < len(pdf.pages):
                img = pdf.pages[page_index].to_image(resolution=200)
                buf = BytesIO()
                img.save(buf, format="PNG")
                return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# 内部：先尝试 pdfplumber 文字提取（有文字层的PDF）
# ---------------------------------------------------------------------------
_FIELD_PATTERNS = {
    "权利人": [r"权\s*利\s*人[：:]\s*(.+?)(?:\s{2,}|$)", r"登记名义人[：:]\s*(.+?)(?:\s{2,}|$)"],
    "坐落":   [r"坐\s*落[：:]\s*(.+?)(?:\s{2,}|$)", r"房屋坐落[：:]\s*(.+?)(?:\s{2,}|$)"],
    "面积":   [r"建筑面积[：:]\s*([\d.]+)\s*(?:平方米|㎡|m²)?",
               r"面\s*积[：:]\s*([\d.]+)\s*(?:平方米|㎡|m²)?"],
}


def _extract_from_text(text: str) -> Dict[str, str]:
    """从文字型PDF提取字段"""
    import re
    result = {"权利人": "", "坐落": "", "面积": ""}
    for field, patterns in _FIELD_PATTERNS.items():
        for pattern in patterns:
            m = re.search(pattern, text, re.MULTILINE)
            if m:
                result[field] = m.group(1).strip()
                break
    return result


# ---------------------------------------------------------------------------
# 内部：调用 LLM Vision 提取
# ---------------------------------------------------------------------------
_VISION_SYSTEM = (
    "你是一位专业的房产证文字识别助手。"
    "请从图片中识别并提取以下信息，以JSON格式输出（不要加任何说明或代码块标记）：\n"
    '{"权利人": "...", "坐落": "...", "面积": "..."}\n'
    "其中面积只需提取数字（不含单位），若某字段识别不到则填空字符串。"
)


def _extract_via_llm_vision(b64_image: str) -> Dict[str, str]:
    """将图片送给 LLM Vision 接口提取字段"""
    empty = {"权利人": "", "坐落": "", "面积": ""}
    try:
        client = llm_client.get_client()
        if client is None:
            return empty

        response = client.chat.completions.create(
            model=llm_client.get_model(),
            messages=[
                {"role": "system", "content": _VISION_SYSTEM},
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{b64_image}"},
                        }
                    ],
                },
            ],
            max_tokens=256,
            temperature=0.0,
        )
        raw = response.choices[0].message.content.strip()
        # 清除可能的 markdown 包裹
        if raw.startswith("```"):
            lines = raw.split("\n")
            raw = "\n".join(lines[1:-1]) if lines[-1].strip() == "```" else "\n".join(lines[1:])
        data = json.loads(raw)
        return {
            "权利人": str(data.get("权利人", "")).strip(),
            "坐落":   str(data.get("坐落", "")).strip(),
            "面积":   str(data.get("面积", "")).strip(),
        }
    except Exception as e:
        print(f"  [property_cert_parser] LLM Vision 调用失败: {e}")
        return empty


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------
def parse(file_paths: List[str]) -> List[Dict[str, Any]]:
    """
    解析房产证 PDF 文件列表。

    Returns:
        [{"文件": str, "权利人": str, "坐落": str, "面积": str}, ...]
    """
    results = []
    for fpath in file_paths:
        fname = os.path.basename(fpath)
        print(f"  → 解析房产证: {fname}")
        item: Dict[str, Any] = {"文件": fname, "权利人": "", "坐落": "", "面积": ""}

        try:
            # Step 1: 先尝试文字提取
            with pdfplumber.open(fpath) as pdf:
                full_text = "\n".join(
                    page.extract_text() or "" for page in pdf.pages
                )

            if full_text.strip():
                extracted = _extract_from_text(full_text)
                item.update(extracted)
                if item["权利人"] or item["坐落"]:
                    print(f"    权利人: {item['权利人']}  坐落: {item['坐落'][:30]}  面积: {item['面积']}")
                    results.append(item)
                    continue

            # Step 2: 无文字层 → LLM Vision
            print(f"    无文字层，转 LLM Vision 识别...")
            b64 = _pdf_page_to_base64(fpath, page_index=0)
            if b64 is None:
                print(f"    [WARN] 无法渲染页面图像，跳过")
                results.append(item)
                continue

            extracted = _extract_via_llm_vision(b64)
            item.update(extracted)
            print(f"    权利人: {item['权利人']}  坐落: {item['坐落'][:30]}  面积: {item['面积']}")

        except Exception as e:
            print(f"    [WARN] 解析失败: {e}")

        results.append(item)

    return results
