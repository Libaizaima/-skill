# -*- coding: utf-8 -*-
"""
Brain Agent — AI大脑，扫描解压后文件树并规划工具调用路线。

工作流程：
  1. 递归扫描解压目录，获取所有文件的相对路径 + 大小
  2. 将文件树 + 工具说明一起发给 LLM
  3. LLM 以 JSON 格式返回「调用计划」
  4. 解析并校验计划，返回给 AgentRunner
"""

import os
import json
import sys
from typing import Dict, List, Any, Optional

# 将 src 目录加入路径，便于导入兄弟模块
_src_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

import llm_client
from agents.tool_registry import get_tool_descriptions_for_prompt, TOOL_SCHEMA_MAP


# ---------------------------------------------------------------------------
# Prompt 模板
# ---------------------------------------------------------------------------
_SYSTEM_PROMPT = """你是一位专业的财务文档分类专家，擅长识别企业财务资料的文件结构。
你的任务是：根据解压后的文件树，制定一份精确的工具调用计划（JSON格式）。

规则：
1. 每个文件只能分配给一个工具（或 skip）
2. 同类型的多个文件合并到同一个 tool 调用中（files 数组）
3. 流水文件：如果文件名或目录名包含"流水"、"明细"、"对账单"、"日记账"、"基本户"等，用 bank_flow_parser
4. 注意：CSV后缀的流水文件也用 bank_flow_parser
5. 进项发票（目录含"进项"或文件名含"进项"）：xlsx用 invoice_parser_in，pdf用 pdf_invoice_parser_in
6. 销项发票（目录含"销项"或文件名含"销项"）：xlsx用 invoice_parser_out，pdf用 pdf_invoice_parser_out
7. 应收（文件名含"应收"）：用 receivable_parser
8. 应付（文件名含"应付"）：用 payable_parser
9. 征信报告（文件名含"征信"或"信用报告"，pdf格式）：用 credit_report_parser
10. 无关文件（.DS_Store、图片、txt说明等）：用 skip
11. 输出纯 JSON，不要加任何解释文字或 Markdown 代码块

输出格式（严格遵守）：
{
  "company_name": "公司名称（从目录名推断）",
  "analysis_notes": "对文件结构的简要说明（1-2句话）",
  "plan": [
    {
      "tool": "工具名称",
      "files": ["相对路径1", "相对路径2"],
      "reason": "为什么用这个工具（简短说明）"
    }
  ]
}"""


def _build_file_tree(extract_dir: str) -> List[Dict[str, Any]]:
    """
    递归扫描解压目录，返回文件信息列表。

    Returns:
        [{"path": "相对路径", "size": 文件大小字节, "ext": 后缀}, ...]
    """
    files = []
    for root, dirs, filenames in os.walk(extract_dir):
        # 跳过隐藏目录
        dirs[:] = [d for d in dirs if not d.startswith('.') and d != '__MACOSX']
        for fname in filenames:
            if fname.startswith('.'):
                continue
            full_path = os.path.join(root, fname)
            rel_path = os.path.relpath(full_path, extract_dir)
            try:
                size = os.path.getsize(full_path)
            except OSError:
                size = 0
            ext = os.path.splitext(fname)[1].lower()
            files.append({
                "path": rel_path,
                "size_kb": round(size / 1024, 1),
                "ext": ext,
            })
    return files


def _format_file_tree_for_prompt(file_list: List[Dict[str, Any]]) -> str:
    """将文件列表格式化为易于 LLM 阅读的文本"""
    if not file_list:
        return "（无文件）"
    lines = []
    for f in file_list:
        lines.append(f'  {f["path"]}  [{f["size_kb"]} KB]')
    return "\n".join(lines)


def _build_user_prompt(file_list: List[Dict[str, Any]], tool_desc: str) -> str:
    """构建发给 LLM 的用户 Prompt"""
    tree_text = _format_file_tree_for_prompt(file_list)
    file_paths_json = json.dumps([f["path"] for f in file_list], ensure_ascii=False)

    return f"""请分析以下解压后的文件目录，制定工具调用计划。

## 文件清单（共 {len(file_list)} 个文件）
{tree_text}

## 所有文件路径（JSON数组，用于生成计划时填写）
{file_paths_json}

## 可用工具说明
{tool_desc}

请输出工具调用计划（JSON格式）。"""


def _parse_llm_response(response_text: str) -> Optional[Dict[str, Any]]:
    """解析 LLM 返回的 JSON，处理可能的格式问题"""
    if not response_text:
        return None

    # 去掉可能的 markdown 代码块
    text = response_text.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        # 去掉首行（```json）和末行（```）
        text = "\n".join(lines[1:-1]) if lines[-1].strip() == "```" else "\n".join(lines[1:])

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"[BrainAgent] JSON解析失败: {e}")
        print(f"[BrainAgent] 原始响应（前500字）: {response_text[:500]}")
        return None


def _validate_and_normalize_plan(plan_data: Dict[str, Any], extract_dir: str) -> Dict[str, Any]:
    """校验并规范化 LLM 返回的计划"""
    valid_tools = set(TOOL_SCHEMA_MAP.keys())

    normalized_plan = []
    skipped_files = []  # 记录被 skip 或未能分类的文件

    for step in plan_data.get("plan", []):
        tool = step.get("tool", "skip")
        if tool not in valid_tools:
            print(f"[BrainAgent] 未知工具 '{tool}'，改为 skip")
            tool = "skip"

        files = step.get("files", [])
        reason = step.get("reason", "")

        # 转为绝对路径，同时检查文件是否存在
        abs_files = []
        for rel_p in files:
            abs_p = os.path.join(extract_dir, rel_p)
            if os.path.exists(abs_p):
                abs_files.append(abs_p)
            else:
                print(f"[BrainAgent] 文件不存在，跳过: {rel_p}")
                skipped_files.append({"file": rel_p, "reason": "文件不存在"})

        if tool == "skip":
            # 大脑主动跳过的文件
            for f in abs_files:
                skipped_files.append({
                    "file": os.path.relpath(f, extract_dir),
                    "reason": reason or "LLM判断为无关文件",
                })
            continue

        if not abs_files:
            continue

        normalized_plan.append({
            "tool": tool,
            "files": abs_files,
            "reason": reason,
        })

    return {
        "company_name": plan_data.get("company_name", ""),
        "analysis_notes": plan_data.get("analysis_notes", ""),
        "plan": normalized_plan,
        "skipped_files": skipped_files,
    }


def _fallback_plan(extract_dir: str, fallback_file_map: Dict) -> Dict[str, Any]:
    """
    当 LLM 不可用时，使用 extractor 的关键词分类结果生成备用计划。
    """
    print("[BrainAgent] 使用 fallback 分类方案（extractor关键词）")
    plan = []

    if fallback_file_map.get("flows"):
        plan.append({"tool": "bank_flow_parser", "files": fallback_file_map["flows"], "reason": "关键词匹配-流水"})
    if fallback_file_map.get("invoices_in"):
        plan.append({"tool": "invoice_parser_in", "files": fallback_file_map["invoices_in"], "reason": "关键词匹配-进项"})
    if fallback_file_map.get("invoices_out"):
        plan.append({"tool": "invoice_parser_out", "files": fallback_file_map["invoices_out"], "reason": "关键词匹配-销项"})
    if fallback_file_map.get("receivable"):
        plan.append({"tool": "receivable_parser", "files": fallback_file_map["receivable"], "reason": "关键词匹配-应收"})
    if fallback_file_map.get("payable"):
        plan.append({"tool": "payable_parser", "files": fallback_file_map["payable"], "reason": "关键词匹配-应付"})
    if fallback_file_map.get("credit_reports"):
        plan.append({"tool": "credit_report_parser", "files": fallback_file_map["credit_reports"], "reason": "关键词匹配-征信"})
    if fallback_file_map.get("invoices_in_pdf"):
        plan.append({"tool": "pdf_invoice_parser_in", "files": fallback_file_map["invoices_in_pdf"], "reason": "关键词匹配-PDF进项"})
    if fallback_file_map.get("invoices_out_pdf"):
        plan.append({"tool": "pdf_invoice_parser_out", "files": fallback_file_map["invoices_out_pdf"], "reason": "关键词匹配-PDF销项"})

    return {
        "company_name": fallback_file_map.get("company_name", ""),
        "analysis_notes": "使用关键词匹配分类（LLM不可用）",
        "plan": plan,
    }


def plan(extract_dir: str, fallback_file_map: Optional[Dict] = None) -> Dict[str, Any]:
    """
    Brain Agent 主入口：扫描文件树，调用 LLM 生成调用计划。

    Args:
        extract_dir: 解压目录绝对路径
        fallback_file_map: extractor 的分类结果，LLM 不可用时使用

    Returns:
        {
            "company_name": str,
            "analysis_notes": str,
            "plan": [{"tool": str, "files": [str, ...], "reason": str}, ...]
        }
    """
    print(f"[BrainAgent] 扫描文件树: {extract_dir}")
    file_list = _build_file_tree(extract_dir)
    print(f"[BrainAgent] 共发现 {len(file_list)} 个文件")

    # 如果 LLM 不可用，使用 fallback
    if not llm_client.is_available():
        if fallback_file_map:
            return _fallback_plan(extract_dir, fallback_file_map)
        return {"company_name": "", "analysis_notes": "LLM不可用且无fallback", "plan": []}

    # 构建 Prompt
    tool_desc = get_tool_descriptions_for_prompt()
    user_prompt = _build_user_prompt(file_list, tool_desc)

    print("[BrainAgent] 调用 LLM 生成调用计划...")
    response = llm_client.chat_json(_SYSTEM_PROMPT, user_prompt)

    if response is None:
        print("[BrainAgent] LLM 调用失败，使用 fallback 方案")
        if fallback_file_map:
            return _fallback_plan(extract_dir, fallback_file_map)
        return {"company_name": "", "analysis_notes": "LLM调用失败", "plan": []}

    plan_data = _parse_llm_response(response)
    if plan_data is None:
        print("[BrainAgent] 响应解析失败，使用 fallback 方案")
        if fallback_file_map:
            return _fallback_plan(extract_dir, fallback_file_map)
        return {"company_name": "", "analysis_notes": "JSON解析失败", "plan": []}

    result = _validate_and_normalize_plan(plan_data, extract_dir)

    # 打印计划摘要
    print(f"[BrainAgent] 公司名称: {result['company_name']}")
    print(f"[BrainAgent] 分析说明: {result['analysis_notes']}")
    print(f"[BrainAgent] 调用计划: 共 {len(result['plan'])} 个工具调用")
    for step in result["plan"]:
        print(f"  → {step['tool']}: {len(step['files'])} 个文件 | {step['reason']}")

    return result
