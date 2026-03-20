# -*- coding: utf-8 -*-
"""
Tool Agent — 工具执行Agent，按照 Brain Agent 的计划逐步调用各解析工具。

职责：
  - 接收 Brain Agent 生成的 plan（JSON）
  - 按顺序调用对应的解析脚本
  - 将各工具结果汇总为统一的 parsed_data 字典
  - 打印每步进度
"""

import os
import sys
import pandas as pd
from typing import Dict, List, Any, Optional

# 将 src 目录加入路径
_src_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

import bank_flow_parser
import invoice_parser
import receivable_payable_parser
import credit_report_parser
import pdf_invoice_parser
import property_cert_parser
import financial_statement_parser
import tax_cert_parser


# ---------------------------------------------------------------------------
# 工具分发表
# ---------------------------------------------------------------------------
def _run_bank_flow_parser(files: List[str], ctx: Dict) -> None:
    """解析对公流水，结果合并到 ctx['flow_df']"""
    print(f"[ToolAgent] 解析对公流水: {len(files)} 个文件")
    try:
        df = bank_flow_parser.parse(files)
        print(f"  → 解析得到 {len(df)} 条记录")
        if ctx.get("flow_df") is None or len(ctx["flow_df"]) == 0:
            ctx["flow_df"] = df
        else:
            ctx["flow_df"] = pd.concat([ctx["flow_df"], df], ignore_index=True)
    except Exception as e:
        print(f"  [WARN] 流水解析失败: {e}")


def _run_invoice_parser_in(files: List[str], ctx: Dict) -> None:
    """解析进项发票，结果合并到 ctx['in_invoice_df']"""
    print(f"[ToolAgent] 解析进项发票: {len(files)} 个文件")
    try:
        in_df, _ = invoice_parser.parse(files, [])
        print(f"  → 解析得到 {len(in_df)} 条记录")
        if ctx.get("in_invoice_df") is None or len(ctx["in_invoice_df"]) == 0:
            ctx["in_invoice_df"] = in_df
        else:
            ctx["in_invoice_df"] = pd.concat([ctx["in_invoice_df"], in_df], ignore_index=True)
    except Exception as e:
        print(f"  [WARN] 进项发票解析失败: {e}")


def _run_invoice_parser_out(files: List[str], ctx: Dict) -> None:
    """解析销项发票，结果合并到 ctx['out_invoice_df']"""
    print(f"[ToolAgent] 解析销项发票: {len(files)} 个文件")
    try:
        _, out_df = invoice_parser.parse([], files)
        print(f"  → 解析得到 {len(out_df)} 条记录")
        if ctx.get("out_invoice_df") is None or len(ctx["out_invoice_df"]) == 0:
            ctx["out_invoice_df"] = out_df
        else:
            ctx["out_invoice_df"] = pd.concat([ctx["out_invoice_df"], out_df], ignore_index=True)
    except Exception as e:
        print(f"  [WARN] 销项发票解析失败: {e}")


def _run_receivable_parser(files: List[str], ctx: Dict) -> None:
    """解析应收账款"""
    print(f"[ToolAgent] 解析应收账款: {len(files)} 个文件")
    try:
        recv_df, _ = receivable_payable_parser.parse(files, [])
        print(f"  → 解析得到 {len(recv_df)} 条记录")
        if ctx.get("recv_df") is None or len(ctx.get("recv_df", [])) == 0:
            ctx["recv_df"] = recv_df
        else:
            ctx["recv_df"] = pd.concat([ctx["recv_df"], recv_df], ignore_index=True)
    except Exception as e:
        print(f"  [WARN] 应收解析失败: {e}")


def _run_payable_parser(files: List[str], ctx: Dict) -> None:
    """解析应付账款"""
    print(f"[ToolAgent] 解析应付账款: {len(files)} 个文件")
    try:
        _, pay_df = receivable_payable_parser.parse([], files)
        print(f"  → 解析得到 {len(pay_df)} 条记录")
        if ctx.get("pay_df") is None or len(ctx.get("pay_df", [])) == 0:
            ctx["pay_df"] = pay_df
        else:
            ctx["pay_df"] = pd.concat([ctx["pay_df"], pay_df], ignore_index=True)
    except Exception as e:
        print(f"  [WARN] 应付解析失败: {e}")


def _run_credit_report_parser(files: List[str], ctx: Dict) -> None:
    """解析征信报告"""
    print(f"[ToolAgent] 解析征信报告: {len(files)} 个文件")
    try:
        credit_data = credit_report_parser.parse_all(files)
        ctx["credit_data"] = credit_data
        if credit_data.get("personal"):
            print(f"  → 法人征信: {credit_data['personal'].get('姓名', '')}")
        if credit_data.get("company"):
            print(f"  → 企业征信: 信贷机构数={credit_data['company'].get('信贷机构数', '?')}")
    except Exception as e:
        print(f"  [WARN] 征信解析失败: {e}")


def _run_pdf_invoice_parser_in(files: List[str], ctx: Dict, save_dir: Optional[str] = None) -> None:
    """解析进项 PDF 发票"""
    print(f"[ToolAgent] 解析进项PDF发票: {len(files)} 个文件")
    try:
        pdf_in_dir = save_dir or os.path.dirname(files[0])
        pdf_df = pdf_invoice_parser.parse_pdf_invoices(files, save_dir=pdf_in_dir)
        pdf_df["方向"] = "in"
        if len(pdf_df) > 0:
            # 与已有进项数据去重合并
            existing = ctx.get("in_invoice_df")
            if existing is not None and "发票号码" in existing.columns and "发票号码" in pdf_df.columns:
                existing_nos = set(existing["发票号码"].astype(str))
                new_rows = pdf_df[~pdf_df["发票号码"].astype(str).isin(existing_nos)]
                print(f"  → PDF进项: {len(pdf_df)}张解析, {len(new_rows)}张新增(去重)")
                if len(new_rows) > 0:
                    ctx["in_invoice_df"] = pd.concat([existing, new_rows], ignore_index=True)
            else:
                print(f"  → PDF进项: {len(pdf_df)} 张")
                ctx["in_invoice_df"] = pdf_df
    except Exception as e:
        print(f"  [WARN] PDF进项发票解析失败: {e}")


def _run_pdf_invoice_parser_out(files: List[str], ctx: Dict, save_dir: Optional[str] = None) -> None:
    """解析销项 PDF 发票"""
    print(f"[ToolAgent] 解析销项PDF发票: {len(files)} 个文件")
    try:
        pdf_out_dir = save_dir or os.path.dirname(files[0])
        pdf_df = pdf_invoice_parser.parse_pdf_invoices(files, save_dir=pdf_out_dir)
        pdf_df["方向"] = "out"
        if len(pdf_df) > 0:
            existing = ctx.get("out_invoice_df")
            if existing is not None and "发票号码" in existing.columns and "发票号码" in pdf_df.columns:
                existing_nos = set(existing["发票号码"].astype(str))
                new_rows = pdf_df[~pdf_df["发票号码"].astype(str).isin(existing_nos)]
                print(f"  → PDF销项: {len(pdf_df)}张解析, {len(new_rows)}张新增(去重)")
                if len(new_rows) > 0:
                    ctx["out_invoice_df"] = pd.concat([existing, new_rows], ignore_index=True)
            else:
                print(f"  → PDF销项: {len(pdf_df)} 张")
                ctx["out_invoice_df"] = pdf_df
    except Exception as e:
        print(f"  [WARN] PDF销项发票解析失败: {e}")


def _run_property_cert_parser(files: List[str], ctx: Dict) -> None:
    """解析房产证歓取权利人/坐落/面积"""
    print(f"[ToolAgent] 解析房产证: {len(files)} 个文件")
    try:
        items = property_cert_parser.parse(files)
        ctx["property_certs"] = ctx.get("property_certs", []) + items
        print(f"  → 解析到 {len(items)} 份房产证")
    except Exception as e:
        print(f"  [WARN] 房产证解析失败: {e}")


def _run_financial_statement_parser(files: List[str], ctx: Dict) -> None:
    """解析财务报表PDF"""
    print(f"[ToolAgent] 解析财务报表PDF: {len(files)} 个文件")
    try:
        items = financial_statement_parser.parse(files)
        ctx["financial_statements"] = ctx.get("financial_statements", []) + items
        print(f"  → 解析到 {len(items)} 份报表")
    except Exception as e:
        print(f"  [WARN] 财务报表解析失败: {e}")


def _run_balance_sheet_parser(files: List[str], ctx: Dict) -> None:
    """解析资产负债表/利润表 XLS"""
    print(f"[ToolAgent] 解析财务报表XLS: {len(files)} 个文件")
    try:
        items = financial_statement_parser.parse(files)
        ctx["financial_statements"] = ctx.get("financial_statements", []) + items
        print(f"  → 解析到 {len(items)} 份报表")
    except Exception as e:
        print(f"  [WARN] 财务表解析失败: {e}")


def _run_tax_cert_parser(files: List[str], ctx: Dict) -> None:
    """解析完税证明"""
    print(f"[ToolAgent] 解析完税证明: {len(files)} 个文件")
    try:
        items = tax_cert_parser.parse(files)
        ctx["tax_certs"] = ctx.get("tax_certs", []) + items
        print(f"  → 解析到 {len(items)} 份证明")
    except Exception as e:
        print(f"  [WARN] 完税证明解析失败: {e}")


# 工具名 → 执行函数 映射
_TOOL_DISPATCH = {
    "bank_flow_parser": _run_bank_flow_parser,
    "invoice_parser_in": _run_invoice_parser_in,
    "invoice_parser_out": _run_invoice_parser_out,
    "receivable_parser": _run_receivable_parser,
    "payable_parser": _run_payable_parser,
    "credit_report_parser": _run_credit_report_parser,
    "pdf_invoice_parser_in": _run_pdf_invoice_parser_in,
    "pdf_invoice_parser_out": _run_pdf_invoice_parser_out,
    "property_cert_parser": _run_property_cert_parser,
    "financial_statement_parser": _run_financial_statement_parser,
    "balance_sheet_parser": _run_balance_sheet_parser,
    "tax_cert_parser": _run_tax_cert_parser,
}


def _empty_df(columns=None) -> pd.DataFrame:
    return pd.DataFrame(columns=columns or [])


def execute(execution_plan: Dict[str, Any], pdf_ocr_dir: Optional[str] = None) -> Dict[str, Any]:
    """
    Tool Agent 主入口：按 Brain Agent 给出的计划逐步执行工具。

    Args:
        execution_plan: brain_agent.plan() 的返回值
        pdf_ocr_dir: PDF识别结果存储目录

    Returns:
        parsed_data: {
            "flow_df": pd.DataFrame,        # 对公流水（合并）
            "in_invoice_df": pd.DataFrame,  # 进项发票
            "out_invoice_df": pd.DataFrame, # 销项发票
            "recv_df": pd.DataFrame,        # 应收账款
            "pay_df": pd.DataFrame,         # 应付账款
            "credit_data": dict,            # 征信数据
        }
    """
    # 初始化上下文（存储各工具的输出）
    ctx: Dict[str, Any] = {
        "flow_df": _empty_df(),
        "in_invoice_df": _empty_df(),
        "out_invoice_df": _empty_df(),
        "recv_df": _empty_df(),
        "pay_df": _empty_df(),
        "credit_data": {},
        "property_certs": [],
        "financial_statements": [],
        "tax_certs": [],
        "skipped_files": [],  # 被扩展名过滤掉的文件
    }

    steps = execution_plan.get("plan", [])
    total = len(steps)
    print(f"[ToolAgent] 开始执行计划，共 {total} 步")

    for i, step in enumerate(steps, 1):
        tool_name = step.get("tool", "skip")
        files = step.get("files", [])
        reason = step.get("reason", "")

        print(f"\n[ToolAgent] Step {i}/{total}: {tool_name}")
        if reason:
            print(f"  原因: {reason}")

        if tool_name == "skip" or not files:
            print(f"  跳过")
            continue

        fn = _TOOL_DISPATCH.get(tool_name)
        if fn is None:
            print(f"  [WARN] 未知工具: {tool_name}，跳过")
            continue

        # 按工具过滤文件扩展名（防止 LLM 将 PDF 错误地归给 Excel 解析器等）
        _TOOL_ACCEPTS = {
            "bank_flow_parser":           {".xls", ".xlsx", ".csv", ".pdf"},
            "invoice_parser_in":          {".xls", ".xlsx"},
            "invoice_parser_out":         {".xls", ".xlsx"},
            "receivable_parser":          {".xls", ".xlsx"},
            "payable_parser":             {".xls", ".xlsx"},
            "credit_report_parser":       {".pdf"},
            "pdf_invoice_parser_in":      {".pdf"},
            "pdf_invoice_parser_out":     {".pdf"},
            "property_cert_parser":       {".pdf"},
            "financial_statement_parser": {".pdf"},
            "balance_sheet_parser":       {".xls", ".xlsx"},
            "tax_cert_parser":            {".pdf"},
        }
        accepted_exts = _TOOL_ACCEPTS.get(tool_name)
        if accepted_exts:
            filtered = [f for f in files if os.path.splitext(f)[1].lower() in accepted_exts]
            ext_skipped = [f for f in files if f not in filtered]
            if ext_skipped:
                print(f"  [过滤] 跳过不支持格式的文件: {[os.path.basename(f) for f in ext_skipped]}")
                for f in ext_skipped:
                    ctx["skipped_files"].append({
                        "file": os.path.basename(f),
                        "reason": f"该工具不支持此格式（将 {os.path.splitext(f)[1]} 分配给 {tool_name}）",
                    })
            files = filtered
            if not files:
                print(f"  无可处理文件，跳过此步骤")
                continue

        # PDF 工具需要传 save_dir
        if tool_name in ("pdf_invoice_parser_in", "pdf_invoice_parser_out") and pdf_ocr_dir:
            sub = "进项" if tool_name == "pdf_invoice_parser_in" else "销项"
            save_dir = os.path.join(pdf_ocr_dir, sub)
            os.makedirs(save_dir, exist_ok=True)
            fn(files, ctx, save_dir=save_dir)
        else:
            fn(files, ctx)


    print(f"\n[ToolAgent] 执行完毕")
    print(f"  流水记录数: {len(ctx['flow_df'])}")
    print(f"  进项发票数: {len(ctx['in_invoice_df'])}")
    print(f"  销项发票数: {len(ctx['out_invoice_df'])}")
    print(f"  应收条目数: {len(ctx['recv_df'])}")
    print(f"  应付条目数: {len(ctx['pay_df'])}")

    return ctx
