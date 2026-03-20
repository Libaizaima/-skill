# -*- coding: utf-8 -*-
"""
工具注册表 — 将所有解析脚本注册为可供 Brain Agent 规划的工具。

每个工具包含：
  - name:        工具唯一标识
  - description: 给大模型看的工具说明（中文）
  - accepts:     可接受的文件后缀列表
  - keywords:    文件名/目录名中常见的关键词（辅助大脑判断）
"""

from typing import Dict, Any, List

# ---------------------------------------------------------------------------
# 工具元信息注册表（给 Brain Agent / LLM 看）
# ---------------------------------------------------------------------------
TOOL_SCHEMAS: List[Dict[str, Any]] = [
    {
        "name": "bank_flow_parser",
        "description": (
            "解析对公银行流水文件（Excel/XLS格式）。"
            "适用于：工行、中行、建行、农行、光大等银行的流水明细、对账单、日记账。"
            "典型文件名包含：流水、流水明细、对公、账户、对账单、日记账、基本户。"
        ),
        "accepts": [".xls", ".xlsx", ".csv"],
        "keywords": ["流水", "对公", "账户", "对账单", "日记账", "基本户", "明细"],
    },
    {
        "name": "invoice_parser_in",
        "description": (
            "解析进项发票汇总表（Excel/XLS格式）。"
            "适用于：供应商开具给本公司的增值税发票明细，文件通常在进项目录下。"
            "典型文件名包含：进项、进项票、进项发票。"
        ),
        "accepts": [".xls", ".xlsx"],
        "keywords": ["进项", "进项票", "进项发票"],
    },
    {
        "name": "invoice_parser_out",
        "description": (
            "解析销项发票汇总表（Excel/XLS格式）。"
            "适用于：本公司开具给客户的增值税发票明细，文件通常在销项目录下。"
            "典型文件名包含：销项、销项票、销项发票、开票。"
        ),
        "accepts": [".xls", ".xlsx"],
        "keywords": ["销项", "销项票", "销项发票", "开票"],
    },
    {
        "name": "receivable_parser",
        "description": (
            "解析应收账款明细表（Excel格式）。"
            "适用于：客户欠款汇总，文件名包含应收。"
        ),
        "accepts": [".xls", ".xlsx"],
        "keywords": ["应收"],
    },
    {
        "name": "payable_parser",
        "description": (
            "解析应付账款明细表（Excel格式）。"
            "适用于：对供应商的欠款汇总，文件名包含应付。"
        ),
        "accepts": [".xls", ".xlsx"],
        "keywords": ["应付"],
    },
    {
        "name": "credit_report_parser",
        "description": (
            "解析企业征信报告或法人个人征信报告（PDF格式）。"
            "典型文件名包含：征信、信用报告。"
        ),
        "accepts": [".pdf"],
        "keywords": ["征信", "信用报告"],
    },
    {
        "name": "pdf_invoice_parser_in",
        "description": (
            "解析进项 PDF 发票原件。适用于单张或多张 PDF 格式的增值税电子发票（进项）。"
            "文件通常在进项子目录下，后缀为 .pdf。"
        ),
        "accepts": [".pdf"],
        "keywords": ["进项"],
    },
    {
        "name": "pdf_invoice_parser_out",
        "description": (
            "解析销项 PDF 发票原件。适用于单张或多张 PDF 格式的增值税电子发票（销项）。"
            "文件通常在销项子目录下，后缀为 .pdf。"
        ),
        "accepts": [".pdf"],
        "keywords": ["销项"],
    },
    {
        "name": "tax_cert_parser",
        "description": (
            "解析完税证明PDF（税务机关打印件）。"
            "提取：纳税人名称、税款明细（税种/时期/金额）、合计金额。"
            "文件名通常包含：完税证明、税收完税。"
        ),
        "accepts": [".pdf"],
        "keywords": ["完税证明", "纳税"],
    },
    {
        "name": "property_cert_parser",
        "description": (
            "解析房产证PDF（含扫描件）。"
            "提取：权利人姓名、坐落位置、建筑面积。"
            "文件名通常包含：房产证、不动产证、产权证。"
        ),
        "accepts": [".pdf"],
        "keywords": ["房产证", "不动产证", "产权证"],
    },
    {
        "name": "financial_statement_parser",
        "description": (
            "解析财务报表PDF（资产负债表+利润表合并PDF）。"
            "适用于税务申报系统导出的会企01/02表PDF文件。"
            "文件名通常包含：财务报表、资产负债、利润。"
        ),
        "accepts": [".pdf"],
        "keywords": ["财务报表", "资产负债", "利润表", "会企"],
    },
    {
        "name": "balance_sheet_parser",
        "description": (
            "解析资产负债表或利润表 Excel 文件（XLS/XLSX格式）。"
            "适用于Excel格式的月度/季度资产负债表和利润表。"
            "文件名通常包含：资产负债表、利润表、损益表。"
        ),
        "accepts": [".xls", ".xlsx"],
        "keywords": ["资产负债表", "利润表", "损益表", "资产负债"],
    },
    {
        "name": "skip",
        "description": (
            "跳过此文件，不作解析。"
            "适用于：无关文件（图片、说明文档、已生成的报告等）。"
        ),
        "accepts": ["*"],
        "keywords": [],
    },
]

# 快速查找：name → schema
TOOL_SCHEMA_MAP: Dict[str, Dict[str, Any]] = {t["name"]: t for t in TOOL_SCHEMAS}


def get_tool_descriptions_for_prompt() -> str:
    """
    生成给 Brain Agent 的工具说明文本（嵌入 LLM Prompt）。
    """
    lines = []
    for t in TOOL_SCHEMAS:
        if t["name"] == "skip":
            lines.append('- "{}": {}'.format(t["name"], t["description"]))
        else:
            exts = ", ".join(t["accepts"])
            lines.append('- "{}": {}（支持格式: {}）'.format(t["name"], t["description"], exts))
    return "\n".join(lines)
