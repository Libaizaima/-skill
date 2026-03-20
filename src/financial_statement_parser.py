# -*- coding: utf-8 -*-
"""
财务报表解析器 — 支持 PDF 格式财务报表 和 XLS 格式资产负债表/利润表。

PDF 格式（如税务申报系统打印件）：同一PDF内含资产负债表+利润表
XLS 格式（如Excel导出）：资产负债表 sheet 或 利润表 sheet

输出统一格式（列表，每个元素为一份报表）：
[
  {
    "期间": "2025-12",
    "报表类型": "资产负债表",   # 或 "利润表"
    "编制单位": str,
    "来源文件": str,
    "科目": [
      {"名称": "货币资金", "期末数": 293975.83, "年初数": 277395.30},
      ...
    ]
  },
  ...
]
"""

import os
import sys
import re
from typing import List, Dict, Any, Optional

_src_dir = os.path.dirname(os.path.abspath(__file__))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------
def _parse_number(s: Any) -> Optional[float]:
    """将字符串清洗为浮点数"""
    if s is None or str(s).strip() in ("", "—", "-", "nan", "None"):
        return None
    cleaned = re.sub(r"[,，\s]", "", str(s)).replace("(", "-").replace("（", "-").replace(")", "").replace("）", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _detect_period(filename: str, header_text: str = "") -> str:
    """从文件名或表头文字中推断报表期间，格式 YYYY-MM"""
    # 文件名中找 202512 / 202601 格式
    m = re.search(r"(20\d{2})(\d{2})", filename)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    # 表头文字中找 "2025-12-31" 或 "2025年12月"
    m = re.search(r"(20\d{2})[.\-年](\d{1,2})", header_text)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}"
    return ""


# ---------------------------------------------------------------------------
# XLS / XLSX 解析
# ---------------------------------------------------------------------------
# 需要提取的资产负债表科目白名单（顺序不重要，按名称匹配）
_BS_KEY_ITEMS = {
    "资产": [
        "货币资金", "应收账款", "预付款项", "存货", "其他应收款",
        "流动资产合计", "固定资产", "非流动资产合计", "资产总计",
    ],
    "负债+权益": [
        "短期借款", "应付账款", "预收款项", "应付职工薪酬", "应交税费",
        "其他应付款", "流动负债合计", "非流动负债合计", "负债合计",
        "实收资本", "未分配利润", "所有者权益合计", "负债和所有者权益合计",
    ],
}
_ALL_BS_ITEMS = {item for group in _BS_KEY_ITEMS.values() for item in group}

_PL_KEY_ITEMS = [
    "营业收入", "营业成本", "税金及附加", "销售费用", "管理费用",
    "财务费用", "营业利润", "利润总额", "所得税费用", "净利润",
]


def _parse_xls_sheet(fpath: str) -> List[Dict[str, Any]]:
    """从 XLS/XLSX 中提取资产负债表和利润表"""
    import xlrd

    results = []
    fname = os.path.basename(fpath)

    try:
        wb = xlrd.open_workbook(fpath)
    except Exception as e:
        print(f"  [financial_statement_parser] 无法打开 {fname}: {e}")
        return results

    for sn in wb.sheet_names():
        ws = wb.sheet_by_name(sn)
        if ws.nrows < 3:
            continue

        sn_lower = sn.strip()
        if "资产负债" in sn_lower:
            rtype = "资产负债表"
            key_items = _ALL_BS_ITEMS
        elif "利润" in sn_lower:
            rtype = "利润表"
            key_items = set(_PL_KEY_ITEMS)
        else:
            # 未知sheet，尝试根据内容判断
            header_row = " ".join(str(ws.cell_value(0, c)) for c in range(ws.ncols))
            if "货币资金" in " ".join(str(ws.cell_value(r, 0)) for r in range(ws.nrows)):
                rtype = "资产负债表"
                key_items = _ALL_BS_ITEMS
            else:
                continue

        # 读取表头（期间）
        header_text = " ".join(str(ws.cell_value(r, c)) for r in range(3) for c in range(ws.ncols))
        period = _detect_period(fname, header_text)

        # 猜测编制单位
        unit = ""
        for r in range(min(3, ws.nrows)):
            row_text = str(ws.cell_value(r, 0)).strip()
            if row_text and "表" not in row_text and "资产" not in row_text:
                unit = row_text
                break

        # 找期末数和年初数列（左侧资产：通常col2=期末数, col3=年初数；右侧负债：col6/col7）
        col_end_l, col_start_l = 2, 3   # 左半部分（资产）
        col_name_r, col_end_r, col_start_r = 4, 6, 7  # 右半部分（负债/权益）
        for r in range(min(5, ws.nrows)):
            for c in range(ws.ncols):
                cell = str(ws.cell_value(r, c)).strip()
                if "期末" in cell and c <= 4:
                    col_end_l = c
                if ("年初" in cell or "上年" in cell) and c <= 5:
                    col_start_l = c

        # 提取科目数据（同时读左右两列组）
        subjects = []
        seen_names = set()

        def _add_subject(name: str, ecol: int, scol: int) -> None:
            name = name.strip().rstrip("：:")
            if not name or name in seen_names:
                return
            if not any(k in name for k in key_items):
                return
            end_val = _parse_number(ws.cell_value(r, ecol) if ecol < ws.ncols else None)
            start_val = _parse_number(ws.cell_value(r, scol) if scol < ws.ncols else None)
            if end_val is None and start_val is None:
                return
            seen_names.add(name)
            subjects.append({"名称": name, "期末数": end_val, "年初数": start_val})

        for r in range(ws.nrows):
            left_name = str(ws.cell_value(r, 0)).strip()
            if left_name and left_name not in ("资产", "项目", "流动资产：", "非流动资产："):
                _add_subject(left_name, col_end_l, col_start_l)

            # 右侧负债/权益列（仅资产负债表格式有8列）
            if rtype == "资产负债表" and ws.ncols >= 7:
                right_name = str(ws.cell_value(r, col_name_r)).strip() if col_name_r < ws.ncols else ""
                if right_name and right_name not in ("负债和所有者（或股东）权益", "流动负债：", "非流动负债：",
                                                      "所有者权益（或股东权益）：", "行次"):
                    _add_subject(right_name, col_end_r, col_start_r)

        if subjects:
            results.append({
                "期间": period,
                "报表类型": rtype,
                "编制单位": unit,
                "来源文件": fname,
                "科目": subjects,
            })

    return results


# ---------------------------------------------------------------------------
# PDF 解析
# ---------------------------------------------------------------------------
# 财务报表PDF中的科目+金额对正则（中文名_空白_数字，数字含逗号和小数）
_AMOUNT_RE = re.compile(r"(-?[\d,，]+\.?\d*)")


def _parse_pdf_page_text(text: str, rtype_hint: str, period: str, unit: str, fname: str) -> Optional[Dict[str, Any]]:
    """从单页文字中提取一种报表的科目列表"""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    subjects = []

    key_items = _ALL_BS_ITEMS if rtype_hint == "资产负债表" else set(_PL_KEY_ITEMS)

    for line in lines:
        # 跳过纯表头行
        if line in ("资产", "负债和所有者权益（或股东权益）", "项目", "行次"):
            continue
        # 找科目名
        matched_item = next((k for k in key_items if k in line), None)
        if not matched_item:
            continue
        # 提取该行内所有数字
        nums = [_parse_number(m) for m in _AMOUNT_RE.findall(line)]
        nums = [n for n in nums if n is not None]
        if not nums:
            continue
        subjects.append({
            "名称": matched_item,
            "期末数": nums[0] if len(nums) >= 1 else None,
            "年初数": nums[1] if len(nums) >= 2 else None,
        })

    if not subjects:
        return None

    return {
        "期间": period,
        "报表类型": rtype_hint,
        "编制单位": unit,
        "来源文件": fname,
        "科目": subjects,
    }


def _parse_pdf(fpath: str) -> List[Dict[str, Any]]:
    """从财务报表 PDF 提取报表数据"""
    import pdfplumber

    results = []
    fname = os.path.basename(fpath)
    period = _detect_period(fname)

    try:
        with pdfplumber.open(fpath) as pdf:
            full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        print(f"  [financial_statement_parser] PDF 读取失败: {e}")
        return results

    # 提取编制单位和期间
    m_unit = re.search(r"编制单位[：:]\s*(.+?)(?:\s{2,}|报送|$)", full_text)
    unit = m_unit.group(1).strip() if m_unit else ""
    if not period:
        m_period = re.search(r"税款所属期[^：:]*[：:]\s*(\d{4})-(\d{2})", full_text)
        if m_period:
            period = f"{m_period.group(1)}-{m_period.group(2)}"

    # 按大块分割资产负债表 / 利润表
    # 策略：找到"资产负债表"标题块 和 "利润表"标题块
    bs_match = re.search(r"资产负债表[\s\S]{0,5000}?(?=利润表|损益表|$)", full_text)
    pl_match = re.search(r"(?:利润表|损益表)[\s\S]+", full_text)

    if bs_match:
        stmt = _parse_pdf_page_text(bs_match.group(), "资产负债表", period, unit, fname)
        if stmt:
            results.append(stmt)

    if pl_match:
        stmt = _parse_pdf_page_text(pl_match.group(), "利润表", period, unit, fname)
        if stmt:
            results.append(stmt)

    # 如果没提取到任何科目，尝试全文作为资产负债表
    if not results:
        stmt = _parse_pdf_page_text(full_text, "资产负债表", period, unit, fname)
        if stmt:
            results.append(stmt)

    return results


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------
def parse(file_paths: List[str]) -> List[Dict[str, Any]]:
    """
    解析财务报表文件（PDF或XLS）。

    Args:
        file_paths: 文件路径列表（可混合PDF和XLS）

    Returns:
        报表数据列表，每元素包含 期间/报表类型/编制单位/科目列表
    """
    all_results = []
    for fpath in file_paths:
        fname = os.path.basename(fpath)
        ext = os.path.splitext(fname)[1].lower()
        print(f"  → 解析财务报表: {fname}")
        try:
            if ext in (".xls", ".xlsx"):
                items = _parse_xls_sheet(fpath)
            elif ext == ".pdf":
                items = _parse_pdf(fpath)
            else:
                print(f"    [SKIP] 不支持的格式: {ext}")
                continue

            for it in items:
                print(f"    {it['报表类型']} ({it['期间']}): {len(it['科目'])} 个科目")
            all_results.extend(items)
        except Exception as e:
            print(f"    [WARN] 解析失败: {e}")

    return all_results
