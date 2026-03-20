# -*- coding: utf-8 -*-
"""
完税证明解析器 — 支持 PDF 格式（税务机关打印导出件）。

提取信息：
  - 纳税人名称
  - 纳税人识别号
  - 证明编号
  - 填发日期
  - 税款明细：[{"税种", "税款所属时期", "入库日期", "实缴金额"}]
  - 金额合计

每个 PDF 可能包含多页（每页/每份为独立完税证明）。
返回 List[Dict]，每元素对应一份证明。
"""

import os
import sys
import re
from typing import List, Dict, Any, Optional

_src_dir = os.path.dirname(os.path.abspath(__file__))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _clean(s: str) -> str:
    """去除多余空格"""
    return re.sub(r'\s+', '', s).strip() if s else ''


def _parse_amount(s: str) -> float:
    """解析金额字符串 ¥1,234.56 → 1234.56"""
    s = re.sub(r'[¥,，\s]', '', str(s))
    # 处理负数括号写法
    s = s.replace('(', '-').replace('（', '-').replace(')', '').replace('）', '')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


# ---------------------------------------------------------------------------
# 核心解析逻辑
# ---------------------------------------------------------------------------
# 税种关键词（用于识别明细行）
_TAX_KEYWORDS = [
    '增值税', '企业所得税', '个人所得税', '城市维护建设税',
    '城建税', '教育费附加', '地方教育附加', '印花税', '房产税',
    '土地增值税', '社会保险费', '残疾人就业保障金', '水利建设基金',
    '工会经费',
]


def _parse_page(text: str, fname: str) -> Optional[Dict[str, Any]]:
    """
    从单页（或一份）完税证明文字中提取结构化数据。
    """
    # --- 证明编号 ---
    cert_no = ''
    m = re.search(r'证明\s+(\d{8,})', text)
    if m:
        cert_no = m.group(1)

    # --- 填发日期 ---
    issue_date = ''
    m = re.search(r'填\s*发\s*日\s*期\s+(\d{4}-\d{2}-\d{2})', text)
    if m:
        issue_date = m.group(1)

    # --- 纳税人名称 ---
    taxpayer = ''
    m = re.search(r'纳税人名称\s+(.+?)\s+纳税人识别号', text)
    if m:
        taxpayer = m.group(1).strip()

    # --- 纳税人识别号 ---
    tax_id = ''
    m = re.search(r'纳税人识别号\s+([A-Z0-9]{15,20})', text)
    if m:
        tax_id = m.group(1)

    # --- 金额合计 ---
    total_amount = 0.0
    m = re.search(r'金额合计[（(][^）)]*[）)]\s+[^\s¥]*\s+¥?([\d,.]+)', text)
    if not m:
        m = re.search(r'¥([\d,]+\.\d{2})\s*$', text, re.MULTILINE)
        if not m:
            m = re.search(r'¥([\d,]+\.\d{2})', text.split('金额合计')[-1] if '金额合计' in text else '')
    if m:
        total_amount = _parse_amount(m.group(1))

    # --- 税款明细行 ---
    # 格式：税种 yyyy-MM-dd至yyyy-MM-dd  yyyy-MM-dd  ¥金额
    details = []
    _TAX_RE = '|'.join(re.escape(k) for k in _TAX_KEYWORDS)
    # 提取含税种关键词的行
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        # 必须包含税种关键词
        tax_type = next((k for k in _TAX_KEYWORDS if k in line), None)
        if not tax_type:
            continue
        # 提取日期范围（税款所属时期）
        period_match = re.search(r'(\d{4}-\d{2}-\d{2})至(\d{4}-\d{2}-\d{2})', line)
        if not period_match:
            continue
        period_start = period_match.group(1)
        period_end = period_match.group(2)
        # 提取入库日期（至之后的 yyyy-MM-dd）
        rest = line[period_match.end():]
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', rest)
        pay_date = date_match.group(1) if date_match else ''
        # 提取金额
        amount_match = re.search(r'¥([\d,]+\.?\d*)', rest)
        amount = _parse_amount(amount_match.group(1)) if amount_match else 0.0

        details.append({
            '税种': tax_type,
            '税款所属时期': f'{period_start}至{period_end}',
            '入库日期': pay_date,
            '实缴金额': amount,
        })

    # 如果没有解析到有效内容则跳过
    if not details and not taxpayer:
        return None

    # 若 total_amount 未从合计行提取到，则累加明细
    if total_amount == 0.0 and details:
        total_amount = sum(d['实缴金额'] for d in details)

    return {
        '纳税人名称': taxpayer,
        '纳税人识别号': tax_id,
        '证明编号': cert_no,
        '填发日期': issue_date,
        '税款明细': details,
        '合计金额': total_amount,
        '来源文件': fname,
    }


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------
def parse(file_paths: List[str]) -> List[Dict[str, Any]]:
    """
    解析完税证明 PDF 列表。

    一个 PDF 文件可能包含多份证明（多页合并打印），
    每页提取为独立记录。

    Returns:
        [{"纳税人名称", "证明编号", "填发日期", "税款明细": [...], "合计金额"}, ...]
    """
    try:
        import pdfplumber
    except ImportError:
        print('[WARN] 未安装 pdfplumber，无法解析完税证明')
        return []

    all_results: List[Dict[str, Any]] = []

    for fpath in file_paths:
        fname = os.path.basename(fpath)
        print(f'  → 解析完税证明: {fname}')
        try:
            with pdfplumber.open(fpath) as pdf:
                # 尝试按页解析（每页是独立证明）
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ''
                    if '完税证明' not in text and '实缴' not in text:
                        continue
                    result = _parse_page(text, fname)
                    if result:
                        result['页码'] = i + 1
                        all_results.append(result)
                        n_details = len(result['税款明细'])
                        print(f'    第{i+1}页 — 证明#{result["证明编号"]} '
                              f'({n_details}条明细, 合计¥{result["合计金额"]:.2f})')

        except Exception as e:
            print(f'    [WARN] 解析失败: {e}')

    return all_results
