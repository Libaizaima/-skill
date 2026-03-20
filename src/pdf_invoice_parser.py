# -*- coding: utf-8 -*-
"""PDF 发票解析模块 — 使用 pdfplumber 从电子发票 PDF 提取结构化数据"""

import os
import pdfplumber
import re
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Optional
import warnings

warnings.filterwarnings('ignore', category=UserWarning)


def parse_pdf_invoices(pdf_paths: List[str], save_dir: Optional[str] = None) -> pd.DataFrame:
    """
    批量解析 PDF 发票文件。

    Args:
        pdf_paths: PDF 文件路径列表
        save_dir: 可选，保存 PDF 识别原始文本的目录

    Returns:
        DataFrame 列: ['开票日期', '发票号码', '销方名称', '购买方名称',
                       '金额', '税额', '价税合计', '发票状态', '发票类型']
    """
    records = []
    for fpath in pdf_paths:
        try:
            rec = _parse_single_pdf(fpath, save_dir=save_dir)
            if rec:
                records.append(rec)
        except Exception as e:
            print(f"[WARN] 解析PDF发票失败: {fpath.split('/')[-1]}, {e}")

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)

    # 转换日期
    df['开票日期'] = df['开票日期'].apply(_parse_date)

    # 转换金额
    for col in ['金额', '税额', '价税合计']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    # 默认状态为正常
    df['发票状态'] = '正常'

    return df


def _parse_single_pdf(fpath: str, save_dir: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """解析单个 PDF 发票，提取关键字段"""
    with pdfplumber.open(fpath) as pdf:
        if not pdf.pages:
            return None

        page = pdf.pages[0]
        text = page.extract_text() or ''
        tables = page.extract_tables() or []

        if len(text) < 50:
            return None

    # 保存 PDF 识别的原始文本
    if save_dir and text:
        txt_name = os.path.splitext(os.path.basename(fpath))[0] + '.txt'
        txt_path = os.path.join(save_dir, txt_name)
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(f"=== 文件: {os.path.basename(fpath)} ===\n\n")
            f.write(f"--- 提取文本 ---\n{text}\n\n")
            if tables:
                f.write(f"--- 提取表格 ({len(tables)} 个) ---\n")
                for i, table in enumerate(tables):
                    f.write(f"\n表格 {i+1}:\n")
                    for row in table:
                        f.write('  | '.join(str(c) if c else '' for c in row) + '\n')

    rec = {
        '发票号码': '',
        '开票日期': '',
        '购买方名称': '',
        '购方税号': '',
        '销方名称': '',
        '销方税号': '',
        '金额': 0.0,
        '税额': 0.0,
        '价税合计': 0.0,
        '发票类型': '',
        '币种': 'CNY',
        '外币金额': 0.0,
        '汇率': 1.0,
        '_source': fpath.split('/')[-1],
    }

    # ── 发票类型 ──
    if '专用发票' in text:
        rec['发票类型'] = '增值税专用发票'
    elif '普通发票' in text:
        rec['发票类型'] = '增值税普通发票'
    else:
        rec['发票类型'] = '其他'

    # ── 发票号码 ──
    m = re.search(r'发票号码[：:]\s*(\d+)', text)
    if m:
        rec['发票号码'] = m.group(1)

    # ── 开票日期 ──
    m = re.search(r'开票日期[：:]\s*(\d{4}年\d{2}月\d{2}日)', text)
    if m:
        rec['开票日期'] = m.group(1)
    else:
        m = re.search(r'开票日期[：:]\s*([\d\-/]+)', text)
        if m:
            rec['开票日期'] = m.group(1)

    # ── 购方/销方 ──
    # 电子发票常见布局: "购 名称：xxx 销 名称：yyy"
    # 或 "购买方名称：xxx" "销售方名称：yyy"
    lines = text.split('\n')
    for line in lines:
        # 购方名称
        m = re.search(r'(?:购买方|购\s*方?)\s*名称[：:]\s*(.+?)(?:\s+销|\s*$)', line)
        if m and not rec['购买方名称']:
            rec['购买方名称'] = m.group(1).strip()

        # 销方名称
        m = re.search(r'(?:销售方|销\s*方?)\s*名称[：:]\s*(.+?)(?:\s*$)', line)
        if m and not rec['销方名称']:
            rec['销方名称'] = m.group(1).strip()

        # 购方税号
        m = re.search(r'(?:购|买).*?(?:统一社会信用代码|纳税人识别号)[：:]\s*(\S+)', line)
        if m and not rec['购方税号']:
            rec['购方税号'] = m.group(1).strip()

        # 销方税号
        m = re.search(r'(?:销|售).*?(?:统一社会信用代码|纳税人识别号)[：:]\s*(\S+)', line)
        if m and not rec['销方税号']:
            rec['销方税号'] = m.group(1).strip()

    # ── 金额/税额/价税合计 ──
    # 方式1: 从文本匹配 "合计 ¥xxx  ¥xxx"
    # 方式2: 从文本匹配 "价税合计（大写）... ¥xxx"
    # 方式3: 从表格提取

    # 价税合计（最可靠）
    m = re.search(r'价税合计[（(]大写[）)]\s*.*?[¥￥]\s*([\d,]+\.?\d*)', text)
    if m:
        rec['价税合计'] = m.group(1).replace(',', '')
    else:
        # 小写金额
        m = re.search(r'价税合计.*?(\d[\d,]*\.?\d*)\s*$', text, re.MULTILINE)
        if m:
            rec['价税合计'] = m.group(1).replace(',', '')

    # 合计金额和税额 — 找 "合计" 行
    for line in lines:
        if line.strip().startswith('合') and '价税' not in line:
            # "合 计 ¥707.96 ¥92.04" 或 "合计 3000.00 390.00"
            amounts = re.findall(r'[¥￥]?\s*([\d,]+\.?\d+)', line)
            if len(amounts) >= 2:
                rec['金额'] = amounts[-2].replace(',', '')
                rec['税额'] = amounts[-1].replace(',', '')
            elif len(amounts) == 1:
                rec['金额'] = amounts[0].replace(',', '')

    # 如果没找到合计行，从表格找
    if not rec['金额'] and tables:
        for table in tables:
            for row in table:
                if not row:
                    continue
                row_str = [str(c).strip() if c else '' for c in row]
                if '合计' in ' '.join(row_str) and '价税' not in ' '.join(row_str):
                    amounts = []
                    for cell in row_str:
                        cell_clean = cell.replace('¥', '').replace('￥', '').replace(',', '').strip()
                        try:
                            v = float(cell_clean)
                            amounts.append(cell_clean)
                        except:
                            pass
                    if len(amounts) >= 2:
                        rec['金额'] = amounts[-2]
                        rec['税额'] = amounts[-1]
                    elif len(amounts) == 1:
                        rec['金额'] = amounts[0]

    # 如果有价税合计但没有金额/税额，尝试推导
    if rec['价税合计'] and not rec['金额']:
        try:
            ptotal = float(str(rec['价税合计']).replace(',', ''))
            # 尝试从表格或文本找税率
            tax_rate = _detect_tax_rate(text)
            if tax_rate > 0:
                amount = ptotal / (1 + tax_rate)
                rec['金额'] = f"{amount:.2f}"
                rec['税额'] = f"{ptotal - amount:.2f}"
        except:
            pass

    # 如果有金额和税额但没有价税合计
    if rec['金额'] and rec['税额'] and not rec['价税合计']:
        try:
            rec['价税合计'] = str(float(str(rec['金额']).replace(',', '')) +
                                float(str(rec['税额']).replace(',', '')))
        except:
            pass

    # ── 币种信息（从备注提取） ──
    currency_map = {
        '美元': 'USD', '美金': 'USD', 'USD': 'USD',
        '欧元': 'EUR', 'EUR': 'EUR',
        '英镑': 'GBP', 'GBP': 'GBP',
        '日元': 'JPY', 'JPY': 'JPY',
        '港币': 'HKD', '港元': 'HKD', 'HKD': 'HKD',
    }
    m_cur = re.search(r'币别[：:]\s*(\S+)', text)
    if m_cur:
        cur_raw = m_cur.group(1).strip('、，, ')
        code = currency_map.get(cur_raw, '')
        if code:
            rec['币种'] = code
            # 外币金额
            m_amt = re.search(r'外币.*?(?:销售额|金额)[：:]?\s*([\d,.]+)', text)
            if m_amt:
                try:
                    rec['外币金额'] = float(m_amt.group(1).replace(',', ''))
                except ValueError:
                    pass
            # 汇率
            m_rate = re.search(r'汇率[：:]?\s*([\d.]+)', text)
            if m_rate:
                try:
                    rec['汇率'] = float(m_rate.group(1))
                except ValueError:
                    pass

    return rec


def _detect_tax_rate(text: str) -> float:
    """从文本中检测税率"""
    m = re.search(r'(\d+)%', text)
    if m:
        rate = int(m.group(1))
        if rate in [1, 3, 6, 9, 13]:
            return rate / 100.0
    return 0.0


def _parse_date(val) -> Optional[datetime]:
    """解析日期字符串"""
    if val is None or val == '':
        return None
    if isinstance(val, datetime):
        return val

    val_str = str(val).strip()
    for fmt in ['%Y年%m月%d日', '%Y-%m-%d', '%Y/%m/%d']:
        try:
            return datetime.strptime(val_str, fmt)
        except ValueError:
            continue
    return None
