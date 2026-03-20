# -*- coding: utf-8 -*-
"""流水分类引擎 — 基于关键词匹配将交易分类"""

import pandas as pd
import re
from typing import List, Optional


def classify(df: pd.DataFrame, shareholder_names: List[str] = None,
             company_name: str = '') -> pd.DataFrame:
    """
    为每笔交易添加 '分类' 列。

    分类优先级（从高到低）：
      1. 结息
      2. 交税
      3. 金融借贷
      4. 发薪
      5. 交租
      6. 股东往来（用途含"股东"或对方为已知股东个人）
      7. 公司往来（对方为公司本身，但不含股东/金融关键词）
      8. 经营性有效流水（默认）

    Args:
        df: 流水 DataFrame
        shareholder_names: 已知股东姓名列表
        company_name: 公司名称（用于识别公司往来）
    """
    if shareholder_names is None:
        shareholder_names = []

    df = df.copy()
    df['分类'] = df.apply(
        lambda row: _classify_row(row, shareholder_names, company_name), axis=1
    )
    return df


def _classify_row(row, shareholder_names: List[str], company_name: str) -> str:
    """对单行进行分类"""
    purpose = str(row.get('交易用途', '')).strip()
    summary = str(row.get('摘要', '')).strip()
    counterparty = str(row.get('对方户名', '')).strip()

    # ── 1. 结息 ──
    if '结息' in summary or '结息' in purpose:
        return '结息'

    # ── 2. 交税 ──
    if _match_tax(purpose, summary, counterparty):
        return '交税'

    # ── 3. 金融借贷 ──
    if _match_financial(purpose, summary, counterparty):
        return '金融借贷'

    # ── 4. 股东往来（优先于公司往来，因为对公司自身的"股东还款"也是股东往来）──
    if _match_shareholder(purpose, summary, counterparty, shareholder_names):
        return '股东往来'

    # ── 5. 公司往来（对方是公司本身的非股东/非金融交易）──
    if company_name and _match_company_self(counterparty, company_name, purpose):
        return '公司往来'

    # ── 6. 发薪 ──
    if _match_salary(purpose, summary, counterparty):
        return '发薪'

    # ── 7. 交租 ──
    if _match_rent(purpose, summary, counterparty):
        return '交租'

    # ── 8. 默认 ──
    return '经营性有效流水'


# ──────────── 各分类匹配函数 ────────────

def _match_tax(purpose: str, summary: str, counterparty: str) -> bool:
    """判断是否为交税"""
    tax_purpose_kw = ['税费扣缴', '税费', '缴税', '纳税', '报税', '税款']
    tax_cp_kw = ['税务', '国税', '地税', '国家税务', '税局']

    for kw in tax_purpose_kw:
        if kw in purpose:
            return True
    if '税' in purpose and ('扣' in purpose or '缴' in purpose or '费' in purpose):
        return True
    for kw in tax_cp_kw:
        if kw in counterparty:
            return True
    # 摘要为"公共缴费"且对方含"税务"
    if '公共缴费' in summary:
        for kw in tax_cp_kw:
            if kw in counterparty:
                return True
    return False


def _match_financial(purpose: str, summary: str, counterparty: str) -> bool:
    """判断是否为金融借贷"""
    # 摘要明确为贷款
    if summary in ('贷款放款', '贷款还款'):
        return True
    # 用途关键词
    fin_purpose_kw = ['贷款', '放款', '还贷', '按揭']
    for kw in fin_purpose_kw:
        if kw in purpose:
            return True
    return False


def _match_shareholder(purpose: str, summary: str, counterparty: str,
                       shareholder_names: List[str]) -> bool:
    """判断是否为股东往来"""
    # 用途明确含"股东"
    if '股东' in purpose:
        return True
    # 对方户名是已知股东（个人）
    for name in shareholder_names:
        if name and counterparty == name:
            return True
    return False


def _match_company_self(counterparty: str, company_name: str, purpose: str) -> bool:
    """判断是否为公司往来（对方是公司本身的非股东/非金融交易）"""
    if not counterparty or not company_name:
        return False

    # 提取公司核心名
    core = _extract_core_name(company_name)

    # 对方名称中包含公司核心名 → 属于公司内部往来
    if core and len(core) >= 2 and core in counterparty:
        return True
    if len(company_name) >= 4 and company_name in counterparty:
        return True
    return False


def _extract_core_name(company_name: str) -> str:
    """从公司名中提取核心名称，如 '佛瑞森科技' → '佛瑞森'"""
    # 去掉城市前缀
    for prefix in ['深圳市', '深圳', '上海市', '上海', '北京市', '北京', '广州市', '广州',
                   '东莞市', '东莞', '成都市', '成都', '杭州市', '杭州', '无锡市', '无锡',
                   '南京市', '南京', '苏州市', '苏州', '天津市', '天津', '重庆市', '重庆']:
        if company_name.startswith(prefix):
            company_name = company_name[len(prefix):]
            break
    # 去掉后缀
    for suffix in ['有限公司', '股份有限公司', '有限责任公司', '科技', '技术', '电子']:
        if company_name.endswith(suffix):
            company_name = company_name[:-len(suffix)]
            break
    return company_name


def _match_salary(purpose: str, summary: str, counterparty: str) -> bool:
    """判断是否为发薪"""
    salary_kw = ['工资', '发工资', '薪金', '劳务费', '代发']
    for kw in salary_kw:
        if kw in purpose:
            return True
    if '代发' in summary or '代发工资' in summary:
        return True
    # "社保"和"公积金"也算发薪类
    if '社保' in purpose or '公积金' in purpose:
        return True
    return False


def _match_rent(purpose: str, summary: str, counterparty: str) -> bool:
    """判断是否为交租"""
    rent_kw = ['租金', '租赁', '房租', '场地租', '租车费', '物业管理费']
    for kw in rent_kw:
        if kw in purpose:
            return True
    return False


def _has_company_suffix(name: str) -> bool:
    """判断名称是否像公司名"""
    suffixes = ['公司', '企业', '集团', '有限', '股份', '合伙', '中心', '事务所',
                '机构', '银行', '基金', '证券', '保险', '信托']
    for suffix in suffixes:
        if suffix in name:
            return True
    return False


def detect_shareholder_names(df: pd.DataFrame) -> List[str]:
    """
    从流水中自动检测可能的股东姓名。
    规则：交易用途含"股东"且对方户名为个人名（2-4字，无公司后缀）
    """
    shareholders = set()

    # 从用途含"股东"的记录中提取对方个人名
    mask = df['交易用途'].str.contains('股东', na=False)
    for _, row in df[mask].iterrows():
        name = str(row['对方户名']).strip()
        if len(name) >= 2 and len(name) <= 4 and not _has_company_suffix(name):
            shareholders.add(name)

    # 从用途含"借钱给"或"往来"的记录中提取对方个人名
    mask2 = df['交易用途'].str.contains('借钱给', na=False)
    for _, row in df[mask2].iterrows():
        name = str(row['对方户名']).strip()
        if len(name) >= 2 and len(name) <= 4 and not _has_company_suffix(name):
            shareholders.add(name)

    return list(shareholders)
