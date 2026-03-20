# -*- coding: utf-8 -*-
"""征信报告 PDF 解析模块 — 使用 pdfplumber 提取关键信息"""

import pdfplumber
import re
from typing import Dict, Any, List, Optional


def parse_all(credit_report_paths: List[str]) -> Dict[str, Any]:
    """
    解析所有征信报告 PDF。

    Args:
        credit_report_paths: 征信报告文件路径列表

    Returns:
        {
            'personal': {...}  或 None,   # 个人征信
            'company': {...}   或 None,   # 企业征信
        }
    """
    result = {
        'personal': None,
        'company': None,
    }

    for fpath in credit_report_paths:
        try:
            fname = fpath.split('/')[-1].lower()
            with pdfplumber.open(fpath) as pdf:
                full_text = '\n'.join([
                    page.extract_text() or '' for page in pdf.pages
                ])
                all_tables = []
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        all_tables.extend(tables)

                # 判断类型
                if '企业信用报告' in full_text or '企业征信' in fname:
                    result['company'] = _parse_company_credit(full_text, all_tables)
                    result['company']['_source'] = fpath.split('/')[-1]
                elif '个人信用报告' in full_text or '法人征信' in fname or '个人征信' in fname:
                    result['personal'] = _parse_personal_credit(full_text, all_tables)
                    result['personal']['_source'] = fpath.split('/')[-1]
                else:
                    print(f"[WARN] 无法识别征信报告类型: {fpath}")
        except Exception as e:
            print(f"[WARN] 解析征信报告失败: {fpath}, 错误: {e}")

    return result


def _parse_personal_credit(text: str, tables: list) -> Dict[str, Any]:
    """
    解析个人信用报告。

    提取字段：
    - 基本信息：姓名、证件号码、婚姻状况
    - 信用卡：账户数、未结清数、逾期数
    - 贷款：账户数、未结清数、逾期数
    - 贷款明细：机构、额度、余额、状态
    - 查询记录：近1月/3月/6月查询次数
    """
    info = {
        '姓名': '',
        '证件号码': '',
        '婚姻状况': '',
        '征信时间': '',
        # 信用卡
        '信用卡账户数': '',
        '信用卡未结清数': '',
        '信用卡总额度': '',
        '信用卡已用额度': '',
        '信用卡逾期账户数': '',
        '信用卡90天以上逾期数': '',
        # 贷款
        '贷款账户数': '',
        '贷款未结清数': '',
        '贷款逾期账户数': '',
        '贷款90天以上逾期数': '',
        # 汇总
        '贷款总额度': '',
        '贷款未还总本金': '',
        '月供金额': '',
        # 贷款明细列表
        '贷款明细': [],
        # 查询
        '近1月查询次数': '',
        '近3月查询次数': '',
        '近6月查询次数': '',
        '查询记录': [],
    }

    # 基本信息
    m = re.search(r'姓名[：:]\s*(\S+)', text)
    if m:
        info['姓名'] = m.group(1)

    m = re.search(r'证件号码[：:]\s*(\S+)', text)
    if m:
        info['证件号码'] = m.group(1)

    # 婚姻状况（通常跟在证件号码后面）
    for status in ['已婚', '未婚', '离婚', '丧偶']:
        if status in text[:500]:
            info['婚姻状况'] = status
            break

    m = re.search(r'报告时间[：:]\s*(\S+)', text)
    if m:
        info['征信时间'] = m.group(1)

    # 从表格提取信用卡/贷款概要
    for table in tables:
        if not table or len(table) < 3:
            continue
        # 找信息概要表（含"账户数"/"未结清"行）
        headers = [str(c).strip() if c else '' for c in table[0]]
        if '信用卡' in ' '.join(headers):
            _extract_credit_summary(table, info)

    # 从文本提取信用卡总额度和已用额度
    total_card_limit = 0
    total_card_used = 0
    # 合并文本以处理跨行的"已使用\n额度"
    merged_text = text.replace('\n', ' ')
    # 直接匹配"信用额度X，已使用额度Y"模式
    for m in re.finditer(r'信用额度\s*([\d,]+)\s*[，,]?\s*已使用\s*(?:额\s*度\s*)?([\d,]+)', merged_text):
        total_card_limit += int(m.group(1).replace(',', ''))
        total_card_used += int(m.group(2).replace(',', ''))
    if total_card_limit > 0:
        info['信用卡总额度'] = f"{total_card_limit:,}"
    info['信用卡已用额度'] = f"{total_card_used:,}"

    # 提取贷款明细
    info['贷款明细'] = _extract_loan_details(text)

    # 计算贷款汇总
    if info['贷款明细']:
        total_limit = 0
        total_balance = 0
        for loan in info['贷款明细']:
            if loan.get('额度'):
                try:
                    total_limit += float(str(loan['额度']).replace(',', ''))
                except:
                    pass
            if loan.get('余额'):
                try:
                    total_balance += float(str(loan['余额']).replace(',', ''))
                except:
                    pass
        info['贷款总额度'] = f"{total_limit:,.0f}"
        info['贷款未还总本金'] = f"{total_balance:,.0f}"

        # 推算月供金额（基于未结清贷款余额）
        # 平安银行经营贷为可循环使用额度，按LPR+浮动估算年利率约4.5%-5.5%
        # 可循环经营贷通常按月付息、到期还本，月供=余额×月利率
        monthly_payment = 0
        payment_details = []
        for loan in info['贷款明细']:
            if loan.get('状态') in ('正常', '逾期'):
                balance = 0
                try:
                    balance = float(str(loan.get('余额', '0')).replace(',', ''))
                except:
                    pass
                if balance > 0:
                    bank = loan.get('机构', '')
                    # 经营贷/授信按月付息: 月供 = 余额 × 年利率/12
                    annual_rate = 0.045  # 估算年利率4.5%（LPR 3.85% + 浮动）
                    est_monthly = balance * annual_rate / 12
                    monthly_payment += est_monthly
                    payment_details.append(f"{bank[:8]}...余额{balance:,.0f}×{annual_rate*100:.1f}%/12≈{est_monthly:,.0f}")
        if monthly_payment > 0:
            info['月供金额'] = f"≈{monthly_payment:,.0f}（推算）"
            info['月供推算说明'] = '按可循环经营贷月付息方式估算(年利率约4.5%): ' + '; '.join(payment_details)

    # 提取查询记录
    info['查询记录'] = _extract_query_records(text)
    _count_queries(info)

    return info


def _extract_credit_summary(table: list, info: dict):
    """从信息概要表提取信用卡/贷款数据"""
    # 典型格式:
    # Row 0: ['', '信用卡', '贷款', None, '其他业务']
    # Row 1: [None, None, '购房', '其他', None]
    # Row 2: ['账户数', '21', '1', '4', '--']
    # Row 3: ['未结清/未销户账户数', '15', '--', '1', '--']
    # Row 4: ['发生过逾期的账户数', ...]
    # Row 5: ['发生过90天以上逾期的账户数', ...]

    for row in table:
        if not row:
            continue
        label = str(row[0]).strip() if row[0] else ''

        if '账户数' in label and '未结清' not in label and '逾期' not in label:
            # 信用卡账户数
            if len(row) > 1 and row[1] and str(row[1]).strip() != '--':
                info['信用卡账户数'] = str(row[1]).strip()
            # 贷款账户数（购房+其他）
            loan_count = 0
            for i in [2, 3]:
                if i < len(row) and row[i] and str(row[i]).strip() not in ('--', '', 'None'):
                    try:
                        loan_count += int(str(row[i]).strip())
                    except:
                        pass
            if loan_count > 0:
                info['贷款账户数'] = str(loan_count)

        elif '未结清' in label or '未销户' in label:
            if len(row) > 1 and row[1] and str(row[1]).strip() != '--':
                info['信用卡未结清数'] = str(row[1]).strip()
            loan_open = 0
            for i in [2, 3]:
                if i < len(row) and row[i] and str(row[i]).strip() not in ('--', '', 'None'):
                    try:
                        loan_open += int(str(row[i]).strip())
                    except:
                        pass
            if loan_open > 0:
                info['贷款未结清数'] = str(loan_open)

        elif '90天' in label:
            if len(row) > 1 and row[1]:
                info['信用卡90天以上逾期数'] = str(row[1]).strip()
            loan_90 = '--'
            for i in [2, 3]:
                if i < len(row) and row[i] and str(row[i]).strip() not in ('--', '', 'None'):
                    loan_90 = str(row[i]).strip()
            info['贷款90天以上逾期数'] = loan_90

        elif '逾期' in label:
            if len(row) > 1 and row[1]:
                info['信用卡逾期账户数'] = str(row[1]).strip()
            loan_overdue = '--'
            for i in [2, 3]:
                if i < len(row) and row[i] and str(row[i]).strip() not in ('--', '', 'None'):
                    loan_overdue = str(row[i]).strip()
            info['贷款逾期账户数'] = loan_overdue


def _extract_loan_details(text: str) -> list:
    """
    从文本中提取贷款明细。
    使用2步法：先切分出每笔贷款的完整文本，再从中提取字段。
    """
    loans = []

    # 找到贷款段落（在"贷款"标题与下一个章节标题之间）
    # 常见段落有: "从未发生过逾期的账户明细如下：" 和 "发生过逾期的账户明细如下："
    loan_section = ''
    # 找到所有 "N.日期 银行名 发放/授信" 之间的文本
    lines = text.split('\n')
    in_loan = False
    for line in lines:
        stripped = line.strip()
        # 检测贷款段落开始
        if '账户明细如下' in stripped and '信用卡' not in stripped:
            in_loan = True
            continue
        # 检测段落结束（非信贷交易/公共记录/查询记录）
        if in_loan and any(kw in stripped for kw in ['非信贷交易', '公共记录', '查询记录', '查询明细']):
            in_loan = False
            continue
        if in_loan:
            loan_section += stripped + ' '

    if not loan_section:
        # Fallback: 尝试从全文找
        loan_section = text.replace('\n', ' ')

    # 按序号切分每笔贷款: "1.2025年..." "2.2015年..."
    entries = re.split(r'(?=\d+\.\s*\d{4}年\d{2}月\d{2}日)', loan_section)

    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        # 提取序号和日期
        m = re.match(r'(\d+)\.\s*(\d{4}年\d{2}月\d{2}日)(.+)', entry)
        if not m:
            continue

        date = m.group(2)
        body = m.group(3)

        # 跳过信用卡条目（贷记卡/准贷记卡）
        if '贷记卡' in body:
            continue

        # 提取银行名 — 截到"发放"或"授信"之前，再截到"行/公司/中心"
        bank = ''
        bank_match = re.match(r'(.+?)(?:发放|为|授信)', body)
        if bank_match:
            bank = bank_match.group(1).strip()
            for end_word in ['支行', '分行', '银行', '公司', '中心']:
                idx = bank.find(end_word)
                if idx > 0:
                    bank = bank[:idx + len(end_word)]
                    break

        # 提取金额（发放金额）
        amount = ''
        m_amt = re.search(r'发放的\s*([\d,]+)\s*元', body)
        if m_amt:
            amount = m_amt.group(1).replace(',', '')

        # 提取信用额度
        credit_limit = ''
        m_limit = re.search(r'信用额度\s*([\d,]+)', body)
        if m_limit:
            credit_limit = m_limit.group(1).replace(',', '')

        # 提取余额
        balance = ''
        m_bal = re.search(r'余额[为]?\s*([\d,]+)', body)
        if m_bal:
            balance = m_bal.group(1).replace(',', '')

        # 提取状态
        status = ''
        if '已结清' in body:
            status = '已结清'
        elif '当前无逾期' in body:
            status = '正常'
        elif '逾期' in body:
            status = '逾期'

        # 额度取 credit_limit 或 amount
        final_limit = credit_limit or amount

        loans.append({
            '机构': bank,
            '日期': date,
            '额度': final_limit,
            '余额': balance,
            '状态': status,
        })

    return loans


def _extract_query_records(text: str) -> list:
    """提取查询记录"""
    records = []
    # 匹配: "序号 日期 机构 原因"
    query_pattern = re.compile(
        r'(\d+)\s+(\d{4}年\d{2}月\d{2}日)\s+(.+?)\s+(贷款审批|信用卡审批|贷后管理|本人查询|担保资格审查)',
    )
    for m in query_pattern.finditer(text):
        records.append({
            '日期': m.group(2),
            '机构': m.group(3).strip(),
            '原因': m.group(4),
        })
    return records


def _count_queries(info: dict):
    """统计近1/3/6月查询次数"""
    records = info.get('查询记录', [])
    if not records:
        info['近1月查询次数'] = '0'
        info['近3月查询次数'] = '0'
        info['近6月查询次数'] = '0'
        return

    from datetime import datetime, timedelta
    now = datetime.now()
    count_1m = count_3m = count_6m = 0

    for rec in records:
        try:
            dt = datetime.strptime(rec['日期'], '%Y年%m月%d日')
            days = (now - dt).days
            if days <= 30:
                count_1m += 1
            if days <= 90:
                count_3m += 1
            if days <= 180:
                count_6m += 1
        except:
            pass

    info['近1月查询次数'] = str(count_1m)
    info['近3月查询次数'] = str(count_3m)
    info['近6月查询次数'] = str(count_6m)


def _parse_company_credit(text: str, tables: list) -> Dict[str, Any]:
    """
    解析企业信用报告。

    提取字段：
    - 基本信息：企业名称、统一社会信用代码
    - 信贷概要：借贷余额、担保余额、被追偿余额
    - 贷款明细：机构、额度、未还本金、使用周期
    - 查询记录
    """
    info = {
        '企业名称': '',
        '统一社会信用代码': '',
        '征信时间': '',
        # 信贷概要
        '首次信贷年份': '',
        '信贷机构数': '',
        '未结清机构数': '',
        '借贷余额': '',
        '担保余额': '',
        '被追偿余额': '',
        # 贷款明细
        '贷款明细': [],
        '贷款总额度': '',
        '未还总本金': '',
        '月供总金额': '',
        # 查询记录
        '查询记录': [],
    }

    # 基本信息
    m = re.search(r'企业名称[：:]\s*(.+?)(?:\n|$)', text)
    if m:
        info['企业名称'] = m.group(1).strip()

    m = re.search(r'统一社会信用代码[：:]\s*(\S+)', text)
    if m:
        info['统一社会信用代码'] = m.group(1).strip()

    m = re.search(r'报告时间[：:]\s*(\S+)', text)
    if m:
        info['征信时间'] = m.group(1).strip()

    # 从表格提取
    for table in tables:
        if not table or len(table) < 2:
            continue
        for row in table:
            if not row:
                continue
            label = str(row[0]).strip() if row[0] else ''

            if label == '企业名称' and len(row) > 1:
                info['企业名称'] = str(row[1]).strip()
            elif label == '统一社会信用代码' and len(row) > 1:
                info['统一社会信用代码'] = str(row[1]).strip()
            elif '首次有信贷' in label and len(row) > 1:
                info['首次信贷年份'] = str(row[0]).strip() if row[0] else '--'
                # 这种表格是横向的
                vals = [str(c).strip() if c else '--' for c in row]
                if len(vals) >= 4:
                    info['首次信贷年份'] = vals[0] if vals[0] != '首次有信贷交易的年份' else '--'

    # 从概要表提取 (通常第2个表格)
    for table in tables:
        if not table:
            continue
        headers = [str(c).strip() if c else '' for c in table[0]]
        if '首次有信贷交易的年份' in headers:
            if len(table) > 1:
                vals = table[1]
                if len(vals) >= 4:
                    info['首次信贷年份'] = str(vals[0]).strip() if vals[0] else '--'
                    info['信贷机构数'] = str(vals[1]).strip() if vals[1] else '0'
                    info['未结清机构数'] = str(vals[2]).strip() if vals[2] else '0'

    # 从文本提取信贷概要数字
    m = re.search(r'借贷交易.*?余额\s+(\d+)', text, re.DOTALL)
    if m:
        info['借贷余额'] = m.group(1)
    m = re.search(r'担保交易.*?余额\s+(\d+)', text, re.DOTALL)
    if m:
        info['担保余额'] = m.group(1)
    m = re.search(r'被追偿余额\s+(\d+)', text)
    if m:
        info['被追偿余额'] = m.group(1)

    # 提取查询记录
    info['查询记录'] = _extract_company_queries(text)

    return info


def _extract_company_queries(text: str) -> list:
    """提取企业征信查询记录"""
    records = []
    query_pattern = re.compile(
        r'(\d{4}\.\d{2}\.\d{2})\s+(.+?)\s+(贷款审批|贷后管理|担保资格审查|本机构查询|关联企业查询)',
    )
    for m in query_pattern.finditer(text):
        records.append({
            '日期': m.group(1),
            '机构': m.group(2).strip(),
            '原因': m.group(3),
        })
    return records
