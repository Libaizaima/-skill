# -*- coding: utf-8 -*-
"""统计分析引擎 — 对解析后的数据执行各项统计计算"""

import pandas as pd
from datetime import datetime
from typing import Dict, Any, Optional


def analyze_all(flow_df: pd.DataFrame, in_invoice_df: pd.DataFrame,
                out_invoice_df: pd.DataFrame, recv_df: pd.DataFrame,
                pay_df: pd.DataFrame, company_name: str) -> Dict[str, Any]:
    """
    执行全部分析计算，返回结构化的分析结果字典。
    """
    results = {
        'company_name': company_name,
        'overall': overall_stats(flow_df),
        'monthly': monthly_summary(flow_df),
        'salary': salary_stats(flow_df),
        'rent': rent_stats(flow_df),
        'interest': interest_stats(flow_df),
        'category_summary': flow_category_summary(flow_df),
        'top_income': top_n_counterparties(flow_df, n=10, direction='income'),
        'top_expense': top_n_counterparties(flow_df, n=10, direction='expense'),
        'invoice_in_yearly': invoice_yearly(in_invoice_df),
        'invoice_out_yearly': invoice_yearly(out_invoice_df),
        'invoice_in_monthly': invoice_monthly(in_invoice_df),
        'invoice_out_monthly': invoice_monthly(out_invoice_df),
        'invoice_comparison': invoice_comparison_yearly(in_invoice_df, out_invoice_df),
        'export_stats': export_invoice_stats(out_invoice_df),
        'receivable_top': receivable_top_n(recv_df, n=10),
        'payable_top': payable_top_n(pay_df, n=10),
        'receivable_total': receivable_total(recv_df),
        'payable_total': payable_total(pay_df),
    }
    return results


def overall_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """总体统计: 总收入/支出/净额/笔数/余额范围"""
    if df.empty:
        return {}

    total_income = df['收入金额'].sum()
    total_expense = df['支出金额'].sum()
    net = total_income - total_expense
    total_count = len(df)

    # 按分类统计笔数
    biz_mask = df['分类'] == '经营性有效流水'
    biz_count = biz_mask.sum()
    other_count = total_count - biz_count

    balance_min = df['账户余额'].min()
    balance_max = df['账户余额'].max()
    balance_last = df.iloc[-1]['账户余额'] if len(df) > 0 else 0

    return {
        'total_income': total_income,
        'total_expense': total_expense,
        'net': net,
        'total_count': total_count,
        'biz_count': biz_count,
        'other_count': other_count,
        'balance_min': balance_min,
        'balance_max': balance_max,
        'balance_last': balance_last,
    }


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    """月度收入/支出/净流入汇总表"""
    if df.empty:
        return pd.DataFrame()

    df = df.copy()
    df['月份'] = df['交易时间'].apply(lambda x: x.strftime('%Y-%m') if pd.notna(x) else None)
    df = df.dropna(subset=['月份'])

    monthly = df.groupby('月份').agg(
        收入=('收入金额', 'sum'),
        支出=('支出金额', 'sum'),
    ).reset_index()

    monthly['净流入'] = monthly['收入'] - monthly['支出']

    # 转换为万元
    monthly['收入（万元）'] = (monthly['收入'] / 10000).round(2)
    monthly['支出（万元）'] = (monthly['支出'] / 10000).round(2)
    monthly['净流入（万元）'] = (monthly['净流入'] / 10000).round(2)

    return monthly.sort_values('月份')


def salary_stats(df: pd.DataFrame) -> pd.DataFrame:
    """发薪月度统计（笔数 + 金额）"""
    if df.empty or '分类' not in df.columns:
        return pd.DataFrame()

    salary = df[df['分类'] == '发薪'].copy()
    if salary.empty:
        return pd.DataFrame()

    salary['月份'] = salary['交易时间'].apply(lambda x: x.strftime('%Y-%m') if pd.notna(x) else None)
    salary = salary.dropna(subset=['月份'])

    result = salary.groupby('月份').agg(
        发薪笔数=('支出金额', 'count'),
        发薪金额=('支出金额', 'sum'),
    ).reset_index()

    # 金额保留2位小数（元）
    result['发薪金额（元）'] = result['发薪金额'].round(2)

    return result.sort_values('月份')


def rent_stats(df: pd.DataFrame) -> pd.DataFrame:
    """交租月度统计"""
    if df.empty or '分类' not in df.columns:
        return pd.DataFrame()

    rent = df[df['分类'] == '交租'].copy()
    if rent.empty:
        return pd.DataFrame()

    rent['月份'] = rent['交易时间'].apply(lambda x: x.strftime('%Y-%m') if pd.notna(x) else None)
    rent = rent.dropna(subset=['月份'])

    result = rent.groupby('月份').agg(
        交租笔数=('支出金额', 'count'),
        交租金额=('支出金额', 'sum'),
    ).reset_index()

    result['交租金额（元）'] = result['交租金额'].round(2)

    return result.sort_values('月份')


def interest_stats(df: pd.DataFrame) -> pd.DataFrame:
    """结息记录表"""
    if df.empty or '分类' not in df.columns:
        return pd.DataFrame()

    interest = df[df['分类'] == '结息'].copy()
    if interest.empty:
        return pd.DataFrame()

    result = interest[['交易时间', '对方户名', '收入金额', '支出金额', '摘要']].copy()
    result['交易时间'] = result['交易时间'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(x) else '')
    result.columns = ['交易时间', '对方户名', '收入金额', '支出金额', '摘要']

    return result


def flow_category_summary(df: pd.DataFrame) -> pd.DataFrame:
    """各分类汇总表（收入/支出/净额/笔数）"""
    if df.empty or '分类' not in df.columns:
        return pd.DataFrame()

    result = df.groupby('分类').agg(
        收入=('收入金额', 'sum'),
        支出=('支出金额', 'sum'),
        笔数=('交易时间', 'count'),
    ).reset_index()

    result['净额'] = result['收入'] - result['支出']

    # 排序：经营性有效流水排在最前面
    category_order = ['经营性有效流水', '发薪', '股东往来', '交税', '公司往来', '交租', '金融借贷', '结息']
    result['_order'] = result['分类'].apply(lambda x: category_order.index(x) if x in category_order else 99)
    result = result.sort_values('_order').drop(columns=['_order'])

    return result


def top_n_counterparties(df: pd.DataFrame, n: int = 10, direction: str = 'income') -> pd.DataFrame:
    """前N收入/支出对手方"""
    if df.empty or '分类' not in df.columns:
        return pd.DataFrame()

    # 只统计经营性有效流水
    biz = df[df['分类'] == '经营性有效流水'].copy()
    if biz.empty:
        return pd.DataFrame()

    if direction == 'income':
        col = '收入金额'
        result_col = '经营收入合计'
    else:
        col = '支出金额'
        result_col = '经营支出合计'

    # 过滤掉金额为0的行
    biz_filtered = biz[biz[col] > 0].copy()

    result = biz_filtered.groupby('对方户名').agg(
        合计=(col, 'sum'),
    ).reset_index()

    result.columns = ['对方户名', result_col]
    result = result.sort_values(result_col, ascending=False).head(n)

    return result


def invoice_yearly(df: pd.DataFrame) -> pd.DataFrame:
    """发票年度统计（张数 + 金额）"""
    if df is None or df.empty or '开票日期' not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df['年份'] = df['开票日期'].apply(lambda x: x.year if pd.notna(x) else None)
    df = df.dropna(subset=['年份'])

    # 按发票号去重计数（一张发票可能有多行明细）
    # 使用 数电发票号码 去重计算张数
    has_invoice_no = '数电发票号码' in df.columns
    if has_invoice_no:
        count_df = df.drop_duplicates(subset=['数电发票号码', '年份'])
        count_result = count_df.groupby('年份').agg(
            开票张数=('数电发票号码', 'count'),
        ).reset_index()
    else:
        count_result = df.groupby('年份').agg(
            开票张数=('价税合计', 'count'),
        ).reset_index()

    amount_result = df.groupby('年份').agg(
        开票金额=('价税合计', 'sum'),
    ).reset_index()

    result = count_result.merge(amount_result, on='年份', how='outer')
    result['开票金额（万元）'] = (result['开票金额'] / 10000).round(2)

    return result.sort_values('年份')


def invoice_monthly(df: pd.DataFrame) -> pd.DataFrame:
    """发票月度统计"""
    if df is None or df.empty or '开票日期' not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df['月份'] = df['开票日期'].apply(lambda x: x.strftime('%Y-%m') if pd.notna(x) else None)
    df = df.dropna(subset=['月份'])

    has_invoice_no = '数电发票号码' in df.columns
    if has_invoice_no:
        count_df = df.drop_duplicates(subset=['数电发票号码', '月份'])
        count_result = count_df.groupby('月份').agg(
            开票张数=('数电发票号码', 'count'),
        ).reset_index()
    else:
        count_result = df.groupby('月份').agg(
            开票张数=('价税合计', 'count'),
        ).reset_index()

    amount_result = df.groupby('月份').agg(
        开票金额=('价税合计', 'sum'),
    ).reset_index()

    result = count_result.merge(amount_result, on='月份', how='outer')
    result['开票金额（万元）'] = (result['开票金额'] / 10000).round(2)

    return result.sort_values('月份')


def invoice_comparison_yearly(in_df: pd.DataFrame, out_df: pd.DataFrame) -> pd.DataFrame:
    """进销项年度对比表"""
    in_yearly = invoice_yearly(in_df)
    out_yearly = invoice_yearly(out_df)

    if in_yearly.empty and out_yearly.empty:
        return pd.DataFrame()

    # 合并
    if not in_yearly.empty:
        in_yearly = in_yearly[['年份', '开票金额（万元）']].rename(columns={'开票金额（万元）': '进项金额（万元）'})
    else:
        in_yearly = pd.DataFrame(columns=['年份', '进项金额（万元）'])

    if not out_yearly.empty:
        out_yearly = out_yearly[['年份', '开票金额（万元）']].rename(columns={'开票金额（万元）': '销项金额（万元）'})
    else:
        out_yearly = pd.DataFrame(columns=['年份', '销项金额（万元）'])

    result = pd.merge(in_yearly, out_yearly, on='年份', how='outer').fillna(0)
    result['销-进差额（万元）'] = (result['销项金额（万元）'] - result['进项金额（万元）']).round(2)

    return result.sort_values('年份')


def receivable_top_n(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    """前N应收客户"""
    if df is None or df.empty:
        return pd.DataFrame()

    sort_col = '合计(万元)' if '合计(万元)' in df.columns else '合计(外币)'
    if sort_col not in df.columns:
        return pd.DataFrame()

    result = df.sort_values(sort_col, ascending=False).head(n).copy()
    return result


def payable_top_n(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    """前N应付供应商"""
    if df is None or df.empty:
        return pd.DataFrame()

    sort_col = '合计(万元)' if '合计(万元)' in df.columns else '应付余额'
    if sort_col not in df.columns:
        return pd.DataFrame()

    result = df.sort_values(sort_col, ascending=False).head(n).copy()
    return result


def receivable_total(df: pd.DataFrame) -> Dict[str, Any]:
    """应收汇总"""
    if df is None or df.empty:
        return {'总应收(万元)': 0, '币种': 'CNY'}

    currency = df['币种'].iloc[0] if '币种' in df.columns and len(df) > 0 else 'CNY'
    rate = df['预算汇率'].iloc[0] if '预算汇率' in df.columns and len(df) > 0 else 1

    if '合计(万元)' in df.columns:
        total_wan = round(df['合计(万元)'].sum(), 2)
    elif '合计(外币)' in df.columns:
        total_foreign = df['合计(外币)'].sum()
        total_wan = round(total_foreign * rate / 10000, 2) if rate > 0 else round(total_foreign / 10000, 2)
    else:
        total_wan = 0

    result = {
        '总应收(万元)': total_wan,
        '币种': currency,
        '汇率': rate,
    }
    if '合计(外币)' in df.columns:
        result['总应收(外币)'] = round(df['合计(外币)'].sum(), 2)

    return result


def payable_total(df: pd.DataFrame) -> Dict[str, Any]:
    """应付汇总"""
    if df is None or df.empty:
        return {'总应付(万元)': 0}

    if '合计(万元)' in df.columns:
        total = round(df['合计(万元)'].sum(), 2)
    else:
        total = 0

    return {'总应付(万元)': total}


def export_invoice_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """
    出口发票统计：按币种分组汇总外币金额和RMB折算金额。

    Returns:
        {
            'has_export': True/False,
            'summary': [{币种, 张数, 外币金额合计, RMB金额合计, 平均汇率}, ...],
            'total_rmb': 全部出口发票RMB金额,
            'total_rmb_all': 全部销项票RMB金额,
            'export_ratio': 出口占比,
        }
    """
    result = {
        'has_export': False,
        'summary': [],
        'total_rmb': 0,
        'total_rmb_all': 0,
        'export_ratio': 0,
    }

    if df is None or df.empty or '币种' not in df.columns:
        return result

    # 筛选非 CNY 记录
    export_df = df[df['币种'] != 'CNY'].copy()
    if export_df.empty:
        return result

    result['has_export'] = True
    result['total_rmb_all'] = round(df['价税合计'].sum(), 2) if '价税合计' in df.columns else 0

    # 按币种分组
    for currency, group in export_df.groupby('币种'):
        foreign_total = group['外币金额'].sum()
        rmb_total = group['价税合计'].sum() if '价税合计' in group.columns else 0
        avg_rate = group['汇率'].mean()
        # 按数电发票号码去重计算张数
        if '数电发票号码' in group.columns:
            count = group['数电发票号码'].nunique()
        elif '发票号码' in group.columns:
            count = group['发票号码'].nunique()
        else:
            count = len(group)

        result['summary'].append({
            '币种': currency,
            '张数': int(count),
            '外币金额合计': round(foreign_total, 2),
            'RMB金额合计': round(rmb_total, 2),
            '平均汇率': round(avg_rate, 4),
        })
        result['total_rmb'] += rmb_total

    result['total_rmb'] = round(result['total_rmb'], 2)
    if result['total_rmb_all'] > 0:
        result['export_ratio'] = round(result['total_rmb'] / result['total_rmb_all'] * 100, 1)

    return result
