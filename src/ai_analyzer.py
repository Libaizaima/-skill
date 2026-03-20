# -*- coding: utf-8 -*-
"""AI 分析模块 — 调用大模型生成分析报告文字"""

import json
import numpy as np
from typing import Dict, Any, Optional
import llm_client


def _json_default(obj):
    """JSON 序列化时处理 numpy 类型"""
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        return float(obj)
    if isinstance(obj, np.ndarray):
        return obj.tolist()
    return str(obj)


SYSTEM_PROMPT = """你是一位经验丰富的银行信贷审批员和财务分析师。
你的任务是根据提供的结构化数据，撰写专业的分析报告段落。

要求：
1. 语言风格：专业、客观、简洁，适合放入正式的客户分析报告
2. 格式：直接输出分析文字，不要用 Markdown 标记，不要加标题
3. 分析要有洞察力：不仅描述数据，还要指出趋势、风险、亮点
4. 金额用中文习惯：万元单位，保留2位小数
5. 每个分析段落控制在100-200字"""


def analyze_all(results: Dict[str, Any]) -> Dict[str, str]:
    """
    调用大模型生成所有分析文字。

    Args:
        results: 完整的分析结果字典

    Returns:
        ai_texts: {
            'credit_personal': '征信分析文字...',
            'credit_company': '企业征信分析文字...',
            'flow_analysis': '流水分析文字...',
            'invoice_analysis': '发票分析文字...',
            'risk_assessment': '风险评估文字...',
            'summary': '总结建议文字...',
        }
    """
    if not llm_client.is_available():
        print("[INFO] LLM 不可用，使用模板文案生成报告")
        return {}

    print("[INFO] Step 6.5: AI 智能分析...")
    ai_texts = {}

    # 1. 个人征信分析
    credit_data = results.get('credit_data', {})
    personal = credit_data.get('personal')
    if personal:
        text = analyze_credit_personal(personal)
        if text:
            ai_texts['credit_personal'] = text
            print("  ✅ 个人征信分析完成")

    # 2. 企业征信分析
    company = credit_data.get('company')
    if company:
        text = analyze_credit_company(company)
        if text:
            ai_texts['credit_company'] = text
            print("  ✅ 企业征信分析完成")

    # 3. 流水分析
    text = analyze_flow(results.get('overall', {}), results.get('company_name', ''))
    if text:
        ai_texts['flow_analysis'] = text
        print("  ✅ 流水分析完成")

    # 4. 发票分析
    text = analyze_invoice(results)
    if text:
        ai_texts['invoice_analysis'] = text
        print("  ✅ 发票分析完成")

    # 5. 综合风险评估
    text = analyze_risk(results)
    if text:
        ai_texts['risk_assessment'] = text
        print("  ✅ 风险评估完成")

    # 6. 总结建议
    text = generate_summary(results, ai_texts)
    if text:
        ai_texts['summary'] = text
        print("  ✅ 总结建议完成")

    return ai_texts


def analyze_credit_personal(personal: Dict[str, Any]) -> Optional[str]:
    """分析个人征信"""
    data = {
        '姓名': personal.get('姓名', ''),
        '婚姻状况': personal.get('婚姻状况', ''),
        '征信时间': personal.get('征信时间', ''),
        '信用卡账户数': personal.get('信用卡账户数', ''),
        '信用卡未结清': personal.get('信用卡未结清数', ''),
        '信用卡逾期': personal.get('信用卡逾期账户数', ''),
        '信用卡总额度': personal.get('信用卡总额度', ''),
        '信用卡已用额度': personal.get('信用卡已用额度', ''),
        '贷款账户数': personal.get('贷款账户数', ''),
        '贷款未结清': personal.get('贷款未结清数', ''),
        '贷款逾期': personal.get('贷款逾期账户数', ''),
        '贷款总额度': personal.get('贷款总额度', ''),
        '贷款未还本金': personal.get('贷款未还总本金', ''),
        '月供金额': personal.get('月供金额', ''),
        '近1月查询': personal.get('近1月查询次数', ''),
        '近3月查询': personal.get('近3月查询次数', ''),
        '近6月查询': personal.get('近6月查询次数', ''),
        '贷款明细': personal.get('贷款明细', []),
    }

    prompt = f"""请分析以下法人个人征信数据，撰写征信分析段落。
重点关注：逾期情况、负债水平、查询频率是否异常、信用卡使用率。

数据：
{json.dumps(data, ensure_ascii=False, indent=2, default=_json_default)}"""

    return llm_client.chat(SYSTEM_PROMPT, prompt)


def analyze_credit_company(company: Dict[str, Any]) -> Optional[str]:
    """分析企业征信"""
    data = {
        '企业名称': company.get('企业名称', ''),
        '统一信用代码': company.get('统一社会信用代码', ''),
        '信贷机构数': company.get('信贷机构数', ''),
        '未结清机构': company.get('未结清机构数', ''),
        '借贷余额': company.get('借贷余额', ''),
        '担保余额': company.get('担保余额', ''),
    }

    prompt = f"""请分析以下企业征信数据，撰写企业征信分析段落。

数据：
{json.dumps(data, ensure_ascii=False, indent=2, default=_json_default)}"""

    return llm_client.chat(SYSTEM_PROMPT, prompt)


def analyze_flow(overall: Dict[str, Any], company_name: str) -> Optional[str]:
    """分析对公流水"""
    if not overall:
        return None

    data = {
        '公司名称': company_name,
        '总收入': overall.get('total_income', 0),
        '总支出': overall.get('total_expense', 0),
        '净额': overall.get('net', 0),
        '总笔数': overall.get('total_count', 0),
        '经营性有效笔数': overall.get('biz_count', 0),
        '经营性收入': overall.get('biz_income', 0),
        '经营性支出': overall.get('biz_expense', 0),
        '月均收入': overall.get('monthly_avg_income', 0),
        '月均支出': overall.get('monthly_avg_expense', 0),
    }

    prompt = f"""请分析以下企业对公流水统计数据，撰写流水分析段落。
重点关注：经营性收入占比、收支是否平衡、月均流水水平、是否有异常波动。

数据：
{json.dumps(data, ensure_ascii=False, indent=2, default=_json_default)}"""

    return llm_client.chat(SYSTEM_PROMPT, prompt)


def analyze_invoice(results: Dict[str, Any]) -> Optional[str]:
    """分析发票数据"""
    invoice_stats = results.get('invoice_stats', {})
    if not invoice_stats:
        return None

    data = {
        '进项票数量': invoice_stats.get('in_count', 0),
        '进项金额合计': invoice_stats.get('in_total', 0),
        '销项票数量': invoice_stats.get('out_count', 0),
        '销项金额合计': invoice_stats.get('out_total', 0),
        '前5大供应商': invoice_stats.get('top_suppliers', []),
        '前5大客户': invoice_stats.get('top_customers', []),
    }

    prompt = f"""请分析以下企业发票统计数据，撰写发票分析段落。
重点关注：进销比例、客户/供应商集中度、是否存在关联交易嫌疑。

数据：
{json.dumps(data, ensure_ascii=False, indent=2, default=_json_default)}"""

    return llm_client.chat(SYSTEM_PROMPT, prompt)


def analyze_risk(results: Dict[str, Any]) -> Optional[str]:
    """综合风险评估"""
    credit_data = results.get('credit_data', {}) or {}
    personal = credit_data.get('personal', {}) or {}
    company = credit_data.get('company', {}) or {}
    overall = results.get('overall', {})

    data = {
        '公司名称': results.get('company_name', ''),
        '法人姓名': personal.get('姓名', ''),
        '法人贷款未还本金': personal.get('贷款未还总本金', ''),
        '法人月供金额': personal.get('月供金额', ''),
        '法人逾期': personal.get('贷款逾期账户数', ''),
        '企业信贷机构数': company.get('信贷机构数', ''),
        '企业借贷余额': company.get('借贷余额', ''),
        '年总收入': overall.get('total_income', 0),
        '年总支出': overall.get('total_expense', 0),
        '经营性收入': overall.get('biz_income', 0),
        '应收总额': results.get('receivable_total', 0),
        '应付总额': results.get('payable_total', 0),
    }

    prompt = f"""请根据以下企业综合数据，撰写风险评估和授信建议段落。
分析维度：
1. 还款能力（收入 vs 负债）
2. 经营稳定性
3. 信用记录
4. 应收应付风险
5. 总体风险等级（低/中/高）和授信建议

数据：
{json.dumps(data, ensure_ascii=False, indent=2, default=_json_default)}"""

    return llm_client.chat(SYSTEM_PROMPT, prompt)


def generate_summary(results: Dict[str, Any], ai_texts: Dict[str, str]) -> Optional[str]:
    """生成总结建议"""
    prompt = f"""请根据以下分析结果，撰写简短的总结建议段落（100字内）。

企业: {results.get('company_name', '')}
已有分析:
- 征信: {ai_texts.get('credit_personal', '无')[:100]}
- 流水: {ai_texts.get('flow_analysis', '无')[:100]}
- 风险: {ai_texts.get('risk_assessment', '无')[:100]}

请给出1-2句话的总结性建议。"""

    return llm_client.chat(SYSTEM_PROMPT, prompt)
