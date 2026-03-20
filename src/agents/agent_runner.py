# -*- coding: utf-8 -*-
"""
Agent Runner — 双Agent协同编排器。

负责协调 Brain Agent 和 Tool Agent 完成完整的流水分析流程：
  1. extractor.extract() → 解压文件，同时获取 fallback 分类结果
  2. brain_agent.plan() → AI大脑分析文件树，生成调用计划
  3. tool_agent.execute() → 工具Agent执行计划，解析各类文件
  4. 流水筛选/分类 (flow_classifier)
  5. analyzer.analyze_all() → 统计分析
  6. ai_analyzer.analyze_all() → AI智能分析
  7. report_generator.generate() → 生成 DOCX 报告
"""

import os
import sys
import json
import pandas as pd
from typing import Dict, Any, Optional

# 将 src 目录加入路径
_src_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

import extractor
import flow_classifier
import analyzer
import ai_analyzer
import report_generator
from agents import brain_agent, tool_agent


def _save_df(df: pd.DataFrame, path: str, label: str):
    """保存 DataFrame 为 CSV（UTF-8 BOM 以便 Excel 打开）"""
    if df is not None and len(df) > 0:
        df.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"  [保存] {label} → {os.path.basename(path)} ({len(df)} 行)")


def _save_json(data: dict, path: str, label: str):
    """保存字典为 JSON"""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"  [保存] {label} → {os.path.basename(path)}")


def run(zip_path: str, output_path: str, ws: Dict[str, str]) -> None:
    """
    多Agent架构主流程。

    Args:
        zip_path:    输入 ZIP 文件路径
        output_path: 输出 DOCX 报告路径
        ws:          工作目录字典 {root, extract, parsed, pdf_ocr}
    """
    # ── Step 1: 解压（extractor 同时提供 fallback 分类）──
    print("\n[AgentRunner] Step 1: 解压 ZIP 文件...")
    file_map = extractor.extract(zip_path, dest_dir=ws["extract"])
    company_name = file_map["company_name"]
    extract_dir = file_map["extract_dir"]
    print(f"  公司名称: {company_name}")
    print(f"  解压目录: {extract_dir}")

    # ── Step 2: Brain Agent — 分析文件树，生成调用计划 ──
    print("\n[AgentRunner] Step 2: Brain Agent 分析文件结构...")
    execution_plan = brain_agent.plan(extract_dir, fallback_file_map=file_map)

    # 如果大脑未能识别公司名，使用 extractor 的结果
    if not execution_plan.get("company_name"):
        execution_plan["company_name"] = company_name
    else:
        company_name = execution_plan["company_name"]

    # 保存执行计划（便于调试）
    plan_path = os.path.join(ws["parsed"], "brain_plan.json")
    _save_json(execution_plan, plan_path, "Brain Agent 调用计划")

    # ── Step 3: Tool Agent — 按计划执行解析 ──
    print("\n[AgentRunner] Step 3: Tool Agent 执行解析...")
    parsed = tool_agent.execute(execution_plan, pdf_ocr_dir=ws["pdf_ocr"])

    flow_df = parsed["flow_df"]
    in_invoice_df = parsed["in_invoice_df"]
    out_invoice_df = parsed["out_invoice_df"]
    recv_df = parsed["recv_df"]
    pay_df = parsed["pay_df"]
    credit_data = parsed.get("credit_data", {})
    property_certs = parsed.get("property_certs", [])
    financial_statements = parsed.get("financial_statements", [])
    tax_certs = parsed.get("tax_certs", [])

    # ── Step 4: 保存解析结果 ──
    print("\n[AgentRunner] Step 4: 保存解析中间结果...")
    _save_df(flow_df, os.path.join(ws["parsed"], "流水_合并.csv"), "全量流水")
    _save_df(in_invoice_df, os.path.join(ws["parsed"], "进项票.csv"), "进项票")
    _save_df(out_invoice_df, os.path.join(ws["parsed"], "销项票.csv"), "销项票")
    _save_df(recv_df, os.path.join(ws["parsed"], "应收明细.csv"), "应收明细")
    _save_df(pay_df, os.path.join(ws["parsed"], "应付明细.csv"), "应付明细")
    _save_json(credit_data, os.path.join(ws["parsed"], "征信报告.json"), "征信报告")
    if property_certs:
        _save_json(property_certs, os.path.join(ws["parsed"], "房产证.json"), "房产证信息")
    if financial_statements:
        _save_json(financial_statements, os.path.join(ws["parsed"], "财务报表.json"), "财务报表")
    if tax_certs:
        _save_json(tax_certs, os.path.join(ws["parsed"], "完税证明.json"), "完税证明")

    # ── Step 4.5: 未解析文件汇报 ──
    skipped_from_brain = execution_plan.get("skipped_files", [])
    skipped_from_tool  = parsed.get("skipped_files", [])
    all_skipped = skipped_from_brain + skipped_from_tool

    if all_skipped:
        skipped_report = {"total": len(all_skipped), "items": all_skipped}
        _save_json(skipped_report,
                   os.path.join(ws["parsed"], "未解析文件清单.json"),
                   "未解析文件清单")
        print(f"\n[AgentRunner] ⚠️  未解析文件：共 {len(all_skipped)} 个")
        for item in all_skipped:
            print(f"  ✗ {item['file']}：{item['reason']}")
    else:
        print("  ✅ 所有文件均已分配解析工具")

    # ── Step 5: 确定分析年份 & 流水分类 ──
    print("\n[AgentRunner] Step 5: 流水筛选与分类...")
    if len(flow_df) == 0:
        print("  [WARN] 流水数据为空，跳过年份筛选")
        flow_df_year = flow_df
        analysis_year = 2025
    else:
        flow_df["_year"] = flow_df["交易时间"].apply(lambda x: x.year if pd.notna(x) else None)
        year_counts = flow_df["_year"].value_counts()
        candidate_years = sorted(
            [int(y) for y in year_counts.index if pd.notna(y) and year_counts[y] >= 100],
            reverse=True,
        )
        analysis_year = candidate_years[0] if candidate_years else (
            int(year_counts.index[0]) if len(year_counts) > 0 else 2025
        )
        print(f"  分析年份: {analysis_year}  (全量: {len(flow_df)} 条)")

        flow_df_year = flow_df[flow_df["_year"] == analysis_year].copy()
        flow_df_year = flow_df_year.drop(columns=["_year"])
        flow_df = flow_df.drop(columns=["_year"])
        print(f"  {analysis_year}年记录数: {len(flow_df_year)} 条")

        # 股东检测 + 分类
        shareholder_names = flow_classifier.detect_shareholder_names(flow_df)
        print(f"  检测到股东: {shareholder_names}")
        flow_df_year = flow_classifier.classify(flow_df_year, shareholder_names, company_name)
        for cat, count in flow_df_year["分类"].value_counts().items():
            print(f"    {cat}: {count} 笔")

    _save_df(flow_df_year, os.path.join(ws["parsed"], f"流水_{analysis_year}年.csv"), f"{analysis_year}年流水")
    _save_df(flow_df_year, os.path.join(ws["parsed"], "流水_已分类.csv"), "已分类流水")

    # ── Step 6: 统计分析 ──
    print("\n[AgentRunner] Step 6: 统计分析...")
    results = analyzer.analyze_all(
        flow_df_year, in_invoice_df, out_invoice_df,
        recv_df, pay_df, company_name,
    )
    results["analysis_year"] = analysis_year
    results["credit_data"] = credit_data

    # ── Step 6.5: AI 智能分析 ──
    print("\n[AgentRunner] Step 6.5: AI智能分析...")
    ai_texts = ai_analyzer.analyze_all(results)
    results["ai_analysis"] = ai_texts

    # ── Step 7: 生成报告 ──
    print("\n[AgentRunner] Step 7: 生成 DOCX 报告...")
    report_generator.generate(results, output_path)

    print(f"\n✅ [AgentRunner] 全部完成！")
    print(f"  报告: {output_path}")
    print(f"  工作目录: {ws['root']}")
