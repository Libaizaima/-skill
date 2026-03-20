# 黄-流水分析系统 — 维护文档

> **版本**: v2.0（多Agent架构）  
> **最后更新**: 2026-03-20  
> **项目路径**: `src/`

---

## 目录

1. [系统架构](#1-系统架构)
2. [目录结构](#2-目录结构)
3. [配置说明](#3-配置说明)
4. [解析器清单](#4-解析器清单)
5. [如何新增解析器](#5-如何新增解析器)
6. [如何扩展现有解析器](#6-如何扩展现有解析器)
7. [AI 分析模块](#7-ai-分析模块)
8. [常见问题与调试](#8-常见问题与调试)
9. [输出文件说明](#9-输出文件说明)
10. [依赖管理](#10-依赖管理)

---

## 1. 系统架构

本系统采用**多 Agent 协同**架构：

```
main.py
  └── AgentRunner (agents/agent_runner.py)
        ├── Step 1: 解压 ZIP
        ├── Step 2: BrainAgent → 分析文件树，生成调用计划 (brain_plan.json)
        ├── Step 3: ToolAgent  → 按计划调用各解析器，汇总结果
        ├── Step 4: 保存解析中间结果 (02_解析结果/)
        ├── Step 5: 流水筛选与分类
        ├── Step 6: 统计分析 (analyzer.py)
        ├── Step 6.5: AI 智能分析 (ai_analyzer.py + llm_client.py)
        └── Step 7: 生成 DOCX 报告 (report_generator.py)
```

**Brain Agent** (`agents/brain_agent.py`)
- 遍历解压目录，生成文件树
- 调用 LLM（通过 `llm_client.py`）分析各文件，制定调用计划
- 输出：JSON 调用计划，包含工具名、文件列表、跳过理由

**Tool Agent** (`agents/tool_agent.py`)
- 按 Brain Agent 的计划逐步执行各工具
- 维护 `ctx` 字典，在其中积累结果
- 对不支持的文件格式自动过滤并记入未解析清单

---

## 2. 目录结构

```
黄-流水分析skill/
├── config.json               # LLM 配置（API Key、Model、Base URL）
├── input/                    # 放入待分析的 ZIP 文件
├── output/                   # 各公司的分析结果
│   └── {公司名}/
│       ├── 01_解压文件/        # ZIP 解压后的原始文件
│       ├── 02_解析结果/        # 解析器输出的中间结果
│       │   ├── 流水_合并.csv
│       │   ├── 流水_2025年.csv
│       │   ├── 流水_已分类.csv
│       │   ├── 征信报告.json
│       │   ├── 房产证.json
│       │   ├── 财务报表.json
│       │   ├── 完税证明.json
│       │   ├── brain_plan.json  # Brain Agent 的调用计划（调试用）
│       │   └── 未解析文件清单.json
│       └── 客户分析（{公司名}）.docx
└── src/
    ├── main.py               # 入口
    ├── llm_client.py         # LLM 客户端封装
    ├── extractor.py          # ZIP 解压
    ├── bank_flow_parser.py   # 流水解析（XLS/XLSX/CSV/PDF）
    ├── invoice_parser.py     # 发票解析（Excel）
    ├── pdf_invoice_parser.py # 发票解析（PDF）
    ├── receivable_payable_parser.py  # 应收/应付解析
    ├── credit_report_parser.py       # 征信报告解析
    ├── property_cert_parser.py       # 房产证解析（含LLM Vision）
    ├── financial_statement_parser.py # 财务报表解析（PDF+XLS）
    ├── tax_cert_parser.py            # 完税证明解析
    ├── flow_classifier.py    # 流水分类
    ├── analyzer.py           # 统计分析
    ├── ai_analyzer.py        # AI 智能分析
    ├── report_generator.py   # DOCX 报告生成
    └── agents/
        ├── agent_runner.py   # 主编排器
        ├── brain_agent.py    # Brain Agent（LLM决策）
        ├── tool_agent.py     # Tool Agent（工具执行）
        └── tool_registry.py  # 工具注册表（Brain LLM 的工具说明）
```

---

## 3. 配置说明

`config.json`（在项目根目录，需手动创建，不提交 git）：

```json
{
  "api_key": "sk-xxx",
  "base_url": "https://your-proxy.example.com/v1",
  "model": "gpt-4o",
  "temperature": 0.1,
  "max_tokens": 4096
}
```

| 字段 | 说明 |
|------|------|
| `api_key` | OpenAI / 中转 API Key |
| `base_url` | API 端点，支持中转（OpenAI 格式） |
| `model` | 使用的模型名称，推荐 `gpt-4o`（支持 Vision）|
| `temperature` | Brain Agent 规划时的温度，建议 0.1 |
| `max_tokens` | LLM 最大响应长度 |

> ⚠️ **房产证 Vision 识别**需要模型支持图像输入（如 `gpt-4o`）。若模型不支持 Vision，会回退到纯文字提取（扫描件则返回空）。

---

## 4. 解析器清单

| 工具名 | 文件 | 支持格式 | 识别关键词 | 输出 |
|--------|------|----------|------------|------|
| `bank_flow_parser` | `bank_flow_parser.py` | `.xls` `.xlsx` `.csv` `.pdf` | 流水、明细、银行 | `ctx["flow_df"]` → `流水_合并.csv` |
| `invoice_parser_in` | `invoice_parser.py` | `.xls` `.xlsx` | 进项 | `ctx["in_invoice_df"]` |
| `invoice_parser_out` | `invoice_parser.py` | `.xls` `.xlsx` | 销项、开票 | `ctx["out_invoice_df"]` |
| `pdf_invoice_parser_in` | `pdf_invoice_parser.py` | `.pdf` | 进项发票 PDF | `ctx["in_invoice_df"]` |
| `pdf_invoice_parser_out` | `pdf_invoice_parser.py` | `.pdf` | 销项发票 PDF | `ctx["out_invoice_df"]` |
| `receivable_parser` | `receivable_payable_parser.py` | `.xls` `.xlsx` | 应收 | `ctx["recv_df"]` |
| `payable_parser` | `receivable_payable_parser.py` | `.xls` `.xlsx` | 应付 | `ctx["pay_df"]` |
| `credit_report_parser` | `credit_report_parser.py` | `.pdf` | 征信、信用报告 | `ctx["credit_data"]` → `征信报告.json` |
| `property_cert_parser` | `property_cert_parser.py` | `.pdf` | 房产证、不动产证 | `ctx["property_certs"]` → `房产证.json` |
| `financial_statement_parser` | `financial_statement_parser.py` | `.pdf` | 财务报表、利润表 | `ctx["financial_statements"]` → `财务报表.json` |
| `balance_sheet_parser` | `financial_statement_parser.py` | `.xls` `.xlsx` | 资产负债表 | `ctx["financial_statements"]` → `财务报表.json` |
| `tax_cert_parser` | `tax_cert_parser.py` | `.pdf` | 完税证明 | `ctx["tax_certs"]` → `完税证明.json` |

---

## 5. 如何新增解析器

以新增一个**营业执照解析器**为例，需要改动 **4 个文件**：

### 步骤 1：创建解析器模块

新建 `src/business_license_parser.py`：

```python
# -*- coding: utf-8 -*-
"""营业执照解析器"""
from typing import List, Dict, Any

def parse(file_paths: List[str]) -> List[Dict[str, Any]]:
    """
    解析营业执照 PDF。
    Returns: [{"公司名称": ..., "统一社会信用代码": ..., "注册资本": ...}]
    """
    results = []
    for fpath in file_paths:
        # ... 解析逻辑
        results.append({...})
    return results
```

**输出规范**：
- 返回 `List[Dict]`（结构化数据）或 `pd.DataFrame`（流水类）
- 函数签名必须是 `parse(file_paths: List[str]) -> ...`
- 打印进度用 `print(f"  → ...")` 格式

### 步骤 2：在 `tool_registry.py` 注册工具

在 `TOOL_SCHEMAS` 列表中添加一个字典：

```python
{
    "name": "business_license_parser",
    "description": (
        "解析营业执照PDF。"
        "提取：公司名称、统一社会信用代码、注册资本、经营范围。"
        "文件名通常包含：营业执照。"
    ),
    "accepts": [".pdf"],
    "keywords": ["营业执照", "统一社会信用代码"],
},
```

> `keywords` 用于 Brain Agent 识别文件是否应该路由到此工具。

### 步骤 3：在 `tool_agent.py` 添加 dispatch

**①** 在顶部 import 区添加：
```python
import business_license_parser
```

**②** 添加 dispatch 函数（在 `_TOOL_DISPATCH` 字典之前）：
```python
def _run_business_license_parser(files: List[str], ctx: Dict) -> None:
    """解析营业执照"""
    print(f"[ToolAgent] 解析营业执照: {len(files)} 个文件")
    try:
        items = business_license_parser.parse(files)
        ctx["business_licenses"] = ctx.get("business_licenses", []) + items
        print(f"  → 解析到 {len(items)} 份营业执照")
    except Exception as e:
        print(f"  [WARN] 营业执照解析失败: {e}")
```

**③** 在 `_TOOL_DISPATCH` 字典中注册：
```python
"business_license_parser": _run_business_license_parser,
```

**④** 在 `ctx` 初始化中添加 key：
```python
"business_licenses": [],
```

**⑤** 在 `_TOOL_ACCEPTS` 中添加扩展名过滤：
```python
"business_license_parser": {".pdf"},
```

### 步骤 4：在 `agent_runner.py` 保存结果

在 Step 4 保存区添加：
```python
business_licenses = parsed.get("business_licenses", [])
if business_licenses:
    _save_json(business_licenses, os.path.join(ws["parsed"], "营业执照.json"), "营业执照")
```

---

## 6. 如何扩展现有解析器

### 让 `bank_flow_parser` 支持新银行 PDF 格式

在 `_parse_pdf()` 函数的 `_PDF_COL_MAP` 字典中增加该银行的列名映射：

```python
_PDF_COL_MAP = {
    "交易日期": "交易时间",       # 农商行/通用
    "记账日期": "交易时间",       # 新增：工行格式
    "收入金额": "收入金额",       # 新增：某行格式
    # ...
}
```

若列名差异较大，可在 `_parse_pdf()` 中增加银行格式判断分支，或新建独立 `_parse_xxx_pdf()` 函数再在 `_parse_single_file()` 中按文件名路由。

### 让 `tax_cert_parser` 支持新税种

在 `_TAX_KEYWORDS` 列表中添加新税种名称：

```python
_TAX_KEYWORDS = [
    '增值税', '企业所得税', ...,
    '文化事业建设费',  # 新增
]
```

---

## 7. AI 分析模块

### llm_client.py

统一的 LLM 客户端入口，供所有模块调用：

```python
from llm_client import call_llm, get_client, get_model, is_available
```

| 函数 | 说明 |
|------|------|
| `call_llm(messages, ...)` | 发送对话请求，返回文本 |
| `get_client()` | 获取原始 OpenAI client（用于 Vision 等高级用法）|
| `get_model()` | 获取当前配置的模型名 |
| `is_available()` | 检查 LLM 是否已配置 |

### ai_analyzer.py

负责对解析结果进行智能分析，生成各报告章节文字。当 LLM 不可用时，自动降级为模板文字。

**新增 AI 分析章节的步骤：**
1. 在 `ai_analyzer.py` 中添加新分析函数，接收结构化数据，返回文字
2. 在 `agent_runner.py` Step 6.5 中调用该函数
3. 在 `report_generator.py` 中将结果写入对应报告章节

---

## 8. 常见问题与调试

### 🔍 查看 Brain Agent 的决策过程

每次运行后，调用计划保存在：
```
output/{公司名}/02_解析结果/brain_plan.json
```

可检查 Brain Agent 是否正确路由了文件，以及有无文件被跳过。

### 🔍 查看未解析文件

```
output/{公司名}/02_解析结果/未解析文件清单.json
```

每条记录包含文件路径和跳过原因，可据此决定是否新增解析器。

### ❌ 解析器接收到 0 条记录

1. 检查 `tool_registry.py` 中该工具的 `keywords` 是否匹配文件名
2. 检查 `tool_agent.py` 中 `_TOOL_ACCEPTS` 的扩展名是否包含该格式
3. 直接单测解析器：
   ```bash
   .venv/bin/python -c "
   import sys; sys.path.insert(0,'src')
   import my_parser
   print(my_parser.parse(['path/to/file']))
   "
   ```

### ❌ LLM 调用失败

1. 检查 `config.json` 是否存在且 API Key 正确
2. 检查 `base_url` 是否可访问
3. 系统在 LLM 不可用时会降级使用模板，不会崩溃

### ❌ PDF 解析无文字（扫描件）

- **流水类**：目前不支持扫描件流水（需 OCR）
- **房产证**：自动调用 LLM Vision 识别（需模型支持图像）
- **其他**：可在对应解析器中参考 `property_cert_parser.py` 的 Vision 回退逻辑

---

## 9. 输出文件说明

| 文件 | 内容 | 格式 |
|------|------|------|
| `流水_合并.csv` | 所有银行账户合并去重的全量流水 | CSV |
| `流水_2025年.csv` | 筛选出的分析年度流水 | CSV |
| `流水_已分类.csv` | 含分类标签的流水（租金/工资/税费等）| CSV |
| `征信报告.json` | 法人/企业信用报告摘要 | JSON |
| `房产证.json` | 房产证信息：权利人/坐落/面积 | JSON |
| `财务报表.json` | 资产负债表+利润表科目数据 | JSON |
| `完税证明.json` | 各期税款明细及合计 | JSON |
| `brain_plan.json` | Brain Agent 调用计划（调试用）| JSON |
| `未解析文件清单.json` | 未处理的文件及原因 | JSON |
| `客户分析（...）.docx` | 完整客户分析报告 | DOCX |

---

## 10. 依赖管理

主要依赖（在 `.venv` 中）：

| 包 | 用途 |
|----|------|
| `openai` | LLM API 调用 |
| `pdfplumber` | PDF 文字/表格提取 |
| `pandas` | 流水数据处理 |
| `xlrd` | 旧版 `.xls` 读取 |
| `openpyxl` | `.xlsx` 读取 |
| `python-docx` | DOCX 报告生成 |
| `Pillow` | 房产证图像处理（Vision 回退）|

安装/更新：
```bash
# 已在 .venv 中，直接使用
.venv/bin/pip install -r requirements.txt

# 运行主程序
.venv/bin/python src/main.py input/xxx公司.zip
```
