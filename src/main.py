# -*- coding: utf-8 -*-
"""流水征信分析 Skill — 主入口（多Agent架构）"""

import sys
import os
import shutil

# 将 src 目录添加到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from agents import agent_runner


def _setup_workspace(output_dir: str, company_name: str) -> dict:
    """
    创建工作目录结构。

    Returns:
        dirs: {
            'root': 工作目录根路径,
            'extract': 解压文件目录,
            'parsed': 解析结果目录,
            'pdf_ocr': PDF识别结果目录,
        }
    """
    workspace = os.path.join(output_dir, company_name)
    dirs = {
        'root': workspace,
        'extract': os.path.join(workspace, '01_解压文件'),
        'parsed': os.path.join(workspace, '02_解析结果'),
        'pdf_ocr': os.path.join(workspace, '03_PDF识别'),
    }

    # 如果工作目录已存在，先清空
    if os.path.exists(workspace):
        shutil.rmtree(workspace)

    for d in dirs.values():
        os.makedirs(d, exist_ok=True)

    return dirs


def process(zip_path: str, output_path: str):
    """
    主处理流程：ZIP → 多Agent分析 → 生成 DOCX 报告

    Args:
        zip_path: 输入 ZIP 文件路径
        output_path: 输出 DOCX 报告路径
    """
    print(f"[INFO] 开始处理: {zip_path}")
    print(f"[INFO] 架构: Brain Agent + Tool Agent (多Agent协同)")

    # ── 预解压获取公司名，建立工作目录 ──
    import tempfile
    import extractor

    tmp_dir = tempfile.mkdtemp(prefix='flow_pre_')
    try:
        pre_map = extractor.extract(zip_path, dest_dir=tmp_dir)
        company_name = pre_map['company_name']
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    output_dir = os.path.dirname(output_path) or 'output'
    ws = _setup_workspace(output_dir, company_name)
    print(f"[INFO] 工作目录: {ws['root']}")

    # 更新 output_path 到工作目录下
    report_name = os.path.basename(output_path)
    output_path = os.path.join(ws['root'], report_name)

    # ── 交由多Agent架构执行 ──
    agent_runner.run(zip_path, output_path, ws)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python main.py <zip_path> [output_path]")
        print("示例: python main.py input/佛瑞森科技.zip output/客户分析报告.docx")
        sys.exit(1)

    zip_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        basename = os.path.splitext(os.path.basename(zip_path))[0]
        output_path = f"output/客户分析（{basename}）.docx"

    process(zip_path, output_path)
