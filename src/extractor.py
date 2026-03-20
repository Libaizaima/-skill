import os
import zipfile
import tempfile
import shutil
import subprocess
from typing import Dict, List, Optional


def extract(zip_path: str, dest_dir: Optional[str] = None) -> Dict[str, List[str]]:
    """
    解压 ZIP 文件并识别各类数据文件。

    Args:
        zip_path: ZIP 文件路径
        dest_dir: 解压目标目录，默认为临时目录

    Returns:
        file_map: {
            'company_name': str,
            'flows': [str, ...],       # 对公流水文件路径列表
            'invoices_in': [str, ...],  # 进项票文件路径列表
            'invoices_out': [str, ...], # 销项票文件路径列表
            'invoices_detail': [str, ...],  # 全量发票查询导出结果
            'receivable': [str, ...],   # 应收明细文件路径列表
            'payable': [str, ...],      # 应付明细文件路径列表
            'credit_reports': [str, ...],  # 征信报告 PDF
            'other': [str, ...],        # 其他文件
            'extract_dir': str,         # 解压目录
        }
    """
    if dest_dir is None:
        dest_dir = tempfile.mkdtemp(prefix='flow_analysis_')

    # 解压，处理中文编码
    _extract_zip_with_encoding(zip_path, dest_dir)

    # 识别公司名（ZIP 根目录下的第一个文件夹名）
    company_name = _detect_company_name(dest_dir, zip_path)

    # 扫描并分类文件
    file_map = {
        'company_name': company_name,
        'flows': [],
        'invoices_in': [],
        'invoices_out': [],
        'invoices_detail': [],
        'receivable': [],
        'payable': [],
        'credit_reports': [],
        'invoices_in_pdf': [],
        'invoices_out_pdf': [],
        'other': [],
        'extract_dir': dest_dir,
    }

    # 解压内嵌的 .rar 文件
    _extract_rar_archives(dest_dir)

    for root, dirs, files in os.walk(dest_dir):
        for fname in files:
            fpath = os.path.join(root, fname)
            rel_path = os.path.relpath(fpath, dest_dir)
            _classify_file(fpath, fname, rel_path, file_map)

    return file_map


def _extract_rar_archives(dest_dir: str):
    """扫描解压目录中的 .rar 文件并用 unrar 命令行工具展开"""
    rar_files = []
    for root, dirs, files in os.walk(dest_dir):
        for fname in files:
            if fname.lower().endswith('.rar'):
                rar_files.append(os.path.join(root, fname))

    if not rar_files:
        return

    # 检查 unrar 是否可用
    try:
        subprocess.run(['unrar'], capture_output=True, timeout=5)
    except FileNotFoundError:
        print(f"[WARN] 发现 {len(rar_files)} 个 .rar 文件，但未安装 unrar 工具，跳过。")
        print(f"       请运行: brew install --cask rar")
        return

    for rar_path in rar_files:
        rar_name = os.path.splitext(os.path.basename(rar_path))[0]
        extract_to = os.path.join(os.path.dirname(rar_path), rar_name) + os.sep
        os.makedirs(extract_to, exist_ok=True)
        try:
            result = subprocess.run(
                ['unrar', 'x', '-y', '-o+', rar_path, extract_to],
                capture_output=True, text=True, timeout=300
            )
            if result.returncode == 0:
                count = sum(len(fs) for _, _, fs in os.walk(extract_to))
                print(f"[INFO] 已解压 RAR: {os.path.basename(rar_path)} -> {count} 个文件")
                os.remove(rar_path)
            else:
                print(f"[WARN] 解压 RAR 失败: {os.path.basename(rar_path)}, {result.stderr[:200]}")
        except Exception as e:
            print(f"[WARN] 解压 RAR 异常: {os.path.basename(rar_path)}, {e}")


def _extract_zip_with_encoding(zip_path: str, dest_dir: str):
    """解压 ZIP 文件，自动处理中文文件名编码（GBK/UTF-8）"""
    zf = zipfile.ZipFile(zip_path, 'r')

    for info in zf.infolist():
        # 尝试解码文件名
        decoded_name = _decode_filename(info.filename)

        target_path = os.path.join(dest_dir, decoded_name)
        if info.is_dir():
            os.makedirs(target_path, exist_ok=True)
        else:
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with zf.open(info) as source, open(target_path, 'wb') as target:
                target.write(source.read())

    zf.close()


def _decode_filename(filename: str) -> str:
    """尝试将 ZIP 内部文件名从 cp437 → GBK 或 UTF-8 解码"""
    try:
        return filename.encode('cp437').decode('gbk')
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass
    try:
        return filename.encode('cp437').decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass
    return filename


def _detect_company_name(dest_dir: str, zip_path: str) -> str:
    """从解压结果或 ZIP 文件名推断公司名称"""
    # 优先使用 ZIP 内的根文件夹名
    entries = os.listdir(dest_dir)
    for entry in entries:
        if os.path.isdir(os.path.join(dest_dir, entry)):
            return entry

    # 否则从 ZIP 文件名推断
    basename = os.path.splitext(os.path.basename(zip_path))[0]
    return basename


def _classify_file(fpath: str, fname: str, rel_path: str, file_map: dict):
    """根据文件名和路径将文件分类"""
    fname_lower = fname.lower()
    rel_lower = rel_path.lower()

    # 跳过隐藏文件和 macOS 元数据
    if fname.startswith('.') or '__MACOSX' in rel_path:
        return

    # 跳过已有的报告文件
    if '客户分析' in fname and fname.endswith('.docx'):
        file_map['other'].append(fpath)
        return

    # 对公流水
    flow_keywords = ['流水', '对公', '账户', '对账单', '日记账', '基本户']
    if any(kw in rel_lower or kw in fname for kw in flow_keywords):
        if fname_lower.endswith(('.xls', '.xlsx')):
            file_map['flows'].append(fpath)
        return

    # 发票
    if '发票' in rel_lower or '开票' in rel_lower:
        if fname_lower.endswith(('.xls', '.xlsx')):
            if '进项' in fname or '进项' in rel_path:
                file_map['invoices_in'].append(fpath)
            elif '销项' in fname or '销项' in rel_path:
                file_map['invoices_out'].append(fpath)
            elif '全量' in fname or '查询' in fname or '导出' in fname:
                file_map['invoices_detail'].append(fpath)
            else:
                file_map['other'].append(fpath)
        elif fname_lower.endswith('.pdf'):
            # PDF 发票原件，按目录分类
            if '进项' in rel_path:
                file_map['invoices_in_pdf'].append(fpath)
            elif '销项' in rel_path:
                file_map['invoices_out_pdf'].append(fpath)
            else:
                file_map['other'].append(fpath)
        return

    # 应收
    if '应收' in fname:
        if fname_lower.endswith(('.xls', '.xlsx')):
            file_map['receivable'].append(fpath)
        return

    # 应付
    if '应付' in fname:
        if fname_lower.endswith(('.xls', '.xlsx')):
            file_map['payable'].append(fpath)
        return

    # 征信报告
    if '征信' in fname or '信用报告' in fname:
        file_map['credit_reports'].append(fpath)
        return

    # 其他
    file_map['other'].append(fpath)
