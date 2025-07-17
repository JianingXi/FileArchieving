import os
import pdfplumber
import pandas as pd
import docx
import win32com.client
from pathlib import Path

fail_log = []

def log_failure(path, reason):
    print(f"❌ 失败：{path} | 原因：{reason}")
    fail_log.append(f"{path}  |  原因：{reason}")

def convert_pdf_to_txt(path):
    try:
        with pdfplumber.open(str(path)) as pdf:
            text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        return text if text.strip() else ""
    except Exception as e:
        log_failure(path, f"PDF读取错误: {e}")
        return ""

def convert_doc_to_docx(path):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(path))
        new_path = path.with_suffix(".docx")
        doc.SaveAs(str(new_path), FileFormat=16)
        doc.Close()
        word.Quit()
        return new_path
    except Exception as e:
        log_failure(path, f"DOC转DOCX失败: {e}")
        return None

def convert_docx_to_txt(path):
    try:
        doc = docx.Document(str(path))
        return '\n'.join(p.text for p in doc.paragraphs)
    except Exception as e:
        log_failure(path, f"DOCX读取失败: {e}")
        return ""

def convert_xls_to_xlsx(path):
    try:
        df = pd.read_excel(path, sheet_name=None, engine="xlrd")
        new_path = path.with_suffix(".xlsx")
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            for sheet, data in df.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        return new_path
    except Exception as e:
        log_failure(path, f"XLS转XLSX失败: {e}")
        return None

def convert_excel_to_txt(path):
    try:
        df_dict = pd.read_excel(path, sheet_name=None, engine='openpyxl')
        text = ''
        for sheet, df in df_dict.items():
            text += f'【Sheet: {sheet}】\n{df.to_string(index=False)}\n\n'
        return text
    except Exception as e:
        log_failure(path, f"XLSX读取失败: {e}")
        return ""

def convert_and_delete(filepath: Path):
    ext = filepath.suffix.lower()
    txt_content = ""

    if not filepath.exists():
        log_failure(filepath, "文件不存在或路径错误")
        return

    if len(str(filepath)) > 250:
        log_failure(filepath, "路径过长")
        return

    if ext == ".pdf":
        txt_content = convert_pdf_to_txt(filepath)

    elif ext == ".doc":
        new_path = convert_doc_to_docx(filepath)
        if new_path:
            txt_content = convert_docx_to_txt(new_path)
            try: filepath.unlink()  # 删除 .doc
            except: pass
            filepath = new_path

    elif ext == ".docx":
        txt_content = convert_docx_to_txt(filepath)

    elif ext == ".xls":
        new_path = convert_xls_to_xlsx(filepath)
        if new_path:
            txt_content = convert_excel_to_txt(new_path)
            try: filepath.unlink()  # 删除 .xls
            except: pass
            filepath = new_path

    elif ext == ".xlsx":
        txt_content = convert_excel_to_txt(filepath)

    else:
        return

    if txt_content.strip():
        txt_path = filepath.with_suffix(".txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(txt_content)
        try: filepath.unlink()
        except: pass
        print(f"[成功] 转换为 TXT 并删除原文件：{filepath}")
    else:
        log_failure(filepath, "内容为空或提取失败")

def scan_and_convert(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            filepath = Path(os.path.join(dirpath, filename))
            if filepath.suffix.lower() in [".pdf", ".doc", ".docx", ".xls", ".xlsx"]:
                convert_and_delete(filepath)

if __name__ == "__main__":
    target_folder = r"C:\MyDocument\ToDoList\D20_ToHardDisk\关于开展教育教改类科的通知"
    scan_and_convert(target_folder)

    # 写入失败日志
    if fail_log:
        log_path = Path(target_folder) / "fail_log.txt"
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(fail_log))
        print(f"\n⚠️ 共 {len(fail_log)} 个文件处理失败，详情请查看：{log_path}")
    else:
        print("\n✅ 所有文件转换成功")
