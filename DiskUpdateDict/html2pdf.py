import os
import re
import pdfkit
import shutil
from pathlib import Path

# 配置 wkhtmltopdf 路径（根据实际安装路径修改）
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')


def read_file_with_fallback(file_path):
    """
    尝试使用 UTF-8 编码读取文件，失败后使用 GBK 编码。
    返回内容和实际使用的编码。
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return content, 'utf-8'
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='gbk') as f:
                content = f.read()
            return content, 'gbk'
        except Exception as e:
            print(f"读取 {file_path} 时出错: {e}")
            return None, None


def write_file_with_encoding(file_path, content, encoding):
    """
    按指定编码写入文件内容。
    """
    try:
        with open(file_path, 'w', encoding=encoding) as f:
            f.write(content)
    except Exception as e:
        print(f"写入 {file_path} 时出错: {e}")


def insert_base_tag(html_path):
    """
    检查并在HTML文件的<head>标签中插入<base>标签，指向HTML所在目录的绝对路径。
    采用合适的编码方式读取文件。
    """
    content, encoding_used = read_file_with_fallback(html_path)
    if content is None:
        print(f"无法读取 {html_path}，跳过预处理。")
        return

    try:
        # 查找<head>标签
        head_match = re.search(r'<head.*?>', content, flags=re.IGNORECASE)
        if head_match:
            # 如果不存在 <base> 标签，则插入
            if not re.search(r'<base\s', content, flags=re.IGNORECASE):
                abs_dir = Path(html_path).parent.resolve()
                base_tag = f'<base href="file:///{abs_dir.as_posix()}/">'
                pos = head_match.end()
                content = content[:pos] + base_tag + content[pos:]
                write_file_with_encoding(html_path, content, encoding_used)
                print(f"已为 {html_path} 插入 <base> 标签")
    except Exception as e:
        print(f"预处理 {html_path} 时出错: {e}")

import tempfile
import tempfile

def html_to_pdf(html_path, pdf_path):
    """
    临时处理 HTML 内容：保留图片，移除干扰资源，引入 base 标签，避免资源路径错误。
    不修改原始 HTML 文件。
    """
    try:
        # 读取内容及编码
        content, encoding_used = read_file_with_fallback(html_path)
        if content is None:
            return False

        # 仅删除容易出错的标签：<script> 和 <link>
        content = re.sub(r'<script[^>]*?>.*?</script>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<link[^>]*?>', '', content, flags=re.IGNORECASE)

        # 插入 <base> 标签指向 HTML 所在文件夹（仅限未存在时）
        head_match = re.search(r'<head.*?>', content, flags=re.IGNORECASE)
        if head_match and not re.search(r'<base\s', content, flags=re.IGNORECASE):
            abs_dir = Path(html_path).parent.resolve()
            base_tag = f'<base href="file:///{abs_dir.as_posix()}/">'
            content = content[:head_match.end()] + base_tag + content[head_match.end():]

        # 写入临时 HTML 文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding=encoding_used) as tmp_file:
            tmp_file.write(content)
            tmp_html_path = tmp_file.name

        # 设置 wkhtmltopdf 转换选项
        options = {
            'enable-local-file-access': None,
            'load-error-handling': 'ignore',
            'load-media-error-handling': 'ignore'
        }

        # 转换为 PDF
        pdfkit.from_file(tmp_html_path, pdf_path, configuration=config, options=options)

        # 删除临时文件
        os.remove(tmp_html_path)
        return True

    except Exception as e:
        print(f"转换失败 {html_path}: {e}")
        return False



def convert_html_files_in_directory(directory):
    """
    递归遍历目录，转换所有HTML文件为PDF
    """
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith(('.htm', '.html')):
                html_path = os.path.join(root, filename)
                pdf_path = os.path.splitext(html_path)[0] + '.pdf'
                print(f"正在处理: {html_path}")


                if html_to_pdf(html_path, pdf_path):
                    if os.path.exists(pdf_path):
                        try:
                            # 删除原始 HTML 文件
                            os.remove(html_path)
                            print(f"转换成功并已删除原始文件: {html_path}")

                            # 尝试删除资源文件夹（匹配 _files 或 .files）
                            html_dir = os.path.dirname(html_path)
                            html_base = os.path.splitext(os.path.basename(html_path))[0]
                            possible_dirs = [
                                os.path.join(html_dir, f"{html_base}_files"),
                                os.path.join(html_dir, f"{html_base}.files")
                            ]

                            for res_dir in possible_dirs:
                                if os.path.isdir(res_dir):
                                    shutil.rmtree(res_dir)
                                    print(f"已删除资源文件夹: {res_dir}")
                        except Exception as e:
                            print(f"删除原始文件或文件夹失败 {html_path}: {e}")
                    else:
                        print(f"PDF文件未生成: {pdf_path}")


if __name__ == "__main__":
    # 指定要处理的目录
    target_directory = r"G:\作品"

    if not os.path.isdir(target_directory):
        print("错误: 指定的路径不是一个有效的目录!")
        exit(1)

    print(f"开始处理目录: {target_directory}")
    print("将把所有HTML文件(.htm, .html)转换为PDF并删除原始文件")

    convert_html_files_in_directory(target_directory)
    print("处理完成!")
