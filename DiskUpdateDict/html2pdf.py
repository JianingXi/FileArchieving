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




def html_to_pdf(html_path, pdf_path, debug_on_fail=False):
    """
    增强版 HTML 转 PDF 函数：
    - 清理脚本和样式引用
    - 插入 <base> 标签
    - 替换所有非法或不存在的本地图片
    - 保留失败时 HTML 文件供调试
    """
    import re
    import tempfile
    from pathlib import Path

    def replace_nonexistent_images_with_files_support(content: str, html_path: str) -> str:
        html_dir = Path(html_path).parent.resolve()
        html_stem = Path(html_path).stem

        candidate_dirs = [html_dir]
        for suffix in ['_files', '.files']:
            candidate_dirs.append(html_dir / f"{html_stem}{suffix}")
        candidate_dirs += list(html_dir.glob("*_files")) + list(html_dir.glob("*.files"))

        def _replace_img(match):
            tag = match.group(0)
            src_match = re.search(r'src="([^"]+)"', tag, re.IGNORECASE)
            if not src_match:
                return tag
            img_src = src_match.group(1).strip()
            if img_src.startswith(('http://', 'https://', 'data:')):
                return tag
            if not img_src.isascii():
                print(f"[DEBUG] 非 ASCII 图片路径，直接移除: {img_src}")
                return '<div style="width:120px;height:60px;border:1px dashed red;text-align:center;line-height:60px;">[非法路径]</div>'
            img_rel_path = Path(img_src)
            exists = any((cand_dir / img_rel_path).exists() for cand_dir in candidate_dirs)
            if exists:
                return tag
            print(f"[DEBUG] 替换不存在图片: {img_src}")
            return '<div style="width:120px;height:60px;border:1px dashed gray;text-align:center;line-height:60px;color:#666;">[图片缺失]</div>'

        return re.sub(r'<img[^>]+>', _replace_img, content, flags=re.IGNORECASE)

    content, encoding_used = read_file_with_fallback(html_path)
    if content is None:
        return False

    try:
        content = re.sub(r'<link[^>]*?>', '', content, flags=re.IGNORECASE)
        content = re.sub(r'<script[^>]*?>.*?</script>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<iframe[^>]*?>.*?</iframe>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<object[^>]*?>.*?</object>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<embed[^>]*?>', '', content, flags=re.IGNORECASE)
        content = re.sub(r'@font-face\s*{[^}]+}', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'url\(["\']?file:[^)]+["\']?\)', 'none', content, flags=re.IGNORECASE)
        content = replace_nonexistent_images_with_files_support(content, html_path)

        head_match = re.search(r'<head.*?>', content, flags=re.IGNORECASE)
        if head_match:
            head_end = head_match.end()
            if not re.search(r'<base\s', content, flags=re.IGNORECASE):
                abs_dir = Path(html_path).parent.resolve()
                base_tag = f'<base href="file:///{abs_dir.as_posix()}/">'
                content = content[:head_end] + base_tag + content[head_end:]
            if not re.search(r'<style[^>]*?>.*?</style>', content, flags=re.IGNORECASE | re.DOTALL):
                default_style = """
                <style>
                    body { font-family: Arial, sans-serif; font-size: 12pt; line-height: 1.6; }
                    h1, h2, h3 { color: #333; }
                    p { margin-bottom: 10px; }
                    img { max-width: 100%; height: auto; display: block; margin: 10px auto; }
                </style>
                """
                content = content[:head_end] + default_style + content[head_end:]

        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding=encoding_used) as tmp_file:
            tmp_file.write(content)
            tmp_html_path = tmp_file.name

        options = {
            'enable-local-file-access': None,
            'load-error-handling': 'ignore',
            'load-media-error-handling': 'ignore'
        }

        pdfkit.from_file(tmp_html_path, pdf_path, configuration=config, options=options)
        os.remove(tmp_html_path)
        return True

    except Exception as e:
        print(f"转换失败 {html_path}: {e}")
        if debug_on_fail:
            print(f"[DEBUG] 保留临时 HTML 文件用于调试: {tmp_html_path}")
        else:
            try:
                os.remove(tmp_html_path)
            except:
                pass
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
