import os
import re
import pdfkit
import shutil
import tempfile
from pathlib import Path

# 配置 wkhtmltopdf 路径（根据实际安装路径修改）
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')


def read_file_with_fallback(file_path: str):
    """
    智能读取 HTML，尽量和它自己声明的 charset 对上
    """
    raw = Path(file_path).read_bytes()
    head_part = raw[:2048]
    m = re.search(rb'charset=["\']?([\w\-]+)', head_part, re.IGNORECASE)
    guess = None
    if m:
        guess = m.group(1).decode('ascii', 'ignore').lower()
        if guess in ('gb2312', 'gb-2312'):
            guess = 'gbk'
    candidates = []
    if guess:
        candidates.append(guess)
    candidates += ['utf-8-sig', 'utf-8', 'gbk', 'gb18030', 'cp936']
    for enc in candidates:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    print(f"读取 {file_path} 失败，编码不识别")
    return None, None


def write_file_with_encoding(file_path, content, encoding):
    with open(file_path, 'w', encoding=encoding, newline='') as f:
        f.write(content)


def relocate_local_images_to_ascii_dir(content: str, html_path: str):
    """
    ★ 关键函数 ★
    把 HTML 里引用的本地图片（特别是路径里有中文的）拷贝到一个全英文临时目录，
    然后把 <img src="..."> 重写成指向这个临时目录的 file:/// 路径。
    返回：新内容、临时目录路径（方便后面删）
    """
    html_dir = Path(html_path).parent.resolve()
    html_stem = Path(html_path).stem

    # 所有可能的资源目录
    candidate_dirs = [html_dir]
    for suffix in ['_files', '.files']:
        candidate_dirs.append(html_dir / f"{html_stem}{suffix}")
    candidate_dirs += list(html_dir.glob("*_files")) + list(html_dir.glob("*.files"))

    # 建一个全英文的临时目录
    tmp_assets_dir = Path(tempfile.mkdtemp(prefix="wkhtml_assets_"))

    img_counter = 0

    def _is_remote(src: str):
        return src.startswith(("http://", "https://", "data:"))

    def _exists_in_candidates(rel_path: Path):
        for cand in candidate_dirs:
            p = cand / rel_path
            if p.exists():
                return p
        return None

    def _replace(match):
        nonlocal img_counter
        tag = match.group(0)
        src_m = re.search(r'src="([^"]+)"', tag, re.IGNORECASE)
        if not src_m:
            return tag
        src = src_m.group(1).strip()

        # 1) 远程图、data 图不动
        if _is_remote(src):
            return tag

        # 2) 先尝试按原样找一遍
        rel_path = Path(src)
        found = _exists_in_candidates(rel_path)
        if not found:
            # 有些 OA 会把中文 URL encode 一下
            from urllib.parse import unquote
            decoded = unquote(src)
            if decoded != src:
                found = _exists_in_candidates(Path(decoded))

        if not found:
            # 找不到就保持原样，别塞 [非法路径]
            return tag

        # 3) 找到了，但路径里有中文/空格/奇怪字符，就搬到临时目录
        need_copy = any(ord(ch) > 127 for ch in str(found)) or " " in str(found)
        if need_copy:
            img_counter += 1
            ext = found.suffix or ".bin"
            new_name = f"img_{img_counter}{ext}"
            new_path = tmp_assets_dir / new_name
            try:
                shutil.copy2(found, new_path)
            except Exception as e:
                print(f"[IMG] 拷贝图片失败 {found} -> {new_path}: {e}")
                return tag  # 出错就不改
            # 用绝对 file:/// 路径，全部 /
            new_src = f"file:///{new_path.as_posix()}"
            new_tag = tag.replace(src, new_src)
            return new_tag
        else:
            # 路径本身就全 ascii，就不用拷贝，直接用原始相对路径
            return tag

    new_content = re.sub(r'<img[^>]+>', _replace, content, flags=re.IGNORECASE)
    return new_content, tmp_assets_dir


def html_to_pdf(html_path, pdf_path, debug_on_fail=False):
    content, encoding_used = read_file_with_fallback(html_path)
    if content is None:
        return False

    tmp_assets_dir = None  # 留着后面删

    try:
        # 删外链、脚本之类
        content = re.sub(r'<link[^>]*?>', '', content, flags=re.IGNORECASE)
        content = re.sub(r'<script[^>]*?>.*?</script>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<iframe[^>]*?>.*?</iframe>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<object[^>]*?>.*?</object>', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'<embed[^>]*?>', '', content, flags=re.IGNORECASE)
        content = re.sub(r'@font-face\s*{[^}]+}', '', content, flags=re.IGNORECASE | re.DOTALL)
        content = re.sub(r'url\(["\']?file:[^)]+["\']?\)', 'none', content, flags=re.IGNORECASE)

        # ① 图片路径重写到英文临时目录
        content, tmp_assets_dir = relocate_local_images_to_ascii_dir(content, html_path)

        # ② 统一 head 里的 charset 和 base
        head_match = re.search(r'<head.*?>', content, flags=re.IGNORECASE)
        if head_match:
            head_end = head_match.end()
            # 干掉原来的 meta charset
            content = re.sub(r'<meta[^>]+charset[^>]*?>', '', content, flags=re.IGNORECASE)
            # 塞自己的
            content = content[:head_end] + f'<meta charset="{encoding_used}">' + content[head_end:]

            # 塞 base，保证相对资源还能找
            if not re.search(r'<base\s', content, flags=re.IGNORECASE):
                abs_dir = Path(html_path).parent.resolve()
                base_tag = f'<base href="file:///{abs_dir.as_posix()}/">'
                content = content[:head_end] + base_tag + content[head_end:]

            # 没样式就给一个
            if not re.search(r'<style[^>]*?>.*?</style>', content, flags=re.IGNORECASE | re.DOTALL):
                style = """
                <style>
                    body { font-family: Arial, "Microsoft YaHei", "SimSun", sans-serif; font-size: 12pt; line-height: 1.6; }
                    img { max-width: 100%; height: auto; display: block; margin: 10px auto; }
                </style>
                """
                content = content[:head_end] + style + content[head_end:]

        # 写临时 HTML
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w',
                                         encoding=encoding_used, newline='') as tmp_file:
            tmp_file.write(content)
            tmp_html_path = tmp_file.name

        options = {
            'enable-local-file-access': None,
            'load-error-handling': 'ignore',
            'load-media-error-handling': 'ignore'
        }
        pdfkit.from_file(tmp_html_path, pdf_path, configuration=config, options=options)

        # 清理临时 html
        os.remove(tmp_html_path)

        # ✅ 转完再删原 html
        if Path(html_path).exists():
            os.remove(html_path)
            print(f"✅ 已删除 HTML 文件：{html_path}")

        # ✅ 删 _files / .files
        html_stem = Path(html_path).stem
        html_dir = Path(html_path).parent
        for suffix in ['_files', '.files']:
            folder = html_dir / f"{html_stem}{suffix}"
            if folder.exists():
                shutil.rmtree(folder, ignore_errors=True)
                print(f"✅ 已删除文件夹：{folder}")

        # ✅ 删我们刚刚建的英文临时图片目录
        if tmp_assets_dir and tmp_assets_dir.exists():
            shutil.rmtree(tmp_assets_dir, ignore_errors=True)

        return True

    except Exception as e:
        print(f"转换失败 {html_path}: {e}")
        if debug_on_fail:
            print(f"[DEBUG] 保留临时文件用于调试: {html_path}")
        # 出错也要尽量删临时图片目录
        if tmp_assets_dir and tmp_assets_dir.exists():
            shutil.rmtree(tmp_assets_dir, ignore_errors=True)
        return False


def convert_html_files_in_directory(directory):
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith(('.htm', '.html')):
                html_path = os.path.join(root, filename)
                pdf_path = os.path.splitext(html_path)[0] + '.pdf'
                print(f"正在处理: {html_path}")
                ok = html_to_pdf(html_path, pdf_path, False)
                if ok and os.path.exists(pdf_path):
                    print(f"转换成功: {pdf_path}")
                else:
                    print(f"⚠ 转换失败: {html_path}")


if __name__ == "__main__":
    target_directory = r"C:\MyDocument\ToDoList\D20_ToDailyNotice"   # 你自己的目录
    if not os.path.isdir(target_directory):
        print("错误: 指定的路径不是一个有效的目录!")
        exit(1)

    print(f"开始处理目录: {target_directory}")
    print("将把所有HTML文件(.htm, .html)转换为PDF并删除原始文件")
    convert_html_files_in_directory(target_directory)
    print("处理完成!")
