import os
import re

def sanitize_name(name, is_file=False):
    """
    清洗路径名：去除首尾空格，合并空格，将空格和短横线转为下划线。
    文件和文件夹的点不做任何处理，直接保留。
    """
    # 去除首尾空格并合并空格
    name = re.sub(r'\s+', ' ', name.strip())
    # 将空格和短横线替换为下划线
    name = re.sub(r'[\s-]+', '_', name)

    return name

def rename_path(path):
    """
    对路径中的每一部分进行重命名，处理空格、短横线情况。
    """
    path_parts = path.split(os.sep)

    # 判断是否是文件
    is_file = os.path.isfile(path)
    new_parts = []

    for i, part in enumerate(path_parts):
        new_parts.append(sanitize_name(part, is_file=is_file))

    new_path = os.sep.join(new_parts)

    if new_path != path:
        try:
            os.rename(path, new_path)
            print(f"Renamed: {path} -> {new_path}")
        except Exception as e:
            print(f"Error renaming {path}: {e}")

    return new_path

def rename_files_and_folders(directory):
    """
    遍历目录树，自下而上重命名所有文件和文件夹。
    """
    for root, dirs, files in os.walk(directory, topdown=False):
        for filename in files:
            file_path = os.path.join(root, filename)
            rename_path(file_path)

        for dirname in dirs:
            dir_path = os.path.join(root, dirname)
            rename_path(dir_path)

    rename_path(directory)

if __name__ == "__main__":
    target_directory = r'C:\MyDocument\教材撰写\科学社专用MathType6点9以下'  # 修改为你想处理的路径
    rename_files_and_folders(target_directory)
