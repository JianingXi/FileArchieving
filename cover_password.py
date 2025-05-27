import os
import re


def replace_string_in_file(file_path, old_str, new_str):
    """替换文件中的指定字符串"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        new_content = re.sub(r'\b' + re.escape(old_str) + r'\b', new_str, content)

        if new_content != content:
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(new_content)
            print(f"已替换文件: {file_path}")
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")


def process_directory(directory, old_str, new_str):
    """递归处理目录中的所有.py文件"""
    if not os.path.isdir(directory):
        print(f"错误：路径 '{directory}' 不存在或不是文件夹！")
        return

    print(f"处理目录: {directory}")
    print(f"替换规则: '{old_str}' → '{new_str}'")

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.py'):
                file_path = os.path.join(root, file)
                replace_string_in_file(file_path, old_str, new_str)

    print("替换完成")


# 直接调用（示例）
target_dir = r"C:\MyPython\ChaoxingUploads"  # 替换为目标路径
old_string = ""  # 要替换的旧字符串
new_string = "Password****"  # 新字符串

process_directory(target_dir, old_string, new_string)

old_string = ""  # 要替换的旧字符串
new_string = "PhoneNumber"  # 新字符串

process_directory(target_dir, old_string, new_string)