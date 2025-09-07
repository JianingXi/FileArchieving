import os
import shutil

def move_files(src_folder, base_dest, year, sub_path):
    """
    将 src_folder 下的所有内容移动到 base_dest\Store{year}\sub_path
    - 文件遇到同名则覆盖
    - 文件夹遇到同名则合并（不会删除目标已有内容）
    - 空文件夹不会破坏目标文件夹
    """
    dest_folder = os.path.join(base_dest, f"Store{year}", sub_path)
    os.makedirs(dest_folder, exist_ok=True)

    for item in os.listdir(src_folder):
        src_item = os.path.join(src_folder, item)
        dest_item = os.path.join(dest_folder, item)

        try:
            if os.path.isfile(src_item):
                # 文件：覆盖
                shutil.move(src_item, dest_item)
                print(f"已移动文件: {src_item} -> {dest_item}")

            elif os.path.isdir(src_item):
                # 文件夹：合并
                if not os.path.exists(dest_item):
                    shutil.move(src_item, dest_item)
                    print(f"已移动文件夹: {src_item} -> {dest_item}")
                else:
                    # 合并内容
                    for root, _, files in os.walk(src_item):
                        rel_path = os.path.relpath(root, src_item)
                        target_dir = os.path.join(dest_item, rel_path)
                        os.makedirs(target_dir, exist_ok=True)
                        for f in files:
                            src_file = os.path.join(root, f)
                            dest_file = os.path.join(target_dir, f)
                            if os.path.exists(dest_file):
                                os.remove(dest_file)  # 覆盖文件
                            shutil.move(src_file, dest_file)
                            print(f"已移动文件: {src_file} -> {dest_file}")
                    # 清理空源目录
                    shutil.rmtree(src_item, ignore_errors=True)

        except Exception as e:
            print(f"移动失败: {src_item}, 错误: {e}")






if __name__ == "__main__":
    disk_char = "E:"
    year = "2025"  # 可以改

    src = r"C:\MyPython\FileArchieving\DoneFileArchived"
    
    base_dest = r"D:\Alpha\StoreLatestYears"
    base_dest = base_dest.replace("D:", disk_char)
    sub_path = r"M02广医事务性工作"

    move_files(src, base_dest, year, sub_path)
