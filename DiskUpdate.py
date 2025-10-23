from DiskUpdateDict.FileNameAsDate import rename_date
from DiskUpdateDict.UpdateShortcut import update_shortcut_folders, check_and_delete_if_empty
from DiskUpdateDict.DiskUpdateFunc import compress_and_remove_folders, update_commercial2rar_files, move_recent_items
from DiskUpdateDict.html2pdf import convert_html_files_in_directory

from rename_space import rename_files_and_folders


disk_char = "D:"
jizhi = False

# --------- 移动Downloads近n_days天文件至Daily Notice --------- #
src = r"C:\Users\xijia\Downloads"
dst = r"C:\MyDocument\ToDoList\D20_ToDailyNotice"
moved = move_recent_items(src, dst, n_days=3)
for f, path in moved:
    print(f"移动: {f} 到 {path}")


for i_loop in range(5):
    basedir = r'C:\MyDocument\ToDoList\D20_ToDailyNotice'
    rename_date(basedir)
    convert_html_files_in_directory(basedir)
    compress_and_remove_folders(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_ToHardDisk'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_Z_ToEvernote'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_Done'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    # basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    # -------------------------- Paper ---------------------------- #
    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20241214_黄思敏论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20241219_余宇论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250114_孔元元论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250607_梁钲禧论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250910_冯景辉论文2'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250901_中医院学生论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250916_黄颂论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250924_陈子翔论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20251010_曾振华论文'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    basedir = r'C:\MyDocument\ToDoList\D20_ToHardDisk\D20251020_智能医学工程专业开设的事情'
    rename_date(basedir)
    rename_files_and_folders(basedir)

    if jizhi:
        basedir = r'E:\Alpha\J机智\工作业务\Y2025'
        basedir = basedir.replace("D:", disk_char)
        rename_date(basedir)
        rename_files_and_folders(basedir)

# end of loop




# 删除超星学习通的临时文件
cx_folder = r"D:\cxdownload"
cx_folder = cx_folder.replace("D:", disk_char)
# check_and_delete_if_empty(cx_folder)


# 硬盘位置确定与硬盘快捷方式更新
year_str = "2025"
# update_shortcut_folders(disk_char, year_str)


# 商业报告参考模板备份
update_commercial2rar_files(disk_char)


print('done!')

