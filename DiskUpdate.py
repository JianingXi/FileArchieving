from DiskUpdateDict.FileNameAsDate import rename_date
from DiskUpdateDict.UpdateShortcut import update_shortcut_folders, check_and_delete_if_empty
from DiskUpdateDict.DiskUpdateFunc import compress_and_remove_folders, update_commercial2rar_files
from DiskUpdateDict.html2pdf import convert_html_files_in_directory

from rename_space import rename_files_and_folders

disk_char = "D:"





"""
basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform'
rename_date(basedir)
rename_files_and_folders(basedir)
"""

basedir = r'C:\MyDocument\ToDoList\D20_ToDailyNotice'
rename_date(basedir)
convert_html_files_in_directory(basedir)
compress_and_remove_folders(basedir)
rename_files_and_folders(basedir)

basedir = r'C:\MyDocument\ToDoList\D20_Done'
rename_date(basedir)
rename_files_and_folders(basedir)
basedir = r'C:\MyDocument\ToDoList\D20_Z_ToEvernote'
rename_date(basedir)
convert_html_files_in_directory(basedir)
rename_files_and_folders(basedir)
basedir = r'C:\MyDocument\ToDoList\D20_ToHardDisk'
rename_date(basedir)
rename_files_and_folders(basedir)



# Paper
basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20241115_冯景辉论文'
rename_date(basedir)
rename_files_and_folders(basedir)
basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20241214_黄思敏论文'
rename_date(basedir)
rename_files_and_folders(basedir)
basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_论文\D20250523_冯敏华论文'
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

basedir = r'C:\MyDocument\ToDoList\D20_ToHardDisk\D20250703_智能医学工程新专业写材料'
rename_date(basedir)
rename_files_and_folders(basedir)

basedir = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_比赛\24硕黄思敏_中国国际'
rename_date(basedir)
rename_files_and_folders(basedir)



# basedir = r'D:\Alpha\J机智\工作业务\Y2025'
"""
rename_date(basedir)
rename_files_and_folders(basedir)
"""

# 删除超星学习通的临时文件
cx_folder = r"D:\cxdownload"
cx_folder = cx_folder.replace("D:", disk_char)
check_and_delete_if_empty(cx_folder)




# 硬盘位置确定与硬盘快捷方式更新
# update_shortcut_folders(disk_char)


# 商业报告参考模板备份
update_commercial2rar_files(disk_char)

print('done!')

