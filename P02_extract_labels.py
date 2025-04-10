import os
import pickle

def generate_labels_and_filter(file_paths, texts, target_categories):
    """
    根据目标分类对每个文件路径生成标签：
      - 跳过所有 .lnk 文件
      - 如果文件所在文件夹的名称（归一化后）在 target_categories 中，
        则标签为该文件夹名称；
      - 否则标签设为 "其他"。
    同时过滤掉跳过的文件，并返回新的文件路径、文本和标签列表。
    """
    filtered_file_paths = []
    filtered_texts = []
    labels = []

    # 对目标分类进行归一化处理，用于比较
    normalized_target = set(os.path.normcase(os.path.normpath(cat)) for cat in target_categories)

    for path, text in zip(file_paths, texts):
        # 跳过 .lnk 文件
        if path.lower().endswith('.lnk'):
            continue
        # 提取最后一级文件夹名称，并归一化
        folder_name = os.path.basename(os.path.dirname(os.path.normpath(path)))
        folder_name = os.path.normcase(folder_name)

        filtered_file_paths.append(path)
        filtered_texts.append(text)
        # 如果文件所在文件夹在目标分类中，则标签为该文件夹名称，否则设为 "其他"
        if folder_name in normalized_target:
            labels.append(folder_name)
        else:
            labels.append("其他")
    return filtered_file_paths, filtered_texts, labels


def add_labels_to_features(features_file_str, target_categories, save_file_str):
    """
    从指定的 pickle 文件中加载数据（文件路径和文本，忽略已有标签），
    调用 generate_labels_and_filter 生成新的标签（并过滤 .lnk 文件），
    然后保存过滤后的数据到新的 pickle 文件中。
    """
    with open(features_file_str, 'rb') as f:
        data = pickle.load(f)

    if len(data) == 2:
        file_paths, texts = data
    elif len(data) == 3:
        file_paths, texts, _ = data
    else:
        raise ValueError("未知的数据格式，期望包含2或3个元素。")

    new_file_paths, new_texts, labels = generate_labels_and_filter(file_paths, texts, target_categories)

    with open(save_file_str, 'wb') as f:
        pickle.dump((new_file_paths, new_texts, labels), f)

    print("文件路径、文本和标签已成功保存到", save_file_str)



# 示例目标分类列表（这些字符串应与文件路径最后一级文件夹名称完全一致）
target_categories = [
    "产学研_产业化工作",
    "产学研_社科科普",
    "产学研_科研工作",
    "人事工作_人才帽子",
    "人事工作_人才补贴政策",
    "人事工作_出境公开",
    "人事工作_提拔调动升职称",
    "人事工作_教师招聘或教师培训",
    "人事工作_职称评审",
    "外界公司",
    "学校党政行政_红头文件",
    "学院党政行政_红头文件",
    "工会后勤宣传保卫等工作",
    "年终绩效总结",
    "教学人才培养_教学学生竞赛",
    "教学人才培养_教改项目论文",
    "教学人才培养_本科教学",
    "教学人才培养_研究生教学",
    "日常整治与安全检查",
    "评审委员_校内校外",
    "财务_经费提醒"
]

# 调用函数进行处理（请确保 'file_paths_texts_and_labels.pkl' 存在且格式正确）
add_labels_to_features('file_paths_and_texts.pkl', target_categories, 'file_paths_texts_and_labels.pkl')
