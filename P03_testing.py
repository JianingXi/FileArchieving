import os
import string
import shutil
import pickle
import sys
import zipfile
import rarfile
import pandas as pd
import nltk
from docx import Document
from pptx import Presentation
from sklearn.feature_extraction.text import TfidfVectorizer

# 增加递归深度限制
sys.setrecursionlimit(2000)

# 下载 NLTK 的停用词数据（第一次运行时会下载），若下载失败请手动配置 nltk_data
try:
    nltk.download('stopwords', quiet=True)
except Exception as e:
    print("下载 stopwords 出现错误：", e)
from nltk.corpus import stopwords

def preprocess_text(text):
    """
    对文本进行预处理：
      - 转换为小写
      - 去除所有标点符号
      - 去除英文停用词
    """
    text = text.lower()
    text = ''.join(char for char in text if char not in string.punctuation)
    stop_words = set(stopwords.words('english'))
    words = text.split()
    words = [word for word in words if word not in stop_words]
    return ' '.join(words)

def extract_text_from_txt(file_path, len_doc):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read(len_doc)  # 只读取前500个字符
        return content
    except Exception as e:
        print(f"读取 txt 文件 {file_path} 时出错: {e}")
        return ""

def extract_text_from_docx(file_path, len_doc):
    try:
        doc = Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        content = "\n".join(full_text)[:len_doc]
        return content
    except Exception as e:
        print(f"读取 docx 文件 {file_path} 时出错: {e}")
        return ""

def extract_text_from_pptx(file_path, len_doc):
    try:
        prs = Presentation(file_path)
        full_text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    full_text.append(shape.text)
        content = "\n".join(full_text)[:len_doc]
        return content
    except Exception as e:
        print(f"读取 pptx 文件 {file_path} 时出错: {e}")
        return ""

def extract_text_from_excel(file_path, len_doc):
    try:
        df_dict = pd.read_excel(file_path, sheet_name=None)
        texts = []
        for sheet_name, df in df_dict.items():
            texts.append(df.to_string())
        content = "\n".join(texts)[:len_doc]
        return content
    except Exception as e:
        print(f"读取 Excel 文件 {file_path} 时出错: {e}")
        return ""

def extract_text_from_zip(file_path, len_doc):
    """
    对于 ZIP 文件，读取压缩包内所有文件的名称，
    预处理后合并成一个字符串，并只保留前500个字符。
    """
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            names = zip_ref.namelist()
        processed_names = " ".join(preprocess_text(name) for name in names)
        return processed_names[:len_doc]
    except Exception as e:
        print(f"处理 zip 文件 {file_path} 时出错: {e}")
        return ""

def extract_text_from_rar(file_path, len_doc):
    """
    对于 RAR 文件，读取压缩包内所有文件的名称，
    预处理后合并成一个字符串，并只保留前500个字符。
    """
    try:
        with rarfile.RarFile(file_path) as rar_ref:
            names = rar_ref.namelist()
        processed_names = " ".join(preprocess_text(name) for name in names)
        return processed_names[:len_doc]
    except Exception as e:
        print(f"处理 rar 文件 {file_path} 时出错: {e}")
        return ""

def extract_features_from_file(filepath, len_doc):
    """
    根据文件后缀提取文本特征：
      - 对于支持的文本型文件（.txt, .docx, .pptx, .xls, .xlsx），读取内容的前500个字符；
      - 对于压缩包文件（.zip, .rar），读取压缩包内所有文件名称（预处理后合并为一个字符串）；
      - 其他类型文件，仅返回预处理后的文件名。
    最后将文件名与读取的内容组合，并只保留前500个字符。
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".txt":
        content = extract_text_from_txt(filepath, len_doc)
    elif ext == ".docx":
        content = extract_text_from_docx(filepath, len_doc)
    elif ext == ".pptx":
        content = extract_text_from_pptx(filepath, len_doc)
    elif ext in [".xls", ".xlsx"]:
        content = extract_text_from_excel(filepath, len_doc)
    elif ext == ".zip":
        content = extract_text_from_zip(filepath, len_doc)
    elif ext == ".rar":
        content = extract_text_from_rar(filepath, len_doc)
    else:
        content = ""
    # 提取并预处理文件名
    file_name = os.path.basename(filepath).lower().translate(str.maketrans('', '', string.punctuation))
    combined = file_name + " " + preprocess_text(content)
    return combined[:len_doc]

def get_all_file_paths(directory):
    """
    递归遍历指定目录，获取所有文件的完整路径。
    注意：仅扫描指定目录，不会遍历整个磁盘。
    """
    file_paths = []
    for root, dirs, files in os.walk(directory):
        for name in files:
            file_paths.append(os.path.join(root, name))
    return file_paths

def classify_files(source_directory, destination_root, len_doc):
    """
    对 source_directory 内的所有文件进行预测分类，
    特征由文件名和（文本型文件的前500字符/压缩包内文件名称）组成，
    然后将文件移动到 destination_root 下对应类别的文件夹中。
    """
    file_paths = get_all_file_paths(source_directory)
    texts = [extract_features_from_file(fp, len_doc) for fp in file_paths]

    # 加载训练好的模型和 TF-IDF 向量化器
    with open('rf_model.pkl', 'rb') as model_file:
        model = pickle.load(model_file)
    with open('tfidf_vectorizer.pkl', 'rb') as vectorizer_file:
        vectorizer = pickle.load(vectorizer_file)

    X = vectorizer.transform(texts)
    predictions = model.predict(X)

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

    for category in target_categories:
        os.makedirs(os.path.join(destination_root, category), exist_ok=True)

    for fp, pred in zip(file_paths, predictions):
        category = str(pred)
        dest_folder = os.path.join(destination_root, category)
        dest_path = os.path.join(dest_folder, os.path.basename(fp))
        print(f"移动文件 {fp} 到 {dest_path} （分类：{category}）")
        shutil.move(fp, dest_path)

if __name__ == "__main__":
    # 指定待分类的文件夹（请根据实际情况修改）
    source_directory = r"C:\Users\xijia\Desktop\ToDoList\D20_ToDailyNotice"  # 示例：文件列表所在目录
    # 目标根目录为当前目录下的 "DoneFileArchived" 文件夹
    destination_root = os.path.join(os.getcwd(), "DoneFileArchived")
    os.makedirs(destination_root, exist_ok=True)

    classify_files(source_directory, destination_root, 50)
    print("所有文件已根据预测分类并移动到目标文件夹中。")
