import pickle
import numpy as np
from sentence_transformers import SentenceTransformer
from sklearn.ensemble import RandomForestClassifier
from sklearn.multioutput import MultiOutputClassifier
from sklearn.preprocessing import MultiLabelBinarizer
from sklearn.model_selection import KFold
from sklearn.metrics import accuracy_score, f1_score

# 你提供的 22 个标签，顺序固定
TARGET_CATEGORIES = [
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

# 1. 加载数据
with open('file_paths_texts_and_labels.pkl', 'rb') as f:
    file_paths, texts, multi_labels = pickle.load(f)

# 2. 规范化 multi_labels：确保每个样本标签都只保留在 TARGET_CATEGORIES 里
normalized = []
for lab in multi_labels:
    # 如果原来是列表，遍历；否则当作单个字符串
    labs = lab if isinstance(lab, (list,tuple)) else [lab]
    # 过滤，只保留 TARGET_CATEGORIES 中的
    cleaned = [x for x in labs if x in TARGET_CATEGORIES]
    # 如果一个都不在，就保留空列表，让它在训练时对应全零向量
    normalized.append(cleaned)

# 3. 初始化好 classes 的 ML-Binarizer，先 fit 再 transform
mlb = MultiLabelBinarizer(classes=TARGET_CATEGORIES)
mlb.fit([])                 # 此处不传数据，只定死 classes_
Y = mlb.transform(normalized)  # 绝不会再报 unknown class

# 4. SBERT 嵌入
embedder = SentenceTransformer('all-MiniLM-L6-v2')
X = embedder.encode(texts, batch_size=32, show_progress_bar=True)

# 5. 定义多标签随机森林
base_rf = RandomForestClassifier(
    n_estimators=200,
    class_weight='balanced',
    n_jobs=-1,
    random_state=42
)
model = MultiOutputClassifier(base_rf, n_jobs=-1)

# 6. 五折交叉验证评估
kf = KFold(n_splits=5, shuffle=True, random_state=42)
accs, f1s = [], []
for tr_idx, te_idx in kf.split(X):
    model.fit(X[tr_idx], Y[tr_idx])
    Yp = model.predict(X[te_idx])
    accs.append(accuracy_score(Y[te_idx], Yp))
    f1s.append(f1_score(Y[te_idx], Yp, average='micro'))

print("CV 准确率：", accs, "平均：", np.mean(accs))
print("CV 微平均 F1：", f1s, "平均：", np.mean(f1s))

# 7. 全量训练并保存
model.fit(X, Y)
with open('rf_model.pkl', 'wb') as fm:
    pickle.dump(model, fm)
with open('label_binarizer.pkl', 'wb') as fb:
    pickle.dump(mlb, fb)

print("训练完成，已保存多标签模型和标签映射：")
print("  - rf_model.pkl")
print("  - label_binarizer.pkl (classes = TARGET_CATEGORIES)")
