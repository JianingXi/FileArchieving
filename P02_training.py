import pickle
import numpy as np
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import StratifiedKFold, cross_val_score
from sklearn.feature_extraction.text import TfidfVectorizer

# 加载之前保存的数据（文件路径、文本和多标签）
with open('file_paths_texts_and_labels.pkl', 'rb') as f:
    file_paths, texts, multi_labels = pickle.load(f)

# 判断 multi_labels 的维度
y_multi = np.array(multi_labels)
if y_multi.ndim == 1:
    # 如果已经是一维，则直接作为标签（单标签分类）
    y = y_multi
else:
    # 如果是二维，则保留多标签数据（多标签分类）
    y = y_multi

# 将文本数据转化为特征矩阵
vectorizer = TfidfVectorizer()
X = vectorizer.fit_transform(texts)

# 定义随机森林模型，设置随机种子和树的数量
# 注意：RandomForestClassifier 可支持多输出（多标签）任务
rf = RandomForestClassifier(random_state=15, n_estimators=100)

# 如果是多标签分类，使用 StratifiedKFold 可能不适合，
# 此处示例中我们仍用 StratifiedKFold（对于多标签问题可考虑其他交叉验证方法）
try:
    cv = StratifiedKFold(n_splits=10, shuffle=True, random_state=42)
    scores = cross_val_score(rf, X, y, cv=cv, scoring='accuracy')
    print("各折准确率：", scores)
    print("平均准确率：", np.mean(scores))
except Exception as e:
    print("交叉验证评估时出现异常：", e)

# 用所有数据训练最终模型
rf.fit(X, y)

# 保存最终训练好的模型和 TF-IDF 向量化器
with open('rf_model.pkl', 'wb') as f:
    pickle.dump(rf, f)
with open('tfidf_vectorizer.pkl', 'wb') as f:
    pickle.dump(vectorizer, f)

print("随机森林模型和TF-IDF向量化器已成功保存")
