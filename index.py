from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# 複数の文章
docs = [
    "私は犬が好きです",
    "私は猫が好きです",
    "今日は天気がいいです"
]

# TF-IDFで数値化
vectorizer = TfidfVectorizer()
tfidf = vectorizer.fit_transform(docs)
