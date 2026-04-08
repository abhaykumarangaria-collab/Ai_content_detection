import pandas as pd
import re
import nltk
import joblib
import xgboost as xgb
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import TfidfVectorizer

# --- CONFIG ---
DATA_PATH = "AI_Human.csv"
MODEL_PATH = "xgb_ai_model.pkl"
VECTORIZER_PATH = "vectorizer.pkl"
CHUNK_SIZE = 50000 

print("📦 Downloading NLP tools...")
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

def preprocess(text):
    text = str(text).lower()
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    words = [w for w in text.split() if w not in stop_words]
    return " ".join(words)

# STEP 1: VECTORIZER
print("\n📊 Step 1: Fitting Vectorizer (100k row sample)...")
sample = pd.read_csv(DATA_PATH, nrows=100000)
sample.dropna(inplace=True)
sample['text'] = sample['text'].apply(preprocess)
vectorizer = TfidfVectorizer(max_features=10000, ngram_range=(1, 2))
vectorizer.fit(sample['text'])
joblib.dump(vectorizer, VECTORIZER_PATH)
print("✅ Vectorizer saved.")
del sample

# STEP 2: INCREMENTAL TRAINING
print("\n🚀 Step 2: Training on 500k rows...")
params = {
    'objective': 'binary:logistic', 
    'max_depth': 6, 
    'learning_rate': 0.05, 
    'eval_metric': 'logloss',
    'tree_method': 'hist'
}
bst = None 

for i, chunk in enumerate(pd.read_csv(DATA_PATH, chunksize=CHUNK_SIZE)):
    print(f"🔄 Processing Batch {i+1}...")
    chunk.dropna(inplace=True)
    if 'generated' in chunk.columns: 
        chunk.rename(columns={'generated': 'label'}, inplace=True)
    
    X = vectorizer.transform(chunk['text'].apply(preprocess))
    y = chunk['label'].astype(int)
    dtrain = xgb.DMatrix(X, label=y)
    bst = xgb.train(params, dtrain, num_boost_round=10, xgb_model=bst)

joblib.dump(bst, MODEL_PATH)
print("\n✨ Model trained and saved!")