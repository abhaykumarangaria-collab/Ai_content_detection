import pandas as pd
import re
import joblib
import nltk
import os
import sys
from nltk.corpus import stopwords
from sklearn.linear_model import SGDClassifier

# --- CONFIG ---
DATA_PATH = "AI_Human.csv"
ALT_MODEL_PATH = "sgd_ai_model.pkl"
VECTORIZER_PATH = "vectorizer.pkl" 
CHUNK_SIZE = 50000 

print("🔍 Running System Checks...")

# 1. PREREQUISITE CHECKS
if not os.path.exists(DATA_PATH):
    print(f"❌ CRITICAL ERROR: Cannot find '{DATA_PATH}' in the current folder.")
    print("Please make sure your dataset is in the same folder as this script.")
    sys.exit()

if not os.path.exists(VECTORIZER_PATH):
    print(f"❌ CRITICAL ERROR: Cannot find '{VECTORIZER_PATH}'.")
    print("You must run your original training script first to generate the vectorizer.")
    sys.exit()

# 2. SETUP
nltk.download('stopwords', quiet=True)
stop_words = set(stopwords.words('english'))

def preprocess(text):
    text = str(text).lower()
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    words = [w for w in text.split() if w not in stop_words]
    return " ".join(words)

print("📊 Loading existing vectorizer...")
try:
    vectorizer = joblib.load(VECTORIZER_PATH)
except Exception as e:
    print(f"❌ ERROR LOADING VECTORIZER: {e}")
    sys.exit()

sgd_model = SGDClassifier(loss='log_loss', penalty='l2', alpha=1e-4, random_state=42)
classes = [0, 1] 

# 3. TRAINING
print("\n🚀 Training Alternative Model (SGD) on chunks...")
try:
    chunk_count = 0
    for i, chunk in enumerate(pd.read_csv(DATA_PATH, chunksize=CHUNK_SIZE)):
        print(f"🔄 Processing Batch {i+1}...")
        chunk.dropna(subset=['text'], inplace=True) # Only drop rows where 'text' is missing
        
        # Standardize the label column name
        if 'generated' in chunk.columns: 
            chunk.rename(columns={'generated': 'label'}, inplace=True)
            
        if 'label' not in chunk.columns or 'text' not in chunk.columns:
            print("❌ ERROR: Your CSV is missing the 'text' or 'label'/'generated' columns!")
            sys.exit()
            
        X_chunk = vectorizer.transform(chunk['text'].apply(preprocess))
        y_chunk = chunk['label'].astype(int)
        
        sgd_model.partial_fit(X_chunk, y_chunk, classes=classes)
        chunk_count += 1
        
    if chunk_count == 0:
        print("❌ ERROR: The CSV was read, but no valid data was found.")
        sys.exit()

except Exception as e:
    print(f"❌ FATAL ERROR DURING TRAINING: {e}")
    sys.exit()

# 4. SAVE MODEL
try:
    joblib.dump(sgd_model, ALT_MODEL_PATH)
    print(f"\n✨ SUCCESS! Alternative Model trained and saved as {ALT_MODEL_PATH}!")
except Exception as e:
    print(f"❌ ERROR SAVING FILE: {e}")
    sys.exit()