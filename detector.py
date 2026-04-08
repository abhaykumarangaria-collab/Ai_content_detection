import joblib
import re
import nltk
import numpy as np

from nltk.corpus import stopwords

# ------------------------
# LOAD MODEL + VECTORIZER
# ------------------------
model = joblib.load("xgb_ai_model.pkl")
vectorizer = joblib.load("vectorizer.pkl")

nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# ------------------------
# PREPROCESS
# ------------------------
def preprocess(text):
    text = str(text).lower()
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    words = text.split()
    words = [w for w in words if w not in stop_words]
    return " ".join(words)

# ------------------------
# FEATURE ENGINEERING
# ------------------------
def get_features(text):
    words = text.split()

    if len(words) == 0:
        return [0, 0, 0]

    # 1. Avg word length
    avg_word_len = np.mean([len(w) for w in words])

    # 2. Lexical diversity
    diversity = len(set(words)) / len(words)

    # 3. Repetition score
    repetition = 1 - diversity

    return [avg_word_len, diversity, repetition]

# ------------------------
# DETECT FUNCTION
# ------------------------
def detect_ai(text):

    clean_text = preprocess(text)

    # TF-IDF prediction
    X = vectorizer.transform([clean_text])
    model_prob = model.predict_proba(X)[0][1]  # AI probability

    # Feature score
    features = get_features(clean_text)

    avg_len, diversity, repetition = features

    # Heuristic scoring
    heuristic_score = 0

    # AI tends to have moderate word length
    if 4 < avg_len < 6:
        heuristic_score += 0.2

    # AI tends to be less diverse
    if diversity < 0.6:
        heuristic_score += 0.4

    # AI tends to repeat patterns
    if repetition > 0.4:
        heuristic_score += 0.4

    # Final score (combine)
    final_score = (0.7 * model_prob) + (0.3 * heuristic_score)

    return {
        "AI_Probability (%)": round(final_score * 100, 2),
        "Human_Probability (%)": round((1 - final_score) * 100, 2),
        "Model_Score": round(model_prob, 3),
        "Heuristic_Score": round(heuristic_score, 3)
    }

# ------------------------
# TEST INPUT
# ------------------------
if __name__ == "__main__":

    text = input("\nEnter text (PPT / paragraph):\n\n")

    result = detect_ai(text)

    print("\n🧠 Detection Result:\n")
    print(f"AI Probability     : {result['AI_Probability (%)']}%")
    print(f"Human Probability  : {result['Human_Probability (%)']}%")
    print(f"Model Score        : {result['Model_Score']}")
    print(f"Heuristic Score    : {result['Heuristic_Score']}")