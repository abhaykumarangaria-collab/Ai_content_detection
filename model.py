from flask import Flask, request, render_template
from pptx import Presentation
import joblib
import re
import nltk
import xgboost as xgb
import os
from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

nltk.download("stopwords")
nltk.download("punkt")
stop_words = set(stopwords.words("english"))

# LOAD MODEL & VECTORIZER
try:
    model = joblib.load("xgb_ai_model.pkl")
    vectorizer = joblib.load("vectorizer.pkl")
    print("✅ System Ready: Model & Vectorizer Loaded")
except Exception as e:
    print(f"❌ Critical Error: {e}")
    exit()

def preprocess(text):
    text = str(text).lower()
    text = re.sub(r"[^a-zA-Z\s]", "", text)
    words = [w for w in text.split() if w not in stop_words]
    return " ".join(words)

def extract_text_from_ppt(file_path):
    prs = Presentation(file_path)
    full_text, slide_data = "", []
    for i, slide in enumerate(prs.slides):
        slide_text = ""
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                slide_text += shape.text + " "
        
        content = slide_text.strip()
        if content:
            slide_data.append({'slide': i + 1, 'text': content})
            full_text += content + "\n"
    return full_text.strip(), slide_data

def highlight_ai_sentences(text):
    sentences = sent_tokenize(text.replace('\n', '. ')) 
    highlighted = []
    
    for s in sentences:
        original = s.strip()
        word_count = len(original.split())
        
        # 1. Skip tiny phrases (Mark as Neutral)
        if word_count < 6:
            highlighted.append(f"<p style='color:#94a3b8; margin: 8px 0;'>⚪ {original}</p>")
            continue
            
        clean = preprocess(original)
        vec = vectorizer.transform([clean])
        
        # 2. Check if words are recognized
        if vec.nnz == 0:
            prob = 0.0
        else:
            prob = float(model.predict(xgb.DMatrix(vec))[0])
            
        # 3. The 67% Trap (Default guess prevention)
        if 0.66 < prob < 0.68:
            prob = 0.0
        
        color = "#ff4d4d" if prob > 0.70 else ("#ffcc00" if prob > 0.40 else "#00ffcc")
        icon = "🔴" if prob > 0.70 else ("🟡" if prob > 0.40 else "🟢")
            
        score_text = f"({round(prob*100,1)}%)" if prob > 0 else "(Neutral)"
        highlighted.append(f"<p style='color:{color}; margin: 8px 0;'>{icon} {original} {score_text}</p>")
        
    return "".join(highlighted)

@app.route("/")
def home(): return render_template("index.html")

@app.route("/predict_ppt", methods=["POST"])
def predict_ppt():
    file = request.files.get("file")
    if not file or file.filename == "": return "❌ No file selected."
    
    path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(file.filename))
    file.save(path)
    
    try:
        text, slide_data = extract_text_from_ppt(path)
        
        # OVERALL SCORE
        clean_full = preprocess(text)
        vec_full = vectorizer.transform([clean_full])
        prob = 0.0 if vec_full.nnz == 0 else float(model.predict(xgb.DMatrix(vec_full))[0])
        
        # SLIDE-WISE SCORES
        slide_scores = []
        for s in slide_data:
            clean_s = preprocess(s['text'])
            s_vec = vectorizer.transform([clean_s])
            
            # Aggregate analysis: Needs 5+ words to be valid
            if s_vec.nnz > 0 and len(s['text'].split()) > 5:
                s_prob = float(model.predict(xgb.DMatrix(s_vec))[0])
                if 0.66 < s_prob < 0.68: s_prob = 0.0 # Trap
            else:
                s_prob = 0.0 
                
            slide_scores.append({"slide": s['slide'], "score": round(s_prob * 100, 1)})

        return render_template("result.html", 
                               result="AI Content Detected" if prob > 0.5 else "Human Authored",
                               probability=round(prob * 100, 2),
                               highlighted_text=highlight_ai_sentences(text),
                               slide_scores=slide_scores)
    finally:
        if os.path.exists(path): os.remove(path)

if __name__ == "__main__": app.run(debug=True)