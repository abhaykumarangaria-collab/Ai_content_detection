import os
import requests
from dotenv import load_dotenv
from flask import Flask, request, render_template, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
import joblib
import re
import nltk
import xgboost as xgb
import base64
import heapq
from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize, word_tokenize

# --- LOAD ENVIRONMENT VARIABLES FIRST ---
load_dotenv()

# --- IMPORT YOUR NEW AI MODULE ---
try:
    from rag_engine import SlideRAGEngine
except ImportError:
    print("⚠️ Warning: rag_engine.py not found. Chat features may not work.")

# Check for Windows-only libraries (for PDF conversion)
try:
    import pythoncom
    import comtypes.client
    COMTYPES_AVAILABLE = True
except ImportError:
    COMTYPES_AVAILABLE = False
    print("⚠️ Warning: pythoncom/comtypes not found. Windows-specific PDF conversion is disabled.")


# --- APP CONFIGURATION ---
app = Flask(__name__)

app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'default-fallback-key-if-env-fails')

db_url = os.environ.get('DATABASE_URL', 'sqlite:///users.db')
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- DATABASE & LOGIN SETUP ---
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = "Please log in to access this page."

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(256), nullable=False)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

with app.app_context():
    db.create_all()

# --- NLTK SETUP ---
nltk.download("stopwords", quiet=True)
nltk.download("punkt", quiet=True)
stop_words = set(stopwords.words("english"))

# --- LOAD LOCAL ML MODELS (AI Detection) ---
print("📦 Loading AI Ensemble Models...")
try:
    model_xgb = joblib.load("xgb_ai_model.pkl")
    model_sgd = joblib.load("sgd_ai_model.pkl") 
    vectorizer = joblib.load("vectorizer.pkl")
    print("✅ System Ready: AI Ensemble & RAG Active")
except Exception as e:
    print(f"⚠️ Warning: Could not load local ML models: {e}")

# --- HELPER FUNCTIONS ---
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
                text_chunk = shape.text.strip()
                if len(text_chunk.split()) >= 5:
                    slide_text += text_chunk + " "
        content = slide_text.strip()
        if content:
            slide_data.append({'slide': i + 1, 'text': content})
            full_text += content + "\n"
    return full_text.strip(), slide_data

def get_ensemble_score(text):
    try:
        clean = preprocess(text)
        vec = vectorizer.transform([clean])
        if vec.nnz == 0: return 0.0
        prob_xgb = float(model_xgb.predict(xgb.DMatrix(vec))[0])
        prob_sgd = float(model_sgd.predict_proba(vec)[0][1])
        return (prob_xgb * 0.6) + (prob_sgd * 0.4)
    except:
        return 0.5 

def analyze_tone(text):
    text_lower = str(text).lower()
    tones = {
        "Academic 🎓": ["research", "analysis", "furthermore", "methodology", "theory", "hypothesis", "data", "conclusion", "study", "significant"],
        "Persuasive 🚀": ["must", "guarantee", "proven", "discover", "win", "best", "urgent", "critical", "invest", "action", "today"],
        "Professional 💼": ["strategy", "alignment", "execute", "leverage", "metrics", "synergy", "objectives", "revenue", "growth", "stakeholders"],
        "Casual ☕": ["cool", "stuff", "totally", "hey", "like", "maybe", "anyway", "awesome", "guess", "fun", "crazy"]
    }
    scores = {tone: 0 for tone in tones}
    words = word_tokenize(text_lower)
    for word in words:
        for tone, vocab in tones.items():
            if word in vocab: scores[tone] += 1
    dominant_tone = max(scores, key=scores.get)
    if scores[dominant_tone] == 0: return "Neutral / Objective 📊"
    return dominant_tone

def generate_summary(text, num_sentences=3):
    sentences = sent_tokenize(text.replace('\n', '. '))
    if len(sentences) <= num_sentences: return sentences
    words = word_tokenize(text.lower())
    word_frequencies = {}
    for word in words:
        if word not in stop_words and word.isalnum():
            word_frequencies[word] = word_frequencies.get(word, 0) + 1
    if not word_frequencies: return []
    max_frequency = max(word_frequencies.values())
    for word in word_frequencies.keys():
        word_frequencies[word] = (word_frequencies[word] / max_frequency)
    sentence_scores = {}
    for sent in sentences:
        for word in word_tokenize(sent.lower()):
            if word in word_frequencies.keys():
                if len(sent.split(' ')) < 30: 
                    sentence_scores[sent] = sentence_scores.get(sent, 0) + word_frequencies[word]
    return heapq.nlargest(num_sentences, sentence_scores, key=sentence_scores.get)

def highlight_ai_sentences(text):
    sentences = sent_tokenize(text.replace('\n', '. ')) 
    highlighted = []
    for s in sentences:
        original = s.strip()
        if len(original.split()) < 6:
            highlighted.append(f"<p style='color:#94a3b8; margin: 8px 0;'>⚪ {original}</p>")
            continue
        prob = get_ensemble_score(original)
        if 0.66 < prob < 0.68: prob = 0.0
        color = "#ff4d4d" if prob > 0.70 else ("#ffcc00" if prob > 0.40 else "#00ffcc")
        icon = "🔴" if prob > 0.70 else ("🟡" if prob > 0.40 else "🟢")
        score_text = f"({round(prob*100,1)}%)" if prob > 0 else "(Neutral)"
        highlighted.append(f"<p style='color:{color}; margin: 8px 0;'>{icon} {original} {score_text}</p>")
    return "".join(highlighted)

def convert_ppt_to_pdf(input_path, output_path):
    if not COMTYPES_AVAILABLE: raise Exception("comtypes library missing or not on Windows.")
    pythoncom.CoInitialize()
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        deck = powerpoint.Presentations.Open(os.path.abspath(input_path), WithWindow=False)
        deck.SaveAs(os.path.abspath(output_path), 32)
        deck.Close()
        powerpoint.Quit()
    finally:
        pythoncom.CoUninitialize()

def analyze_slide_structure(file_path):
    prs = Presentation(file_path)
    visual_feedback = []
    
    for i, slide in enumerate(prs.slides):
        word_count = 0
        has_image = False
        tiny_text_warnings = 0
        
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                has_image = True
                
            if hasattr(shape, "text") and shape.text:
                words = shape.text.split()
                word_count += len(words)
                
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.font and run.font.size:
                                font_size = run.font.size.pt
                                if font_size < 18:
                                    tiny_text_warnings += 1

        feedback = []
        if word_count > 40:
            feedback.append("⚠️ Too much text. Consider using bullet points and talking tracks.")
        if word_count > 10 and not has_image:
            feedback.append("🖼️ Text-heavy without visuals. Consider adding an image/diagram.")
        if tiny_text_warnings > 0:
            feedback.append(f"👓 Found {tiny_text_warnings} text elements smaller than 18pt. Might be hard to read.")
            
        if not feedback:
            feedback.append("✅ Good structure and readability.")

        visual_feedback.append({
            "slide": i + 1,
            "word_count": word_count,
            "has_image": has_image,
            "feedback": feedback
        })
        
    return visual_feedback

def check_plagiarism(text_summary_list):
    """Mock API call for checking plagiarism on summary sentences."""
    plagiarism_results = []
    
    for sentence in text_summary_list:
        if len(sentence.split()) > 5:
            try:
                # Simulating a hit for demonstration purposes
                if "research" in sentence.lower() or "proven" in sentence.lower():
                    plagiarism_results.append({
                        "sentence": sentence,
                        "flagged": True,
                        "possible_source": "Similarity found in online academic databases."
                    })
                else:
                    plagiarism_results.append({
                        "sentence": sentence,
                        "flagged": False,
                        "possible_source": "Original"
                    })
            except Exception as e:
                print(f"Plagiarism API error: {e}")
                
    flagged_count = sum(1 for item in plagiarism_results if item['flagged'])
    total_checked = len(plagiarism_results)
    
    originality_score = 100 if total_checked == 0 else int(((total_checked - flagged_count) / total_checked) * 100)

    return {
        "score": originality_score,
        "details": plagiarism_results
    }


# ==========================================
# 🔐 AUTHENTICATION ROUTES
# ==========================================
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if User.query.filter_by(username=username).first():
            flash("Username already exists!")
            return redirect(url_for('register'))
            
        hashed_pw = generate_password_hash(password, method='pbkdf2:sha256')
        new_user = User(username=username, password=hashed_pw)
        db.session.add(new_user)
        db.session.commit()
        
        login_user(new_user)
        return redirect(url_for('home'))
        
    return render_template('register.html', active_page='register')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password, password):
            login_user(user)
            return redirect(url_for('home'))
        else:
            flash('Invalid username or password')
            
    return render_template('login.html', active_page='login')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ==========================================
# 🌐 FRONTEND PAGE ROUTES
# ==========================================
@app.route("/")
@app.route("/detector")
def home(): 
    return render_template("detector.html", active_page="detector")

@app.route("/converter")
def converter(): 
    return render_template("converter.html", active_page="converter")

@app.route("/dashboard")
@login_required
def dashboard(): 
    return render_template("dashboard.html", active_page="dashboard")

@app.route("/history")
@login_required
def history(): 
    return render_template("history.html", active_page="history")

@app.route("/about")
def about(): 
    return render_template("about.html", active_page="about")

@app.route("/contact")
def contact(): 
    return render_template("contact.html", active_page="contact")


# ==========================================
# ⚙️ API & PROCESSING ROUTES
# ==========================================
@app.route("/api/convert", methods=["POST"])
def api_convert():
    file = request.files.get("file")
    conversion_type = request.form.get("type")
    
    if not file or file.filename == "": 
        return jsonify({"error": "No file"}), 400
        
    path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(file.filename))
    file.save(path)
    
    try:
        if conversion_type == "ppt-text":
            text, _ = extract_text_from_ppt(path)
            prob = get_ensemble_score(text)
            b64 = base64.b64encode(text.encode("utf-8")).decode("utf-8")
            return jsonify({"status": "success", "ai_probability": round(prob * 100, 2), "human_probability": round((1-prob)*100,2), "file_b64": b64, "mime_type": "text/plain", "extension": ".txt", "scan_complete": True})
        
        elif conversion_type == "text-ppt":
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = [line.strip() for line in f.readlines() if line.strip()]
            prs = Presentation()
            for i in range(0, len(lines), 5):
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"Slide {i//5 + 1}"
                tf = slide.shapes.placeholders[1].text_frame
                for j, line in enumerate(lines[i:i+5]):
                    if j == 0: tf.text = line
                    else: tf.add_paragraph().text = line
            out = os.path.join(app.config["UPLOAD_FOLDER"], "gen.pptx")
            prs.save(out)
            with open(out, "rb") as f: b64 = base64.b64encode(f.read()).decode("utf-8")
            try: os.remove(out)
            except Exception: pass
            return jsonify({"status": "success", "file_b64": b64, "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation", "extension": ".pptx", "scan_complete": False})
        
        elif conversion_type == "ppt-pdf":
            out = os.path.join(app.config["UPLOAD_FOLDER"], "conv.pdf")
            convert_ppt_to_pdf(path, out)
            with open(out, "rb") as f: b64 = base64.b64encode(f.read()).decode("utf-8")
            try: os.remove(out)
            except Exception: pass
            return jsonify({"status": "success", "file_b64": b64, "mime_type": "application/pdf", "extension": ".pdf", "scan_complete": False})
            
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            if os.path.exists(path): os.remove(path)
        except Exception: pass 

@app.route("/api/chat", methods=["POST"])
def api_chat():
    data = request.json
    question = data.get("question")
    slides = data.get("slides") 
    
    if not question or not slides:
        return jsonify({"error": "Missing question or slide data."}), 400

    try:
        rag = SlideRAGEngine()
        success = rag.ingest_slides(slides)
        if not success:
             return jsonify({"error": "Failed to process the slide data."}), 500
        result = rag.generate_answer(question)
        if "error" in result:
            return jsonify({"error": result["error"]}), 500
        return jsonify({"answer": result["answer"]})
    except Exception as e:
        return jsonify({"error": f"Chat engine error: {str(e)}"}), 500

@app.route("/predict_ppt", methods=["POST"])
def predict_ppt():
    file = request.files.get("file")
    if not file: return "No file"
    
    path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(file.filename))
    file.save(path)
    
    try:
        # Standard Processing
        text, slide_data = extract_text_from_ppt(path)
        final_prob = get_ensemble_score(text)
        
        slide_scores = []
        for s in slide_data:
            s_final = get_ensemble_score(s['text']) if len(s['text'].split()) > 5 else 0.0
            slide_scores.append({"slide": s['slide'], "score": round(s_final * 100, 1)})
            
        doc_tone = analyze_tone(text)
        doc_summary = generate_summary(text)
        
        # New Processing
        visual_analysis = analyze_slide_structure(path)
        plagiarism_report = check_plagiarism(doc_summary)
        
        return render_template("result.html", 
                               active_page="detector",
                               result="AI Detected" if final_prob > 0.5 else "Human Authored", 
                               probability=round(final_prob*100,2), 
                               highlighted_text=highlight_ai_sentences(text), 
                               slide_scores=slide_scores,
                               tone=doc_tone,
                               summary=doc_summary,
                               raw_slides=slide_data,
                               visual_analysis=visual_analysis,
                               plagiarism_report=plagiarism_report)
    finally:
        try:
            if os.path.exists(path): os.remove(path)
        except Exception: pass 

if __name__ == '__main__':
    # Get the PORT from the environment variables, default to 5000 if not found
    port = int(os.environ.get("PORT", 5000))
    # Bind to 0.0.0.0 so Render can route outside traffic to it
    app.run(host='0.0.0.0', port=port)