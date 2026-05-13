# Analyses of AI Content Detection 🕵️‍♂️🤖

https://verfai.onrender.com

![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.0.3-black?logo=flask&logoColor=white)
![XGBoost](https://img.shields.io/badge/XGBoost-1.7.6-orange?logo=xgboost&logoColor=white)
![Status](https://img.shields.io/badge/Status-Deployment_Ready-success)

A robust, full-stack web application designed to accurately distinguish between human-written and machine-generated text. By leveraging statistical NLP metrics (Perplexity and Burstiness) and an Extreme Gradient Boosting (XGBoost) ensemble model, this platform provides transparent, sentence-level authenticity verification. 

Beyond text detection, the application includes a dedicated file-conversion backend, PowerPoint slide structural analysis, and a Retrieval-Augmented Generation (RAG) chat engine.

---

## 🌟 Key Features

* **Hybrid Ensemble Detection:** Combines XGBoost and SGD classifiers to calculate a highly accurate AI-generation probability score.
* **Granular Sentence Highlighting:** Breaks down documents line-by-line, color-coding sentences based on their individual AI likelihood (🔴 High AI, 🟡 Mixed, 🟢 Human).
* **Tone & Summary Analysis:** Automatically detects the dominant tone of the text (Academic, Professional, Casual, etc.) and generates a weighted summary.
* **Advanced Presentation Processing:**
    * Extracts text directly from `.pptx` files for AI analysis.
    * Analyzes slide structure (word count, image presence, font size readability).
    * Converts PPT to PDF (Note: Requires a Windows environment with `comtypes` locally; safely bypassed on Linux cloud servers).
* **Slide RAG Engine:** Users can interactively chat with their uploaded presentation data using an integrated LLM.
* **Secure Authentication:** Full user registration and session management powered by `Flask-Login` and `SQLAlchemy`.

---

## 🛠️ Technology Stack

**Backend & Core Logic**
* **Framework:** Python 3.11 / Flask
* **Database:** SQLite (Local) / PostgreSQL (Production) via SQLAlchemy
* **Authentication:** Werkzeug Security, Flask-Login

**Machine Learning & NLP**
* **Models:** XGBoost, Scikit-Learn (SGD Classifier, TF-IDF Vectorizer)
* **NLP Processing:** NLTK (Tokenization, Stopwords)
* **RAG Engine:** Custom `SlideRAGEngine` integration

**Frontend**
* HTML5, CSS3, JavaScript
* Jinja2 Templating

---

## 🚀 Local Installation & Setup

To run this project on your local machine, follow these steps:

### 1. Clone the Repository
```bash
git clone [https://github.com/abhaykumarangaria-collab/Ai_content_detection.git](https://github.com/abhaykumarangaria-collab/Ai_content_detection.git)
cd Ai_content_detection
