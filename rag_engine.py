import os
import re
import logging
import numpy as np
from typing import List, Dict, Any
from google import genai
from openai import OpenAI
from dotenv import load_dotenv

# Set up basic logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

class SlideRAGEngine:
    def __init__(self):
        """Initializes the API Clients."""
        load_dotenv()
        gemini_key = os.getenv("GEMINI_API_KEY")
        chat_key = os.getenv("GROQ_API_KEY") 
        
        self.gemini_client = None
        self.chat_client = None
        self.slide_data: List[Dict[str, Any]] = []
        self.slide_embeddings: np.ndarray = None
        
        # 1. Initialize Gemini (Strictly for Embeddings)
        if not gemini_key:
            logging.error("GEMINI_API_KEY is missing from .env")
        else:
            try:
                self.gemini_client = genai.Client(api_key=gemini_key)
                logging.info("Gemini Client initialized for embeddings.")
            except Exception as e:
                logging.error(f"Gemini Client failed: {e}")

        # 2. Initialize Groq (Strictly for Chat Generation)
        if not chat_key:
            logging.error("GROQ_API_KEY is missing from .env")
        else:
            try:
                self.chat_client = OpenAI(
                    api_key=chat_key,
                    base_url="https://api.groq.com/openai/v1",
                )
                logging.info("Groq Chat Client initialized.")
            except Exception as e:
                logging.error(f"Groq Chat Client failed: {e}")

    def ingest_slides(self, slide_data: List[Dict[str, Any]]) -> bool:
        """Embeds and caches slide data using Gemini."""
        if not self.gemini_client: return False

        self.slide_data = [s for s in slide_data if s.get('text', '').strip()]
        if not self.slide_data: return False

        texts_to_embed = [s['text'] for s in self.slide_data]
        
        try:
            logging.info("Generating embeddings for slides with Gemini...")
            all_embeddings = []
            
            for text in texts_to_embed:
                response = self.gemini_client.models.embed_content(
                    model='gemini-embedding-001', 
                    contents=text
                )
                all_embeddings.append(response.embeddings[0].values)

            self.slide_embeddings = np.array(all_embeddings)
            logging.info(f"Successfully embedded {len(self.slide_data)} slides.")
            return True
        except Exception as e:
            logging.error(f"Embedding API failed: {e}")
            return False

    def retrieve_relevant_slides(self, question: str, top_k: int = 3, threshold: float = 0.45) -> str:
        """Finds the most relevant slides using Smart Detection + Cosine Similarity."""
        if not self.gemini_client or self.slide_embeddings is None: return ""

        context = ""
        
        # --- SMART DETECTOR: Explicitly catch "Slide X" requests ---
        match = re.search(r'slide\s*(\d+)', question.lower())
        if match:
            explicit_slide_num = int(match.group(1))
            found_exact_slide = False
            
            for slide in self.slide_data:
                if slide.get('slide') == explicit_slide_num:
                    context += f"--- Slide {slide['slide']} (Direct Match) ---\n{slide['text']}\n\n"
                    found_exact_slide = True
                    break
            
            if found_exact_slide:
                # If we found the specific slide, RETURN ONLY THAT SLIDE. 
                return context
            else:
                # If the slide was filtered out (no text / images only), stop the AI from guessing.
                return f"SYSTEM INSTRUCTION: The user is asking about Slide {explicit_slide_num}. Inform the user that this specific slide contains no readable text (it is likely a visual-only slide with charts/images) or it does not exist in the deck."

        # --- RAG LOGIC: Vector math for topic-based questions ---
        try:
            response = self.gemini_client.models.embed_content(
                model='gemini-embedding-001', 
                contents=question
            )
            query_embedding = np.array(response.embeddings[0].values)

            dot_products = np.dot(self.slide_embeddings, query_embedding)
            norms_slides = np.linalg.norm(self.slide_embeddings, axis=1)
            norm_query = np.linalg.norm(query_embedding)
            similarities = dot_products / (norms_slides * norm_query)

            top_indices = np.argsort(similarities)[::-1][:top_k]

            for idx in top_indices:
                sim_score = similarities[idx]
                if sim_score > threshold:
                    slide = self.slide_data[idx]
                    context += f"--- Slide {slide.get('slide', 'Unknown')} (Relevance: {sim_score:.2f}) ---\n{slide['text']}\n\n"
            
            return context
        except Exception as e:
            logging.error(f"Retrieval failed: {e}")
            return context 

    def generate_answer(self, question: str) -> Dict[str, str]:
        """Generates the final answer using Groq Llama 3."""
        if not self.chat_client:
            return {"error": "Chat API client is not initialized. Check GROQ_API_KEY."}
        if self.slide_embeddings is None:
            return {"error": "Slides have not been ingested."}

        context = self.retrieve_relevant_slides(question)
        if not context:
            context = "No specific slides contain highly relevant information. Rely on general knowledge, but inform the user."

        system_prompt = f"""You are an expert Presentation Assistant. Your task is to answer the user's question based strictly on the provided slide context.

RULES:
1. Base your answer primarily on the provided context.
2. If the answer is found in the context, explicitly mention which slide it came from (e.g., "According to Slide 4...").
3. If the context does not contain the answer, you may use your general knowledge, but you MUST state clearly: "This information is not explicitly covered in the presentation deck."
4. Keep the response clear, professional, and concise.

CONTEXT (Relevant Slides / System Instructions):
{context}"""
        
        try:
            response = self.chat_client.chat.completions.create(
                model="llama-3.3-70b-versatile", 
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": question}
                ]
            )
            return {"answer": response.choices[0].message.content}
        except Exception as e:
            logging.error(f"Chat generation failed: {e}")
            return {"error": str(e)}