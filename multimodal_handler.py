import io
import os
import zipfile
import logging
from typing import Optional, List
from google import genai
from google.genai import types
import fitz  # PyMuPDF

import config

logger = logging.getLogger("MedusaMultimodal")

SYSTEM_INSTRUCTION = (
    "You are Medusa AI, an intelligent, highly perceptive visual and document analysis assistant. "
    "When shown an image, your primary strength is accurate identification and deep recognition. "
    "Specifically identify and name any people, anime/manga/game/movie/comic characters, celebrities, "
    "mythological figures, landmarks, plants, animals, or objects. "
    "If the user asks 'who is she?', 'who is he?', 'who is this?', or asks to identify a subject: "
    "Directly state their exact name, the franchise/series/origin they are from, and their background or role. "
    "CRITICAL RULE: Never disclose or mention any underlying model names (e.g., Gemini, Google, Groq, "
    "Llama, Qwen), model versions, or backend infrastructure. Always identify yourself solely as Medusa AI."
)

FALLBACK_MODELS = [
    config.GEMINI_MODEL,
    "gemini-3.8-flash",
    "gemini-2.5-flash",
]


class MultimodalAnalyzer:
    """Handles multimodal document, image, audio, and archive analysis."""

    def __init__(self, api_key: str = config.GEMINI_API_KEY):
        self.api_key = api_key
        self._client: Optional[genai.Client] = None

    @property
    def client(self) -> genai.Client:
        if self._client is None:
            if not self.api_key:
                raise ValueError("GEMINI_API_KEY is not configured.")
            self._client = genai.Client(api_key=self.api_key)
        return self._client

    def _call_gemini_with_fallback(self, contents: list, temperature: float = 0.3) -> str:
        """Calls Gemini with automatic fallback to alternate models on error."""
        last_error = None
        # Deduplicate models preserving order
        models_to_try = []
        for m in FALLBACK_MODELS:
            if m and m not in models_to_try:
                models_to_try.append(m)

        for model_name in models_to_try:
            try:
                response = self.client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        temperature=temperature,
                    ),
                )
                if response and response.text:
                    return response.text.strip()
            except Exception as e:
                logger.warning(f"Model {model_name} failed: {e}. Trying next fallback...")
                last_error = e

        if last_error:
            raise last_error
        return "I was unable to analyze the content."

    def analyze_image(self, image_bytes: bytes, mime_type: str = "image/jpeg", user_prompt: Optional[str] = None) -> str:
        """
        Analyzes an image. Specifically optimizes for recognizing people,
        characters (anime, gaming, pop-culture), objects, and scenes.
        """
        cleaned_prompt = user_prompt.strip() if user_prompt else ""

        if not cleaned_prompt:
            prompt = (
                "Identify and analyze who or what is shown in this image in detail. "
                "Specifically identify and name any characters (anime, manga, video games, movies, comics), "
                "people, celebrities, landmarks, or objects. "
                "State their exact name, their origin or franchise, and key details about them."
            )
        else:
            # If the user asked "who is she?", "who is this?", etc., make it explicit
            prompt = (
                f"User Question: '{cleaned_prompt}'\n\n"
                "Task: Accurately identify the person, character, or object in this image to answer the user question. "
                "State their exact name, franchise/origin, and background details."
            )

        part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)
        return self._call_gemini_with_fallback([prompt, part], temperature=0.2)

    def analyze_pdf(self, pdf_bytes: bytes, filename: str, user_prompt: Optional[str] = None) -> str:
        """Extracts text from PDF or analyzes it directly."""
        prompt = user_prompt.strip() if user_prompt else f"Summarize this PDF document '{filename}', highlighting its main points and takeaways."

        text_content = ""
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            extracted_pages = []
            for page_idx in range(min(len(doc), 50)):
                page_text = doc[page_idx].get_text()
                if page_text.strip():
                    extracted_pages.append(f"--- Page {page_idx + 1} ---\n{page_text}")
            text_content = "\n\n".join(extracted_pages)
            doc.close()
        except Exception as e:
            logger.warning(f"PyMuPDF text extraction failed: {e}")

        if text_content and len(text_content.strip()) > 50:
            full_content = f"Document: {filename}\n\nContent:\n{text_content[:80000]}\n\nTask: {prompt}"
            return self._call_gemini_with_fallback([full_content], temperature=0.3)

        part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
        return self._call_gemini_with_fallback([prompt, part], temperature=0.3)

    def analyze_zip(self, zip_bytes: bytes, filename: str, user_prompt: Optional[str] = None) -> str:
        """Inspects, extracts, and summarizes the contents of a ZIP file."""
        prompt = user_prompt.strip() if user_prompt else f"Analyze this ZIP archive '{filename}', explain its structure, purpose, and key components."

        file_tree = []
        code_snippets = []
        total_snippet_chars = 0
        max_snippet_chars = 70000

        text_extensions = {
            ".py", ".js", ".jsx", ".ts", ".tsx", ".html", ".css", ".json",
            ".md", ".txt", ".yaml", ".yml", ".env", ".sql", ".sh", ".bat",
            ".ps1", ".c", ".cpp", ".h", ".rs", ".go", ".java", ".xml", ".csv"
        }

        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                info_list = z.infolist()
                for info in info_list:
                    file_tree.append(f"{info.filename} ({info.file_size} bytes)")
                    _, ext = os.path.splitext(info.filename)
                    if ext.lower() in text_extensions and info.file_size < 100_000:
                        if total_snippet_chars < max_snippet_chars:
                            try:
                                with z.open(info) as f:
                                    raw = f.read().decode("utf-8", errors="replace")
                                    truncated = raw[:8000]
                                    code_snippets.append(f"=== File: {info.filename} ===\n{truncated}")
                                    total_snippet_chars += len(truncated)
                            except Exception:
                                pass

            tree_str = "\n".join(file_tree[:100])
            if len(file_tree) > 100:
                tree_str += f"\n... and {len(file_tree) - 100} more files."

            snippets_str = "\n\n".join(code_snippets)

            full_prompt = (
                f"You are analyzing a ZIP archive named '{filename}'.\n\n"
                f"📁 **Archive File Manifest:**\n{tree_str}\n\n"
                f"📄 **Extracted Source & Text Files:**\n{snippets_str}\n\n"
                f"🎯 **User Request:**\n{prompt}"
            )
            return self._call_gemini_with_fallback([full_prompt], temperature=0.3)

        except zipfile.BadZipFile:
            return "⚠️ The uploaded file is not a valid ZIP archive or is corrupted."

    def analyze_audio(self, audio_bytes: bytes, mime_type: str, user_prompt: Optional[str] = None) -> str:
        """Transcribes and analyzes voice notes or audio files."""
        prompt = user_prompt.strip() if user_prompt else "Listen to this audio carefully, transcribe its contents, and respond to what was spoken."
        part = types.Part.from_bytes(data=audio_bytes, mime_type=mime_type)
        return self._call_gemini_with_fallback([prompt, part], temperature=0.3)

    def analyze_text_doc(self, text: str, filename: str, user_prompt: Optional[str] = None) -> str:
        """Analyzes plain text or code document files."""
        prompt = user_prompt.strip() if user_prompt else f"Review and summarize this document '{filename}'."
        content = f"File: {filename}\n\nContent:\n{text[:80000]}\n\nTask: {prompt}"
        return self._call_gemini_with_fallback([content], temperature=0.3)
