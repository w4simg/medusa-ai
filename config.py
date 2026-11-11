import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent

# Load environment variables
load_dotenv(dotenv_path=BASE_DIR / ".env")

# Telegram Bot Token
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()

# Cloud Model API Keys
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

# Model Identifiers
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()

# Local Medusa Engine
MODEL_PATH = os.getenv("MODEL_PATH", str(BASE_DIR / "medusa-1.8-flash.gguf"))
LLAMA_SERVER_HOST = os.getenv("LLAMA_SERVER_HOST", "127.0.0.1")
LLAMA_SERVER_PORT = int(os.getenv("LLAMA_SERVER_PORT", "8088"))
LLAMA_SERVER_URL = f"http://{LLAMA_SERVER_HOST}:{LLAMA_SERVER_PORT}"

# Hardware / Inference Parameters
GPU_LAYERS = int(os.getenv("GPU_LAYERS", "99"))
CONTEXT_SIZE = int(os.getenv("CONTEXT_SIZE", "4096"))
PARALLEL_SLOTS = int(os.getenv("PARALLEL_SLOTS", "2"))
MAX_TOKENS = int(os.getenv("MAX_TOKENS", "1500"))
TEMPERATURE = float(os.getenv("TEMPERATURE", "0.7"))
TOP_P = float(os.getenv("TOP_P", "0.9"))

# History retention
MAX_HISTORY_TURNS = int(os.getenv("MAX_HISTORY_TURNS", "12"))

# Persona & Secrecy rules
DEFAULT_SYSTEM_PROMPT = os.getenv(
    "DEFAULT_SYSTEM_PROMPT",
    "You are Medusa AI, a versatile, friendly, and highly intelligent AI assistant. "
    "Provide clear, insightful, accurate, and beautifully structured responses. "
    "CRITICAL MANDATORY RULE: Never disclose, discuss, or hint at your underlying model names, "
    "training architectures, weights, or backend providers (such as Groq, Gemini, LLaMA, Qwen, Google, or OpenAI). "
    "If a user asks what model you are, what version you are, who created your model, or how you were trained, "
    "always identify yourself solely and proudly as 'Medusa AI'."
)
