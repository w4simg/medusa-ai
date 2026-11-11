import asyncio
import logging
from typing import List, Dict, Any, Optional
import httpx
from groq import AsyncGroq

import config
from search_tool import should_search_web, search_web

logger = logging.getLogger("MedusaLLMClient")


class LLMClient:
    """Handles inference queries for both Default Mode (Groq + Web Search) and Medusa Mode (Local GGUF)."""

    def __init__(self):
        self.groq_api_key = config.GROQ_API_KEY
        self.groq_model = config.GROQ_MODEL
        self._groq_client: Optional[AsyncGroq] = None
        self._http_client: Optional[httpx.AsyncClient] = None

    @property
    def groq_client(self) -> AsyncGroq:
        if self._groq_client is None:
            if not self.groq_api_key:
                raise ValueError("GROQ_API_KEY is not configured.")
            self._groq_client = AsyncGroq(api_key=self.groq_api_key)
        return self._groq_client

    async def get_http_client(self) -> httpx.AsyncClient:
        if self._http_client is None or self._http_client.is_closed:
            self._http_client = httpx.AsyncClient(timeout=90.0)
        return self._http_client

    async def close(self):
        if self._http_client and not self._http_client.is_closed:
            await self._http_client.aclose()

    async def generate_default_response(
        self,
        messages: List[Dict[str, str]],
        user_query: str,
        temperature: float = config.TEMPERATURE,
        max_tokens: int = config.MAX_TOKENS,
    ) -> str:
        """
        Default Mode: Uses Groq for ultra-fast, long-context text responses.
        Augmented with real-time web search whenever real-time/current info is needed.
        """
        prepared_messages = list(messages)

        # 1. Check if real-time web search should be performed
        if should_search_web(user_query):
            logger.info("Real-time information requested. Querying live web...")
            loop = asyncio.get_running_loop()
            search_snippets = await loop.run_in_executor(None, search_web, user_query)
            if search_snippets:
                search_context = (
                    "=== REAL-TIME WEB SEARCH RESULTS ===\n"
                    f"{search_snippets}\n"
                    "=== INSTRUCTIONS ===\n"
                    "Use the above live web search results to answer the user's question accurately with up-to-date facts. "
                    "Do not mention that you received web snippets unless relevant to cite the source."
                )
                prepared_messages.insert(-1, {"role": "system", "content": search_context})

        try:
            res = await self.groq_client.chat.completions.create(
                model=self.groq_model,
                messages=prepared_messages,
                temperature=temperature,
                max_tokens=max_tokens,
            )
            choice = res.choices[0].message
            content = choice.content or ""
            return content.strip()
        except Exception as e:
            logger.error(f"Error querying Groq default model ({self.groq_model}): {e}")
            # Try fallback Groq model if the primary failed
            fallback_models = ["qwen/qwen3.8-27b", "openai/gpt-oss-120b", "openai/gpt-oss-20b"]
            for fb in fallback_models:
                if fb != self.groq_model:
                    try:
                        logger.info(f"Attempting fallback Groq model: {fb}")
                        fb_res = await self.groq_client.chat.completions.create(
                            model=fb,
                            messages=prepared_messages,
                            temperature=temperature,
                            max_tokens=max_tokens,
                        )
                        fb_content = fb_res.choices[0].message.content or ""
                        if fb_content.strip():
                            return fb_content.strip()
                    except Exception as fb_err:
                        logger.warning(f"Fallback model {fb} failed: {fb_err}")
            raise RuntimeError("Unable to generate response from default AI engine.")

    async def generate_medusa_response(
        self,
        messages: List[Dict[str, str]],
        temperature: float = config.TEMPERATURE,
        max_tokens: int = config.MAX_TOKENS,
    ) -> str:
        """
        Medusa Mode: Uses the local llama-server running medusa-1.8-flash.gguf.
        If the server is offline or unreachable, raises ConnectionError.
        """
        endpoint = f"{config.LLAMA_SERVER_URL}/v1/chat/completions"
        client = await self.get_http_client()

        payload = {
            "messages": messages,
            "temperature": temperature,
            "top_p": config.TOP_P,
            "max_tokens": max_tokens,
            "stream": False,
        }

        try:
            response = await client.post(endpoint, json=payload, timeout=90.0)
            response.raise_for_status()
            data = response.json()
            reply = data["choices"][0]["message"]["content"]
            return reply.strip()
        except (httpx.ConnectError, httpx.ConnectTimeout):
            logger.warning("Local Medusa server is offline or unreachable.")
            raise ConnectionError("MEDUSA_OFFLINE")
        except httpx.HTTPStatusError as e:
            logger.error(f"HTTP error from local Medusa server: {e.response.status_code}")
            raise ConnectionError("MEDUSA_OFFLINE")
        except Exception as e:
            logger.error(f"Unexpected error querying Medusa model: {e}")
            raise ConnectionError("MEDUSA_OFFLINE")
