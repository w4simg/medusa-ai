import logging
import re
from typing import Optional

try:
    from ddgs import DDGS
except ImportError:
    from duckduckgo_search import DDGS

logger = logging.getLogger("MedusaSearch")

# Keywords indicative of requiring current/real-time information
REALTIME_PATTERNS = [
    r"\b(today|tonight|yesterday|tomorrow|this week|this month|this year|now|currently)\b",
    r"\b(latest|recent|newest|breaking|upcoming|update|updates)\b",
    r"\b(news|weather|price|stock|crypto|bitcoin|btc|eth|market)\b",
    r"\b(score|match|game|election|winner|standing|fixtures)\b",
    r"\b(search for|search the web|search online|look up|google|browse)\b",
    r"\b(who is currently|who is the current|what is the current|what happened)\b",
    r"\b(release date|released date|patch notes|status)\b",
    r"\b(2025|2026|2027)\b",
]

COMPILED_PATTERNS = [re.compile(p, re.IGNORECASE) for p in REALTIME_PATTERNS]


def should_search_web(prompt: str) -> bool:
    """Detects whether a user prompt likely requires real-time web search."""
    cleaned = prompt.strip()
    # Explicit search prefixes
    if cleaned.lower().startswith(("/search", "!search", "search:", "search ")):
        return True

    for pattern in COMPILED_PATTERNS:
        if pattern.search(cleaned):
            return True

    return False


def clean_search_query(prompt: str) -> str:
    """Strips conversational noise from search queries."""
    query = re.sub(r"^(please\s+)?(search\s+for|search|look\s+up|find\s+information\s+about|tell\s+me\s+about)\s+", "", prompt, flags=re.IGNORECASE)
    return query.strip()


def search_web(query: str, max_results: int = 5) -> Optional[str]:
    """Queries DuckDuckGo and returns a structured summary of web snippets."""
    try:
        search_query = clean_search_query(query)
        logger.info(f"Conducting web search for: '{search_query}'")

        with DDGS() as ddgs:
            results = list(ddgs.text(search_query, max_results=max_results))

        if not results:
            return None

        formatted_parts = []
        for i, item in enumerate(results, 1):
            title = item.get("title", "").strip()
            body = item.get("body", "").strip()
            url = item.get("href", "").strip()
            if title and body:
                formatted_parts.append(
                    f"[{i}] {title}\nSummary: {body}\nSource: {url}"
                )

        if not formatted_parts:
            return None

        return "\n\n".join(formatted_parts)

    except Exception as e:
        logger.warning(f"Web search error for '{query}': {e}")
        return None

