import re
import html
import logging

logger = logging.getLogger("MedusaFormatter")


def markdown_to_telegram_html(text: str) -> str:
    """
    Converts standard Markdown text into Telegram-supported HTML.
    Prevents raw *** or ** from displaying, converts headers, bold,
    italic, inline code, code blocks, and links cleanly.
    """
    if not text:
        return ""

    # 1. Protect code blocks ```lang\ncode```
    code_blocks = []
    def save_block(match):
        code_blocks.append(match.group(0))
        return f"___BLOCK_CODE_{len(code_blocks)-1}___"

    text = re.sub(r"```([a-zA-Z0-9_\-]*)\n?(.*?)```", save_block, text, flags=re.DOTALL)

    # 2. Protect inline code `code`
    inline_codes = []
    def save_inline(match):
        inline_codes.append(match.group(1))
        return f"___INLINE_CODE_{len(inline_codes)-1}___"

    text = re.sub(r"`([^`\n]+)`", save_inline, text)

    # 3. Escape HTML characters in remaining text
    text = html.escape(text, quote=False)

    # 4. Convert Markdown headers (# Header -> <b>Header</b>)
    text = re.sub(r"^(#{1,6})\s+(.+)$", r"<b>\2</b>", text, flags=re.MULTILINE)

    # 5. Bold + Italic: ***text*** or ___text___ -> <b><i>text</i></b>
    text = re.sub(r"\*\*\*([^\*\n]+?)\*\*\*", r"<b><i>\1</i></b>", text)
    text = re.sub(r"___([^_\n]+?)___", r"<b><i>\1</i></b>", text)

    # 6. Bold: **text** or __text__ -> <b>text</b>
    text = re.sub(r"\*\*([^\*\n]+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"__([^_\n]+?)__", r"<b>\1</b>", text)

    # 7. Italic: *text* or _text_ -> <i>text</i>
    text = re.sub(r"(?<!\w)\*([^\*\n]+?)\*(?!\w)", r"<i>\1</i>", text)
    text = re.sub(r"(?<!\w)_([^_\n]+?)_(?!\w)", r"<i>\1</i>", text)

    # 8. Strikethrough: ~~text~~ -> <s>text</s>
    text = re.sub(r"~~([^~\n]+?)~~", r"<s>\1</s>", text)

    # 9. Links: [text](url) -> <a href="url">text</a>
    text = re.sub(r"\[([^\]]+)\]\((https?://[^\s\)]+)\)", r'<a href="\2">\1</a>', text)

    # 10. Clean up any leftover stray asterisks or markdown artifacts
    text = re.sub(r"\*{2,}", "", text)

    # 11. Restore inline code
    for i, code in enumerate(inline_codes):
        escaped_code = html.escape(code, quote=False)
        text = text.replace(f"___INLINE_CODE_{i}___", f"<code>{escaped_code}</code>")

    # 12. Restore code blocks
    for i, block in enumerate(code_blocks):
        m = re.match(r"```([a-zA-Z0-9_\-]*)\n?(.*?)```", block, flags=re.DOTALL)
        if m:
            lang = m.group(1).strip()
            content = html.escape(m.group(2), quote=False)
            cls_attr = f' class="language-{lang}"' if lang else ""
            text = text.replace(
                f"___BLOCK_CODE_{i}___",
                f"<pre><code{cls_attr}>{content}</code></pre>"
            )

    return text

