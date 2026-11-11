import os
import re
import asyncio
import logging
from collections import defaultdict
from typing import Dict, List, Optional, Any

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    constants,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

import config
from server_manager import LlamaServerManager
from llm_client import LLMClient
from multimodal_handler import MultimodalAnalyzer
from formatter import markdown_to_telegram_html

# Logging configuration
logging.basicConfig(
    format="%(asctime)s - [%(name)s] - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("MedusaBot")

# Modes
MODE_DEFAULT = "default"  # High-speed cloud intelligence + live web search
MODE_MEDUSA = "medusa"    # Local secondary neural engine

# Patterns that indicate a user is asking about an image/character
IMAGE_QUERY_PATTERNS = [
    r"\bwho\s+is\s+(she|he|this|that|it|they)\b",
    r"\bwhat\s+(is\s+)?(her|his|their)\s+name\b",
    r"\bwhat\s+(character|anime|game|movie|series|franchise|celebrity)\b",
    r"\bidentify\s+(this|her|him|them)\b",
    r"\bwhich\s+(character|person|anime)\b",
    r"\bwhere\s+is\s+(she|he|this)\s+from\b",
    r"\b(tell\s+me\s+about\s+)?(this|the)\s+(picture|photo|image|art|character)\b",
]
COMPILED_IMAGE_PATTERNS = [re.compile(p, re.IGNORECASE) for p in IMAGE_QUERY_PATTERNS]


def is_image_followup(text: str) -> bool:
    """Detects if a text message is asking about a recent image or character."""
    cleaned = text.strip()
    for pattern in COMPILED_IMAGE_PATTERNS:
        if pattern.search(cleaned):
            return True
    return False


class MedusaBot:
    def __init__(self):
        self.server_manager = LlamaServerManager()
        self.llm_client = LLMClient()
        self.multimodal = MultimodalAnalyzer()

        # State per chat
        self.modes: Dict[int, str] = defaultdict(lambda: MODE_DEFAULT)
        self.history: Dict[int, List[Dict[str, str]]] = defaultdict(list)
        self.system_prompts: Dict[int, str] = {}
        self.locks: Dict[int, asyncio.Lock] = defaultdict(asyncio.Lock)

        # Cache last uploaded photo per chat for seamless follow-up Q&A
        self.last_photo: Dict[int, bytes] = {}

    def get_mode(self, chat_id: int) -> str:
        return self.modes[chat_id]

    def set_mode(self, chat_id: int, mode: str):
        self.modes[chat_id] = mode

    def get_system_prompt(self, chat_id: int) -> str:
        return self.system_prompts.get(chat_id, config.DEFAULT_SYSTEM_PROMPT)

    def reset_chat(self, chat_id: int):
        if chat_id in self.history:
            self.history[chat_id].clear()
        if chat_id in self.last_photo:
            del self.last_photo[chat_id]

    async def _send_typing_loop(self, bot, chat_id: int, stop_event: asyncio.Event, action=constants.ChatAction.TYPING):
        """Sends periodic chat action status while generating."""
        while not stop_event.is_set():
            try:
                await bot.send_chat_action(chat_id=chat_id, action=action)
            except Exception:
                pass
            try:
                await asyncio.wait_for(stop_event.wait(), timeout=4.0)
            except asyncio.TimeoutError:
                pass

    @staticmethod
    def split_text(text: str, max_chars: int = 3800) -> List[str]:
        """Splits long text into manageable chunks."""
        if len(text) <= max_chars:
            return [text]

        chunks = []
        while text:
            if len(text) <= max_chars:
                chunks.append(text)
                break

            split_idx = text.rfind("\n\n", 0, max_chars)
            if split_idx == -1:
                split_idx = text.rfind("\n", 0, max_chars)
            if split_idx == -1:
                split_idx = text.rfind(" ", 0, max_chars)
            if split_idx == -1:
                split_idx = max_chars

            chunks.append(text[:split_idx].strip())
            text = text[split_idx:].strip()

        return chunks

    async def safe_reply(self, update: Update, raw_markdown: str):
        """
        Sends formatted response to Telegram.
        Converts Markdown to Telegram-safe HTML, properly rendering bold, italic,
        and code without raw asterisks (***).
        """
        chunks = self.split_text(raw_markdown)
        for chunk in chunks:
            formatted_html = markdown_to_telegram_html(chunk)
            try:
                await update.effective_message.reply_text(
                    formatted_html,
                    parse_mode=constants.ParseMode.HTML,
                    disable_web_page_preview=True,
                )
            except Exception as e:
                logger.warning(f"HTML parse failed ({e}), falling back to plain text.")
                await update.effective_message.reply_text(
                    chunk,
                    disable_web_page_preview=True,
                )

    # --- Command Handlers ---

    async def cmd_start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Displays welcome menu and mode selector."""
        user = update.effective_user
        first_name = user.first_name if user else "Friend"
        chat_id = update.effective_chat.id
        current_mode = self.get_mode(chat_id)
        mode_label = "⚡ Default Mode" if current_mode == MODE_DEFAULT else "🐍 Medusa Mode"

        welcome_text = (
            f"👋 <b>Hello, {first_name}!</b>\n\n"
            f"I am <b>Medusa AI</b>, your advanced, versatile personal AI assistant.\n\n"
            f"🌟 <b>Capabilities:</b>\n"
            f"• <b>Image & Character Recognition:</b> Send any photo to identify people, anime/game characters, landmarks, or objects.\n"
            f"• <b>Real-Time Web Knowledge:</b> Ask about current events, live news, and recent information.\n"
            f"• <b>Document Analysis:</b> Send PDFs, ZIP archives, or voice notes for detailed breakdowns.\n"
            f"• <b>Dual Engine Modes:</b> High-speed Default Mode and secondary Medusa Mode.\n\n"
            f"📌 <b>Current Engine:</b> <code>{mode_label}</code>\n\n"
            f"💬 <i>Send me any question or photo to begin!</i>"
        )

        keyboard = [
            [
                InlineKeyboardButton("⚡ Default Mode", callback_data="set_default"),
                InlineKeyboardButton("🐍 Medusa Mode", callback_data="set_medusa"),
            ],
            [
                InlineKeyboardButton("ℹ️ Help & Commands", callback_data="action_help"),
                InlineKeyboardButton("🔄 Reset Chat", callback_data="action_reset"),
            ],
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)

        await update.effective_message.reply_text(
            welcome_text,
            reply_markup=reply_markup,
            parse_mode=constants.ParseMode.HTML,
        )

    async def cmd_help(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Displays usage guide and commands."""
        text = (
            "📖 <b>Medusa AI — User Guide</b>\n\n"
            "<b>Commands:</b>\n"
            "• <code>/default</code> — Switch to Default Mode (real-time web search & fast chats)\n"
            "• <code>/medusa</code> — Switch to Medusa Secondary Mode\n"
            "• <code>/reset</code> or <code>/new</code> — Clear conversation memory\n"
            "• <code>/help</code> — Show this guide\n"
            "• <code>/system &lt;prompt&gt;</code> — Set a custom persona or role\n"
            "• <code>/system reset</code> — Restore default persona\n\n"
            "🖼️ <b>Image & Character Recognition:</b>\n"
            "• Send any photo directly or reply to a photo with <i>'who is she?'</i> or <i>'who is this?'</i>\n"
            "• Medusa AI will accurately identify characters, anime, people, landmarks, and objects!\n\n"
            "📁 <b>Supported Files:</b>\n"
            "• <b>PDF Documents:</b> In-depth summaries, document Q&A\n"
            "• <b>ZIP Archives:</b> Complete directory manifest and code analysis\n"
            "• <b>Audio & Voice:</b> Automatic transcription and responses"
        )
        await self.safe_reply(update, text)

    async def cmd_default(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Switches the active chat to Default Mode."""
        chat_id = update.effective_chat.id
        self.set_mode(chat_id, MODE_DEFAULT)
        self.reset_chat(chat_id)
        await update.effective_message.reply_text(
            "⚡ <b>Switched to Default Mode!</b>\n\n"
            "Equipped with real-time web search and fast conversation memory. How can I help you?",
            parse_mode=constants.ParseMode.HTML,
        )

    async def cmd_medusa(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Switches the active chat to secondary Medusa Mode."""
        chat_id = update.effective_chat.id

        if not self.server_manager.is_running():
            logger.info("Medusa mode requested but server is offline. Attempting quick start...")
            started = self.server_manager.start(timeout_sec=15)
            if not started:
                await update.effective_message.reply_text(
                    "⚠️ <b>Models are offline now.</b>\n\n"
                    "The Medusa engine is currently offline. You can continue chatting in default mode by using /default.",
                    parse_mode=constants.ParseMode.HTML,
                )
                return

        self.set_mode(chat_id, MODE_MEDUSA)
        self.reset_chat(chat_id)
        await update.effective_message.reply_text(
            "🐍 <b>Switched to Medusa Mode!</b>\n\n"
            "The secondary Medusa engine is active and ready. What would you like to discuss?",
            parse_mode=constants.ParseMode.HTML,
        )

    async def cmd_reset(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Clears conversational memory."""
        chat_id = update.effective_chat.id
        self.reset_chat(chat_id)
        await update.effective_message.reply_text(
            "🔄 <b>Conversation memory cleared!</b> Starting a fresh session.",
            parse_mode=constants.ParseMode.HTML,
        )

    async def cmd_system(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Sets or views custom persona."""
        chat_id = update.effective_chat.id
        if not context.args:
            current = self.get_system_prompt(chat_id)
            await update.effective_message.reply_text(
                f"🧠 <b>Current Persona:</b>\n<i>{current}</i>\n\n"
                f"To change: <code>/system &lt;your custom instruction&gt;</code>\n"
                f"To reset: <code>/system reset</code>",
                parse_mode=constants.ParseMode.HTML,
            )
            return

        arg_text = " ".join(context.args).strip()
        if arg_text.lower() == "reset":
            if chat_id in self.system_prompts:
                del self.system_prompts[chat_id]
            self.reset_chat(chat_id)
            await update.effective_message.reply_text(
                "✅ Restored default Medusa AI persona.",
                parse_mode=constants.ParseMode.HTML,
            )
        else:
            self.system_prompts[chat_id] = arg_text
            self.reset_chat(chat_id)
            await update.effective_message.reply_text(
                f"✅ <b>Persona updated:</b>\n<i>{arg_text}</i>",
                parse_mode=constants.ParseMode.HTML,
            )

    async def handle_callback(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handles inline keyboard button interactions."""
        query = update.callback_query
        await query.answer()
        data = query.data

        if data == "set_default":
            await self.cmd_default(update, context)
        elif data == "set_medusa":
            await self.cmd_medusa(update, context)
        elif data == "action_help":
            await self.cmd_help(update, context)
        elif data == "action_reset":
            await self.cmd_reset(update, context)

    # --- Message Processing ---

    async def handle_text(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handles incoming text messages with vision follow-up intelligence."""
        if not update.effective_message or not update.effective_message.text:
            return

        chat = update.effective_chat
        message = update.effective_message
        text = message.text.strip()
        is_private = (chat.type == constants.ChatType.PRIVATE)

        # In groups, enforce mention/reply requirement
        if not is_private:
            bot_user = await context.bot.get_me()
            username = f"@{bot_user.username}" if bot_user.username else ""
            is_reply = (
                message.reply_to_message
                and message.reply_to_message.from_user
                and message.reply_to_message.from_user.id == bot_user.id
            )
            is_mentioned = username and (username.lower() in text.lower())

            if not (is_reply or is_mentioned):
                return

            if username:
                text = text.replace(username, "").strip()

        if not text:
            return

        chat_id = chat.id

        # ============================================================
        # 1. SPECIAL CASE: USER REPLIED DIRECTLY TO A PHOTO MESSAGE
        # ============================================================
        replied_photo = None
        if message.reply_to_message and message.reply_to_message.photo:
            try:
                photo_obj = message.reply_to_message.photo[-1]
                photo_file = await photo_obj.get_file()
                replied_photo = bytes(await photo_file.download_as_bytearray())
                self.last_photo[chat_id] = replied_photo
            except Exception as e:
                logger.warning(f"Could not retrieve replied photo: {e}")

        # ============================================================
        # 2. SPECIAL CASE: FOLLOW-UP QUESTION ON A RECENT PHOTO
        # (e.g. user replies to bot or types "who is she?" / "her name")
        # ============================================================
        should_use_photo = replied_photo is not None
        target_photo = replied_photo

        if not should_use_photo and (chat_id in self.last_photo):
            is_replying_to_bot = (
                message.reply_to_message
                and message.reply_to_message.from_user
                and message.reply_to_message.from_user.is_bot
            )
            if is_replying_to_bot or is_image_followup(text):
                should_use_photo = True
                target_photo = self.last_photo[chat_id]

        if should_use_photo and target_photo:
            async with self.locks[chat_id]:
                stop_typing = asyncio.Event()
                typing_task = asyncio.create_task(
                    self._send_typing_loop(context.bot, chat_id, stop_typing, constants.ChatAction.UPLOAD_PHOTO)
                )
                try:
                    loop = asyncio.get_running_loop()
                    analysis = await loop.run_in_executor(
                        None,
                        self.multimodal.analyze_image,
                        target_photo,
                        "image/jpeg",
                        text
                    )
                    # Update chat memory
                    history = self.history[chat_id]
                    history.append({"role": "user", "content": text})
                    history.append({"role": "assistant", "content": analysis})
                    max_records = config.MAX_HISTORY_TURNS * 2
                    if len(history) > max_records:
                        self.history[chat_id] = history[-max_records:]

                    await self.safe_reply(update, analysis)
                    return
                except Exception as e:
                    logger.error(f"Error analyzing image follow-up: {e}", exc_info=True)
                    # If vision fails, proceed to normal text handling below
                finally:
                    stop_typing.set()
                    await typing_task

        # ============================================================
        # 3. STANDARD CONVERSATIONAL TEXT HANDLING
        # ============================================================
        async with self.locks[chat_id]:
            stop_typing = asyncio.Event()
            typing_task = asyncio.create_task(
                self._send_typing_loop(context.bot, chat_id, stop_typing)
            )

            try:
                mode = self.get_mode(chat_id)
                sys_prompt = self.get_system_prompt(chat_id)
                history = self.history[chat_id]

                messages = [{"role": "system", "content": sys_prompt}]
                messages.extend(history)
                messages.append({"role": "user", "content": text})

                reply = ""
                if mode == MODE_MEDUSA:
                    if not self.server_manager.is_running():
                        await update.effective_message.reply_text(
                            "⚠️ <b>Models are offline now.</b>\n\n"
                            "The Medusa engine is currently offline. Please switch to default mode using /default.",
                            parse_mode=constants.ParseMode.HTML,
                        )
                        return
                    try:
                        reply = await self.llm_client.generate_medusa_response(messages)
                    except ConnectionError:
                        await update.effective_message.reply_text(
                            "⚠️ <b>Models are offline now.</b>\n\n"
                            "The Medusa engine is currently offline. Please switch to default mode using /default.",
                            parse_mode=constants.ParseMode.HTML,
                        )
                        return
                else:
                    # Default Mode: Groq + Live Web Search
                    reply = await self.llm_client.generate_default_response(messages, user_query=text)

                if reply:
                    history.append({"role": "user", "content": text})
                    history.append({"role": "assistant", "content": reply})
                    max_records = config.MAX_HISTORY_TURNS * 2
                    if len(history) > max_records:
                        self.history[chat_id] = history[-max_records:]

                    await self.safe_reply(update, reply)

            except Exception as e:
                logger.error(f"Error handling text for chat {chat_id}: {e}", exc_info=True)
                await update.effective_message.reply_text(
                    "⚠️ An unexpected error occurred while generating a response. Please try again in a moment."
                )
            finally:
                stop_typing.set()
                await typing_task

    # --- Multimodal File / Media Handlers ---

    async def handle_photo(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handles incoming photo uploads."""
        message = update.effective_message
        if not message.photo:
            return

        chat_id = update.effective_chat.id
        caption = message.caption or ""

        async with self.locks[chat_id]:
            stop_typing = asyncio.Event()
            typing_task = asyncio.create_task(
                self._send_typing_loop(context.bot, chat_id, stop_typing, constants.ChatAction.UPLOAD_PHOTO)
            )
            try:
                # Download highest resolution photo
                photo_file = await message.photo[-1].get_file()
                photo_bytes = bytes(await photo_file.download_as_bytearray())

                # Cache photo for seamless follow-up questions ("who is she?", etc.)
                self.last_photo[chat_id] = photo_bytes

                loop = asyncio.get_running_loop()
                analysis = await loop.run_in_executor(
                    None,
                    self.multimodal.analyze_image,
                    photo_bytes,
                    "image/jpeg",
                    caption
                )

                # Persist in conversation history so future questions know about this image!
                history = self.history[chat_id]
                user_record = f"[User uploaded a photo{f': {caption}' if caption else ''}]"
                history.append({"role": "user", "content": user_record})
                history.append({"role": "assistant", "content": analysis})
                max_records = config.MAX_HISTORY_TURNS * 2
                if len(history) > max_records:
                    self.history[chat_id] = history[-max_records:]

                await self.safe_reply(update, analysis)
            except Exception as e:
                logger.error(f"Error analyzing photo: {e}", exc_info=True)
                await update.effective_message.reply_text(
                    "⚠️ Could not analyze the image. Please verify file format and try again."
                )
            finally:
                stop_typing.set()
                await typing_task

    async def handle_document(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handles document uploads (PDF, ZIP, Code, Text)."""
        message = update.effective_message
        doc = message.document
        if not doc:
            return

        chat_id = update.effective_chat.id
        filename = doc.file_name or "document"
        caption = message.caption or f"Please analyze and summarize this file: {filename}"
        _, ext = os.path.splitext(filename.lower())

        async with self.locks[chat_id]:
            stop_typing = asyncio.Event()
            typing_task = asyncio.create_task(
                self._send_typing_loop(context.bot, chat_id, stop_typing, constants.ChatAction.UPLOAD_DOCUMENT)
            )

            try:
                tg_file = await doc.get_file()
                file_bytes = bytes(await tg_file.download_as_bytearray())
                loop = asyncio.get_running_loop()

                analysis = ""
                if ext == ".pdf":
                    analysis = await loop.run_in_executor(
                        None,
                        self.multimodal.analyze_pdf,
                        file_bytes,
                        filename,
                        caption
                    )
                elif ext == ".zip":
                    analysis = await loop.run_in_executor(
                        None,
                        self.multimodal.analyze_zip,
                        file_bytes,
                        filename,
                        caption
                    )
                else:
                    try:
                        text_content = file_bytes.decode("utf-8")
                        analysis = await loop.run_in_executor(
                            None,
                            self.multimodal.analyze_text_doc,
                            text_content,
                            filename,
                            caption
                        )
                    except UnicodeDecodeError:
                        analysis = await loop.run_in_executor(
                            None,
                            self.multimodal.analyze_pdf,
                            file_bytes,
                            filename,
                            caption
                        )

                # Persist in conversation memory
                history = self.history[chat_id]
                history.append({"role": "user", "content": f"[User uploaded file '{filename}': {caption}]"})
                history.append({"role": "assistant", "content": analysis})
                max_records = config.MAX_HISTORY_TURNS * 2
                if len(history) > max_records:
                    self.history[chat_id] = history[-max_records:]

                await self.safe_reply(update, analysis)

            except Exception as e:
                logger.error(f"Error analyzing document '{filename}': {e}", exc_info=True)
                await update.effective_message.reply_text(
                    f"⚠️ Unable to analyze '{filename}'. Please ensure the file is not corrupted."
                )
            finally:
                stop_typing.set()
                await typing_task

    async def handle_audio(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handles incoming voice notes and audio clips."""
        message = update.effective_message
        audio_obj = message.voice or message.audio
        if not audio_obj:
            return

        chat_id = update.effective_chat.id
        caption = message.caption or "Transcribe and respond to this audio."
        mime = audio_obj.mime_type or "audio/ogg"

        async with self.locks[chat_id]:
            stop_typing = asyncio.Event()
            typing_task = asyncio.create_task(
                self._send_typing_loop(context.bot, chat_id, stop_typing, constants.ChatAction.RECORD_VOICE)
            )
            try:
                tg_file = await audio_obj.get_file()
                audio_bytes = bytes(await tg_file.download_as_bytearray())
                loop = asyncio.get_running_loop()
                analysis = await loop.run_in_executor(
                    None,
                    self.multimodal.analyze_audio,
                    audio_bytes,
                    mime,
                    caption
                )

                history = self.history[chat_id]
                history.append({"role": "user", "content": f"[User sent a voice message: {caption}]"})
                history.append({"role": "assistant", "content": analysis})
                max_records = config.MAX_HISTORY_TURNS * 2
                if len(history) > max_records:
                    self.history[chat_id] = history[-max_records:]

                await self.safe_reply(update, analysis)
            except Exception as e:
                logger.error(f"Error analyzing audio: {e}", exc_info=True)
                await update.effective_message.reply_text(
                    "⚠️ Could not transcribe or analyze the audio. Please try again."
                )
            finally:
                stop_typing.set()
                await typing_task

    def run(self):
        """Initializes and runs the Telegram bot."""
        if not config.TELEGRAM_BOT_TOKEN:
            print("❌ ERROR: TELEGRAM_BOT_TOKEN is not set.")
            return

        print("\n🚀 Checking local Medusa model server...")
        try:
            self.server_manager.start(timeout_sec=20)
        except Exception as e:
            logger.warning(f"Medusa server could not be started at launch: {e}")

        print("🤖 Initializing Telegram Bot application...")
        app = Application.builder().token(config.TELEGRAM_BOT_TOKEN).build()

        # Commands
        app.add_handler(CommandHandler("start", self.cmd_start))
        app.add_handler(CommandHandler("help", self.cmd_help))
        app.add_handler(CommandHandler(["default", "normal"], self.cmd_default))
        app.add_handler(CommandHandler("medusa", self.cmd_medusa))
        app.add_handler(CommandHandler(["reset", "clear", "new"], self.cmd_reset))
        app.add_handler(CommandHandler("system", self.cmd_system))

        # Inline button callbacks
        app.add_handler(CallbackQueryHandler(self.handle_callback))

        # Media handlers
        app.add_handler(MessageHandler(filters.PHOTO, self.handle_photo))
        app.add_handler(MessageHandler(filters.Document.ALL, self.handle_document))
        app.add_handler(MessageHandler(filters.VOICE | filters.AUDIO, self.handle_audio))

        # Text handler
        app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, self.handle_text))

        print("\n" + "=" * 60)
        print("✅ Medusa AI Telegram Bot is online and active!")
        print("Vision & Character Recognition: Fully enabled.")
        print("Photo Replies & Follow-up Q&A: Fully enabled.")
        print("Default Mode (Groq + Real-Time Web Search): Active.")
        print("Secondary Mode (/medusa): Ready.")
        print("=" * 60 + "\n")

        try:
            app.run_polling(drop_pending_updates=True)
        finally:
            print("\nShutting down bot and server...")
            self.server_manager.stop()
            print("Done. Goodbye!")


if __name__ == "__main__":
    bot = MedusaBot()
    bot.run()
