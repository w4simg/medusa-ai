import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

OUTPUT_PPTX = r"c:\Users\unixb\OneDrive\Desktop\medusa_ai\Medusa_AI_Presentation.pptx"

# Color Palette
DARK_NAVY = RGBColor(15, 23, 42)      # #0F172A
DEEP_SLATE = RGBColor(30, 41, 59)     # #1E293B
ACCENT_BLUE = RGBColor(37, 99, 235)   # #2563EB
CYAN_ACCENT = RGBColor(14, 165, 233)  # #0EA5E9
EMERALD = RGBColor(16, 185, 129)      # #10B981
CARD_BG = RGBColor(248, 250, 252)     # #F8FAFC
CARD_BORDER = RGBColor(226, 232, 240) # #E2E8F0
TEXT_DARK = RGBColor(15, 23, 42)      # #0F172A
TEXT_MUTED = RGBColor(71, 85, 105)    # #475569
TEXT_WHITE = RGBColor(255, 255, 255)
CARD_WHITE = RGBColor(255, 255, 255)


def add_slide_header(slide, title_text, category="MEDUSA AI PROJECT"):
    """Adds a modern header to a slide."""
    # Category badge
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.4))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = ACCENT_BLUE

    # Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Arial"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = DARK_NAVY


def add_card(slide, left, top, width, height, bg_color=CARD_WHITE, border_color=CARD_BORDER):
    """Draws a rounded rectangular card background."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)
    return shape


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (Dark Navy Hero Theme)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_NAVY
    bg1.line.fill.background()

    # Title & Subtitle box
    tbox = s1.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(11.3), Inches(2.3))
    tf1 = tbox.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "MINOR PROJECT PRESENTATION | B.TECH CSE"
    p0.font.name = "Arial"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = CYAN_ACCENT
    p0.space_after = Pt(10)

    p1 = tf1.add_paragraph()
    p1.text = "MEDUSA AI"
    p1.font.name = "Arial"
    p1.font.size = Pt(40)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_after = Pt(8)

    p2 = tf1.add_paragraph()
    p2.text = "An Advanced Multimodal Conversational AI System with Dual-Engine Inference, Live Web Reasoning, and Distributed Telegram & CLI Deployment"
    p2.font.name = "Arial"
    p2.font.size = Pt(16)
    p2.font.color.rgb = RGBColor(203, 213, 225)

    # Info Cards Container (Bottom)
    c_left = Inches(1.0)
    c_top = Inches(3.9)
    c_w1 = Inches(6.8)
    c_w2 = Inches(4.3)
    c_h = Inches(2.8)

    card1 = add_card(s1, c_left, c_top, c_w1, c_h, bg_color=DEEP_SLATE, border_color=RGBColor(51, 65, 85))
    tf_c1 = card1.text_frame
    tf_c1.word_wrap = True
    p_c1_t = tf_c1.paragraphs[0]
    p_c1_t.text = "PROJECT TEAM MEMBERS"
    p_c1_t.font.name = "Arial"
    p_c1_t.font.size = Pt(11)
    p_c1_t.font.bold = True
    p_c1_t.font.color.rgb = CYAN_ACCENT
    p_c1_t.space_after = Pt(6)

    members = [
        "Wasim Akram (Reg. No: 2301348256)",
        "Satyajit Samantaray (Reg. No: 2301348203)",
        "Susanta Kumar Parida (Reg. No: 2421348051)",
        "Somanath Sahoo (Reg. No: 2421348050)",
        "Walter Pradhan (Reg. No: 2301348255)",
    ]
    for m in members:
        pm = tf_c1.add_paragraph()
        pm.text = f"•  {m}"
        pm.font.name = "Arial"
        pm.font.size = Pt(11)
        pm.font.color.rgb = TEXT_WHITE
        pm.space_after = Pt(3)

    card2 = add_card(s1, c_left + c_w1 + Inches(0.2), c_top, c_w2, c_h, bg_color=DEEP_SLATE, border_color=RGBColor(51, 65, 85))
    tf_c2 = card2.text_frame
    tf_c2.word_wrap = True
    p_c2_t = tf_c2.paragraphs[0]
    p_c2_t.text = "ACADEMIC INSTITUTION"
    p_c2_t.font.name = "Arial"
    p_c2_t.font.size = Pt(11)
    p_c2_t.font.bold = True
    p_c2_t.font.color.rgb = CYAN_ACCENT
    p_c2_t.space_after = Pt(8)

    inst_lines = [
        ("Guide:", "Prof. Achinta Kumar Palit"),
        ("Dept:", "Computer Science & Engineering"),
        ("College:", "GIET, Bhubaneswar"),
        ("Affiliation:", "BPUT, Odisha (2023 - 2027)"),
    ]
    for lbl, val in inst_lines:
        p_in = tf_c2.add_paragraph()
        p_in.text = f"{lbl} {val}"
        p_in.font.name = "Arial"
        p_in.font.size = Pt(11)
        p_in.font.color.rgb = TEXT_WHITE
        p_in.space_after = Pt(4)

    # =========================================================================
    # SLIDE 2: INTRODUCTION & OVERVIEW
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_header(s2, "1. Introduction & Project Overview", "EXECUTIVE SUMMARY")

    # 3 Cards Layout
    cards_data = [
        ("🤖 What is Medusa AI?",
         "An enterprise-grade, privacy-respecting conversational artificial intelligence system.",
         "• Bridges high-velocity cloud inference with private local edge computation.\n"
         "• Features real-time live web search augmentation.\n"
         "• Eliminates dependency on expensive, slow web-based GUI wrappers.",
         ACCENT_BLUE),
        ("🌐 Dual-Access Deployment",
         "Distributed across both mobile messaging and developer workstations.",
         "• Telegram Bot (@aimedusaroobot): Public access worldwide for 1-on-1 chats and groups.\n"
         "• Terminal CLI (medusaai): Pip-installable global package for developers with zero GUI overhead.",
         CYAN_ACCENT),
        ("⚡ Core Innovations",
         "Modern breakthroughs integrated into a single cohesive architecture.",
         "• Dual Engine: Cloud Groq LPU + Local NVIDIA RTX GPU llama-server.\n"
         "• Character Recognition: Names anime, gaming, and movie figures accurately.\n"
         "• Strict Privacy: Exclusively identifies as Medusa AI with zero model leaks.",
         EMERALD),
    ]

    for i, (title, sub, bullets, color) in enumerate(cards_data):
        c_x = Inches(0.8 + i * 3.95)
        card = add_card(s2, c_x, Inches(1.8), Inches(3.75), Inches(5.0))
        tf = card.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = color
        pt.space_after = Pt(8)

        psub = tf.add_paragraph()
        psub.text = sub
        psub.font.name = "Arial"
        psub.font.size = Pt(11.5)
        psub.font.bold = True
        psub.font.color.rgb = TEXT_DARK
        psub.space_after = Pt(14)

        for line in bullets.split("\n"):
            p_b = tf.add_paragraph()
            p_b.text = line
            p_b.font.name = "Arial"
            p_b.font.size = Pt(11)
            p_b.font.color.rgb = TEXT_MUTED
            p_b.space_after = Pt(6)

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & MOTIVATION
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_header(s3, "2. Problem Statement & Motivation", "THE CHALLENGE")

    probs = [
        ("🔒 Privacy & Vendor Lock-In",
         "Cloud-Only Centralization",
         "Most modern chatbots require transmitting private conversations and confidential files to proprietary third-party servers. Users lose ownership of data, and service terminates when offline.",
         RGBColor(239, 68, 68)),
        ("📅 Knowledge Horizon Staleness",
         "Outdated Information",
         "Pre-trained LLMs have static knowledge cutoffs. They cannot answer questions regarding breaking news, live sports, today's stock valuations, or recent events without expensive vector databases.",
         RGBColor(245, 158, 11)),
        ("🖼️ Multimodal Blindspots",
         "Superficial Vision",
         "When shown an image, standard bots provide superficial artistic commentary (colors, lighting) rather than recognizing specific characters (anime/gaming/celebrity) or handling photo follow-ups.",
         RGBColor(168, 85, 247)),
        ("💻 Heavy GUI Frameworks",
         "Terminal Neglect",
         "Electron-based desktop wrappers consume 1-2 GB of RAM just to render chat bubbles. Developers working inside terminal command-lines lack lightweight, installable AI tools.",
         RGBColor(59, 130, 246)),
    ]

    for i, (title, tag, desc, col) in enumerate(probs):
        col_idx = i % 2
        row_idx = i // 2
        x = Inches(0.8 + col_idx * 5.95)
        y = Inches(1.8 + row_idx * 2.55)

        card = add_card(s3, x, y, Inches(5.75), Inches(2.35))
        tf = card.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(4)

        ptag = tf.add_paragraph()
        ptag.text = f"KEY ISSUE: {tag.upper()}"
        ptag.font.name = "Arial"
        ptag.font.size = Pt(10)
        ptag.font.bold = True
        ptag.font.color.rgb = TEXT_MUTED
        ptag.space_after = Pt(8)

        pdesc = tf.add_paragraph()
        pdesc.text = desc
        pdesc.font.name = "Arial"
        pdesc.font.size = Pt(11)
        pdesc.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 4: SYSTEM OBJECTIVES
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_header(s4, "3. Key System Objectives", "DESIGN GOALS")

    objs = [
        ("Dual-Engine Hybrid Inference",
         "Seamlessly toggle between ultra-fast cloud inference (Groq LPUs) and 100% offline edge neural inference (NVIDIA RTX GPU with GGUF quantization)."),
        ("Live Web Search Augmentation",
         "Dynamically detect real-time informational intent and fetch up-to-the-minute web snippets via DuckDuckGo without requiring persistent vector databases."),
        ("Deep Multimodal Intelligence",
         "Accurately recognize people, anime/game characters, and landmarks from photos; parse multi-page PDFs, inspect ZIP software archives, and transcribe audio voice clips."),
        ("Reliable Communication & Formatting",
         "Provide a robust Telegram Bot with safe Markdown-to-HTML conversion that eliminates raw asterisks (***) and respects group privacy through mention filtering."),
        ("Global CLI Software Packaging",
         "Distribute the terminal client as an independent Python package (`medusaai`) installable with pip, featuring ASCII Figlet banners and persistent API key management."),
        ("Complete Model Privacy & Secrecy",
         "Strictly protect model identity so that backend architectures, weights, and providers are never revealed to end users, consistently maintaining the Medusa AI persona."),
    ]

    for i, (title, desc) in enumerate(objs):
        col_idx = i % 2
        row_idx = i // 2
        x = Inches(0.8 + col_idx * 5.95)
        y = Inches(1.8 + row_idx * 1.7)

        card = add_card(s4, x, y, Inches(5.75), Inches(1.5))
        tf = card.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = f"0{i+1}.  {title}"
        pt.font.name = "Arial"
        pt.font.size = Pt(13.5)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_BLUE
        pt.space_after = Pt(4)

        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.name = "Arial"
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 5: SYSTEM ARCHITECTURE & DUAL ENGINE
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_header(s5, "4. System Architecture & Dual-Engine Workflow", "SYSTEM DESIGN")

    # 2 Comparison Columns
    col1 = add_card(s5, Inches(0.8), Inches(1.8), Inches(5.75), Inches(5.0))
    tf1 = col1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "⚡ DEFAULT ENGINE (Cloud Accelerated)"
    p1.font.name = "Arial"
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = ACCENT_BLUE
    p1.space_after = Pt(10)

    e1_points = [
        ("Core Technology:", "Groq Language Processing Unit (LPU) API"),
        ("Active Model:", "Qwen / LLaMA high-capacity reasoning model"),
        ("Token Velocity:", "Extremely fast (>300 tokens/second)"),
        ("Live Web Search:", "Integrated via DuckDuckGo Real-Time Search"),
        ("Activation Command:", "/default (Primary mode for all chats)"),
        ("Best Suited For:", "Complex reasoning, long multi-turn dialogues, current news, live facts, and instant responses."),
    ]
    for lbl, val in e1_points:
        p = tf1.add_paragraph()
        p.text = f"•  {lbl} {val}"
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(6)

    col2 = add_card(s5, Inches(6.78), Inches(1.8), Inches(5.75), Inches(5.0))
    tf2 = col2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "🐍 MEDUSA ENGINE (Local Neural Model)"
    p2.font.name = "Arial"
    p2.font.size = Pt(15)
    p2.font.bold = True
    p2.font.color.rgb = EMERALD
    p2.space_after = Pt(10)

    e2_points = [
        ("Core Technology:", "llama-server C++ Subprocess on Port 8088"),
        ("Active Model:", "medusa-1.8-flash.gguf (Q4_K Medium Quantization)"),
        ("Hardware Offload:", "100% layer offloading to NVIDIA RTX 2050 GPU"),
        ("Data Sovereignty:", "100% Private, zero cloud telemetry or logs"),
        ("Activation Command:", "/medusa (Secondary local mode)"),
        ("Offline Failover:", "If server is stopped, bot gracefully warns: '⚠️ Models are offline now. Please switch with /default.'"),
    ]
    for lbl, val in e2_points:
        p = tf2.add_paragraph()
        p.text = f"•  {lbl} {val}"
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(6)

    # =========================================================================
    # SLIDE 6: REAL-TIME WEB SEARCH AUGMENTATION
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_header(s6, "5. Real-Time Web Search Augmentation", "ZERO-VECTOR RAG")

    top_card = add_card(s6, Inches(0.8), Inches(1.8), Inches(11.73), Inches(1.4))
    tf_top = top_card.text_frame
    tf_top.word_wrap = True
    pt = tf_top.paragraphs[0]
    pt.text = "💡 Dynamic Zero-Vector Web Retrieval Pipeline"
    pt.font.name = "Arial"
    pt.font.size = Pt(14)
    pt.font.bold = True
    pt.font.color.rgb = ACCENT_BLUE
    pt.space_after = Pt(4)

    pd = tf_top.add_paragraph()
    pd.text = (
        "Unlike traditional RAG systems that require heavy embedding databases, Medusa AI employs an on-demand, "
        "heuristic-driven search pipeline. Incoming user prompts are parsed by regex intent classifiers; if real-time information "
        "is required, the system queries the live internet and synthesizes facts directly into the prompt context."
    )
    pd.font.name = "Arial"
    pd.font.size = Pt(11)
    pd.font.color.rgb = TEXT_MUTED

    steps = [
        ("Step 1: Heuristic Detection",
         "Compiled regex matches temporal triggers: 'today', 'latest', 'current', 'news', 'price', 'weather', 'who won'."),
        ("Step 2: Query Normalization",
         "Conversational fluff ('please search for', 'tell me about') is stripped to form clean, high-precision search keywords."),
        ("Step 3: DuckDuckGo Live Query",
         "The engine fetches top 4 factual snippets, titles, and source URLs asynchronously in an executor threadpool."),
        ("Step 4: Context Injection",
         "Extracted web snippets are injected into an ephemeral system prompt, allowing the LLM to generate cited, up-to-date facts."),
    ]

    for i, (title, desc) in enumerate(steps):
        x = Inches(0.8 + i * 2.98)
        c = add_card(s6, x, Inches(3.4), Inches(2.78), Inches(3.4))
        tf = c.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = "Arial"
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = DARK_NAVY
        p_t.space_after = Pt(8)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.name = "Arial"
        p_d.font.size = Pt(10.5)
        p_d.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 7: MULTIMODAL VISION & CHARACTER RECOGNITION
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_header(s7, "6. Multimodal Vision & Character Recognition", "VISUAL INTELLIGENCE")

    v_cards = [
        ("🎯 Character Recognition Breakthrough",
         "Naming the subject instead of generic art critique.",
         "• Engineered specialized prompts instructing Gemini to prioritize identifying characters from anime, manga, video games, comics, movies, and pop culture.\n"
         "• Example: An anime girl with braided orange hair, yellow concentric eyes, and a halo is immediately recognized as 'Makima from Chainsaw Man', complete with background lore.",
         ACCENT_BLUE),
        ("💬 Contextual Photo Replies",
         "Seamless multi-turn image follow-ups.",
         "• Detects when a user replies to any photo message with 'who is she?' or 'what's her name?'.\n"
         "• Caches the last uploaded photo (`self.last_photo`) per chat, enabling natural follow-up questions without re-uploading imagery.",
         CYAN_ACCENT),
        ("📄 Document & Archive Analysis",
         "Deep multi-page extraction & code review.",
         "• PDF Documents: Reads multi-page research papers and reports via PyMuPDF (fitz) up to 50 pages.\n"
         "• ZIP Archives: Unpacks file manifests, inspects code files (.py, .js, .json, .md), and generates structural architectural summaries.",
         EMERALD),
    ]

    for i, (title, sub, bullets, col) in enumerate(v_cards):
        x = Inches(0.8 + i * 3.95)
        c = add_card(s7, x, Inches(1.8), Inches(3.75), Inches(5.0))
        tf = c.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(6)

        psub = tf.add_paragraph()
        psub.text = sub
        psub.font.name = "Arial"
        psub.font.size = Pt(11)
        psub.font.bold = True
        psub.font.color.rgb = TEXT_DARK
        psub.space_after = Pt(12)

        for line in bullets.split("\n"):
            pb = tf.add_paragraph()
            pb.text = line
            pb.font.name = "Arial"
            pb.font.size = Pt(10.5)
            pb.font.color.rgb = TEXT_MUTED
            pb.space_after = Pt(6)

    # =========================================================================
    # SLIDE 8: TELEGRAM BOT IMPLEMENTATION & UX
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_header(s8, "7. Telegram Bot Implementation & UX", "FRONTEND DEPLOYMENT")

    bot_features = [
        ("⚡ python-telegram-bot v22 Async",
         "Built on a non-blocking asynchronous event loop with continuous typing actions (`ChatAction.TYPING`) to provide immediate visual feedback while inference executes."),
        ("🎛️ Interactive Inline Keyboards",
         "Interactive buttons embedded into `/start` allow users to switch engines, trigger quick help, or clear memory with a single tap, bypassing manual command typing."),
        ("👥 Intelligent Group Privacy",
         "In Telegram group chats, the bot stays silent during standard chatter. It activates exclusively when directly @mentioned or replied to, keeping groups clean."),
        ("🛡️ Deterministic HTML Formatter",
         "Converts markdown (`**bold**`, `***bold-italic***`, headers, code blocks) into Telegram HTML. Eliminates ugly unrendered asterisks (***) with safe plain-text fallback."),
        ("✂️ 4096-Character Message Chunking",
         "Telegram limits messages to 4,096 characters. Medusa AI automatically splits long essays and code listings by paragraphs, delivering complete replies seamlessly."),
        ("🧠 Multi-Turn Memory Sliding Window",
         "Maintains conversational context per chat up to 12 turns, allowing users to ask natural follow-up questions while preventing token budget exhaustion."),
    ]

    for i, (title, desc) in enumerate(bot_features):
        col_idx = i % 2
        row_idx = i // 2
        x = Inches(0.8 + col_idx * 5.95)
        y = Inches(1.8 + row_idx * 1.7)

        c = add_card(s8, x, y, Inches(5.75), Inches(1.5))
        tf = c.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_BLUE
        pt.space_after = Pt(4)

        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.name = "Arial"
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 9: STANDALONE CLI PACKAGE (medusaai)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_header(s9, "8. Standalone CLI Tool (medusaai)", "DEVELOPER DISTRIBUTION")

    c_left = add_card(s9, Inches(0.8), Inches(1.8), Inches(5.75), Inches(5.0))
    tf_l = c_left.text_frame
    tf_l.word_wrap = True
    pt = tf_l.paragraphs[0]
    pt.text = "📦 Independent Python Package"
    pt.font.name = "Arial"
    pt.font.size = Pt(15)
    pt.font.bold = True
    pt.font.color.rgb = ACCENT_BLUE
    pt.space_after = Pt(10)

    pkg_details = [
        ("Package Name:", "medusaai (Version 1.0.0)"),
        ("Installation:", "pip install medusaai (or pip install .)"),
        ("Console Commands:", "medusaai or simply medusa"),
        ("Directory:", "Completely isolated in C:\\...\\medusaai"),
        ("Build Artifacts:", "Built .whl and .tar.gz in dist/ ready for PyPI"),
        ("Credential Storage:", "Stored once in ~/.medusaai/config.json with zero repeat prompts."),
    ]
    for lbl, val in pkg_details:
        p = tf_l.add_paragraph()
        p.text = f"•  {lbl} {val}"
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    c_right = add_card(s9, Inches(6.78), Inches(1.8), Inches(5.75), Inches(5.0), bg_color=DARK_NAVY, border_color=DEEP_SLATE)
    tf_r = c_right.text_frame
    tf_r.word_wrap = True
    pt2 = tf_r.paragraphs[0]
    pt2.text = "💻 Terminal REPL Interface"
    pt2.font.name = "Arial"
    pt2.font.size = Pt(15)
    pt2.font.bold = True
    pt2.font.color.rgb = CYAN_ACCENT
    pt2.space_after = Pt(10)

    cli_features = [
        "• PyFiglet ASCII Banner: Renders bold 'Medusa' branding in crimson red.",
        "• Live Timestamp & Author: Displays centered Date, Time, and 'Author: Tsukiziro'.",
        "• Mode Indicator: Displays active engine status (MODE: DEFAULT / MEDUSA).",
        "• Real-Time Web Triggers: Automatically searches the web for current questions.",
        "• Colorama Styling: Highlighted terminal text without raw asterisks (***).",
        "• Commands: /default, /medusa, /key <api_key>, /clear, /exit.",
    ]
    for feat in cli_features:
        p = tf_r.add_paragraph()
        p.text = feat
        p.font.name = "Consolas"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 10: PERFORMANCE BENCHMARKS & TESTING
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_header(s10, "9. Performance Benchmarks & Testing", "EMPIRICAL RESULTS")

    metrics = [
        ("Cloud Generation Speed", "310 tokens/sec", "Groq LPU accelerated inference", ACCENT_BLUE),
        ("Local Model Velocity", "27.9 tokens/sec", "NVIDIA RTX 2050 (All layers GPU offloaded)", EMERALD),
        ("Cloud Response Latency", "340 ms", "Sub-second Time-to-First-Token", CYAN_ACCENT),
        ("VRAM Utilization", "2.28 GB", "Comfortably fits within 4GB consumer VRAM", RGBColor(168, 85, 247)),
    ]

    for i, (title, val, note, col) in enumerate(metrics):
        x = Inches(0.8 + i * 2.98)
        c = add_card(s10, x, Inches(1.8), Inches(2.78), Inches(1.8))
        tf = c.text_frame
        tf.word_wrap = True
        pt = tf.paragraphs[0]
        pt.text = title.upper()
        pt.font.name = "Arial"
        pt.font.size = Pt(10)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_MUTED

        pv = tf.add_paragraph()
        pv.text = val
        pv.font.name = "Arial"
        pv.font.size = Pt(22)
        pv.font.bold = True
        pv.font.color.rgb = col
        pv.space_after = Pt(2)

        pn = tf.add_paragraph()
        pn.text = note
        pn.font.name = "Arial"
        pn.font.size = Pt(9.5)
        pn.font.color.rgb = TEXT_DARK

    # Bottom Table Container
    b_card = add_card(s10, Inches(0.8), Inches(3.9), Inches(11.73), Inches(2.9))
    tf_b = b_card.text_frame
    tf_b.word_wrap = True
    pt = tf_b.paragraphs[0]
    pt.text = "🧪 Verification & Test Case Highlights (16 / 16 PASSED)"
    pt.font.name = "Arial"
    pt.font.size = Pt(13)
    pt.font.bold = True
    pt.font.color.rgb = DARK_NAVY
    pt.space_after = Pt(6)

    tc_highlights = [
        "✔ TC-01 & TC-02: Slash command registration and inline keyboard callback handling verified.",
        "✔ TC-04 & TC-05: Local Medusa engine connection and graceful offline failover message confirmed.",
        "✔ TC-06 & TC-07: Real-time news and live crypto price retrieval tested via DuckDuckGo scraping.",
        "✔ TC-09 & TC-10: Anime character identification (Makima) and photo reply follow-ups verified.",
        "✔ TC-13 & TC-14: Deterministic HTML rendering (zero asterisks) and 4096-char chunk splitting verified.",
        "✔ TC-15 & TC-16: Model secrecy guardrails and global pip CLI execution verified.",
    ]
    for tc in tc_highlights:
        p = tf_b.add_paragraph()
        p.text = tc
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_MUTED
        p.space_after = Pt(4)

    # =========================================================================
    # SLIDE 11: KEY RESULTS & DEMONSTRATION HIGHLIGHTS
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_header(s11, "10. Key Results & Practical Demonstrations", "VALIDATION")

    res_cards = [
        ("📱 Live Telegram Bot",
         "Active Worldwide Access",
         "• Bot Username: @aimedusaroobot\n"
         "• Responds with typing indicators in <1 second.\n"
         "• Handles dual-engine switching, inline buttons, and multi-user concurrent sessions effortlessly.\n"
         "• Zero formatting errors: headers, bold text, and code blocks render with native elegance.",
         ACCENT_BLUE),
        ("🖼️ Visual Recognition",
         "Accurate Character Identification",
         "• Successfully recognized complex anime characters from raw user imagery.\n"
         "• Accurately parsed PDF research papers, returning structured chapter summaries.\n"
         "• Unpacked software ZIP archives, analyzing multi-file codebases in seconds.",
         EMERALD),
        ("🔒 Persona Confidentiality",
         "Strict Model Anonymity",
         "• Complete protection against prompt leaks: The model consistently identifies exclusively as Medusa AI.\n"
         "• When asked 'what model are you?', it never exposes Groq, Gemini, Qwen, or LLaMA.\n"
         "• The /model command was permanently decommissioned.",
         CYAN_ACCENT),
    ]

    for i, (title, sub, bullets, col) in enumerate(res_cards):
        x = Inches(0.8 + i * 3.95)
        c = add_card(s11, x, Inches(1.8), Inches(3.75), Inches(5.0))
        tf = c.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(6)

        psub = tf.add_paragraph()
        psub.text = sub
        psub.font.name = "Arial"
        psub.font.size = Pt(11)
        psub.font.bold = True
        psub.font.color.rgb = TEXT_DARK
        psub.space_after = Pt(12)

        for line in bullets.split("\n"):
            pb = tf.add_paragraph()
            pb.text = line
            pb.font.name = "Arial"
            pb.font.size = Pt(10.5)
            pb.font.color.rgb = TEXT_MUTED
            pb.space_after = Pt(6)

    # =========================================================================
    # SLIDE 12: CONCLUSION & FUTURE SCOPE (Closing Theme)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    bg12 = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg12.fill.solid()
    bg12.fill.fore_color.rgb = DARK_NAVY
    bg12.line.fill.background()

    # Header on Dark
    tbox = s12.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
    tf12 = tbox.text_frame
    p_cat = tf12.paragraphs[0]
    p_cat.text = "PROJECT WRAP-UP"
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = CYAN_ACCENT

    p_t = tf12.add_paragraph()
    p_t.text = "11. Conclusion & Future Roadmap"
    p_t.font.name = "Arial"
    p_t.font.size = Pt(24)
    p_t.font.bold = True
    p_t.font.color.rgb = TEXT_WHITE

    # Left: Conclusion Card
    c_concl = add_card(s12, Inches(0.8), Inches(1.8), Inches(5.75), Inches(4.8), bg_color=DEEP_SLATE, border_color=RGBColor(51, 65, 85))
    tf_c = c_concl.text_frame
    tf_c.word_wrap = True
    pt = tf_c.paragraphs[0]
    pt.text = "🌟 Summary of Contributions"
    pt.font.name = "Arial"
    pt.font.size = Pt(15)
    pt.font.bold = True
    pt.font.color.rgb = EMERALD
    pt.space_after = Pt(10)

    c_points = [
        "• Successfully engineered a hybrid edge-cloud conversational ecosystem.",
        "• Solved the model privacy dilemma by combining local RTX GPU execution with zero cloud telemetry.",
        "• Implemented zero-vector live web search for up-to-date factual intelligence.",
        "• Resolved multimodal character identification breakdowns for anime and gaming imagery.",
        "• Packaged a standalone pip library (`medusaai`) for developers worldwide.",
    ]
    for cp in c_points:
        p = tf_c.add_paragraph()
        p.text = cp
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(8)

    # Right: Future Scope & Q&A
    c_future = add_card(s12, Inches(6.78), Inches(1.8), Inches(5.75), Inches(4.8), bg_color=DEEP_SLATE, border_color=RGBColor(51, 65, 85))
    tf_f = c_future.text_frame
    tf_f.word_wrap = True
    pt = tf_f.paragraphs[0]
    pt.text = "🚀 Future Enhancements"
    pt.font.name = "Arial"
    pt.font.size = Pt(15)
    pt.font.bold = True
    pt.font.color.rgb = CYAN_ACCENT
    pt.space_after = Pt(10)

    f_points = [
        "1. Bidirectional Voice: Real-time speech-to-speech with Whisper.cpp and Edge-TTS.",
        "2. Fully Offline Vision: Integrating quantized MiniCPM-V GGUF into llama-server.",
        "3. Webhook Production Scaling: Deploying behind Cloudflare reverse proxies.",
        "4. Persistent Long-Term Memory: Vector-based user conversation recall.",
    ]
    for fp in f_points:
        p = tf_f.add_paragraph()
        p.text = fp
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(8)

    # Thank you footer inside right card
    p_ty = tf_f.add_paragraph()
    p_ty.text = "\nThank You!  Questions & Discussion Welcome."
    p_ty.font.name = "Arial"
    p_ty.font.size = Pt(12)
    p_ty.font.bold = True
    p_ty.font.color.rgb = RGBColor(253, 224, 71)  # Yellow

    # Save presentation
    prs.save(OUTPUT_PPTX)
    print(f"Presentation saved successfully with {len(prs.slides)} slides to {OUTPUT_PPTX}!")


if __name__ == "__main__":
    create_presentation()
