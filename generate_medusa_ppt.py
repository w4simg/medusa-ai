import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import win32com.client

BASE_DIR = r"c:\Users\unixb\OneDrive\Desktop\medusa_ai"
OUTPUT_PPTX = os.path.join(BASE_DIR, "Medusa_AI_Presentation.pptx")
OUTPUT_PDF = os.path.join(BASE_DIR, "Medusa_AI_Presentation.pdf")
IMG_TERMINAL = os.path.join(BASE_DIR, "Screenshot 2026-11-11 150945.png")
IMG_TELEGRAM = os.path.join(BASE_DIR, "Screenshot 2026-11-11 151327.png")

# Premium Modern Color Palette
DARK_NAVY = RGBColor(15, 23, 42)       # #0F172A
DEEP_SLATE = RGBColor(30, 41, 59)      # #1E293B
ACCENT_BLUE = RGBColor(37, 99, 235)    # #2563EB
CYAN_ACCENT = RGBColor(14, 165, 233)   # #0EA5E9
EMERALD = RGBColor(16, 185, 129)       # #10B981
CARD_BG = RGBColor(248, 250, 252)      # #F8FAFC
CARD_BORDER = RGBColor(226, 232, 240)  # #E2E8F0
TEXT_DARK = RGBColor(15, 23, 42)       # #0F172A
TEXT_MUTED = RGBColor(71, 85, 105)     # #475569
TEXT_WHITE = RGBColor(255, 255, 255)
CARD_WHITE = RGBColor(255, 255, 255)
AMBER_ACCENT = RGBColor(245, 158, 11)  # #F59E0B
ROSE_ACCENT = RGBColor(239, 68, 68)    # #EF4444


def add_slide_header(slide, title_text, category="MEDUSA AI", subtitle=None):
    """Adds a clean, modern header to a slide."""
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.73), Inches(1.1))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Category Pill / Tag
    p_cat = tf.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(10.5)
    p_cat.font.bold = True
    p_cat.font.color.rgb = ACCENT_BLUE
    p_cat.space_after = Pt(2)

    # Slide Title
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = "Arial"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = DARK_NAVY
    p_title.space_after = Pt(2)

    # Subtitle (if provided)
    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.name = "Arial"
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = TEXT_MUTED


def add_card(slide, left, top, width, height, bg_color=CARD_WHITE, border_color=CARD_BORDER, anchor_top=True):
    """Draws a rounded rectangular card background with top vertical anchor."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)
    tf = shape.text_frame
    if anchor_top:
        tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_top = Inches(0.24)
    tf.margin_bottom = Inches(0.24)
    tf.margin_left = Inches(0.26)
    tf.margin_right = Inches(0.26)
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
    p0.text = "MINOR PROJECT PRESENTATION | B.TECH CSE (2023 - 2027)"
    p0.font.name = "Arial"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = CYAN_ACCENT
    p0.space_after = Pt(8)

    p1 = tf1.add_paragraph()
    p1.text = "MEDUSA AI"
    p1.font.name = "Arial"
    p1.font.size = Pt(42)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_after = Pt(8)

    p2 = tf1.add_paragraph()
    p2.text = "A Smart, Dual-Engine Conversational AI Telegram Bot with Live Web Search, Character Vision, and Private Local Inference"
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
    p_c1_t.space_after = Pt(8)

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
        pm.font.size = Pt(11.5)
        pm.font.color.rgb = TEXT_WHITE
        pm.space_after = Pt(4)

    card2 = add_card(s1, c_left + c_w1 + Inches(0.2), c_top, c_w2, c_h, bg_color=DEEP_SLATE, border_color=RGBColor(51, 65, 85))
    tf_c2 = card2.text_frame
    tf_c2.word_wrap = True
    p_c2_t = tf_c2.paragraphs[0]
    p_c2_t.text = "ACADEMIC DETAILS"
    p_c2_t.font.name = "Arial"
    p_c2_t.font.size = Pt(11)
    p_c2_t.font.bold = True
    p_c2_t.font.color.rgb = CYAN_ACCENT
    p_c2_t.space_after = Pt(8)

    inst_lines = [
        ("Project Guide:", "Prof. Achinta Kumar Palit"),
        ("Department:", "Computer Science & Engineering"),
        ("College:", "GIET, Bhubaneswar"),
        ("Affiliation:", "BPUT, Odisha"),
    ]
    for lbl, val in inst_lines:
        p_in = tf_c2.add_paragraph()
        p_in.text = f"{lbl} {val}"
        p_in.font.name = "Arial"
        p_in.font.size = Pt(11.5)
        p_in.font.color.rgb = TEXT_WHITE
        p_in.space_after = Pt(5)

    # =========================================================================
    # SLIDE 2: WHAT IS MEDUSA AI? (OVERVIEW)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_header(s2, "What is Medusa AI?", "01 | PROJECT OVERVIEW",
                     "An intelligent Telegram assistant uniting fast cloud intelligence with 100% private local hardware.")

    cards_data_s2 = [
        ("⚡ Dual Engine Modes",
         "The Best of Both Worlds",
         [
             "• High-speed Cloud Mode: Delivers answers in milliseconds for everyday questions.",
             "• 100% Private Local Mode: Runs directly on local GPU without sending data to the cloud.",
             "• Seamlessly toggle between engines using interactive inline keyboard buttons.",
         ],
         ACCENT_BLUE),
        ("🌐 Live Web Knowledge",
         "Always Current & Accurate",
         [
             "• Automatically queries the live internet for breaking news, stock prices, or sports.",
             "• Overcomes static model cutoffs without requiring heavy vector databases.",
             "• Cites reliable sources and synthesizes factual answers directly in chat.",
         ],
         CYAN_ACCENT),
        ("📱 Native Telegram Bot",
         "Convenient Worldwide Access",
         [
             "• Available on iOS, Android, macOS, and Windows via Telegram (@aimedusaroobot).",
             "• Zero installation required for users; simply open chat and start talking.",
             "• Interactive buttons, live typing indicators, and clean formatted messages.",
         ],
         EMERALD),
    ]

    for i, (title, tag, points, color) in enumerate(cards_data_s2):
        x = Inches(0.8 + i * 3.95)
        card = add_card(s2, x, Inches(1.8), Inches(3.75), Inches(5.1))
        tf = card.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(17)
        pt.font.bold = True
        pt.font.color.rgb = color
        pt.space_after = Pt(4)

        ptag = tf.add_paragraph()
        ptag.text = tag
        ptag.font.name = "Arial"
        ptag.font.size = Pt(11.5)
        ptag.font.bold = True
        ptag.font.color.rgb = TEXT_DARK
        ptag.space_after = Pt(14)

        for pt_text in points:
            pb = tf.add_paragraph()
            pb.text = pt_text
            pb.font.name = "Arial"
            pb.font.size = Pt(11.5)
            pb.font.color.rgb = TEXT_MUTED
            pb.space_after = Pt(10)

    # =========================================================================
    # SLIDE 3: PROBLEMS WITH EXISTING CHATBOTS
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_header(s3, "Why Did We Build Medusa AI?", "02 | THE PROBLEM",
                     "Today's AI chatbots suffer from four major limitations that impact privacy, accuracy, and user convenience.")

    probs = [
        ("🔒 Cloud-Only & No Privacy",
         "Standard chatbots transmit all personal chats and confidential files to proprietary cloud servers. If internet access goes down, the chatbot fails completely.",
         ROSE_ACCENT),
        ("📅 Outdated Knowledge",
         "Pre-trained models are frozen in time with static knowledge cutoffs. They cannot tell you today's sports scores, live stock valuations, or recent breaking news.",
         AMBER_ACCENT),
        ("🖼️ Generic Vision AI",
         "When shown an image, standard bots provide superficial artistic commentary ('a person with hair') rather than recognizing specific anime, game, or pop culture characters.",
         ACCENT_BLUE),
        ("🌐 Clunky Web Interfaces",
         "Web-based chatbots require signing in through browsers, solving captchas, and consuming excessive RAM just to chat. Users prefer instant messaging apps like Telegram.",
         EMERALD),
    ]

    for i, (title, desc, col) in enumerate(probs):
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
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(6)

        pdesc = tf.add_paragraph()
        pdesc.text = desc
        pdesc.font.name = "Arial"
        pdesc.font.size = Pt(12)
        pdesc.font.color.rgb = TEXT_MUTED
        pdesc.line_spacing = 1.25

    # =========================================================================
    # SLIDE 4: KEY SYSTEM OBJECTIVES
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_header(s4, "Key Project Goals", "03 | OBJECTIVES",
                     "Practical goals focused on building a dependable, private, and capable Telegram assistant.")

    objs = [
        ("⚡ Hybrid Dual-Engine Inference",
         "Provide users with both ultra-fast cloud LPUs (>300 tokens/sec) and 100% private, offline neural models running on local NVIDIA GPUs.",
         ACCENT_BLUE),
        ("🔍 Zero-Vector Live Web Search",
         "Automatically detect when questions require current facts, fetch live DuckDuckGo results, and synthesize up-to-date answers in real time.",
         CYAN_ACCENT),
        ("🎯 Character-Level Visual Intelligence",
         "Accurately recognize anime, game, and celebrity figures from photos, and maintain photo reply memory so users can ask natural follow-up questions.",
         EMERALD),
        ("🚀 Frictionless Telegram Experience",
         "Design an intuitive Telegram bot with one-tap inline buttons, clean markdown-to-HTML formatting, and respectful group chat privacy.",
         DARK_NAVY),
    ]

    for i, (title, desc, col) in enumerate(objs):
        col_idx = i % 2
        row_idx = i // 2
        x = Inches(0.8 + col_idx * 5.95)
        y = Inches(1.8 + row_idx * 2.55)

        card = add_card(s4, x, y, Inches(5.75), Inches(2.35))
        tf = card.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = f"Goal 0{i+1}: {title}"
        pt.font.name = "Arial"
        pt.font.size = Pt(15.5)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(6)

        pdesc = tf.add_paragraph()
        pdesc.text = desc
        pdesc.font.name = "Arial"
        pdesc.font.size = Pt(12)
        pdesc.font.color.rgb = TEXT_MUTED
        pdesc.line_spacing = 1.25

    # =========================================================================
    # SLIDE 5: SYSTEM ARCHITECTURE & BACKEND SERVER (SCREENSHOT 1)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_header(s5, "System Architecture & Dual-Engine Workflow", "04 | SYSTEM ARCHITECTURE",
                     "A unified backend that intelligently routes Telegram prompts between cloud LPUs and local GPU models.")

    # Left Column: Dual-Engine Architecture Card
    c_arch = add_card(s5, Inches(0.8), Inches(1.8), Inches(5.1), Inches(5.1))
    tf_arch = c_arch.text_frame
    tf_arch.word_wrap = True

    p_a_title = tf_arch.paragraphs[0]
    p_a_title.text = "⚡ Dual Engine Architecture"
    p_a_title.font.name = "Arial"
    p_a_title.font.size = Pt(16)
    p_a_title.font.bold = True
    p_a_title.font.color.rgb = ACCENT_BLUE
    p_a_title.space_after = Pt(8)

    arch_points = [
        ("Default Cloud Engine:",
         "Powered by Groq LPUs for rapid conversational flow (>300 t/s). Features DuckDuckGo live web search for up-to-date facts."),
        ("Medusa Local Engine:",
         "Runs llama-server on port 8088 offloaded to an NVIDIA RTX 2050 GPU. Uses quantized GGUF weights for 100% offline privacy."),
        ("Smart Offline Failover:",
         "If the local server is stopped, the bot politely notifies the user to switch back to Default mode, ensuring zero crashed chats."),
    ]
    for lbl, desc in arch_points:
        p_lbl = tf_arch.add_paragraph()
        p_lbl.text = f"• {lbl}"
        p_lbl.font.name = "Arial"
        p_lbl.font.size = Pt(12)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = DARK_NAVY
        p_lbl.space_after = Pt(2)

        p_desc = tf_arch.add_paragraph()
        p_desc.text = desc
        p_desc.font.name = "Arial"
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = TEXT_MUTED
        p_desc.space_after = Pt(8)

    # Right Column: Framed Backend Terminal Screenshot Card
    c_img5 = add_card(s5, Inches(6.1), Inches(1.8), Inches(6.43), Inches(5.1))

    # Textbox inside card for clean title & subtitle without collision
    h_box5 = s5.shapes.add_textbox(Inches(6.3), Inches(1.95), Inches(6.03), Inches(0.65))
    tf_h5 = h_box5.text_frame
    tf_h5.word_wrap = True
    tf_h5.margin_left = tf_h5.margin_top = tf_h5.margin_right = tf_h5.margin_bottom = 0

    p_h5 = tf_h5.paragraphs[0]
    p_h5.text = "🖥️ Live VS Code Terminal & Server Health"
    p_h5.font.name = "Arial"
    p_h5.font.size = Pt(13.5)
    p_h5.font.bold = True
    p_h5.font.color.rgb = ACCENT_BLUE
    p_h5.space_after = Pt(1)

    p_h5_sub = tf_h5.add_paragraph()
    p_h5_sub.text = "Port 8088 Health Check 200 OK • Medusa Local Server Ready • Telegram Polling"
    p_h5_sub.font.name = "Arial"
    p_h5_sub.font.size = Pt(9.5)
    p_h5_sub.font.color.rgb = TEXT_MUTED

    # Embed Terminal Screenshot inside card below header
    if os.path.exists(IMG_TERMINAL):
        img_w = Inches(6.03)
        img_h = Inches(3.60)  # 1535x916 ratio -> 6.03 / 1.6757 = 3.60
        pic5 = s5.shapes.add_picture(IMG_TERMINAL, Inches(6.3), Inches(2.68), width=img_w, height=img_h)
        pic5.line.color.rgb = CARD_BORDER
        pic5.line.width = Pt(1)

    # =========================================================================
    # SLIDE 6: REAL-TIME WEB SEARCH
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_header(s6, "Real-Time Web Search Augmentation", "05 | WEB RETRIEVAL",
                     "A lightweight, zero-vector pipeline that fetches live facts from the web whenever needed.")

    # Top Overview Card
    top_c = add_card(s6, Inches(0.8), Inches(1.8), Inches(11.73), Inches(1.3))
    tf_top = top_c.text_frame
    tf_top.word_wrap = True
    p_tt = tf_top.paragraphs[0]
    p_tt.text = "💡 How Medusa Answers Questions About Today's Events"
    p_tt.font.name = "Arial"
    p_tt.font.size = Pt(14)
    p_tt.font.bold = True
    p_tt.font.color.rgb = ACCENT_BLUE
    p_tt.space_after = Pt(2)

    p_td = tf_top.add_paragraph()
    p_td.text = "Instead of complex vector databases, Medusa AI uses an on-demand web search heuristic. When a question mentions time-sensitive words, Medusa searches the live internet and blends fresh facts directly into its answer."
    p_td.font.name = "Arial"
    p_td.font.size = Pt(11.5)
    p_td.font.color.rgb = TEXT_MUTED

    # 4 Simple Step Cards
    steps = [
        ("Step 1", "Detect Intent",
         "The system checks for keywords like 'today', 'latest', 'current', 'news', 'price', or 'who won'."),
        ("Step 2", "Clean Query",
         "Strips polite conversational words ('please tell me about') to produce clean, high-impact search keywords."),
        ("Step 3", "DuckDuckGo Search",
         "Fetches the top 4 factual snippets, news headlines, and source links asynchronously in the background."),
        ("Step 4", "Smart Synthesis",
         "Injects search results into the model context to deliver an accurate, up-to-date answer with citations."),
    ]

    for i, (step_num, title, desc) in enumerate(steps):
        x = Inches(0.8 + i * 2.98)
        c = add_card(s6, x, Inches(3.3), Inches(2.78), Inches(3.6))
        tf = c.text_frame
        tf.word_wrap = True

        ps = tf.paragraphs[0]
        ps.text = step_num.upper()
        ps.font.name = "Arial"
        ps.font.size = Pt(11)
        ps.font.bold = True
        ps.font.color.rgb = CYAN_ACCENT
        ps.space_after = Pt(4)

        pt = tf.add_paragraph()
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = DARK_NAVY
        pt.space_after = Pt(8)

        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.name = "Arial"
        pd.font.size = Pt(11.5)
        pd.font.color.rgb = TEXT_MUTED
        pd.line_spacing = 1.25

    # =========================================================================
    # SLIDE 7: MULTIMODAL VISION & CHARACTER RECOGNITION
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_header(s7, "Multimodal Vision & Character Recognition", "06 | COMPUTER VISION",
                     "Advanced visual intelligence designed to name characters and inspect files accurately.")

    v_cards = [
        ("🎯 Character Identification",
         "Naming the Figure Accurately",
         [
             "• Standard AI models only give vague art descriptions (e.g., 'a red-haired anime girl').",
             "• Medusa AI identifies the exact character name, franchise, and background lore.",
             "• Example: Instantly recognizes 'Makima from Chainsaw Man' and explains her role.",
         ],
         ACCENT_BLUE),
        ("💬 Photo Reply Memory",
         "Natural Image Conversations",
         [
             "• Remembers the last uploaded photo per chat in active memory.",
             "• If a user replies to an image with 'who is she?' or 'what anime is this?', Medusa understands immediately.",
             "• Eliminates the frustration of having to re-upload photos for every follow-up question.",
         ],
         CYAN_ACCENT),
        ("📑 Document & Code Inspection",
         "PDFs, ZIP Archives & Audio",
         [
             "• PDF Analysis: Reads multi-page research papers and notes via PyMuPDF (fitz).",
             "• ZIP Archives: Inspects project structure and reviews Python/JS code automatically.",
             "• Audio Notes: Transcribes voice messages into clean text for quick answers.",
         ],
         EMERALD),
    ]

    for i, (title, tag, points, col) in enumerate(v_cards):
        x = Inches(0.8 + i * 3.95)
        c = add_card(s7, x, Inches(1.8), Inches(3.75), Inches(5.1))
        tf = c.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(15.5)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(4)

        ptag = tf.add_paragraph()
        ptag.text = tag
        ptag.font.name = "Arial"
        ptag.font.size = Pt(11.5)
        ptag.font.bold = True
        ptag.font.color.rgb = TEXT_DARK
        ptag.space_after = Pt(14)

        for pt_text in points:
            pb = tf.add_paragraph()
            pb.text = pt_text
            pb.font.name = "Arial"
            pb.font.size = Pt(11.5)
            pb.font.color.rgb = TEXT_MUTED
            pb.space_after = Pt(10)

    # =========================================================================
    # SLIDE 8: TELEGRAM BOT INTERFACE (SCREENSHOT 2)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_header(s8, "Telegram Bot Interface & User Experience", "07 | TELEGRAM USER INTERFACE",
                     "A polished mobile and desktop messaging experience featuring interactive buttons and clean formatting.")

    # Left Column: Bot Feature Highlights
    c_bot = add_card(s8, Inches(0.8), Inches(1.8), Inches(5.1), Inches(5.1))
    tf_bot = c_bot.text_frame
    tf_bot.word_wrap = True

    p_b_title = tf_bot.paragraphs[0]
    p_b_title.text = "📱 Interactive Bot Features"
    p_b_title.font.name = "Arial"
    p_b_title.font.size = Pt(16)
    p_b_title.font.bold = True
    p_b_title.font.color.rgb = ACCENT_BLUE
    p_b_title.space_after = Pt(8)

    bot_features = [
        ("One-Tap Inline Buttons:",
         "Switch between Default and Medusa modes, trigger help, or reset memory with a single click, eliminating manual commands."),
        ("Live Typing Indicators:",
         "Displays continuous typing feedback while the model generates answers, so users know the bot is working."),
        ("Clean HTML Formatting:",
         "Converts markdown to native Telegram HTML, completely eliminating ugly raw asterisks (***)."),
        ("Smart Group Privacy:",
         "Only responds when directly @mentioned or replied to, keeping group chats uncluttered."),
    ]
    for lbl, desc in bot_features:
        p_lbl = tf_bot.add_paragraph()
        p_lbl.text = f"• {lbl}"
        p_lbl.font.name = "Arial"
        p_lbl.font.size = Pt(12)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = DARK_NAVY
        p_lbl.space_after = Pt(2)

        p_desc = tf_bot.add_paragraph()
        p_desc.text = desc
        p_desc.font.name = "Arial"
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = TEXT_MUTED
        p_desc.space_after = Pt(8)

    # Right Column: Framed Telegram UI Screenshot Card
    c_img8 = add_card(s8, Inches(6.1), Inches(1.8), Inches(6.43), Inches(5.1))

    # Textbox inside card for clean title & subtitle without collision
    h_box8 = s8.shapes.add_textbox(Inches(6.3), Inches(1.95), Inches(6.03), Inches(0.65))
    tf_h8 = h_box8.text_frame
    tf_h8.word_wrap = True
    tf_h8.margin_left = tf_h8.margin_top = tf_h8.margin_right = tf_h8.margin_bottom = 0

    p_h8 = tf_h8.paragraphs[0]
    p_h8.text = "💬 Live Telegram Chat with @aimedusaroobot"
    p_h8.font.name = "Arial"
    p_h8.font.size = Pt(13.5)
    p_h8.font.bold = True
    p_h8.font.color.rgb = ACCENT_BLUE
    p_h8.space_after = Pt(1)

    p_h8_sub = tf_h8.add_paragraph()
    p_h8_sub.text = "Interactive Buttons [⚡ Default] [🐍 Medusa] • Real-Time Mode Switching"
    p_h8_sub.font.name = "Arial"
    p_h8_sub.font.size = Pt(9.5)
    p_h8_sub.font.color.rgb = TEXT_MUTED

    # Embed Telegram Screenshot inside card below header
    if os.path.exists(IMG_TELEGRAM):
        img_w = Inches(6.03)
        img_h = Inches(3.81)  # 1396x882 ratio -> 6.03 / 1.5828 = 3.81
        pic8 = s8.shapes.add_picture(IMG_TELEGRAM, Inches(6.3), Inches(2.68), width=img_w, height=img_h)
        pic8.line.color.rgb = CARD_BORDER
        pic8.line.width = Pt(1)

    # =========================================================================
    # SLIDE 9: TELEGRAM COMMANDS & USER INTERACTION FLOW
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_header(s9, "Telegram Commands & User Interaction Flow", "08 | TELEGRAM CONTROL",
                     "An intuitive command suite and automated interaction rules designed for smooth messaging.")

    # Left Card: Command Reference
    c_cmd = add_card(s9, Inches(0.8), Inches(1.8), Inches(5.75), Inches(5.1))
    tf_cmd = c_cmd.text_frame
    tf_cmd.word_wrap = True

    p_cmd_title = tf_cmd.paragraphs[0]
    p_cmd_title.text = "🤖 Supported Telegram Commands"
    p_cmd_title.font.name = "Arial"
    p_cmd_title.font.size = Pt(16)
    p_cmd_title.font.bold = True
    p_cmd_title.font.color.rgb = ACCENT_BLUE
    p_cmd_title.space_after = Pt(10)

    commands_list = [
        ("/start", "Launches the bot welcome menu, capabilities overview, and inline mode switch buttons."),
        ("/default", "Switches active engine to Cloud Mode (high-speed Groq LPUs with live DuckDuckGo web search)."),
        ("/medusa", "Switches active engine to Local Mode (100% private offline neural model on local GPU)."),
        ("/help", "Displays a quick-start guide, available features, and tips for photo & document analysis."),
        ("/reset", "Clears the active conversation context window to start a fresh dialogue session."),
    ]
    for cmd, desc in commands_list:
        p_c = tf_cmd.add_paragraph()
        p_c.text = f"•  {cmd} :"
        p_c.font.name = "Consolas"
        p_c.font.size = Pt(11.5)
        p_c.font.bold = True
        p_c.font.color.rgb = DARK_NAVY
        p_c.space_after = Pt(1)

        p_cd = tf_cmd.add_paragraph()
        p_cd.text = f"    {desc}"
        p_cd.font.name = "Arial"
        p_cd.font.size = Pt(10.5)
        p_cd.font.color.rgb = TEXT_MUTED
        p_cd.space_after = Pt(6)

    # Right Card: Smart Telegram Interaction Rules
    c_rules = add_card(s9, Inches(6.78), Inches(1.8), Inches(5.75), Inches(5.1))
    tf_rules = c_rules.text_frame
    tf_rules.word_wrap = True

    p_r_title = tf_rules.paragraphs[0]
    p_r_title.text = "⚙️ Automated Interaction Rules"
    p_r_title.font.name = "Arial"
    p_r_title.font.size = Pt(16)
    p_r_title.font.bold = True
    p_r_title.font.color.rgb = CYAN_ACCENT
    p_r_title.space_after = Pt(10)

    rules = [
        ("🎛️ Inline Keyboard Callbacks:",
         "Tapping [⚡ Default Mode] or [🐍 Medusa Mode] buttons instantly switches engines without having to type commands manually."),
        ("👥 Group Chat Privacy Filter:",
         "In group chats, Medusa stays silent during casual chatter. It only responds when directly @mentioned or replied to, keeping groups clean."),
        ("✂️ 4,096-Character Message Chunking:",
         "Telegram limits messages to 4,096 characters. Medusa automatically splits long essays, code files, and research summaries by paragraphs."),
        ("🧠 12-Turn Context Sliding Window:",
         "Maintains conversation history per user up to 12 turns, enabling natural multi-turn discussions while preventing token overflows."),
    ]
    for r_title, r_desc in rules:
        pr_t = tf_rules.add_paragraph()
        pr_t.text = f"• {r_title}"
        pr_t.font.name = "Arial"
        pr_t.font.size = Pt(12)
        pr_t.font.bold = True
        pr_t.font.color.rgb = DARK_NAVY
        pr_t.space_after = Pt(2)

        pr_d = tf_rules.add_paragraph()
        pr_d.text = r_desc
        pr_d.font.name = "Arial"
        pr_d.font.size = Pt(11)
        pr_d.font.color.rgb = TEXT_MUTED
        pr_d.space_after = Pt(8)

    # =========================================================================
    # SLIDE 10: PERFORMANCE BENCHMARKS & TESTING
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_header(s10, "Performance & Verification Results", "09 | PERFORMANCE & TESTING",
                     "Empirical benchmarks demonstrating rapid token generation and 100% test case validation.")

    # 4 Metric Cards
    metrics = [
        ("Cloud Generation Speed", "310 tokens/s", "Groq LPU ultra-fast inference", ACCENT_BLUE),
        ("Local Model Speed", "28 tokens/s", "NVIDIA RTX 2050 GPU offload", EMERALD),
        ("Cloud Latency (TTFT)", "340 ms", "Sub-second response time", CYAN_ACCENT),
        ("GPU VRAM Footprint", "2.3 GB", "Fits easily on 4GB consumer laptops", RGBColor(168, 85, 247)),
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
        pn.font.size = Pt(10)
        pn.font.color.rgb = TEXT_DARK

    # Bottom Verification Highlights Card
    b_card = add_card(s10, Inches(0.8), Inches(3.9), Inches(11.73), Inches(3.0))
    tf_b = b_card.text_frame
    tf_b.word_wrap = True

    pt_b = tf_b.paragraphs[0]
    pt_b.text = "🧪 System Testing Summary (16 / 16 Test Cases Passed - 100%)"
    pt_b.font.name = "Arial"
    pt_b.font.size = Pt(14)
    pt_b.font.bold = True
    pt_b.font.color.rgb = DARK_NAVY
    pt_b.space_after = Pt(8)

    tc_items = [
        ("✔ Dual-Engine Switching:", "Verified smooth toggling between /default (cloud) and /medusa (local)."),
        ("✔ Offline Failover:", "Gracefully warns user when the local server is stopped, preventing crashes."),
        ("✔ Live Web Retrieval:", "Accurately retrieved current live sports scores, weather, and news via DuckDuckGo."),
        ("✔ Anime & Character Vision:", "Accurately identified Makima from Chainsaw Man and answered follow-up queries."),
        ("✔ Formatting & Clean Output:", "Eliminated raw markdown asterisks (***), outputting clean Telegram HTML."),
        ("✔ Group Chat Privacy Filter:", "Confirmed bot stays silent in groups until directly @mentioned or replied to."),
    ]
    for lbl, desc in tc_items:
        p = tf_b.add_paragraph()
        p.text = f"{lbl}  {desc}"
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MUTED
        p.space_after = Pt(4)

    # =========================================================================
    # SLIDE 11: SYSTEM DEMONSTRATION & PRACTICAL VALIDATION
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_header(s11, "Live System Demonstration Highlights", "10 | SYSTEM VALIDATION",
                     "Real-world validation confirming end-to-end functionality across mobile and desktop Telegram sessions.")

    demo_cards = [
        ("⚡ Dual Engine Demo",
         "Cloud Speed & Local Privacy",
         [
             "• Click [⚡ Default Mode]: Bot switches instantly and delivers sub-second answers with web search.",
             "• Click [🐍 Medusa Mode]: Switches to local GPU llama-server, processing prompts 100% offline.",
             "• Real-time feedback via interactive Telegram buttons.",
         ],
         ACCENT_BLUE),
        ("🖼️ Visual Recognition Demo",
         "Character & Photo Follow-Ups",
         [
             "• Upload an anime/game image: Medusa accurately identifies the character name and franchise.",
             "• Reply 'who is she?': Contextual memory recalls the previous photo seamlessly.",
             "• Verified with complex anime illustrations and pop culture figures.",
         ],
         EMERALD),
        ("📑 File & Document Demo",
         "Multi-Page PDFs & Code Archives",
         [
             "• Upload a research PDF: Extracted and summarized chapters in seconds via PyMuPDF.",
             "• Send software ZIP archives: Analyzed directory structure and code quality automatically.",
             "• Audio voice notes transcribed into clean text replies.",
         ],
         CYAN_ACCENT),
    ]

    for i, (title, tag, points, col) in enumerate(demo_cards):
        x = Inches(0.8 + i * 3.95)
        c = add_card(s11, x, Inches(1.8), Inches(3.75), Inches(5.1))
        tf = c.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.name = "Arial"
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(4)

        ptag = tf.add_paragraph()
        ptag.text = tag
        ptag.font.name = "Arial"
        ptag.font.size = Pt(11.5)
        ptag.font.bold = True
        ptag.font.color.rgb = TEXT_DARK
        ptag.space_after = Pt(14)

        for pt_text in points:
            pb = tf.add_paragraph()
            pb.text = pt_text
            pb.font.name = "Arial"
            pb.font.size = Pt(11.5)
            pb.font.color.rgb = TEXT_MUTED
            pb.space_after = Pt(10)

    # =========================================================================
    # SLIDE 12: CONCLUSION & FUTURE ROADMAP (Dark Hero Theme)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    bg12 = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg12.fill.solid()
    bg12.fill.fore_color.rgb = DARK_NAVY
    bg12.line.fill.background()

    # Header on Dark
    tbox = s12.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.73), Inches(1.0))
    tf12 = tbox.text_frame
    p_cat = tf12.paragraphs[0]
    p_cat.text = "PROJECT WRAP-UP"
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = CYAN_ACCENT

    p_t = tf12.add_paragraph()
    p_t.text = "Conclusion & Future Roadmap"
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
    pt.font.size = Pt(16)
    pt.font.bold = True
    pt.font.color.rgb = EMERALD
    pt.space_after = Pt(10)

    c_points = [
        "• Engineered a hybrid AI assistant uniting fast cloud LPUs with private local GPU models.",
        "• Implemented zero-vector live web search for up-to-date factual answers.",
        "• Solved character recognition and photo follow-up questions for anime, gaming, and pop culture.",
        "• Delivered a smooth Telegram Bot (@aimedusaroobot) with interactive buttons and clean HTML formatting.",
        "• Created a reliable, privacy-first personal assistant accessible worldwide on mobile & desktop.",
    ]
    for cp in c_points:
        p = tf_c.add_paragraph()
        p.text = cp
        p.font.name = "Arial"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(8)

    # Right: Future Scope & Q&A
    c_future = add_card(s12, Inches(6.78), Inches(1.8), Inches(5.75), Inches(4.8), bg_color=DEEP_SLATE, border_color=RGBColor(51, 65, 85))
    tf_f = c_future.text_frame
    tf_f.word_wrap = True

    pt_f = tf_f.paragraphs[0]
    pt_f.text = "🚀 Future Roadmap"
    pt_f.font.name = "Arial"
    pt_f.font.size = Pt(16)
    pt_f.font.bold = True
    pt_f.font.color.rgb = CYAN_ACCENT
    pt_f.space_after = Pt(10)

    f_points = [
        "1. Real-Time Voice: Bidirectional speech conversations using Whisper.cpp and Edge-TTS.",
        "2. Fully Offline Vision: Local quantized vision models (MiniCPM-V) running on GPU.",
        "3. Cloud Production Scaling: Webhook deployment behind Cloudflare reverse proxies.",
        "4. Persistent Long-Term Memory: Cross-session recall for personalized user experiences.",
    ]
    for fp in f_points:
        p = tf_f.add_paragraph()
        p.text = fp
        p.font.name = "Arial"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(8)

    # Thank you footer inside right card
    p_ty = tf_f.add_paragraph()
    p_ty.text = "\nThank You!  Questions & Discussion Welcome."
    p_ty.font.name = "Arial"
    p_ty.font.size = Pt(13)
    p_ty.font.bold = True
    p_ty.font.color.rgb = RGBColor(253, 224, 71)  # Yellow Accent

    # Save presentation
    prs.save(OUTPUT_PPTX)
    print(f"Presentation saved successfully with {len(prs.slides)} slides to {OUTPUT_PPTX}!")

    # Export to PDF via PowerPoint COM
    export_to_pdf(OUTPUT_PPTX, OUTPUT_PDF)


def export_to_pdf(pptx_path, pdf_path):
    print("Exporting presentation to PDF via PowerPoint COM...")
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        deck = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        # Format 32 = ppSaveAsPDF
        deck.SaveAs(os.path.abspath(pdf_path), 32)
        deck.Close()
        powerpoint.Quit()
        print(f"PDF exported successfully to {pdf_path}!")
    except Exception as e:
        print(f"Error exporting PDF via PowerPoint COM: {e}")


if __name__ == "__main__":
    create_presentation()
