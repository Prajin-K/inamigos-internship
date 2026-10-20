import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    # 16:9 widescreen layout
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # Colors
    NAVY = RGBColor(24, 76, 120)       # #184C78
    DARK_BG = RGBColor(15, 23, 42)      # #0F172A
    TEAL = RGBColor(13, 148, 136)      # #0D9488
    LIGHT_BG = RGBColor(248, 250, 252) # #F8FAFC
    CARD_BG = RGBColor(255, 255, 255)  # White
    CARD_BORDER = RGBColor(226, 232, 240)
    TEXT_DARK = RGBColor(30, 41, 59)   # Slate 800
    TEXT_MUTED = RGBColor(100, 116, 139) # Slate 500
    WHITE = RGBColor(255, 255, 255)

    def add_background(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_badge=None):
        add_background(slide, LIGHT_BG)

        # Top Accent Line
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = TEAL
        top_bar.line.fill.background()

        # Category Badge if available
        if category_badge:
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(3.2), Inches(0.4))
            badge.fill.solid()
            badge.fill.fore_color.rgb = RGBColor(224, 242, 254)
            badge.line.fill.background()
            bp = badge.text_frame.paragraphs[0]
            bp.alignment = PP_ALIGN.CENTER
            br = bp.add_run()
            br.text = category_badge.upper()
            br.font.size = Pt(11)
            br.font.bold = True
            br.font.color.rgb = NAVY

        # Slide Title
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.85 if category_badge else 0.5), Inches(11.7), Inches(0.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title_text
        r.font.size = Pt(26)
        r.font.bold = True
        r.font.color.rgb = NAVY

    # ==========================
    # SLIDE 1: Title Slide (Dark Theme)
    # ==========================
    s1 = prs.slides.add_slide(blank_layout)
    add_background(s1, DARK_BG)

    # Accent Top Stripe
    accent = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18))
    accent.fill.solid()
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()

    # Title Container
    s1_title_box = s1.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11.0), Inches(3.2))
    tf1 = s1_title_box.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.space_after = Pt(8)
    r0 = p0.add_run()
    r0.text = "INAMIGOS FOUNDATION INTERNSHIP  •  TASK 1"
    r0.font.size = Pt(13)
    r0.font.bold = True
    r0.font.color.rgb = TEAL

    p1 = tf1.add_paragraph()
    p1.space_after = Pt(12)
    r1 = p1.add_run()
    r1.text = "AI Tools Discovery & Exploration"
    r1.font.size = Pt(40)
    r1.font.bold = True
    r1.font.color.rgb = WHITE

    p2 = tf1.add_paragraph()
    r2 = p2.add_run()
    r2.text = "Analyzing 5 Transformative AI Solutions Shaping Education, Productivity, Design & Media"
    r2.font.size = Pt(16)
    r2.font.color.rgb = RGBColor(203, 213, 225)

    # Candidate Meta Box
    meta_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(5.3), Inches(10.9), Inches(1.2))
    meta_box.fill.solid()
    meta_box.fill.fore_color.rgb = RGBColor(30, 41, 59)
    meta_box.line.color.rgb = RGBColor(51, 65, 85)
    
    mtf = meta_box.text_frame
    mp = mtf.paragraphs[0]
    mp.alignment = PP_ALIGN.LEFT
    mr1 = mp.add_run()
    mr1.text = "Presented by: "
    mr1.font.size = Pt(13)
    mr1.font.color.rgb = RGBColor(148, 163, 184)
    mr2 = mp.add_run()
    mr2.text = "Prajin K   |   "
    mr2.font.size = Pt(14)
    mr2.font.bold = True
    mr2.font.color.rgb = WHITE
    mr3 = mp.add_run()
    mr3.text = "Task: "
    mr3.font.size = Pt(13)
    mr3.font.color.rgb = RGBColor(148, 163, 184)
    mr4 = mp.add_run()
    mr4.text = "Task 1: AI Tools Discovery   |   "
    mr4.font.size = Pt(14)
    mr4.font.bold = True
    mr4.font.color.rgb = WHITE
    mr5 = mp.add_run()
    mr5.text = "Platform: "
    mr5.font.size = Pt(13)
    mr5.font.color.rgb = RGBColor(148, 163, 184)
    mr6 = mp.add_run()
    mr6.text = "InAmigos Foundation"
    mr6.font.size = Pt(14)
    mr6.font.bold = True
    mr6.font.color.rgb = TEAL

    # ==========================
    # SLIDE 2: Executive Overview
    # ==========================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Executive Overview: Navigating the Modern AI Ecosystem")

    # 3 Pillar Cards
    pillars = [
        ("Why AI Exploration?", "AI tools have transitioned from technical curiosities into core infrastructure across industries. Understanding specialized tools is now a foundational competency for all professionals."),
        ("Categories Explored", "To ensure holistic discovery, this study analyzes 5 distinct domains: Academic Research, Workspace Productivity, Creative Image Generation, Synthetic Video, and Presentation Systems."),
        ("Key Evaluation Criteria", "Each tool is scrutinized on its core purpose, real-world practical use case, friction reduction, and democratization of technical skills.")
    ]

    for idx, (head, desc) in enumerate(pillars):
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + idx * 3.95), Inches(2.0), Inches(3.75), Inches(4.5))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_left = Inches(0.3)
        ctf.margin_right = Inches(0.3)
        ctf.margin_top = Inches(0.4)

        cp0 = ctf.paragraphs[0]
        cp0.space_after = Pt(14)
        cpr0 = cp0.add_run()
        cpr0.text = f"0{idx+1}"
        cpr0.font.size = Pt(28)
        cpr0.font.bold = True
        cpr0.font.color.rgb = TEAL

        cp1 = ctf.add_paragraph()
        cp1.space_after = Pt(10)
        cpr1 = cp1.add_run()
        cpr1.text = head
        cpr1.font.size = Pt(16)
        cpr1.font.bold = True
        cpr1.font.color.rgb = NAVY

        cp2 = ctf.add_paragraph()
        cpr2 = cp2.add_run()
        cpr2.text = desc
        cpr2.font.size = Pt(12)
        cpr2.font.color.rgb = TEXT_DARK

    # ==========================
    # TOOL SLIDES (5 SLIDES)
    # ==========================
    tools = [
        {
            "name": "Perplexity AI",
            "category": "Education & Academic Research",
            "tagline": "The Citation-Backed AI Search & Discovery Engine",
            "purpose": "A conversational AI search engine designed to provide comprehensive, factual answers backed by verifiable real-time web citations and academic repositories.",
            "use_case_title": "Academic Research & Literature Synthesis",
            "use_case_desc": "A student investigating renewable energy policies queries Perplexity. Within seconds, it synthesizes recent whitepapers and policy journals into an executive summary with interactive numeric citations linking directly to verified publications.",
            "highlights": [
                "Real-Time Web Indexing with source links",
                "Academic Filter for peer-reviewed papers",
                "Collections feature for structured project files",
                "Direct PDF/Document analysis"
            ],
            "impact": "Cuts research time by 70% while mitigating hallucinations."
        },
        {
            "name": "Notion AI",
            "category": "Productivity & Workspace Management",
            "tagline": "Connected Intelligence for Notes, Docs & Projects",
            "purpose": "An integrated workflow AI assistant embedded natively inside Notion, automating writing, project tracking, synthesis, and cross-workspace knowledge retrieval.",
            "use_case_title": "Meeting Minutes & Task Automation",
            "use_case_desc": "After a team sync, an intern pastes raw bullet points into Notion. Notion AI extracts key decisions, creates a table of action items assigned to teammates, and drafts a stakeholder update email with one click.",
            "highlights": [
                "Instant Meeting Notes Summarization",
                "Workspace Q&A across databases & Slack",
                "Tone Refactoring & Copy Editing",
                "Automated Project Property Autofill"
            ],
            "impact": "Eliminates documentation overhead and streamlines team alignment."
        },
        {
            "name": "Midjourney",
            "category": "Image Generation & Visual Arts",
            "tagline": "State-of-the-Art Diffusion Text-to-Image Synthesis",
            "purpose": "A cutting-edge generative AI model that translates descriptive natural language prompts into high-resolution, photorealistic, and stylized visual artwork.",
            "use_case_title": "Campaign Concept Art & Mood Boarding",
            "use_case_desc": "A creative intern designing an awareness campaign for InAmigos Foundation generates vivid photorealistic concept artwork depicting community impact and futuristic green initiatives in seconds, replacing costly stock photos.",
            "highlights": [
                "Unmatched Aesthetic Quality & Photorealism",
                "Precise Aspect Ratio & Camera Controls",
                "Inpainting (Vary Region) & Canvas Expansion",
                "Style Referencing & Consistent Character Prompting"
            ],
            "impact": "Democratizes high-end visual design for non-artists."
        },
        {
            "name": "Synthesia",
            "category": "Video Creation & AI Avatars",
            "tagline": "Text-to-Video Production with Realistic Digital Avatars",
            "purpose": "An enterprise-grade synthetic media platform that converts written scripts into studio-grade presenter videos featuring expressive AI avatars in 120+ languages.",
            "use_case_title": "Multilingual Volunteer Onboarding Modules",
            "use_case_desc": "An organization creates a global volunteer training series. Instead of booking actors, studios, and cameras, they input their training script into Synthesia, which generates polished video tutorials in English, Hindi, and Spanish.",
            "highlights": [
                "150+ Diverse Photorealistic Human Avatars",
                "Voice Synthesis in 120+ Regional Languages",
                "Turnkey Script-to-Video Workflow in minutes",
                "Easy updates without costly re-shoots"
            ],
            "impact": "Reduces corporate video production costs and timelines by 80%."
        },
        {
            "name": "Gamma AI (Gamma.app)",
            "category": "Presentations & Interactive Decks",
            "tagline": "Prompt-to-Presentation with Dynamic Responsive Layouts",
            "purpose": "An AI-powered presentation and document generator that produces polished, beautifully structured, interactive card decks from brief outlines or text notes.",
            "use_case_title": "Instant Pitch Deck & Report Formatting",
            "use_case_desc": "Needing to present project results on short notice, an intern submits an outline into Gamma AI. The tool builds a 10-slide responsive presentation complete with modern typography, card layouts, and visual hierarchy ready to share.",
            "highlights": [
                "One-Click Prompt/Doc to Presentation",
                "Fluid Card UI breaking rigid slide limitations",
                "Interactive Embeds (Websites, Figma, Video)",
                "One-Click Restyling & Responsive Formatting"
            ],
            "impact": "Allows presenters to focus on storytelling rather than alignment."
        }
    ]

    for tool in tools:
        slide = prs.slides.add_slide(blank_layout)
        add_header(slide, tool['name'], tool['category'])

        # Subtitle / Tagline
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.45), Inches(11.7), Inches(0.4))
        tt = tag_box.text_frame
        tp = tt.paragraphs[0]
        tr = tp.add_run()
        tr.text = tool['tagline']
        tr.font.size = Pt(13)
        tr.font.italic = True
        tr.font.color.rgb = TEAL

        # Left Column: Purpose & Practical Use Case
        left_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(7.4), Inches(4.8))
        left_card.fill.solid()
        left_card.fill.fore_color.rgb = CARD_BG
        left_card.line.color.rgb = CARD_BORDER

        ltf = left_card.text_frame
        ltf.word_wrap = True
        ltf.margin_left = Inches(0.35)
        ltf.margin_right = Inches(0.35)
        ltf.margin_top = Inches(0.35)

        # Core Purpose
        p_phead = ltf.paragraphs[0]
        p_phead.space_after = Pt(4)
        r_phead = p_phead.add_run()
        r_phead.text = "CORE PURPOSE"
        r_phead.font.size = Pt(11)
        r_phead.font.bold = True
        r_phead.font.color.rgb = NAVY

        p_ptxt = ltf.add_paragraph()
        p_ptxt.space_after = Pt(14)
        r_ptxt = p_ptxt.add_run()
        r_ptxt.text = tool['purpose']
        r_ptxt.font.size = Pt(12)
        r_ptxt.font.color.rgb = TEXT_DARK

        # Practical Use Case
        p_uhead = ltf.add_paragraph()
        p_uhead.space_after = Pt(4)
        r_uhead = p_uhead.add_run()
        r_uhead.text = f"PRACTICAL USE CASE: {tool['use_case_title'].upper()}"
        r_uhead.font.size = Pt(11)
        r_uhead.font.bold = True
        r_uhead.font.color.rgb = TEAL

        p_utxt = ltf.add_paragraph()
        p_utxt.space_after = Pt(12)
        r_utxt = p_utxt.add_run()
        r_utxt.text = tool['use_case_desc']
        r_utxt.font.size = Pt(12)
        r_utxt.font.color.rgb = TEXT_DARK

        # Right Column: Highlights & Impact
        right_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.45), Inches(2.0), Inches(4.05), Inches(4.8))
        right_card.fill.solid()
        right_card.fill.fore_color.rgb = CARD_BG
        right_card.line.color.rgb = CARD_BORDER

        rtf = right_card.text_frame
        rtf.word_wrap = True
        rtf.margin_left = Inches(0.3)
        rtf.margin_right = Inches(0.3)
        rtf.margin_top = Inches(0.35)

        p_khead = rtf.paragraphs[0]
        p_khead.space_after = Pt(8)
        r_khead = p_khead.add_run()
        r_khead.text = "KEY CAPABILITIES"
        r_khead.font.size = Pt(11)
        r_khead.font.bold = True
        r_khead.font.color.rgb = NAVY

        for hl in tool['highlights']:
            hp = rtf.add_paragraph()
            hp.space_after = Pt(6)
            hr = hp.add_run()
            hr.text = f"•  {hl}"
            hr.font.size = Pt(11.5)
            hr.font.color.rgb = TEXT_DARK

        # Impact Badge
        p_ihead = rtf.add_paragraph()
        p_ihead.space_before = Pt(14)
        p_ihead.space_after = Pt(4)
        r_ihead = p_ihead.add_run()
        r_ihead.text = "WORKFLOW IMPACT"
        r_ihead.font.size = Pt(11)
        r_ihead.font.bold = True
        r_ihead.font.color.rgb = TEAL

        p_itxt = rtf.add_paragraph()
        r_itxt = p_itxt.add_run()
        r_itxt.text = tool['impact']
        r_itxt.font.size = Pt(11.5)
        r_itxt.font.italic = True
        r_itxt.font.color.rgb = TEXT_DARK

    # ==========================
    # SLIDE 8: Cross-Tool Comparison Matrix
    # ==========================
    s_matrix = prs.slides.add_slide(blank_layout)
    add_header(s_matrix, "Comparative Analysis: AI Tools Across Categories")

    rows, cols = 6, 4
    left, top, width, height = Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8)
    table_shape = s_matrix.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # Column widths
    table.columns[0].width = Inches(2.3)
    table.columns[1].width = Inches(2.7)
    table.columns[2].width = Inches(3.8)
    table.columns[3].width = Inches(2.9)

    headers = ["Tool Name", "Category", "Core Value Proposition", "Target User"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = WHITE

    matrix_data = [
        ("Perplexity AI", "Education & Research", "Source-verified real-time answer synthesis", "Students, Researchers & Analysts"),
        ("Notion AI", "Productivity & Ops", "Native workspace summarization & custom Q&A", "Project Managers, Teams, Interns"),
        ("Midjourney", "Image Generation", "Unmatched photorealistic visual asset generation", "Creative Designers, Marketers, Artists"),
        ("Synthesia", "Video Creation", "Turnkey script-to-video with multilingual avatars", "Corporate Trainers, HR, Educators"),
        ("Gamma AI", "Presentations", "Outline-to-responsive slide deck generation", "Speakers, Founders, Presenters")
    ]

    for row_idx, row_data in enumerate(matrix_data, start=1):
        bg_col = CARD_BG if row_idx % 2 == 1 else RGBColor(241, 245, 249)
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = text
            r.font.size = Pt(11)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = NAVY
            else:
                r.font.color.rgb = TEXT_DARK

    # ==========================
    # SLIDE 9: Strategic Takeaways & Learnings
    # ==========================
    s_takeaways = prs.slides.add_slide(blank_layout)
    add_header(s_takeaways, "Key Insights & Practical Internship Takeaways")

    takeaway_cards = [
        ("1. Specialized AI Over Generic Chatbots", "While general-purpose LLMs are versatile, specialized domain tools (e.g. Gamma for slides, Perplexity for citations) offer superior workflow acceleration and lower friction."),
        ("2. Multi-Modal Workflow Orchestration", "The modern competitive advantage lies in chaining AI tools: using Perplexity for research, Notion AI for structuring, Midjourney for artwork, and Synthesia for presentations."),
        ("3. The Human-in-the-Loop Imperative", "AI provides powerful 10x drafting leverage, but critical thinking, verification, empathetic judgment, and ethical oversight remain irreplaceable human responsibilities.")
    ]

    for idx, (t_title, t_body) in enumerate(takeaway_cards):
        card = s_takeaways.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8 + idx * 1.65), Inches(11.7), Inches(1.4))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_left = Inches(0.35)
        ctf.margin_top = Inches(0.2)

        p1 = ctf.paragraphs[0]
        p1.space_after = Pt(4)
        r1 = p1.add_run()
        r1.text = t_title
        r1.font.size = Pt(14)
        r1.font.bold = True
        r1.font.color.rgb = NAVY

        p2 = ctf.add_paragraph()
        r2 = p2.add_run()
        r2.text = t_body
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = TEXT_DARK

    # ==========================
    # SLIDE 10: Conclusion & Thank You (Dark Theme)
    # ==========================
    s_end = prs.slides.add_slide(blank_layout)
    add_background(s_end, DARK_BG)

    accent_end = s_end.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18))
    accent_end.fill.solid()
    accent_end.fill.fore_color.rgb = TEAL
    accent_end.line.fill.background()

    end_box = s_end.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10.33), Inches(3.5))
    etf = end_box.text_frame
    etf.word_wrap = True

    ep0 = etf.paragraphs[0]
    ep0.space_after = Pt(8)
    er0 = ep0.add_run()
    er0.text = "TASK 1 COMPLETED"
    er0.font.size = Pt(14)
    er0.font.bold = True
    er0.font.color.rgb = TEAL

    ep1 = etf.add_paragraph()
    ep1.space_after = Pt(14)
    er1 = ep1.add_run()
    er1.text = "Thank You!"
    er1.font.size = Pt(44)
    er1.font.bold = True
    er1.font.color.rgb = WHITE

    ep2 = etf.add_paragraph()
    ep2.space_after = Pt(20)
    er2 = ep2.add_run()
    er2.text = "Special thanks to InAmigos Foundation for fostering an environment of active experimentation and technological curiosity."
    er2.font.size = Pt(15)
    er2.font.color.rgb = RGBColor(203, 213, 225)

    ep3 = etf.add_paragraph()
    er3 = ep3.add_run()
    er3.text = "Candidate: Prajin K   •   Internship Domain: AI & Emerging Tech"
    er3.font.size = Pt(13)
    er3.font.color.rgb = RGBColor(148, 163, 184)

    prs.save("Task1_AI_Tools_Discovery.pptx")
    print("Task1_AI_Tools_Discovery.pptx generated successfully!")

if __name__ == "__main__":
    create_presentation()
