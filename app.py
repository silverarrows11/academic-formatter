import io
import time
import re
import os
import streamlit as st
import openai
import stripe
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. Page Configuration
st.set_page_config(
    page_title="FormatForge | Academic, AI Cleaner & ATS Resume Suite",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Unified High-Trust SaaS Styling
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
}
h1, h2, h3, h4 {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    letter-spacing: -0.03em !important;
}
#MainMenu, header, footer { visibility: hidden; }

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 940px;
}

.hero-card {
    background: radial-gradient(circle at top right, #1E293B, #090D16 70%);
    padding: 46px 32px;
    border-radius: 18px;
    color: #FFFFFF !important;
    margin-bottom: 24px;
    text-align: center;
    border: 1px solid #334155;
    box-shadow: 0 16px 36px -8px rgba(9, 13, 22, 0.45);
}
.hero-pill {
    display: inline-block;
    background: rgba(16, 185, 129, 0.15);
    color: #34D399;
    border: 1px solid rgba(52, 211, 153, 0.3);
    padding: 5px 14px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 16px;
}
.hero-card h1 {
    color: #FFFFFF !important;
    font-size: 38px;
    font-weight: 800;
    margin: 0 0 14px 0;
    line-height: 1.12;
}
.hero-card p {
    color: #94A3B8 !important;
    font-size: 15.5px;
    margin: 0 auto;
    max-width: 620px;
    line-height: 1.6;
}

.badge {
    display: inline-block;
    padding: 6px 14px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.badge-pro {
    background: rgba(16, 185, 129, 0.2);
    color: #34D399;
    border: 1px solid #10B981;
}
.badge-free {
    background: rgba(148, 163, 184, 0.15);
    color: #CBD5E1;
    border: 1px solid #475569;
}

.stButton > button[kind="primary"] {
    background: linear-gradient(180deg, #10B981 0%, #059669 100%) !important;
    color: #052E16 !important;
    border: 1px solid #34D399 !important;
    border-radius: 12px !important;
    padding: 14px 20px !important;
    font-weight: 800 !important;
    font-size: 14.5px !important;
    transition: all 0.15s ease !important;
    box-shadow: 0 4px 16px rgba(16, 185, 129, 0.25) !important;
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 22px rgba(16, 185, 129, 0.4) !important;
}
.stButton > button[kind="secondary"] {
    border-radius: 12px !important;
    padding: 13px 18px !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    border: 1px solid rgba(148, 163, 184, 0.3) !important;
    background: rgba(30, 41, 59, 0.5) !important;
    color: #E2E8F0 !important;
    transition: all 0.15s ease !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #10B981 !important;
    color: #34D399 !important;
    background: rgba(16, 185, 129, 0.08) !important;
}

.service-spotlight {
    background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
    border: 1px solid #334155;
    border-left: 4px solid #10B981;
    border-radius: 14px;
    padding: 20px 24px;
    margin-top: 8px;
    margin-bottom: 20px;
    color: #F8FAFC;
    box-shadow: 0 10px 28px -8px rgba(0, 0, 0, 0.35);
}
.spotlight-top {
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    margin-bottom: 8px;
}
.spotlight-title {
    font-size: 18px;
    font-weight: 800;
    color: #FFFFFF;
}
.spotlight-desc {
    font-size: 13.5px;
    color: #94A3B8;
    line-height: 1.5;
    margin-bottom: 14px;
}
.benefit-pills {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
}
.benefit-pill {
    background: rgba(16, 185, 129, 0.12);
    border: 1px solid rgba(52, 211, 153, 0.28);
    color: #34D399;
    padding: 5px 12px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 700;
}

/* Authentic White Word Page Preview Sheet */
.word-preview-sheet {
    background: #FFFFFF;
    color: #111827;
    font-family: 'Times New Roman', Times, serif;
    padding: 48px 56px;
    border-radius: 6px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    margin: 16px 0;
    max-height: 420px;
    overflow-y: auto;
    font-size: 14.5px;
    line-height: 2.0;
    border: 1px solid #CBD5E1;
}

.feature-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
    gap: 16px;
    margin-top: 28px;
    margin-bottom: 36px;
}
.feature-card {
    background: rgba(15, 23, 42, 0.6);
    border: 1px solid #334155;
    padding: 22px;
    border-radius: 14px;
}
.pricing-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 20px;
    margin-top: 16px;
    margin-bottom: 36px;
}
.price-card-free {
    background: rgba(15, 23, 42, 0.5);
    border: 1px solid #334155;
    padding: 28px;
    border-radius: 16px;
}
.price-card-pro {
    background: linear-gradient(145deg, #0F172A 0%, #1E293B 100%);
    color: #FFFFFF;
    border: 2px solid #10B981;
    padding: 28px;
    border-radius: 16px;
    box-shadow: 0 12px 30px -8px rgba(16, 185, 129, 0.25);
}

.blur-container {
    position: relative;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 22px;
    background: #0F172A;
    color: #F8FAFC;
    overflow: hidden;
    margin-top: 12px;
}
.blurred-body {
    filter: blur(5px);
    user-select: none;
    pointer-events: none;
    opacity: 0.45;
    font-size: 14px;
    line-height: 1.6;
    margin-top: 8px;
    color: #94A3B8;
}
.blur-cta-overlay {
    background: rgba(15, 23, 42, 0.92);
    backdrop-filter: blur(6px);
    border: 1px solid #10B981;
    border-radius: 12px;
    padding: 22px;
    text-align: center;
    margin-top: -55px;
    position: relative;
    z-index: 10;
    box-shadow: 0 12px 32px rgba(0, 0, 0, 0.45);
}
.cta-btn {
    background: linear-gradient(180deg, #10B981 0%, #059669 100%);
    color: #052E16 !important;
    font-weight: 800;
    padding: 12px 24px;
    border-radius: 10px;
    text-decoration: none !important;
    font-size: 14px;
    display: inline-block;
    transition: transform 0.15s ease;
}
.cta-btn:hover {
    transform: scale(1.02);
    background: #34D399;
}
.trust-strip {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 24px;
    font-size: 12px;
    color: #94A3B8;
    font-weight: 600;
    margin-top: 16px;
    padding-top: 12px;
    border-top: 1px solid rgba(148, 163, 184, 0.15);
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# 3. Configuration Limits & Credentials
FREE_CHAR_LIMIT = 1200
PRO_MAX_CHARS = 35000
STRIPE_PAYMENT_URL = "https://buy.stripe.com/cNidRbdxg5l4c9902l5Ne01"
DELIMITER = "---AUDIT_AND_FEEDBACK---"

stripe_key = os.getenv("STRIPE_SECRET_KEY")
if stripe_key:
    stripe.api_key = stripe_key

TOOL_CONFIG = {
    "academic": {
        "tab_label": "📚 Academic Paper Forge",
        "short_name": "Academic Paper Forge",
        "headline": "📚 Academic Citation & Manuscript Engine",
        "tagline": "Transform rough drafts and messy bibliographies into strict university-rubric Word documents.",
        "benefits": [
            "✓ Auto Hanging Indents (.docx)",
            "✓ Missing In-Text Citation Audit",
            "✓ APA 7, MLA 9, Harvard, Chicago & IEEE",
            "✓ Optional Auto Cover Page"
        ],
        "button_text": "✨ Format Paper & Run Citation Audit",
        "placeholder": "Paste your unformatted essay, rough draft, or messy bibliography right here...",
        "sample": (
            "The Impact of Screen Time on Teen Sleep Patterns\n\n"
            "Recent empirical studies demonstrate that screen-based media consumption before bed disrupts circadian biology. "
            "According to smith (2021), blue light emitted by smart devices suppresses nocturnal melatonin release, making it "
            "difficult for adolescents to initiate restful sleep cycles. Furthermore, incoming social alerts frequently induce "
            "sleep fragmentation (Johnson and Lee 2019). Moreover, Davis (2022) argues that late-night scrolling correlates with "
            "daytime fatigue. Consequently, public health guidelines should advise device cutoffs 60 minutes prior to bedtime.\n\n"
            "References:\n"
            "Smith, John. (2021). Blue light and circadian rhythms. Journal of Sleep Health, 15(2), 104-112.\n"
            "johnson, m., & Lee, T. 2019. Social media addiction in secondary school students. Adolescent Psychology Review 8(4): 45-59."
        )
    },
    "humanizer": {
        "tab_label": "🛡️ AI-Risk Auditor & Cleaner",
        "short_name": "AI-Risk Auditor & Cleaner",
        "headline": "🛡️ Turnitin AI-Trigger Auditor & Word Cleaner",
        "tagline": "Eliminate robotic LLM clichés, strip copy-paste ChatGPT markdown artifacts, and flag uncited claims.",
        "benefits": [
            "✓ Strips ### & ** Markdown Clutter",
            "✓ Replaces Robotic AI Trigger Words",
            "✓ Uncited Fact & Plagiarism Risk Scan",
            "✓ Exports Clean, Natural .docx"
        ],
        "button_text": "🛡️ Scan AI Triggers & Clean Document",
        "placeholder": "Paste your ChatGPT/Claude draft here to strip Markdown symbols (###, **), fix robotic AI clichés, and flag uncited claims...",
        "sample": (
            "### The Crucial Role of Renewable Energy in Modern Society\n\n"
            "In today's rapidly evolving landscape, it is crucial to delve into the multifaceted tapestry of renewable energy. "
            "**Solar and wind power** serve as a testament to human ingenuity. Furthermore, studies show that global emissions "
            "dropped by 14% in sectors that adopted micro-grids, paving the way for a sustainable future. Moreover, it is important "
            "to note that battery storage technology continues to revolutionize grid reliability.\n\n"
            "* **Key Takeaway 1:** Cost reduction in photovoltaic cells.\n"
            "* **Key Takeaway 2:** Enhanced grid resilience across urban centers."
        )
    },
    "resume": {
        "tab_label": "💼 ATS Resume Forge",
        "short_name": "ATS Resume Forge",
        "headline": "💼 Wall Street & Harvard 1-Page ATS Resume Forge",
        "tagline": "Convert rough work experience or LinkedIn bios into a high-impact, single-page .docx resume built to pass ATS filters.",
        "benefits": [
            "✓ Wall Street XYZ Impact Bullet Rewrites",
            "✓ Strict 1-Page Margin & Font Lock",
            "✓ Target Job Keyword Optimization",
            "✓ 100% ATS-Parseable .docx Export"
        ],
        "button_text": "💼 Forge 1-Page Wall Street / Harvard Resume",
        "placeholder": "Paste your messy resume, LinkedIn bio, or rough work/project bullet points here...",
        "sample": (
            "Alex Rivera | alex.rivera@email.edu | (555) 234-5678 | New York, NY\n\n"
            "Education:\n"
            "State University, B.S. in Finance & Data Science, Expected May 2027. GPA 3.7. Relevant coursework: Corporate Finance, Python.\n\n"
            "Experience:\n"
            "Summer Analyst at Apex Capital (June 2025 - Aug 2025)\n"
            "- Helped the team with financial models in Excel for retail companies.\n"
            "- Looked at 15 potential acquisition targets and made slide decks for senior partners.\n"
            "- Automated weekly reporting scripts using Python which saved a lot of time.\n\n"
            "Project Lead, University Investment Club (Sept 2024 - Present)\n"
            "- Managed part of the student endowment fund and pitched 3 tech stocks.\n\n"
            "Skills: Python, SQL, Excel (DCF, LBO), Bloomberg Terminal, Spanish."
        )
    }
}

# 4. Session State Initialization
defaults = {
    "unlocked_tools": [],
    "free_uses": {"academic": 0, "humanizer": 0, "resume": 0},
    "active_tool": "academic",
    "input_text": "",
    "job_target": "",
    "last_uploaded_file": None,
    "clean_text": None,
    "editorial_notes": None,
    "selected_style": "APA 7th Edition",
    "current_view": "landing"
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# 5. Crash-Proof URL Pass Persistence (Idea #1)
# Check URL query params for session verification or persisted unlock state
if "session_id" in st.query_params:
    try:
        checkout_session = stripe.checkout.Session.retrieve(st.query_params["session_id"])
        if checkout_session.payment_status == "paid":
            paid_tool = getattr(checkout_session, "client_reference_id", None) or st.session_state.active_tool
            if paid_tool not in TOOL_CONFIG:
                paid_tool = st.session_state.active_tool
            if paid_tool not in st.session_state.unlocked_tools:
                st.session_state.unlocked_tools.append(paid_tool)
            st.session_state.active_tool = paid_tool
            st.session_state.current_view = "editor"
            # Keep a clean URL parameter so refreshes preserve access without re-querying Stripe
            st.query_params["unlocked"] = paid_tool
            st.toast(f"⚡ {TOOL_CONFIG[paid_tool]['short_name']} Unlocked!", icon="✅")
    except Exception as e:
        st.error(f"Stripe Verification Error: {e}")
        st.query_params.clear()
elif "unlocked" in st.query_params:
    persisted_tool = st.query_params["unlocked"]
    if persisted_tool in TOOL_CONFIG and persisted_tool not in st.session_state.unlocked_tools:
        st.session_state.unlocked_tools.append(persisted_tool)
        st.session_state.active_tool = persisted_tool

is_pro = st.session_state.active_tool in st.session_state.unlocked_tools

def get_stripe_url(tool_key):
    sep = "&" if "?" in STRIPE_PAYMENT_URL else "?"
    return f"{STRIPE_PAYMENT_URL}{sep}client_reference_id={tool_key}"

def switch_tool(new_tool):
    if st.session_state.active_tool != new_tool:
        st.session_state.active_tool = new_tool
        st.session_state.clean_text = None
        st.session_state.editorial_notes = None
        st.session_state.input_text = ""

# 6. Helper Functions
def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    return uploaded_file.read().decode("utf-8", errors="ignore")

def add_formatted_runs(paragraph, text, font_name="Times New Roman", font_size=12):
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*|_.*?_)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith(('*', '_')) and part.endswith(('*', '_')) and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            run = paragraph.add_run(part)
        run.font.name = font_name
        run.font.size = Pt(font_size)

def create_smart_docx(content, tool_mode, style_name="APA 7th Edition", title_info=None):
    doc = Document()
    margin_size = Inches(0.6) if tool_mode == "resume" else Inches(1.0)
    for section in doc.sections:
        section.top_margin = margin_size
        section.bottom_margin = margin_size
        section.left_margin = margin_size
        section.right_margin = margin_size

    if tool_mode == "academic" and title_info and title_info.get("include") and title_info.get("title"):
        for _ in range(3):
            doc.add_paragraph()
        
        t_para = doc.add_paragraph()
        t_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t_para.paragraph_format.line_spacing = 2.0
        t_run = t_para.add_run(title_info["title"])
        t_run.bold = True
        t_run.font.name = "Times New Roman"
        t_run.font.size = Pt(12)

        doc.add_paragraph()
        for field in ["author", "institution", "course", "instructor", "date"]:
            val = title_info.get(field, "").strip()
            if val:
                m_para = doc.add_paragraph()
                m_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                m_para.paragraph_format.line_spacing = 2.0
                m_run = m_para.add_run(val)
                m_run.font.name = "Times New Roman"
                m_run.font.size = Pt(12)
        
        doc.add_page_break()

    paragraphs = [p.strip() for p in content.split('\n') if p.strip()]

    if tool_mode == "resume":
        section_headers = ["education", "experience", "professional experience", "work experience", "projects", "leadership", "skills", "technical skills", "summary", "certifications"]
        for idx, para_text in enumerate(paragraphs):
            clean_line = re.sub(r'[*_#]', '', para_text).strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            
            if idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(clean_line)
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(15)
            elif idx == 1 and ("|" in clean_line or "@" in clean_line):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(8)
                run = p.add_run(clean_line)
                run.font.name = "Times New Roman"
                run.font.size = Pt(10.5)
            elif clean_line.rstrip(':').lower() in section_headers or (clean_line.isupper() and len(clean_line) < 35):
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(clean_line.rstrip(':').upper())
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(11)
            elif para_text.startswith(('-', '•', '* ')):
                p.paragraph_format.left_indent = Inches(0.25)
                p.paragraph_format.first_line_indent = Inches(-0.15)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                bullet_text = re.sub(r'^[-•*]\s*', '• ', para_text)
                add_formatted_runs(p, bullet_text, font_size=10.5)
            else:
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(1)
                add_formatted_runs(p, para_text, font_size=11)
    else:
        in_references = False
        ref_headers = ["references", "reference list", "works cited", "bibliography"]

        for idx, para_text in enumerate(paragraphs):
            clean_header = re.sub(r'[*_#:]', '', para_text).strip().lower()
            
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 2.0
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)

            if clean_header in ref_headers:
                in_references = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(re.sub(r'[*_#]', '', para_text).strip())
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                continue

            if idx == 0 and not in_references and len(para_text) < 120 and not para_text.endswith('.'):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(re.sub(r'[*_#]', '', para_text).strip())
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                continue

            if in_references:
                p.paragraph_format.left_indent = Inches(0.5)
                p.paragraph_format.first_line_indent = Inches(-0.5)
            else:
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.first_line_indent = Inches(0.5)

            add_formatted_runs(p, re.sub(r'^#+\s*', '', para_text))

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# 7. Legal Modals
@st.dialog("Privacy Policy", width="large")
def privacy_modal():
    st.markdown(
        "**Effective Date: November 2026**\n\n"
        "**1. Zero-Retention Architecture:** FormatForge processes your text strictly in volatile memory. "
        "Your drafts and resumes are permanently destroyed the moment your formatting finishes or you close your browser tab.\n\n"
        "**2. Payment Security:** All payments are processed directly by Stripe.\n\n"
        "**3. No Tracking:** We only use temporary session state to keep your unlocked tool active."
    )

@st.dialog("Terms of Service", width="large")
def terms_modal():
    st.markdown(
        "**Effective Date: November 2026**\n\n"
        "**1. Scope of Utility:** FormatForge is an automated formatting and document export utility.\n\n"
        "**2. Final Review Responsibility:** Always verify your exported `.docx` document before submitting.\n\n"
        "**3. Per-Tool Session Pass Policy:** Each $1.99 pass grants unlimited re-runs and `.docx` exports for the selected tool during your session."
    )


# ==============================================================================
# VIEW 1: MARKETING LANDING PAGE + SEO ACCORDION (Idea #3)
# ==============================================================================
if st.session_state.current_view == "landing":
    unlocked_count = len(st.session_state.unlocked_tools)
    status_badge = f'<span class="badge badge-pro">⚡ {unlocked_count} Tool(s) Active</span>' if unlocked_count > 0 else '<span class="badge badge-free">Free Preview Available</span>'
    
    nav_col1, nav_col2 = st.columns([3, 1.2], vertical_alignment="center")
    with nav_col1:
        st.markdown(f'<div style="display:flex; align-items:center; gap:12px;"><span style="font-weight:800; font-size:22px; letter-spacing:-0.03em;">✦ FormatForge</span> {status_badge}</div>', unsafe_allow_html=True)
    with nav_col2:
        if st.button("Launch Studio →", use_container_width=True):
            st.session_state.current_view = "editor"
            st.rerun()

    st.markdown("<div style='margin-bottom: 14px;'></div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="hero-card">'
        '<div class="hero-pill">⚡ 3 Specialized Document Engines • $1.99 Per Tool Pass</div>'
        '<h1>Messy draft in.<br>Submission-ready Word doc out.</h1>'
        '<p>Whether you have an 11:59 PM paper with broken citations, a ChatGPT draft full of robotic phrasing, or a messy internship resume—forge a rubric-perfect .docx in 10 seconds.</p>'
        '</div>',
        unsafe_allow_html=True
    )

    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.button("📚 Academic Paper Forge", type="primary", use_container_width=True):
            st.session_state.active_tool = "academic"
            st.session_state.current_view = "editor"
            st.rerun()
    with l_col2:
        if st.button("🛡️ AI-Risk Auditor & Cleaner", type="primary", use_container_width=True):
            st.session_state.active_tool = "humanizer"
            st.session_state.current_view = "editor"
            st.rerun()
    with l_col3:
        if st.button("💼 ATS Resume Forge", type="primary", use_container_width=True):
            st.session_state.active_tool = "resume"
            st.session_state.current_view = "editor"
            st.rerun()

    st.markdown(
        '<div class="feature-grid">'
        '<div class="feature-card">'
        '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">'
        '<span style="font-size:24px;">📚</span><span class="badge badge-pro">$1.99 Pass</span></div>'
        '<div style="font-weight:800; font-size:16px; margin-bottom:6px;">Academic Citation & .docx Forge</div>'
        '<div style="font-size:13px; color:#94A3B8; line-height:1.55;">Fixes APA 7th, MLA 9th, Harvard, Chicago & IEEE citations. Cross-checks missing references and exports a .docx with hanging indents & cover page.</div>'
        '</div>'
        '<div class="feature-card">'
        '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">'
        '<span style="font-size:24px;">🛡️</span><span class="badge badge-pro">$1.99 Pass</span></div>'
        '<div style="font-weight:800; font-size:16px; margin-bottom:6px;">AI-Risk Auditor & Word Cleaner</div>'
        '<div style="font-size:13px; color:#94A3B8; line-height:1.55;">Flags robotic AI trigger phrases ("delve", "tapestry") and uncited claims. Strips ChatGPT markdown symbols and exports a natural, clean .docx.</div>'
        '</div>'
        '<div class="feature-card">'
        '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">'
        '<span style="font-size:24px;">💼</span><span class="badge badge-pro">$1.99 Pass</span></div>'
        '<div style="font-weight:800; font-size:16px; margin-bottom:6px;">1-Page ATS Resume Forge</div>'
        '<div style="font-size:13px; color:#94A3B8; line-height:1.55;">Rewrites weak bullets using the Wall Street XYZ impact formula, optimizes keywords for your target job, and exports a 1-page Harvard .docx resume.</div>'
        '</div>'
        '</div>',
        unsafe_allow_html=True
    )

    # SEO FAQ Accordion (Idea #3)
    with st.expander("📖 Frequently Asked Questions & Supported Rubrics (SEO Hub)"):
        st.markdown("""
        - **How does FormatForge fix APA 7th hanging indents?** FormatForge automatically detects your reference list and applies strict 0.5-inch hanging indents to every reference item in the exported Microsoft Word (`.docx`) file.
        - **Is FormatForge Turnitin-Safe?** Yes. FormatForge uses a volatile zero-retention memory architecture. Your text is processed instantly in memory and never stored in any database or indexed by AI training models.
        - **What styling manuals are supported?** We support APA 7th Edition, MLA 9th Edition, Harvard Reference Style, Chicago 17th Edition, and IEEE standards.
        - **How does the AI-Risk Auditor work?** It scans drafts for robotic LLM patterns, uniform sentence structures, and overused AI transition words that trigger AI detectors, while also scanning for un-cited statistical claims.
        """)

# ==============================================================================
# VIEW 2: FORMATTING STUDIO + VISUAL WORD SHEET PREVIEW (Idea #3)
# ==============================================================================
else:
    active_tool = st.session_state.active_tool
    is_tool_pro = active_tool in st.session_state.unlocked_tools
    tool_info = TOOL_CONFIG[active_tool]
    stripe_tool_url = get_stripe_url(active_tool)

    status_badge = (
        f'<span class="badge badge-pro">⚡ {tool_info["short_name"]}: UNLOCKED</span>'
        if is_tool_pro else
        f'<span class="badge badge-free">🔒 {tool_info["short_name"]}: FREE PREVIEW</span>'
    )

    nav_c1, nav_c2, nav_c3 = st.columns([1.2, 2.2, 1.6], vertical_alignment="center")
    with nav_c1:
        if st.button("← Home", use_container_width=True):
            st.session_state.current_view = "landing"
            st.rerun()
    with nav_c2:
        st.markdown('<div style="text-align:center; font-weight:800; font-size:20px; letter-spacing:-0.03em;">✦ FormatForge Studio</div>', unsafe_allow_html=True)
    with nav_c3:
        st.markdown(f'<div style="text-align:right;">{status_badge}</div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-top: 10px; margin-bottom: 6px; font-size: 11px; font-weight: 800; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.06em;'>Step 1 — Select Your Formatting Engine:</div>", unsafe_allow_html=True)

    tab_c1, tab_c2, tab_c3 = st.columns(3)
    with tab_c1:
        label_1 = f"📚 Academic Paper {'(⚡ Active)' if 'academic' in st.session_state.unlocked_tools else '($1.99)'}"
        if st.button(label_1, type="primary" if active_tool == "academic" else "secondary", use_container_width=True):
            switch_tool("academic")
            st.rerun()
    with tab_c2:
        label_2 = f"🛡️ AI Auditor & Cleaner {'(⚡ Active)' if 'humanizer' in st.session_state.unlocked_tools else '($1.99)'}"
        if st.button(label_2, type="primary" if active_tool == "humanizer" else "secondary", use_container_width=True):
            switch_tool("humanizer")
            st.rerun()
    with tab_c3:
        label_3 = f"💼 ATS Resume Forge {'(⚡ Active)' if 'resume' in st.session_state.unlocked_tools else '($1.99)'}"
        if st.button(label_3, type="primary" if active_tool == "resume" else "secondary", use_container_width=True):
            switch_tool("resume")
            st.rerun()

    pills_html = "".join([f'<span class="benefit-pill">{b}</span>' for b in tool_info["benefits"]])
    unlock_cta_html = (
        '<span class="badge badge-pro">✓ Unlimited Session Edits Unlocked</span>'
        if is_tool_pro else
        f'<a href="{stripe_tool_url}" target="_blank" class="cta-btn" style="padding:7px 14px; font-size:12px;">⚡ Unlock Full {tool_info["short_name"]} ($1.99) →</a>'
    )

    st.markdown(
        '<div class="service-spotlight">'
        '<div class="spotlight-top">'
        f'<div class="spotlight-title">{tool_info["headline"]}</div>'
        f'<div>{unlock_cta_html}</div>'
        '</div>'
        f'<div class="spotlight-desc">{tool_info["tagline"]}</div>'
        f'<div class="benefit-pills">{pills_html}</div>'
        '</div>',
        unsafe_allow_html=True
    )

    with st.container(border=True):
        ctrl_col1, ctrl_col2 = st.columns([2.8, 1.2], vertical_alignment="bottom")
        
        with ctrl_col1:
            if active_tool == "academic":
                style_options = ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago 17th Edition", "IEEE"]
                if is_tool_pro:
                    selected_style = st.selectbox("Formatting Rubric", style_options, index=0)
                else:
                    selected_style = st.selectbox(
                        "Formatting Rubric", 
                        ["APA 7th Edition (Free Preview)", "🔒 MLA 9th Edition ($1.99 Pass)", "🔒 Harvard Style ($1.99 Pass)", "🔒 Chicago 17th Edition ($1.99 Pass)", "🔒 IEEE ($1.99 Pass)"],
                        index=0
                    )
                    if "🔒" in selected_style:
                        st.warning("Unlock the $1.99 Academic Pass to format with MLA, Harvard, Chicago, or IEEE.")
                    selected_style = "APA 7th Edition"
                st.session_state.selected_style = selected_style
            elif active_tool == "humanizer":
                tone_mode = st.selectbox("Cleanup & Humanizer Standard", [
                    "Natural Academic Prose (Strip AI Clichés & Markdown)",
                    "Concise Executive Polish (High Clarity & Flow)"
                ])
                st.session_state.selected_style = tone_mode
            else:
                resume_style = st.selectbox("ATS Template Standard", [
                    "Wall Street / Harvard 1-Page Standard",
                    "Tech & Engineering ATS Standard"
                ])
                st.session_state.selected_style = resume_style

        with ctrl_col2:
            if st.button("🧪 Load Sample Text", use_container_width=True):
                st.session_state.input_text = tool_info["sample"]
                st.rerun()

        user_text = st.text_area(
            "Document Input",
            value=st.session_state.input_text,
            height=240,
            placeholder=tool_info["placeholder"],
            label_visibility="collapsed"
        )
        st.session_state.input_text = user_text

        if active_tool == "resume":
            st.session_state.job_target = st.text_input(
                "🎯 Target Job Title or Paste Job Description Keywords (Optional)",
                value=st.session_state.job_target,
                placeholder="e.g., Investment Banking Summer Analyst / Data Science Intern (Python, SQL, Financial Modeling)"
            )

        opt_col1, opt_col2 = st.columns(2)
        tp_title = tp_author = tp_inst = tp_course = tp_instructor = tp_date = ""
        with opt_col1:
            with st.expander("📎 Or Upload a .docx / .txt File"):
                uploaded_doc = st.file_uploader("Upload file", type=["docx", "txt"], label_visibility="collapsed")
                if uploaded_doc is not None and uploaded_doc.name != st.session_state.last_uploaded_file:
                    st.session_state.input_text = extract_text_from_file(uploaded_doc)
                    st.session_state.last_uploaded_file = uploaded_doc.name
                    st.rerun()
        with opt_col2:
            if active_tool == "academic":
                with st.expander("📄 Add Auto-Formatted Cover Page (Optional)"):
                    tp_title = st.text_input("Paper Title", placeholder="e.g., The Impact of Screen Time on Sleep")
                    tp_author = st.text_input("Your Full Name", placeholder="e.g., Alex Rivera")
                    tp_inst = st.text_input("University / College", placeholder="e.g., State University")
                    tp_course = st.text_input("Course Name", placeholder="e.g., PSY 101")
                    tp_instructor = st.text_input("Instructor Name", placeholder="e.g., Dr. Sarah Jenkins")
                    tp_date = st.text_input("Due Date", placeholder="e.g., November 6, 2026")

        char_len = len(user_text)
        word_count = len(user_text.split()) if user_text.strip() else 0
        tool_free_used = st.session_state.free_uses.get(active_tool, 0)
        
        meta_col1, meta_col2 = st.columns(2)
        with meta_col1:
            if not is_tool_pro:
                st.caption(f"📏 **{word_count} words** ({char_len}/{FREE_CHAR_LIMIT} free preview chars) • Previews remaining: **{max(0, 1 - tool_free_used)}**")
            else:
                st.caption(f"📏 **{word_count} words** ({char_len:,}/{PRO_MAX_CHARS:,} chars) • **⚡ Unlimited Session Edits Active**")
        with meta_col2:
            st.markdown('<div style="text-align:right; font-size:12px; color:#34D399; font-weight:700;">🛡️ Zero-Retention Memory • 100% Turnitin-Safe</div>', unsafe_allow_html=True)

        if st.button(tool_info["button_text"], type="primary", use_container_width=True):
            if not user_text.strip():
                st.warning("Paste your text above or click '🧪 Load Sample Text' to test the engine!")
            elif not is_tool_pro and tool_free_used >= 1:
                st.error(f"You've used your free preview for {tool_info['short_name']}! Unlock unlimited runs & Word (.docx) export for $1.99 below.")
            elif not is_tool_pro and char_len > FREE_CHAR_LIMIT:
                st.error(f"Your text is {char_len:,} characters (Free preview limit is {FREE_CHAR_LIMIT:,}). Unlock the $1.99 {tool_info['short_name']} Pass below to process up to 35,000 characters with unlimited edits!")
            else:
                client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
                with st.spinner(f"Running {tool_info['short_name']} engine..."):
                    if active_tool == "academic":
                        system_message = (
                            f"You are an elite university academic copyeditor. Reformat the user's input strictly according to {st.session_state.selected_style} standards.\n"
                            "- Fix all in-text citations and alphabetize the reference list.\n"
                            "- Wrap journal or book titles in single asterisks (*Title*) for proper Word italicization.\n"
                            "- Keep body paragraphs intact while fixing citation syntax.\n\n"
                            "Output Format:\n"
                            "1. Output ONLY the clean, formatted academic paper first. No markdown code fences.\n"
                            f"2. On a new line by itself, output this exact separator:\n{DELIMITER}\n"
                            "3. Below the separator, provide a bulleted 'Citation & Rubric Audit':\n"
                            "- Flag any in-text citation missing from the Reference list (or vice versa).\n"
                            "- List specific formatting fixes applied and 2 quick academic tone improvements."
                        )
                    elif active_tool == "humanizer":
                        system_message = (
                            "You are an academic AI-Risk Auditor and Document Cleaner. Transform raw AI/ChatGPT drafts into natural, authentic human academic prose.\n"
                            "- Strip all Markdown artifacts (### headers, **bold** clutter, bullet dumps) into clean, cohesive paragraphs.\n"
                            "- Replace overused LLM trigger words ('delve', 'tapestry', 'testament', 'crucial', 'furthermore', 'multifaceted') with natural sentence variation.\n\n"
                            "Output Format:\n"
                            "1. Output ONLY the cleaned, natural human-sounding document first. No markdown code fences.\n"
                            f"2. On a new line by itself, output this exact separator:\n{DELIMITER}\n"
                            "3. Below the separator, provide a bulleted 'AI-Risk & Originality Audit':\n"
                            "- Estimate the initial AI-Pattern Risk (High/Medium/Low) and list the exact robotic phrases you removed.\n"
                            "- Flag any specific factual claims, statistics, or numbers in the text that lack an in-text citation (plagiarism risk)."
                        )
                    else:
                        target_ctx = f"Target Role/Keywords: {st.session_state.job_target}" if st.session_state.job_target else "General competitive internship/corporate role"
                        system_message = (
                            f"You are an elite Wall Street & Tech ATS Resume writer. Reformat the user's input into a strict 1-page Harvard/Wall Street ATS Resume ({target_ctx}).\n"
                            "- First line: Candidate Name. Second line: Contact info separated by ' | '.\n"
                            "- Use clean uppercase section headers: EDUCATION, EXPERIENCE, PROJECTS, SKILLS.\n"
                            "- Rewrite weak experience bullets using the XYZ impact formula (Accomplished [X] as measured by [Y], by doing [Z]) with strong action verbs.\n\n"
                            "Output Format:\n"
                            "1. Output ONLY the formatted resume text first. Start bullet points with '• '. No markdown code fences.\n"
                            f"2. On a new line by itself, output this exact separator:\n{DELIMITER}\n"
                            "3. Below the separator, provide a bulleted 'ATS & Impact Audit':\n"
                            "- Highlight weak bullets you upgraded into quantified impact statements.\n"
                            "- List 3 high-value ATS keywords recommended for their target role."
                        )

                    try:
                        response = client.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[
                                {"role": "system", "content": system_message},
                                {"role": "user", "content": user_text}
                            ],
                            timeout=60.0
                        )
                        output_content = response.choices[0].message.content
                        
                        if DELIMITER in output_content:
                            clean, notes = output_content.split(DELIMITER, 1)
                            st.session_state.clean_text = clean.strip()
                            st.session_state.editorial_notes = notes.strip()
                        else:
                            st.session_state.clean_text = output_content.strip()
                            st.session_state.editorial_notes = "• Document standardized and formatted.\n• Structural audit completed."
                            
                        if not is_tool_pro:
                            st.session_state.free_uses[active_tool] = tool_free_used + 1
                        
                        st.toast("Document processed & audited!", icon="📝")
                        
                    except Exception as e:
                        st.error(f"OpenAI Error: {e}")

        st.markdown(
            '<div class="trust-strip">'
            '<span>🔒 256-Bit Stripe Checkout</span>'
            '<span>♻️ Unlimited Re-Edits Included ($1.99)</span>'
            '<span>📄 Native Microsoft Word (.docx) Engine</span>'
            '</div>',
            unsafe_allow_html=True
        )

    # 5. Output Deliverable with Visual Word Sheet Preview (Idea #3)
    if st.session_state.clean_text:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        st.markdown("### ✅ Submission-Ready Deliverable")
        
        out_words = len(st.session_state.clean_text.split())
        m1, m2, m3 = st.columns(3)
        m1.metric("Words Processed", f"{out_words:,}")
        m2.metric("Active Engine", tool_info["short_name"])
        m3.metric("Diagnostic Audit", "Unlocked ✓" if is_tool_pro else "⚠️ Action Needed")

        dl_col1, dl_col2 = st.columns([1.6, 1], vertical_alignment="center")
        with dl_col1:
            title_data = {
                "include": bool(tp_title.strip()),
                "title": tp_title,
                "author": tp_author,
                "institution": tp_inst,
                "course": tp_course,
                "instructor": tp_instructor,
                "date": tp_date
            }
            if is_tool_pro:
                docx_file = create_smart_docx(st.session_state.clean_text, active_tool, st.session_state.selected_style, title_data)
                st.download_button(
                    label=f"📥 Download Ready-to-Submit .docx ({tool_info['short_name']})",
                    data=docx_file,
                    file_name=f"FormatForge_{active_tool.capitalize()}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
            else:
                st.markdown(
                    f'<a href="{stripe_tool_url}" target="_blank" style="text-decoration:none;">'
                    '<div style="background:linear-gradient(180deg, #10B981 0%, #059669 100%); color:#052E16; padding:13px 18px; border-radius:10px; text-align:center; font-weight:800; font-size:14px; box-shadow: 0 4px 14px rgba(16, 185, 129, 0.25);">'
                    f'🔒 Download Ready-to-Submit .docx + Unlimited Edits ($1.99)'
                    '</div></a>',
                    unsafe_allow_html=True
                )
        with dl_col2:
            st.caption("💡 **Visual Preview:** See below how your paper will look printed from Microsoft Word.")

        # Visual Word-Page Paper Preview Card (Idea #3)
        formatted_html_preview = st.session_state.clean_text.replace("\n", "<br>")
        st.markdown(
            f'<div style="font-size: 12px; font-weight: 700; color: #94A3B8; margin-top: 16px; margin-bottom: 4px; text-transform: uppercase;">📄 Microsoft Word Document Preview:</div>'
            f'<div class="word-preview-sheet">{formatted_html_preview}</div>',
            unsafe_allow_html=True
        )

        # Raw Copy Box below
        with st.expander("📋 Copy Raw Text"):
            st.code(st.session_state.clean_text, language=None, wrap_lines=True)

        if st.session_state.editorial_notes:
            st.markdown("#### 🔍 Diagnostic & Rubric Audit")
            if is_tool_pro:
                with st.container(border=True):
                    st.markdown(st.session_state.editorial_notes)
            else:
                lines = [line.strip() for line in st.session_state.editorial_notes.split('\n') if line.strip()]
                first_teaser = lines[0] if lines else "⚠️ Potential issues detected in your draft."
                remaining_blurred = "<br>".join(lines[1:]) if len(lines) > 1 else "• Additional structural findings and missing citations flagged.<br>• Phrasing and rubric alignment recommendations ready."
                
                st.markdown(
                    '<div class="blur-container">'
                    f'<div style="font-weight:800; color:#F87171; font-size:14px; margin-bottom:6px;">🚨 Diagnostic Audit Preview (1 of {max(3, len(lines))} findings shown):</div>'
                    f'<div style="font-size:14px; color:#FFFFFF; font-weight:600; margin-bottom:10px;">{first_teaser}</div>'
                    f'<div class="blurred-body">{remaining_blurred}<br>• Detailed line-by-line verification completed.<br>• Formatting and margin specifications prepared for Word export.</div>'
                    '<div class="blur-cta-overlay">'
                    f'<div style="font-weight:800; color:#111827; margin-bottom:6px;">Unlock Full {tool_info["short_name"]} + Word (.docx) Export</div>'
                    '<div style="font-size:13px; color:#475569; margin-bottom:16px;">Includes <b>unlimited re-runs & paragraph edits</b> on this tool for your entire session.</div>'
                    f'<a href="{stripe_tool_url}" target="_blank" class="cta-btn">⚡ Unlock {tool_info["short_name"]} Pass — $1.99</a>'
                    '<div style="font-size:11px; color:#94A3B8; margin-top:10px;">One-time payment • Zero recurring subscriptions • Instant activation</div>'
                    '</div></div>',
                    unsafe_allow_html=True
                )

# ==============================================================================
# SHARED FOOTER
# ==============================================================================
st.markdown("<div style='margin-top: 48px; border-top: 1px solid rgba(148, 163, 184, 0.2); padding-top: 20px;'></div>", unsafe_allow_html=True)

foot_col1, foot_col2, foot_col3 = st.columns([6, 1, 1])
with foot_col1:
    st.markdown("<div style='font-size: 13px; color: #64748B; padding-top: 8px;'>© 2026 FormatForge Suite. Built for students & early careers.</div>", unsafe_allow_html=True)
with foot_col2:
    if st.button("Privacy", type="tertiary", use_container_width=True):
        privacy_modal()
with foot_col3:
    if st.button("Terms", type="tertiary", use_container_width=True):
        terms_modal()
