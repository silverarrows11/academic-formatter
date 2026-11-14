import io
import re
import os
import sys
import hmac
import hashlib
import html
import streamlit as st
import openai
import stripe
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def _make_favicon():
    """Green rounded square with a dark F, drawn at start-up (Pillow ships with Streamlit)."""
    try:
        from PIL import Image, ImageDraw
        size = 256
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=56, fill=(16, 185, 129, 255))
        dark = (5, 46, 22, 255)
        d.rectangle([84, 64, 112, 192], fill=dark)
        d.rectangle([84, 64, 184, 90], fill=dark)
        d.rectangle([84, 112, 160, 136], fill=dark)
        return img
    except Exception:
        return "⚡"

# 1. Page Configuration
st.set_page_config(
    page_title="FormatForge | Academic, AI Cleaner & ATS Resume Suite",
    page_icon=_make_favicon(),
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Clean, Restrained SaaS Styling (One Primary CTA Per View, Zero Truncation)
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
[data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"], .stDeployButton { display: none !important; }

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 880px;
}

/* Hero Card */
.hero-card {
    background: radial-gradient(circle at top right, #1E293B, #090D16 70%);
    padding: 44px 32px;
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
    max-width: 600px;
    line-height: 1.6;
}

/* Badges */
.badge {
    display: inline-block;
    padding: 5px 12px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}
.badge-pro {
    background: rgba(16, 185, 129, 0.18);
    color: #34D399;
    border: 1px solid #10B981;
}
.badge-free {
    background: rgba(148, 163, 184, 0.12);
    color: #CBD5E1;
    border: 1px solid #475569;
}

/* Buttons: every button shares one size and shape; green is reserved for the single primary action */
.stButton > button[kind="primary"], .stButton > button[kind="secondary"], .stDownloadButton > button {
    min-height: 48px !important;
    border-radius: 10px !important;
    padding: 10px 18px !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    transition: all 0.15s ease !important;
}
.stButton > button p, .stDownloadButton > button p {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
    line-height: 1.25 !important;
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(180deg, #10B981 0%, #059669 100%) !important;
    color: #052E16 !important;
    border: 1px solid #34D399 !important;
    box-shadow: 0 4px 16px rgba(16, 185, 129, 0.25) !important;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 22px rgba(16, 185, 129, 0.38) !important;
}
.stButton > button[kind="secondary"] {
    border: 1px solid #334155 !important;
    background: #1E293B !important;
    color: #E2E8F0 !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #64748B !important;
    background: #334155 !important;
}

/* Guided Stepper Bar */
.stepper-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: #0F172A;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 12px 20px;
    margin-bottom: 22px;
}
.step-item {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 13px;
    font-weight: 600;
    color: #8FA0B8;
}
.step-item.active {
    color: #FFFFFF;
    font-weight: 800;
}
.step-item.done {
    color: #34D399;
}
.step-num {
    width: 22px;
    height: 22px;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    font-weight: 800;
    background: #1E293B;
    border: 1px solid #475569;
}
.step-item.active .step-num {
    background: #10B981;
    color: #052E16;
    border-color: #10B981;
}
.step-item.done .step-num {
    background: rgba(16, 185, 129, 0.2);
    color: #34D399;
    border-color: #10B981;
}

/* Step 1: Tool Selection Cards */
.tool-choice-card {
    background: #0F172A;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 22px;
    min-height: 340px;
    box-sizing: border-box;
    margin-bottom: 10px;
}
.tool-choice-title {
    font-size: 17px;
    font-weight: 800;
    color: #FFFFFF;
    margin-bottom: 6px;
}
.tool-choice-desc {
    font-size: 13px;
    color: #94A3B8;
    line-height: 1.5;
    margin-bottom: 14px;
    min-height: 78px;
}
.tool-choice-list {
    font-size: 12.5px;
    color: #CBD5E1;
    line-height: 1.85;
}

/* Step 2: Compact Active Tool Banner */
.compact-banner {
    background: #0F172A;
    border: 1px solid #334155;
    border-left: 4px solid #10B981;
    border-radius: 12px;
    padding: 14px 18px;
    margin-bottom: 16px;
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 10px;
}

/* Step 3: Non-Truncating Summary Pills (Replaces broken st.metric) */
.summary-strip {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 12px;
    margin-bottom: 18px;
}
.summary-box {
    background: #0F172A;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 14px 18px;
}
.summary-label {
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    color: #94A3B8;
    letter-spacing: 0.04em;
    margin-bottom: 4px;
}
.summary-val {
    font-size: 16px;
    font-weight: 800;
    color: #FFFFFF;
    white-space: normal;
    line-height: 1.3;
}

/* Simulated Microsoft Word (.docx) Live Paper Preview */
.word-preview-wrapper {
    background: #0F172A;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 16px;
    margin-top: 14px;
    margin-bottom: 18px;
}
.word-toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding-bottom: 12px;
    margin-bottom: 14px;
    border-bottom: 1px solid #334155;
    font-size: 12px;
    color: #94A3B8;
    font-weight: 700;
}
.word-page {
    background: #FFFFFF;
    color: #0F172A;
    font-family: 'Times New Roman', Times, serif;
    font-size: 15px;
    line-height: 2.0;
    padding: 36px 44px;
    border-radius: 6px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.45);
    max-height: 380px;
    overflow-y: auto;
}

/* Dark-Mode Safe Blurred Audit Card */
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
    background: rgba(15, 23, 42, 0.95);
    backdrop-filter: blur(6px);
    border: 1px solid #10B981;
    border-radius: 12px;
    padding: 22px;
    text-align: center;
    margin-top: -45px;
    position: relative;
    z-index: 10;
}
.cta-btn {
    background: linear-gradient(180deg, #10B981 0%, #059669 100%);
    color: #052E16 !important;
    font-weight: 800;
    padding: 13px 26px;
    border-radius: 10px;
    text-decoration: none !important;
    font-size: 14.5px;
    display: inline-block;
    box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3);
}
.cta-btn:hover {
    background: #34D399;
}

/* Pricing Grid on Landing Page */
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
    padding: 26px;
    border-radius: 16px;
}
.price-card-pro {
    background: linear-gradient(145deg, #0F172A 0%, #1E293B 100%);
    color: #FFFFFF;
    border: 2px solid #10B981;
    padding: 26px;
    border-radius: 16px;
}

/* Locked-result strip (free preview) */
.lock-strip {
    background: linear-gradient(145deg, #0F172A 0%, #1E293B 100%);
    border: 1px solid #10B981;
    border-radius: 14px;
    padding: 22px;
    text-align: center;
    margin-bottom: 18px;
}

/* Brand mark, tool icons, stepper connectors */
.brand-mark {
    display: inline-flex; align-items: center; justify-content: center;
    width: 26px; height: 26px; border-radius: 7px; margin-right: 9px; vertical-align: middle;
    background: linear-gradient(180deg, #10B981, #059669); color: #052E16; font-weight: 800; font-size: 15px;
}
.tool-card-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
.tool-icon {
    display: inline-flex; align-items: center; justify-content: center; width: 36px; height: 36px; border-radius: 10px;
    background: rgba(16, 185, 129, 0.12); border: 1px solid rgba(52, 211, 153, 0.25); color: #34D399;
}
.tool-icon svg { width: 20px; height: 20px; }
.tool-eyebrow { font-size: 11px; font-weight: 800; color: #34D399; text-transform: uppercase; letter-spacing: 0.04em; }
.stepper-bar { flex-wrap: wrap; gap: 8px 0; margin-top: 10px; }
.price-card-free .benefit, .price-card-pro .benefit { margin-bottom: 0; padding-top: 4px; padding-bottom: 4px; line-height: 1.5; }
.step-sep { flex: 1; height: 1px; background: #334155; margin: 0 14px; min-width: 14px; }
.badge-free, .badge-pro { white-space: nowrap; }
.tool-choice-list { line-height: 1.6; }
.benefit { padding-left: 1.25em; text-indent: -1.25em; margin-bottom: 4px; }

/* One consistent dark look, whatever theme the visitor's browser or Streamlit would otherwise use */
:root { color-scheme: dark; }
.stApp, [data-testid="stAppViewContainer"] { background: #0B1220 !important; color: #E2E8F0; }
[data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] p, label { color: #E2E8F0; }
[data-testid="stCaptionContainer"] { color: #94A3B8 !important; }
[data-baseweb="textarea"], [data-baseweb="input"], [data-baseweb="select"] > div {
    background: #0F172A !important; border-color: #334155 !important; border-radius: 10px !important;
}
textarea, input { color: #E2E8F0 !important; -webkit-text-fill-color: #E2E8F0 !important; }
textarea::placeholder, input::placeholder { color: #64748B !important; -webkit-text-fill-color: #64748B !important; }
[data-baseweb="textarea"]:focus-within, [data-baseweb="input"]:focus-within, [data-baseweb="select"] > div:focus-within {
    border-color: #10B981 !important; box-shadow: 0 0 0 1px #10B981 !important;
}
[data-testid="stExpander"] details { background: #0F172A; border: 1px solid #334155 !important; border-radius: 12px; }
[data-testid="stVerticalBlockBorderWrapper"] { border-color: #334155 !important; border-radius: 14px !important; }

/* Tablet-width windows: columns get narrow, so reserve room for wrapped benefit lines */
@media (min-width: 641px) and (max-width: 900px) {
    .tool-choice-card { min-height: 440px; }
}

/* Small screens */
@media (max-width: 640px) {
    .hero-card { padding: 32px 20px; }
    .hero-card h1 { font-size: 28px; }
    .stepper-bar { padding: 12px 14px; gap: 8px 14px; }
    .step-sep { display: none; }
    .lbl-x { display: none; }
    .step-item { font-size: 11.5px; gap: 6px; }
    .step-num { width: 20px; height: 20px; }
    .stepper-bar { justify-content: space-between; }
    .word-page { padding: 24px 20px; }
    .tool-choice-card { min-height: 0; }
    .tool-choice-desc { min-height: 0; }
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# 3. Configuration Limits & Credentials
FREE_INPUT_LIMIT = 6000    # characters a free user may submit for their one free preview run
FREE_TEASER_CHARS = 900    # how much of the formatted result free users see
PRO_MAX_CHARS = 35000

# Set these in Railway. UNLOCK_SECRET is required for passes to survive a page refresh.
UNLOCK_SECRET = os.getenv("UNLOCK_SECRET", "")
CONTACT_EMAIL = os.getenv("CONTACT_EMAIL", "info@formatforge.help").strip()
try:
    REFUND_DAYS = int(os.getenv("REFUND_DAYS", "0"))
except ValueError:
    REFUND_DAYS = 0

CONTACT_MD = f"\n\n**Contact:** {CONTACT_EMAIL}" if CONTACT_EMAIL else ""
if CONTACT_EMAIL and REFUND_DAYS > 0:
    REFUND_MD = (f"\n\n**4. Refunds:** If the tool does not work for you, email {CONTACT_EMAIL} within "
                 f"{REFUND_DAYS} days of purchase with your Stripe receipt and we will refund the $1.99.")
    REFUND_NOTE = f"{REFUND_DAYS}-day refund promise: email {html.escape(CONTACT_EMAIL)}"
else:
    REFUND_MD = ""
    REFUND_NOTE = "One-time payment • No subscription"

FOOTER_TEXT = "© 2026 FormatForge • AI-assisted formatting: review results before you submit"
if CONTACT_EMAIL:
    FOOTER_TEXT += (f' • <a href="mailto:{html.escape(CONTACT_EMAIL)}" style="color:#94A3B8;">'
                    f'{html.escape(CONTACT_EMAIL)}</a>')
STRIPE_PAYMENT_URL = "https://buy.stripe.com/cNidRbdxg5l4c9902l5Ne01"
DELIMITER = "---AUDIT_AND_FEEDBACK---"

stripe_key = os.getenv("STRIPE_SECRET_KEY")
if stripe_key:
    stripe.api_key = stripe_key

TOOL_CONFIG = {
    "academic": {
        "short_name": "Academic Paper Forge",
        "headline": "Academic Paper & Citation Forge",
        "tagline": "Fix broken in-text citations, alphabetize bibliographies, and export a formatted Word (.docx) file.",
        "benefits": [
            "✓ Auto hanging indents (.docx)",
            "✓ Missing-reference cross-check",
            "✓ APA, MLA, Harvard, Chicago, IEEE",
            "✓ Optional cover page"
        ],
        "steps": [
            "Parsing body paragraphs & bibliography entries...",
            "Standardizing citation syntax & alphabetizing references...",
            "Cross-checking in-text citations & building .docx layout..."
        ],
        "button_text": "Format Paper & Run Citation Audit →",
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
        "short_name": "AI Cleaner & Auditor",
        "headline": "AI-Cliché Cleaner & Citation-Gap Auditor",
        "tagline": "Strip copy-paste ChatGPT markdown clutter, replace robotic LLM transition words, and flag uncited claims.",
        "benefits": [
            "✓ Strips ### & ** markdown clutter",
            "✓ Replaces robotic AI transitions",
            "✓ Flags uncited facts & statistics",
            "✓ Exports a clean .docx"
        ],
        "steps": [
            "Scanning draft for Markdown symbols & repetitive LLM phrasing...",
            "Refining transitions & checking factual claims for missing citations...",
            "Preparing clean Microsoft Word (.docx) manuscript..."
        ],
        "button_text": "Audit AI Clichés & Clean Document →",
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
        "short_name": "ATS Resume Forge",
        "headline": "Wall Street & Harvard Style ATS Resume",
        "tagline": "Turn messy work history or LinkedIn bios into a clean, single-column .docx resume built for one page.",
        "benefits": [
            "✓ XYZ-formula bullet rewrites",
            "✓ 0.6-inch margins for 1 page",
            "✓ Target-keyword suggestions",
            "✓ Single-column .docx export"
        ],
        "steps": [
            "Analyzing experience bullets & target role keywords...",
            "Rewriting bullet points into quantified XYZ impact statements...",
            "Compiling single-column 0.6-inch margin .docx layout..."
        ],
        "button_text": "Format Resume & Run ATS Audit →",
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

_SVG = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
        'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{}</svg>')
TOOL_ICONS = {
    "academic": '<span class="tool-icon">' + _SVG.format('<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>') + '</span>',
    "humanizer": '<span class="tool-icon">' + _SVG.format('<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>') + '</span>',
    "resume": '<span class="tool-icon">' + _SVG.format('<rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/>') + '</span>',
}

def tool_card_head(t_key):
    tag = "Unlocked" if t_key in st.session_state.unlocked_tools else "Free preview"
    return f'<div class="tool-card-head">{TOOL_ICONS[t_key]}<span class="tool-eyebrow">{tag}</span></div>'


# 4. Session State Initialization (With Guided Stepper State)
defaults = {
    "unlocked_tools": [],
    "free_uses": {"academic": 0, "humanizer": 0, "resume": 0},
    "active_tool": "academic",
    "studio_step": 1,  # 1 = Choose Tool, 2 = Input & Configure, 3 = Review & Export
    "input_text": "",
    "job_target": "",
    "last_uploaded_file": None,
    "clean_text": None,
    "editorial_notes": None,
    "selected_style": "APA 7th Edition",
    "current_view": "landing",
    "deeplink_done": False,
    "run_in_words": 0,
    "just_unlocked": None
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# 5. Signed Pass Recovery, Stripe Verification & Deep Links
TOOL_ALIASES = {"cleaner": "humanizer"}  # lets marketing links use friendlier names

def _pass_signature(tool_key):
    msg = f"formatforge:pass:v1:{tool_key}".encode()
    return hmac.new(UNLOCK_SECRET.encode(), msg, hashlib.sha256).hexdigest()[:32]

def build_pass_param(tools):
    """Signed value for the ?pass= URL parameter ('' when no secret is configured)."""
    if not UNLOCK_SECRET:
        return ""
    return ",".join(f"{t}.{_pass_signature(t)}" for t in tools if t in TOOL_CONFIG)

def read_pass_param(value):
    """Return only the tools whose signature checks out. Typed-in or edited values unlock nothing."""
    if not UNLOCK_SECRET or not value:
        return []
    tools = []
    for item in str(value).split(","):
        tool, _, sig = item.strip().partition(".")
        if tool in TOOL_CONFIG and sig and hmac.compare_digest(sig.encode(), _pass_signature(tool).encode()):
            tools.append(tool)
    return tools

@st.cache_resource
def _warn_missing_secret():
    print("WARNING: UNLOCK_SECRET is not set. Passes last for the current session only and cannot be restored from the URL.", file=sys.stderr)
    return True

if not UNLOCK_SECRET:
    _warn_missing_secret()

if "pass" in st.query_params:
    for t in read_pass_param(st.query_params["pass"]):
        if t not in st.session_state.unlocked_tools:
            st.session_state.unlocked_tools.append(t)
if "unlocked" in st.query_params:
    del st.query_params["unlocked"]  # the old unsigned parameter is no longer honoured

# Deep link from marketing pages, e.g. ?tool=resume&utm_source=seo (first load of a session only)
if not st.session_state.deeplink_done:
    st.session_state.deeplink_done = True
    _requested = st.query_params.get("tool")
    _requested = TOOL_ALIASES.get(_requested, _requested)
    if _requested in TOOL_CONFIG and "session_id" not in st.query_params:
        st.session_state.active_tool = _requested
        st.session_state.current_view = "editor"
        st.session_state.studio_step = 2

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
            st.session_state.just_unlocked = paid_tool
            # Drop user directly into Step 3 if they already have output, else Step 2
            st.session_state.studio_step = 3 if st.session_state.clean_text else 2
            tool_title = TOOL_CONFIG[paid_tool]["short_name"]
            st.toast(f"{tool_title} unlocked. Unlimited edits and .docx exports are active.", icon="✅")
        else:
            st.warning("We could not confirm that payment yet. If you were charged, please refresh in a minute."
                       + (f" Still stuck? Email {CONTACT_EMAIL}." if CONTACT_EMAIL else ""))

        st.query_params.clear()
        signed_pass = build_pass_param(st.session_state.unlocked_tools)
        if signed_pass:
            st.query_params["pass"] = signed_pass
    except Exception as e:
        print(f"Stripe verification failed: {type(e).__name__}: {e}", file=sys.stderr)
        st.error("We could not verify your payment automatically. If you were charged, "
                 + (f"email {CONTACT_EMAIL} with your Stripe receipt and we will sort it out." if CONTACT_EMAIL
                    else "please contact support with your Stripe receipt."))
        st.query_params.clear()

def get_stripe_url(tool_key):
    sep = "&" if "?" in STRIPE_PAYMENT_URL else "?"
    return f"{STRIPE_PAYMENT_URL}{sep}client_reference_id={tool_key}"

# 6. Helper Functions
def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    return uploaded_file.read().decode("utf-8", errors="ignore")

def make_teaser(text, max_chars=FREE_TEASER_CHARS):
    """Opening of the result for free users, cut at a sentence/paragraph boundary.
    Returns (teaser_text, words_not_shown). The full text never leaves the server."""
    paragraphs = [p for p in text.split("\n") if p.strip()]
    kept, used = [], 0
    for p in paragraphs:
        room = max_chars - used
        if len(p) <= room:
            kept.append(p)
            used += len(p) + 1
            continue
        if room > 200 or not kept:
            budget = room if room > 200 else max_chars
            cut = p[:budget]
            end = max(cut.rfind(". "), cut.rfind("? "), cut.rfind("! "))
            piece = cut[:end + 1] if end > budget // 3 else (cut.rsplit(" ", 1)[0] if " " in cut else cut)
            kept.append(piece.strip())
        break
    teaser = "\n".join(kept)
    return teaser, max(0, len(text.split()) - len(teaser.split()))

def split_audit_lines(notes):
    """Individual audit items (bullets if the model used them, otherwise non-empty lines)."""
    lines = [l.strip() for l in notes.split("\n") if l.strip()]
    bullets = [l for l in lines if re.match(r"^([-•*]|\d+[.)])\s", l)]
    return bullets or lines

def render_word_preview_html(content, tool_mode):
    paragraphs = [p.strip() for p in content.split('\n') if p.strip()]
    html_lines = []
    in_refs = False
    ref_headers = ["references", "reference list", "works cited", "bibliography"]

    for idx, p in enumerate(paragraphs):
        safe_p = html.escape(re.sub(r'^#+\s*', '', p))
        safe_p = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', safe_p)
        safe_p = re.sub(r'\*(.*?)\*', r'<i>\1</i>', safe_p)
        clean_low = re.sub(r'[*_#:]', '', p).strip().lower()

        if tool_mode == "resume":
            if idx == 0:
                html_lines.append(f'<div style="text-align:center; font-size:18px; font-weight:bold; line-height:1.3;">{safe_p}</div>')
            elif idx == 1 and ("|" in safe_p or "@" in safe_p):
                html_lines.append(f'<div style="text-align:center; font-size:12px; margin-bottom:10px; border-bottom:1px solid #CBD5E1; padding-bottom:6px; line-height:1.3;">{safe_p}</div>')
            elif clean_low in ["education", "experience", "professional experience", "work experience", "projects", "leadership", "skills", "technical skills"]:
                html_lines.append(f'<div style="font-weight:bold; text-transform:uppercase; font-size:13px; margin-top:10px; border-bottom:1px solid #0F172A; line-height:1.4;">{safe_p}</div>')
            elif p.startswith(('-', '•', '* ')):
                html_lines.append(f'<div style="padding-left:18px; font-size:13px; line-height:1.4; margin-top:2px;">{safe_p}</div>')
            else:
                html_lines.append(f'<div style="font-size:13.5px; font-weight:600; margin-top:6px; line-height:1.4;">{safe_p}</div>')
        else:
            if clean_low in ref_headers:
                in_refs = True
                html_lines.append(f'<div style="text-align:center; font-weight:bold; margin-top:14px;">{safe_p}</div>')
            elif idx == 0 and len(p) < 120 and not p.endswith('.'):
                html_lines.append(f'<div style="text-align:center; font-weight:bold; margin-bottom:8px;">{safe_p}</div>')
            elif in_refs:
                html_lines.append(f'<div style="padding-left:36px; text-indent:-36px;">{safe_p}</div>')
            else:
                html_lines.append(f'<div style="text-indent:36px;">{safe_p}</div>')

    return "".join(html_lines)

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

def add_page_number(paragraph, font_name="Times New Roman", font_size=12):
    """Insert a live PAGE field; Word and LibreOffice fill in the number."""
    run = paragraph.add_run()
    run.font.name = font_name
    run.font.size = Pt(font_size)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, end):
        run._r.append(el)

def add_bottom_border(paragraph):
    """Thin rule under a resume section heading (inserted in schema order so Word opens it cleanly)."""
    pPr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, val in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "000000")):
        bottom.set(qn(key), val)
    borders.append(bottom)
    pPr.insert_element_before(borders, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
                              "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
                              "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
                              "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
                              "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                              "w:pPrChange")

def create_smart_docx(content, tool_mode, style_name="APA 7th Edition", title_info=None):
    doc = Document()
    margin_size = Inches(0.6) if tool_mode == "resume" else Inches(1.0)
    for section in doc.sections:
        section.top_margin = margin_size
        section.bottom_margin = margin_size
        section.left_margin = margin_size
        section.right_margin = margin_size
        
        if tool_mode == "academic" and "IEEE" not in style_name:
            # Page number top right on every page; MLA also puts the author's last name before it.
            hp = section.header.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            author = ((title_info or {}).get("author") or "").strip()
            if "MLA" in style_name and author:
                lead = hp.add_run(f"{author.split()[-1]} ")
                lead.font.name = "Times New Roman"
                lead.font.size = Pt(12)
            add_page_number(hp)

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
                add_bottom_border(p)
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
        is_ieee = "IEEE" in style_name

        for idx, para_text in enumerate(paragraphs):
            clean_header = re.sub(r'[*_#:]', '', para_text).strip().lower()
            
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.15 if is_ieee else 2.0
            p.paragraph_format.space_after = Pt(6) if is_ieee else Pt(0)
            p.paragraph_format.space_before = Pt(0)

            if clean_header in ref_headers:
                in_references = True
                if not is_ieee:
                    p.paragraph_format.page_break_before = True  # reference list starts on its own page
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
                p.paragraph_format.first_line_indent = Inches(0.25 if is_ieee else 0.5)

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
        "**1. Your documents:** FormatForge does not store your documents. The text you submit and the result we generate "
        "are held in server memory only while your browser session is open, and FormatForge does not write them to a database or disk.\n\n"
        "**2. AI processing:** To produce your result, your text is sent to OpenAI's API. OpenAI processes it under its own API terms "
        "and privacy policy, which govern how long OpenAI may keep it. FormatForge does not use your text to train AI models.\n\n"
        "**3. Payments:** Payments are processed by Stripe. FormatForge never sees or stores your card details.\n\n"
        "**4. Your pass:** Unlocked tools are remembered through a signed token in your browser's address bar. "
        "FormatForge has no accounts and runs no customer database of its own.\n\n"
        "**5. Server logs:** Our hosting provider may keep standard technical logs (such as IP addresses and request times). "
        "We do not sell user data."
        + CONTACT_MD
    )

@st.dialog("Terms of Service", width="large")
def terms_modal():
    st.markdown(
        "**Effective Date: November 2026**\n\n"
        "**1. What FormatForge is:** An automated formatting and editing tool that uses AI to reformat text you provide. "
        "It does not check text against plagiarism or AI-detection tools, does not verify that any source exists, and makes no promise "
        "about grades, interviews, hiring decisions or any third-party score.\n\n"
        "**2. You are responsible for the result:** AI output can contain mistakes. Always review the result and check it against your "
        "instructor's, publisher's or employer's requirements before you submit it.\n\n"
        f"**3. Single-Tool Pass:** Each $1.99 pass is a one-time payment (not a subscription) that unlocks unlimited re-runs and `.docx` exports "
        f"for the one tool you purchased, up to {PRO_MAX_CHARS:,} characters per document."
        + REFUND_MD
        + CONTACT_MD
    )


# ==============================================================================
# VIEW 1: MARKETING LANDING PAGE
# ==============================================================================
if st.session_state.current_view == "landing":
    unlocked_count = len(st.session_state.unlocked_tools)
    status_badge = f'<span class="badge badge-pro">{unlocked_count} Unlocked</span>' if unlocked_count > 0 else '<span class="badge badge-free">Free Preview Available</span>'
    
    nav_col1, nav_col2 = st.columns([3, 1.2], vertical_alignment="center")
    with nav_col1:
        st.markdown(f'<div style="display:flex; align-items:center; gap:12px;"><span style="font-weight:800; font-size:22px; letter-spacing:-0.03em;"><span class="brand-mark">F</span>FormatForge</span> {status_badge}</div>', unsafe_allow_html=True)
    with nav_col2:
        if st.button("Open Studio →", use_container_width=True):
            st.session_state.current_view = "editor"
            st.session_state.studio_step = 1
            st.rerun()

    st.markdown("<div style='margin-bottom: 14px;'></div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="hero-card">'
        '<div class="hero-pill">3 document tools • $1.99 one-time pass per tool</div>'
        '<h1>Messy draft in.<br>Clean, formatted Word doc out.</h1>'
        '<p>Whether you have an 11:59 PM paper with broken citations, a ChatGPT draft full of Markdown clutter, or a messy internship resume—turn it into a cleanly formatted .docx in 3 guided steps. Free preview first.</p>'
        '</div>',
        unsafe_allow_html=True
    )

    # Single Primary CTA on Landing Hero (Eliminates multiple green buttons)
    cta_c1, cta_c2, cta_c3 = st.columns([0.8, 2.4, 0.8])
    with cta_c2:
        if st.button("Start free preview →", type="primary", use_container_width=True):
            st.session_state.current_view = "editor"
            st.session_state.studio_step = 1
            st.rerun()
        st.markdown('<div style="text-align:center; font-size:12px; color:#94A3B8; margin-top:6px;">Choose Academic Paper, AI Cleaner, or ATS Resume inside • No signup required</div>', unsafe_allow_html=True)

    # 3 Service Cards (Secondary buttons to jump straight to Step 2 for a specific tool)
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    for col, t_key in zip([c1, c2, c3], ["academic", "humanizer", "resume"]):
        info = TOOL_CONFIG[t_key]
        benefits_str = "".join(f'<div class="benefit">{b}</div>' for b in info["benefits"])
        with col:
            st.markdown(
                '<div class="tool-choice-card">'
                f'{tool_card_head(t_key)}'
                f'<div class="tool-choice-title">{info["short_name"]}</div>'
                f'<div class="tool-choice-desc">{info["tagline"]}</div>'
                f'<div class="tool-choice-list">{benefits_str}</div>'
                '</div>',
                unsafe_allow_html=True
            )
            if st.button("Choose this tool →", key=f"land_{t_key}", use_container_width=True):
                st.session_state.active_tool = t_key
                st.session_state.current_view = "editor"
                st.session_state.studio_step = 2
                st.rerun()

    st.markdown('<h2 style="text-align: center; margin-top: 40px; margin-bottom: 6px; font-size: 24px; font-weight: 800;">Simple, Pay-As-You-Go Pricing</h2>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #94A3B8; font-size: 14px; margin-bottom: 20px;">Zero subscriptions. Unlimited re-runs & paragraph edits on your unlocked tool.</p>', unsafe_allow_html=True)

    st.markdown(
        '<div class="pricing-grid">'
        '<div class="price-card-free">'
        '<div style="font-weight: 700; font-size: 13px; color: #94A3B8; text-transform: uppercase;">Free Preview (All 3 Tools)</div>'
        '<div style="font-size: 32px; font-weight: 800; margin: 8px 0 16px 0;">$0 <span style="font-size: 14px; font-weight: 500; color: #94A3B8;">/ test any tool</span></div>'
        '<div style="font-size: 14px; line-height: 2.0; color: #CBD5E1;">'
        '<div class="benefit">✓ 1 free preview run per tool (up to 6,000 characters)</div>'
        '<div class="benefit">✓ Live simulated Microsoft Word page preview</div>'
        '<div class="benefit">✓ See the opening of your result and the first audit finding</div>'
        '<div class="benefit">✓ Zero account signup required</div>'
        '</div></div>'
        '<div class="price-card-pro">'
        '<div style="display:flex; justify-content:space-between; align-items:center;">'
        '<span style="font-weight: 800; font-size: 13px; color: #34D399; text-transform: uppercase;">Single-Tool Unlimited Pass</span>'
        '<span style="background:#10B981; color:#052E16; font-size:11px; font-weight:800; padding:3px 8px; border-radius:99px;">NO SUBSCRIPTION</span>'
        '</div>'
        '<div style="font-size: 32px; font-weight: 800; margin: 8px 0 16px 0;">$1.99 <span style="font-size: 14px; font-weight: 500; color: #94A3B8;">/ per tool pass</span></div>'
        '<div style="font-size: 14px; line-height: 2.0; color: #E2E8F0;">'
        '<div class="benefit">✓ <b>Unlimited re-runs & edits</b> on that tool</div>'
        '<div class="benefit">✓ Up to <b>35,000 characters</b> per document</div>'
        '<div class="benefit">✓ <b>Word (.docx) download</b> & one-click copy</div>'
        '<div class="benefit">✓ Full Citation, AI-Cliché, or ATS Audit unlocked</div>'
        '</div></div></div>',
        unsafe_allow_html=True
    )

    st.markdown('<h3 style="text-align: center; margin-top: 24px; margin-bottom: 16px; font-size: 20px; font-weight: 800;">Frequently Asked Questions</h3>', unsafe_allow_html=True)
    with st.expander("Is the $1.99 pass a recurring subscription?"):
        st.write("No. It is a strict one-time payment of $1.99 via Stripe (Apple Pay, Google Pay, or card). You will never be billed again automatically.")
    with st.expander("What if I notice a typo or forgot a paragraph after paying $1.99?"):
        st.write("You get unlimited re-runs on the tool you unlocked! Your unlocked pass is kept in your browser's address bar as a signed link, so a refresh will not lock you out. Bookmark the page to come back to it later.")
    with st.expander("Are my documents stored?"):
        st.write("FormatForge does not store your documents: your text is processed in server memory and is not saved to a database or disk by us. To generate your result, the text is sent to OpenAI's API (see Privacy at the bottom of the page for details).")
    with st.expander("Can I trust the AI's output?"):
        st.write("It is AI-assisted and can make mistakes. FormatForge reformats and flags what is in your text; it does not verify that your sources exist or are accurate. Always review the result before you submit it.")
    if REFUND_MD:
        with st.expander("What if it does not work for me?"):
            st.write(f"Email {CONTACT_EMAIL} within {REFUND_DAYS} days of purchase with your Stripe receipt and we will refund the $1.99.")


# ==============================================================================
# VIEW 2: 3-STEP GUIDED FORMATTING STUDIO (Zero Clutter, Step-by-Step Flow)
# ==============================================================================
else:
    active_tool = st.session_state.active_tool
    is_tool_pro = active_tool in st.session_state.unlocked_tools
    tool_info = TOOL_CONFIG[active_tool]
    stripe_tool_url = get_stripe_url(active_tool)
    step = st.session_state.studio_step

    status_badge = (
        f'<span class="badge badge-pro">{tool_info["short_name"]}: Unlocked</span>'
        if is_tool_pro else
        f'<span class="badge badge-free">{tool_info["short_name"]}: Free preview</span>'
    )

    if step == 1:
        status_badge = ""  # no tool chosen yet, so no tool badge

    # Top Header Bar
    nav_c1, nav_c2, nav_c3 = st.columns([1.1, 2.1, 1.8], vertical_alignment="center")
    with nav_c1:
        if st.button("← Home", use_container_width=True):
            st.session_state.current_view = "landing"
            st.rerun()
    with nav_c2:
        st.markdown('<div style="text-align:center; font-weight:800; font-size:20px; letter-spacing:-0.03em;"><span class="brand-mark">F</span>FormatForge Studio</div>', unsafe_allow_html=True)
    with nav_c3:
        st.markdown(f'<div style="text-align:right;">{status_badge}</div>', unsafe_allow_html=True)

    # Visual 3-Step Guided Progress Bar
    s1_cls = "active" if step == 1 else "done"
    s2_cls = "active" if step == 2 else ("done" if step > 2 else "")
    s3_cls = "active" if step == 3 else ""

    st.markdown(
        '<div class="stepper-bar">'
        f'<div class="step-item {s1_cls}"><span class="step-num">{"✓" if step > 1 else "1"}</span><span>Choose<span class="lbl-x"> tool</span></span></div>'
        '<div class="step-sep"></div>'
        f'<div class="step-item {s2_cls}"><span class="step-num">{"✓" if step > 2 else "2"}</span><span>Add<span class="lbl-x"> your</span> draft</span></div>'
        '<div class="step-sep"></div>'
        f'<div class="step-item {s3_cls}"><span class="step-num">3</span><span>Review<span class="lbl-x"> & export</span></span></div>'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------------------------
    # STEP 1: SELECT YOUR SERVICE ENGINE
    # --------------------------------------------------------------------------
    if step == 1:
        st.markdown("<h3 style='text-align:center; margin-bottom:4px;'>Which tool do you need today?</h3>", unsafe_allow_html=True)
        st.markdown("<p style='text-align:center; color:#94A3B8; font-size:14px; margin-bottom:20px;'>Pick a tool to get started. You can switch at any time.</p>", unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        for col, t_key in zip([col1, col2, col3], ["academic", "humanizer", "resume"]):
            info = TOOL_CONFIG[t_key]
            benefits_html = "".join(f'<div class="benefit">{b}</div>' for b in info["benefits"])
            with col:
                st.markdown(
                    '<div class="tool-choice-card">'
                    f'{tool_card_head(t_key)}'
                    f'<div class="tool-choice-title">{info["short_name"]}</div>'
                    f'<div class="tool-choice-desc">{info["tagline"]}</div>'
                    f'<div class="tool-choice-list">{benefits_html}</div>'
                    '</div>',
                    unsafe_allow_html=True
                )
                if st.button("Choose this tool →", key=f"s1_{t_key}", use_container_width=True):
                    if st.session_state.active_tool != t_key:
                        st.session_state.active_tool = t_key
                        st.session_state.clean_text = None
                        st.session_state.editorial_notes = None
                        st.session_state.input_text = ""
                    st.session_state.studio_step = 2
                    st.rerun()

    # --------------------------------------------------------------------------
    # STEP 2: ADD YOUR DRAFT & RUN ENGINE (Single Primary Green CTA)
    # --------------------------------------------------------------------------
    elif step == 2:
        if st.session_state.just_unlocked == active_tool and is_tool_pro:
            st.success("Payment confirmed: this tool is unlocked. FormatForge does not store your text, so paste your draft again to get the full result.")
        # Compact header bar showing chosen tool + quick switch button
        top_b1, top_b2 = st.columns([3.2, 1.2], vertical_alignment="center")
        with top_b1:
            pills_inline = " &nbsp;•&nbsp; ".join(tool_info["benefits"][:3])
            st.markdown(
                '<div class="compact-banner">'
                f'<div><div style="font-weight:800; font-size:16px; color:#FFFFFF;">{tool_info["headline"]}</div>'
                f'<div style="font-size:12px; color:#34D399; font-weight:600; margin-top:3px;">{pills_inline}</div></div>'
                '</div>',
                unsafe_allow_html=True
            )
        with top_b2:
            if st.button("Switch tool", use_container_width=True):
                st.session_state.studio_step = 1
                st.rerun()

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
                            ["APA 7th Edition (Free Preview)", "MLA 9th Edition (pass required)", "Harvard Style (pass required)", "Chicago 17th Edition (pass required)", "IEEE (pass required)"],
                            index=0
                        )
                        if "pass required" in selected_style:
                            st.info("MLA, Harvard, Chicago and IEEE need the $1.99 pass. The free preview runs in APA 7th.")
                        selected_style = "APA 7th Edition"
                    st.session_state.selected_style = selected_style
                elif active_tool == "humanizer":
                    tone_mode = st.selectbox("Cleanup & Tone Standard", [
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
                if st.button("Load sample text", use_container_width=True):
                    st.session_state.input_text = tool_info["sample"]
                    st.rerun()

            user_text = st.text_area(
                "Document Input",
                value=st.session_state.input_text,
                height=230,
                placeholder=tool_info["placeholder"],
                label_visibility="collapsed"
            )
            st.session_state.input_text = user_text

            if active_tool == "resume":
                st.session_state.job_target = st.text_input(
                    "Target job title or job-description keywords (optional)",
                    value=st.session_state.job_target,
                    placeholder="e.g., Investment Banking Summer Analyst / Data Science Intern (Python, SQL, Financial Modeling)"
                )

            opt_col1, opt_col2 = st.columns(2)
            with opt_col1:
                with st.expander("Or upload a .docx / .txt file"):
                    uploaded_doc = st.file_uploader("Upload file", type=["docx", "txt"], label_visibility="collapsed")
                    if uploaded_doc is not None and uploaded_doc.name != st.session_state.last_uploaded_file:
                        st.session_state.input_text = extract_text_from_file(uploaded_doc)
                        st.session_state.last_uploaded_file = uploaded_doc.name
                        st.rerun()
            with opt_col2:
                if active_tool == "academic":
                    with st.expander("Add a cover page (optional)"):
                        st.session_state["tp_title"] = st.text_input("Paper Title", value=st.session_state.get("tp_title", ""), placeholder="e.g., The Impact of Screen Time on Sleep")
                        st.session_state["tp_author"] = st.text_input("Your Full Name", value=st.session_state.get("tp_author", ""), placeholder="e.g., Alex Rivera")
                        st.session_state["tp_inst"] = st.text_input("University / College", value=st.session_state.get("tp_inst", ""), placeholder="e.g., State University")
                        st.session_state["tp_course"] = st.text_input("Course Name", value=st.session_state.get("tp_course", ""), placeholder="e.g., PSY 101")
                        st.session_state["tp_instructor"] = st.text_input("Instructor Name", value=st.session_state.get("tp_instructor", ""), placeholder="e.g., Dr. Sarah Jenkins")
                        st.session_state["tp_date"] = st.text_input("Due Date", value=st.session_state.get("tp_date", ""), placeholder="e.g., November 7, 2026")

            char_len = len(user_text)
            word_count = len(user_text.split()) if user_text.strip() else 0
            tool_free_used = st.session_state.free_uses.get(active_tool, 0)
            
            meta_col1, meta_col2 = st.columns(2)
            with meta_col1:
                if not is_tool_pro:
                    st.caption(f"**{word_count} words** ({char_len:,}/{FREE_INPUT_LIMIT:,} free preview chars) • Free previews left: **{max(0, 1 - tool_free_used)}**")
                else:
                    st.caption(f"**{word_count} words** ({char_len:,}/{PRO_MAX_CHARS:,} chars) • **Unlimited edits active**")
            with meta_col2:
                st.markdown('<div style="text-align:right; font-size:12px; color:#94A3B8; font-weight:600;">Not stored by FormatForge</div>', unsafe_allow_html=True)

            # THE ONLY GREEN BUTTON ON STEP 2
            if st.button(tool_info["button_text"], type="primary", use_container_width=True):
                if not user_text.strip():
                    st.warning("Paste your text above or click 'Load sample text' first!")
                elif not is_tool_pro and tool_free_used >= 1:
                    st.error(f"You've used your free preview for {tool_info['short_name']}!")
                    st.markdown(f'<a href="{stripe_tool_url}" target="_blank" class="cta-btn">Unlock {tool_info["short_name"]} — $1.99 →</a>', unsafe_allow_html=True)
                elif not is_tool_pro and char_len > FREE_INPUT_LIMIT:
                    st.error(f"Your draft is {char_len:,} characters. The free preview accepts up to {FREE_INPUT_LIMIT:,}. Unlock the $1.99 Pass to format up to {PRO_MAX_CHARS:,} characters.")
                    st.markdown(f'<a href="{stripe_tool_url}" target="_blank" class="cta-btn">Unlock full length + .docx export — $1.99 →</a>', unsafe_allow_html=True)
                elif is_tool_pro and char_len > PRO_MAX_CHARS:
                    st.error(f"Your draft is {char_len:,} characters. The maximum per document is {PRO_MAX_CHARS:,}. Please split it into two parts.")
                else:
                    with st.status(f"Running {tool_info['short_name']}...", expanded=True) as status_box:
                        st.write(f"{tool_info['steps'][0]}")
                        if active_tool == "academic":
                            system_message = (
                                f"You are an elite university academic copyeditor. Reformat the user's input strictly according to {st.session_state.selected_style} standards.\n"
                                "- Fix all in-text citations and alphabetize the reference list.\n"
                                "- Wrap journal or book titles in single asterisks (*Title*) for proper Word italicization.\n"
                                "- Keep body paragraphs intact while fixing citation syntax.\n"
                                "- Never invent or change sources, authors, years, titles, journals, or page numbers. If a reference is incomplete, keep it as written and flag it in the audit.\n\n"
                                "Output Format:\n"
                                "1. Output ONLY the clean, formatted academic paper first. No markdown code fences.\n"
                                f"2. On a new line by itself, output this exact separator:\n{DELIMITER}\n"
                                "3. Below the separator, provide a bulleted 'Citation & Rubric Audit':\n"
                                "- Flag any in-text citation missing from the Reference list (or vice versa).\n"
                                "- List specific formatting fixes applied and 2 quick academic tone improvements."
                            )
                        elif active_tool == "humanizer":
                            system_message = (
                                "You are an academic AI-Cliché Cleaner and Citation-Gap Auditor. Clean up raw AI/ChatGPT drafts into clear, natural academic prose without changing their meaning.\n"
                                "- Strip all Markdown artifacts (### headers, **bold** clutter, bullet dumps) into clean, cohesive paragraphs.\n"
                                "- Replace overused LLM transition clichés ('delve', 'tapestry', 'testament', 'crucial', 'furthermore', 'multifaceted') with natural sentence variation.\n"
                                "- Do not add facts, statistics, or citations that are not in the input.\n\n"
                                "Output Format:\n"
                                "1. Output ONLY the cleaned document first. No markdown code fences.\n"
                                f"2. On a new line by itself, output this exact separator:\n{DELIMITER}\n"
                                "3. Below the separator, provide a bulleted 'AI-Cliché & Citation-Gap Audit':\n"
                                "- List the exact robotic AI phrases and Markdown clutter you removed.\n"
                                "- Flag any specific factual claims, statistics, or numbers in the text that lack an in-text citation."
                            )
                        else:
                            target_ctx = f"Target Role/Keywords: {st.session_state.job_target}" if st.session_state.job_target else "General competitive internship/corporate role"
                            system_message = (
                                f"You are an elite Wall Street & Tech ATS Resume writer. Reformat the user's input into a strict 1-page Harvard/Wall Street ATS Resume ({target_ctx}).\n"
                                "- First line: Candidate Name. Second line: Contact info separated by ' | '.\n"
                                "- Use clean uppercase section headers: EDUCATION, EXPERIENCE, PROJECTS, SKILLS.\n"
                                "- Put each role or project heading line (title, organization, dates) in **double asterisks** so it prints in bold.\n"
                                "- Rewrite weak experience bullets using the XYZ impact formula (Accomplished [X] as measured by [Y], by doing [Z]) with strong action verbs.\n"
                                "- Never invent numbers, employers, titles, dates, degrees, or achievements. Use only metrics that appear in the input. If a bullet has no metric, keep it factual without one and list it in the audit as needing a real number from the candidate.\n\n"
                                "Output Format:\n"
                                "1. Output ONLY the formatted resume text first. Start bullet points with '• '. No markdown code fences.\n"
                                f"2. On a new line by itself, output this exact separator:\n{DELIMITER}\n"
                                "3. Below the separator, provide a bulleted 'ATS & Impact Audit':\n"
                                "- Highlight weak bullets you rewrote, and list bullets that still need a real number from the candidate.\n"
                                "- List 3 high-value ATS keywords recommended for their target role."
                            )

                        try:
                            st.write(f"{tool_info['steps'][1]}")
                            client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
                            response = client.chat.completions.create(
                                model="gpt-4o-mini",
                                messages=[
                                    {"role": "system", "content": system_message},
                                    {"role": "user", "content": user_text}
                                ],
                                timeout=60.0
                            )
                            st.write(f"{tool_info['steps'][2]}")
                            output_content = response.choices[0].message.content
                            
                            if DELIMITER in output_content:
                                clean, notes = output_content.split(DELIMITER, 1)
                                st.session_state.clean_text = clean.strip()
                                st.session_state.editorial_notes = notes.strip()
                            else:
                                st.session_state.clean_text = output_content.strip()
                                st.session_state.editorial_notes = "• Document standardized and formatted.\n• Structural audit completed."
                                
                            st.session_state.run_in_words = word_count
                            st.session_state.just_unlocked = None
                            if not is_tool_pro:
                                st.session_state.free_uses[active_tool] = tool_free_used + 1
                            
                            status_box.update(label="Done. Opening your result...", state="complete", expanded=False)
                            st.session_state.studio_step = 3
                            st.rerun()
                            
                        except Exception as e:
                            status_box.update(label="Error during processing", state="error", expanded=True)
                            print(f"OpenAI error: {type(e).__name__}: {e}", file=sys.stderr)
                            st.error("Something went wrong while processing your text. Your draft is still here, so please try again in a moment."
                                     + (f" If it keeps happening, email {CONTACT_EMAIL}." if CONTACT_EMAIL else ""))

    # --------------------------------------------------------------------------
    # STEP 3: REVIEW DELIVERABLE & EXPORT .DOCX (No Scrolling, Zero Truncation)
    # --------------------------------------------------------------------------
    elif step == 3:
        if not st.session_state.clean_text:
            st.session_state.studio_step = 2
            st.rerun()

        # Navigation Row at Top of Step 3
        act_c1, act_c2 = st.columns([1.5, 2.5], vertical_alignment="center")
        with act_c1:
            if st.button("← Edit Draft / Run Again", use_container_width=True):
                st.session_state.studio_step = 2
                st.rerun()
        with act_c2:
            edit_hint = ("Need to add a forgotten paragraph? Click <b>Edit Draft</b>—unlimited re-runs are included with your pass."
                         if is_tool_pro else "Your free preview run is used. Unlock this tool to edit and re-run without limits.")
            st.markdown(f'<div style="text-align:right; font-size:13px; color:#94A3B8;">{edit_hint}</div>', unsafe_allow_html=True)

        st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

        # Custom Non-Truncating Summary Strip (Fixes the "Academic Pap..." overflow bug)
        out_words = len(st.session_state.clean_text.split())
        audit_status_html = (
            '<span style="color:#34D399;">✓ Full Audit Unlocked</span>'
            if is_tool_pro else
            '<span style="color:#FBBF24;">First finding shown</span>'
        )
        st.markdown(
            '<div class="summary-strip">'
            f'<div class="summary-box"><div class="summary-label">Words (in → out)</div><div class="summary-val">{st.session_state.run_in_words:,} → {out_words:,}</div></div>'
            f'<div class="summary-box"><div class="summary-label">Active tool</div><div class="summary-val">{tool_info["short_name"]}</div></div>'
            f'<div class="summary-box"><div class="summary-label">Diagnostic Audit</div><div class="summary-val">{audit_status_html}</div></div>'
            '</div>',
            unsafe_allow_html=True
        )

        # Primary Export Button (Only 1 Green Button on Step 3)
        title_data = {
            "include": bool(st.session_state.get("tp_title", "").strip()),
            "title": st.session_state.get("tp_title", ""),
            "author": st.session_state.get("tp_author", ""),
            "institution": st.session_state.get("tp_inst", ""),
            "course": st.session_state.get("tp_course", ""),
            "instructor": st.session_state.get("tp_instructor", ""),
            "date": st.session_state.get("tp_date", "")
        }

        if is_tool_pro:
            docx_file = create_smart_docx(st.session_state.clean_text, active_tool, st.session_state.selected_style, title_data)
            st.download_button(
                label="Download Word (.docx)",
                data=docx_file,
                file_name=f"FormatForge_{active_tool.capitalize()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )

        if is_tool_pro:
            st.caption("Bookmark this page to keep your pass. There are no accounts, so the link is your key.")
        st.caption("AI-assisted result. Review it before you submit or send it. FormatForge reformats and flags what is in your text; it does not verify that sources exist.")

        # Simulated Microsoft Word (.docx) Sheet Preview
        preview_text = st.session_state.clean_text
        locked_words = 0
        if not is_tool_pro:
            preview_text, locked_words = make_teaser(preview_text)
        preview_html = render_word_preview_html(preview_text, active_tool)

        st.markdown(
            '<div class="word-preview-wrapper">'
            '<div class="word-toolbar">'
            '<span>Word (.docx) preview</span>'
            f'<span>{st.session_state.selected_style}</span>'
            '</div>'
            f'<div class="word-page">{preview_html}</div>'
            '</div>',
            unsafe_allow_html=True
        )

        if not is_tool_pro:
            more_line = (f"<b>{locked_words:,} more words</b> are formatted and ready."
                         if locked_words > 0 else "Your full result is ready.")
            # THE ONLY GREEN CTA FOR FREE USERS ON STEP 3
            st.markdown(
                '<div class="lock-strip">'
                f'<div style="font-size:16px; color:#FFFFFF; margin-bottom:4px;">{more_line}</div>'
                '<div style="font-size:13px; color:#CBD5E1; margin-bottom:16px;">Unlock this tool to read the full result, see the full audit, and download the Word (.docx) file.</div>'
                f'<a href="{stripe_tool_url}" target="_blank" class="cta-btn">Unlock full result — $1.99</a>'
                f'<div style="font-size:11px; color:#94A3B8; margin-top:10px;">{REFUND_NOTE}</div>'
                '</div>',
                unsafe_allow_html=True
            )

        if is_tool_pro:
            with st.expander("Copy text"):
                st.code(st.session_state.clean_text, language=None, wrap_lines=True)

        # Diagnostic & Rubric Audit Section
        if st.session_state.editorial_notes:
            st.markdown("#### Audit findings")
            if is_tool_pro:
                with st.container(border=True):
                    st.markdown(st.session_state.editorial_notes)
            else:
                findings = split_audit_lines(st.session_state.editorial_notes)
                total = len(findings)
                if findings:
                    first_finding = html.escape(re.sub(r"\*\*|__", "", re.sub(r"^([-•*]|\d+[.)])\s*", "", findings[0])))
                else:
                    first_finding = "No issues were flagged in this preview."
                remaining = max(0, total - 1)
                lock_line = (f"{remaining} more item{'s' if remaining != 1 else ''} in the full audit."
                             if remaining > 0 else "The full audit and the .docx download come with the pass.")
                st.markdown(
                    '<div class="blur-container">'
                    f'<div style="font-weight:800; color:#FBBF24; font-size:14px; margin-bottom:6px;">Audit preview: first of {total} item{"s" if total != 1 else ""}</div>'
                    f'<div style="font-size:14px; color:#FFFFFF; font-weight:600; margin-bottom:12px;">{first_finding}</div>'
                    f'<div style="font-size:13px; color:#94A3B8;">{lock_line}</div>'
                    '</div>',
                    unsafe_allow_html=True
                )

# ==============================================================================
# SHARED FOOTER
# ==============================================================================
st.markdown("<div style='margin-top: 48px; border-top: 1px solid rgba(148, 163, 184, 0.2); padding-top: 20px;'></div>", unsafe_allow_html=True)

foot_col1, foot_col2, foot_col3 = st.columns([6, 1, 1])
with foot_col1:
    st.markdown(f"<div style='font-size: 13px; color: #94A3B8; padding-top: 8px;'>{FOOTER_TEXT}</div>", unsafe_allow_html=True)
with foot_col2:
    if st.button("Privacy", type="tertiary", use_container_width=True):
        privacy_modal()
with foot_col3:
    if st.button("Terms", type="tertiary", use_container_width=True):
        terms_modal()
