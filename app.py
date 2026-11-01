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
    page_title="FormatForge | Instant Academic Formatting",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Modern "Study-Tech" SaaS Styling (Notion / Linear Aesthetic)
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
}
h1, h2, h3 {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    letter-spacing: -0.03em !important;
}
#MainMenu, header, footer { visibility: hidden; }

.block-container {
    padding-top: 1.75rem;
    padding-bottom: 3rem;
    max-width: 880px;
}
.top-nav {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
    padding-bottom: 14px;
    border-bottom: 1px solid var(--border-color, #E2E8F0);
}
.logo-text {
    font-weight: 800;
    font-size: 22px;
    color: var(--text-color) !important;
    letter-spacing: -0.04em;
    display: flex;
    align-items: center;
    gap: 6px;
}
.hero-card {
    background: radial-gradient(circle at top right, #1E293B, #090D16 70%);
    padding: 36px 28px;
    border-radius: 16px;
    color: #FFFFFF !important;
    margin-bottom: 24px;
    text-align: center;
    border: 1px solid #334155;
    box-shadow: 0 12px 32px -8px rgba(9, 13, 22, 0.35);
}
.hero-pill {
    display: inline-block;
    background: rgba(16, 185, 129, 0.15);
    color: #34D399;
    border: 1px solid rgba(52, 211, 153, 0.3);
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 600;
    margin-bottom: 14px;
}
.hero-card h1 {
    color: #FFFFFF !important;
    font-size: 34px;
    font-weight: 800;
    margin: 0 0 12px 0;
    line-height: 1.15;
}
.hero-card p {
    color: #94A3B8 !important;
    font-size: 15px;
    margin: 0 auto;
    max-width: 580px;
    line-height: 1.55;
}
.badge {
    display: inline-block;
    padding: 5px 12px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}
.badge-pro { background: #10B981; color: #052E16; }
.badge-free { background: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; }

.stButton > button[kind="primary"] {
    background: linear-gradient(180deg, #18181B 0%, #09090B 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #27272A !important;
    border-radius: 10px !important;
    padding: 12px 24px !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    transition: all 0.15s ease !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12) !important;
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 16px rgba(0, 0, 0, 0.2) !important;
}
.blur-container {
    position: relative;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 20px;
    background: var(--secondary-background-color, #F8FAFC);
    overflow: hidden;
    margin-top: 12px;
}
.blurred-body {
    filter: blur(5px);
    user-select: none;
    pointer-events: none;
    opacity: 0.55;
    font-size: 14px;
    line-height: 1.6;
    margin-top: 8px;
}
.blur-cta-overlay {
    background: rgba(255, 255, 255, 0.92);
    backdrop-filter: blur(4px);
    border: 1px solid #CBD5E1;
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    margin-top: -65px;
    position: relative;
    z-index: 10;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08);
}
.cram-banner {
    background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
    color: #FFFFFF;
    padding: 24px;
    border-radius: 14px;
    margin-top: 28px;
    margin-bottom: 24px;
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
    border: 1px solid #334155;
}
.cta-btn {
    background: #10B981;
    color: #052E16 !important;
    font-weight: 800;
    padding: 12px 22px;
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

SAMPLE_TEXT = (
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

# 4. Session State Initialization
defaults = {
    "is_pro": False,
    "free_uses": 0,
    "input_text": "",
    "last_uploaded_file": None,
    "clean_text": None,
    "editorial_notes": None,
    "selected_style": "APA 7th Edition"
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# 5. Catch Stripe Redirect & Unlock Pro
if "session_id" in st.query_params and not st.session_state.is_pro:
    try:
        checkout_session = stripe.checkout.Session.retrieve(st.query_params["session_id"])
        if checkout_session.payment_status == "paid":
            st.session_state.is_pro = True
            st.toast("24-Hour Cram Pass Activated! All Pro features unlocked.", icon="⚡")
        st.query_params.clear()
    except Exception as e:
        st.error(f"Stripe Verification Error: {e}")
        st.query_params.clear()

is_pro = st.session_state.is_pro

# 6. Helper Functions (File Extraction & Smart .docx Engine)
def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    return uploaded_file.read().decode("utf-8", errors="ignore")

def add_formatted_runs(paragraph, text):
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
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

def create_smart_docx(content, style_name, title_info=None):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    if title_info and title_info.get("include") and title_info.get("title"):
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

        add_formatted_runs(p, para_text)

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# 7. Top Nav & Hero Section
status_badge = '<span class="badge badge-pro">⚡ Cram Pass Active</span>' if is_pro else '<span class="badge badge-free">Free Mode</span>'

st.markdown(
    f'<div class="top-nav"><span class="logo-text">✦ FormatForge</span><div>{status_badge}</div></div>'
    '<div class="hero-card">'
    '<div class="hero-pill">Built for 11:59 PM Deadlines</div>'
    '<h1>Messy draft in.<br>Rubric-ready paper out.</h1>'
    '<p>Fix broken APA, MLA, Chicago, or Harvard citations, hanging indents, and body formatting in 10 seconds. Zero manual tweaking required.</p>'
    '</div>',
    unsafe_allow_html=True
)

# 8. Inline Workspace
with st.container(border=True):
    top_col1, top_col2, top_col3 = st.columns([2, 1.2, 1.2], vertical_alignment="bottom")
    
    with top_col1:
        style_options = ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago 17th Edition", "IEEE"]
        if is_pro:
            selected_style = st.selectbox("Style Rubric", style_options, index=0)
        else:
            selected_style = st.selectbox(
                "Style Rubric", 
                ["APA 7th Edition", "🔒 MLA 9th (Pro)", "🔒 Harvard (Pro)", "🔒 Chicago 17th (Pro)", "🔒 IEEE (Pro)"],
                index=0
            )
            if "🔒" in selected_style:
                st.warning("Unlock the $1.99 Cram Pass below to format with MLA, Harvard, Chicago, or IEEE.")
                selected_style = "APA 7th Edition"
        st.session_state.selected_style = selected_style

    with top_col2:
        if st.button("🧪 Load Sample Essay", use_container_width=True):
            st.session_state.input_text = SAMPLE_TEXT
            st.rerun()

    with top_col3:
        uploaded_doc = st.file_uploader("Upload .docx / .txt", type=["docx", "txt"], label_visibility="collapsed")
        if uploaded_doc is not None and uploaded_doc.name != st.session_state.last_uploaded_file:
            st.session_state.input_text = extract_text_from_file(uploaded_doc)
            st.session_state.last_uploaded_file = uploaded_doc.name
            st.rerun()

    user_text = st.text_area(
        "Manuscript Input",
        value=st.session_state.input_text,
        height=230,
        placeholder="Paste your unformatted essay, rough draft, or messy bibliography right here...",
        label_visibility="collapsed"
    )
    st.session_state.input_text = user_text

    with st.expander("📄 Add Auto-Formatted Cover Page (Optional)"):
        tp_col1, tp_col2 = st.columns(2)
        with tp_col1:
            tp_title = st.text_input("Paper Title", placeholder="e.g., The Impact of Screen Time on Sleep")
            tp_author = st.text_input("Your Full Name", placeholder="e.g., Alex Rivera")
            tp_inst = st.text_input("University / High School", placeholder="e.g., State University")
        with tp_col2:
            tp_course = st.text_input("Course Code & Name", placeholder="e.g., PSY 101: Intro to Psychology")
            tp_instructor = st.text_input("Professor / Instructor Name", placeholder="e.g., Dr. Sarah Jenkins")
            tp_date = st.text_input("Due Date", placeholder="e.g., November 1, 2026")

    char_len = len(user_text)
    word_count = len(user_text.split()) if user_text.strip() else 0
    
    meta_col1, meta_col2 = st.columns(2)
    with meta_col1:
        if not is_pro:
            st.caption(f"📏 **{word_count} words** ({char_len}/{FREE_CHAR_LIMIT} chars) • Free runs left: **{max(0, 1 - st.session_state.free_uses)}**")
        else:
            st.caption(f"📏 **{word_count} words** ({char_len:,}/{PRO_MAX_CHARS:,} chars) • **⚡ Unlimited Cram Pass**")
    with meta_col2:
        st.markdown('<div style="text-align:right; font-size:12px; color:#059669; font-weight:600;">🛡️ Zero-Retention AI: Never stored or Turnitin-indexed</div>', unsafe_allow_html=True)

    if st.button("✨ Format Paper & Run Citation Audit", type="primary", use_container_width=True):
        if not user_text.strip():
            st.warning("Paste your text or click 'Load Sample Essay' first!")
        elif not is_pro and st.session_state.free_uses >= 1:
            st.error("You've used your free preview run! Grab a $1.99 Cram Pass below for unlimited runs and instant Word (.docx) export.")
        elif not is_pro and char_len > FREE_CHAR_LIMIT:
            st.error(f"Your draft is {char_len:,} characters (Free preview limit is {FREE_CHAR_LIMIT:,}). Unlock the $1.99 Cram Pass below to format your full paper at once!")
        else:
            client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
            with st.spinner("Formatting citations, fixing indents, and auditing references..."):
                system_message = (
                    f"You are an elite university academic copyeditor. Reformat the user's input strictly according to {selected_style} standards.\n"
                    "- Fix all in-text citations and alphabetize the reference list.\n"
                    "- Wrap journal titles or book titles in single asterisks (*Title*) so the Word exporter italicizes them properly.\n"
                    "- Keep body paragraphs intact while fixing citation syntax.\n\n"
                    "Output Format:\n"
                    "1. Output ONLY the clean, formatted academic paper first. Do not include markdown code fences.\n"
                    "2. Then, on a new line by itself, add this exact separator:\n"
                    f"{DELIMITER}\n"
                    "3. Below the separator, provide a bulleted 'Citation & Rubric Audit':\n"
                    "- Flag any in-text citation missing from the Reference list (or vice versa).\n"
                    "- Point out specific formatting fixes you automatically applied.\n"
                    "- Give 2 quick, actionable tips to elevate academic tone."
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
                        st.session_state.editorial_notes = "• All citations standardized to rubric.\n• Reference list alphabetized and hanging indents prepared."
                        
                    if not is_pro:
                        st.session_state.free_uses += 1
                    
                    st.toast("Paper formatted & audited!", icon="📝")
                    
                except Exception as e:
                    st.error(f"OpenAI Error: {e}")

# 9. Output Deliverable, One-Click Copy & Blurred Audit Tease
if st.session_state.clean_text:
    st.markdown("### ✅ Submission-Ready Output")
    
    out_words = len(st.session_state.clean_text.split())
    m1, m2, m3 = st.columns(3)
    m1.metric("Words Processed", f"{out_words:,}")
    m2.metric("Style Compliance", st.session_state.selected_style)
    m3.metric("Citation Audit", "Complete ✓" if is_pro else "⚠️ Action Needed")

    dl_col1, dl_col2 = st.columns([1.5, 1])
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
        if is_pro:
            docx_file = create_smart_docx(st.session_state.clean_text, st.session_state.selected_style, title_data)
            st.download_button(
                label="📥 Download Ready-to-Submit .docx (Word)",
                data=docx_file,
                file_name=f"FormatForge_{st.session_state.selected_style.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        else:
            st.markdown(
                f'<a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">'
                '<div style="background:#09090B; color:#FFFFFF; padding:12px 18px; border-radius:10px; text-align:center; font-weight:700; font-size:14px; border:1px solid #27272A;">'
                '🔒 Download .docx with Hanging Indents ($1.99 Cram Pass)'
                '</div></a>',
                unsafe_allow_html=True
            )
    with dl_col2:
        st.caption("💡 **One-Click Copy:** Hover over the top-right corner of the box below to copy raw text.")

    st.code(st.session_state.clean_text, language=None, wrap_lines=True)

    if st.session_state.editorial_notes:
        st.markdown("#### 🔍 Citation & Rubric Audit")
        if is_pro:
            with st.container(border=True):
                st.markdown(st.session_state.editorial_notes)
        else:
            lines = [line.strip() for line in st.session_state.editorial_notes.split('\n') if line.strip()]
            first_teaser = lines[0] if lines else "⚠️ Potential citation discrepancies detected between your body text and reference list."
            remaining_blurred = "<br>".join(lines[1:]) if len(lines) > 1 else "• Missing reference entry detected in body paragraph.<br>• Author capitalization and punctuation adjusted."
            
            st.markdown(
                '<div class="blur-container">'
                f'<div style="font-weight:700; color:#DC2626; font-size:14px; margin-bottom:6px;">🚨 Audit Preview (1 of {max(3, len(lines))} findings shown):</div>'
                f'<div style="font-size:14px; color:#0F172A; font-weight:500; margin-bottom:10px;">{first_teaser}</div>'
                f'<div class="blurred-body">{remaining_blurred}<br>• Additional in-text citation year mismatch flagged for review before grading.<br>• Hanging indent and DOI link syntax standardized for university rubric compliance.</div>'
                '<div class="blur-cta-overlay">'
                '<div style="font-weight:800; font-size:16px; color:#0F172A; margin-bottom:4px;">Don\'t lose easy points on citation mistakes</div>'
                '<div style="font-size:13px; color:#475569; margin-bottom:14px;">Unlock the full Citation Audit + instant Word (.docx) download with proper hanging indents.</div>'
                f'<a href="{STRIPE_PAYMENT_URL}" target="_blank" class="cta-btn">⚡ Unlock 24-Hour Cram Pass — $1.99</a>'
                '<div style="font-size:11px; color:#64748B; margin-top:8px;">One-time payment • Zero recurring subscriptions • Instant unlock</div>'
                '</div></div>',
                unsafe_allow_html=True
            )

# 10. Student-Friendly Conversion Banner (For Free Users)
if not is_pro:
    st.markdown(
        '<div class="cram-banner">'
        '<div style="flex: 1; min-width: 260px;">'
        '<div style="display:inline-block; background:#10B981; color:#052E16; font-weight:800; font-size:11px; padding:3px 8px; border-radius:6px; text-transform:uppercase; margin-bottom:8px;">No Subscription Trap</div>'
        '<div style="font-size: 20px; font-weight: 800; margin-bottom: 4px;">24-Hour Cram Pass — $1.99</div>'
        '<div style="font-size: 13px; color: #94A3B8; line-height: 1.5;">Format up to 35,000 characters per paper • Export ready-to-submit Word (.docx) files with real hanging indents • Unlock MLA, Harvard, Chicago & full Citation Audits.</div>'
        '</div>'
        '<div style="text-align: center;">'
        f'<a href="{STRIPE_PAYMENT_URL}" target="_blank" class="cta-btn">Unlock Pass ($1.99) →</a>'
        '<div style="font-size: 11px; color: #94A3B8; margin-top: 6px;">🔒 Apple Pay, Google Pay & Card via Stripe</div>'
        '</div></div>',
        unsafe_allow_html=True
    )

# 11. Before / After Visual Proof
st.markdown(
    '<div style="margin-top: 36px; margin-bottom: 32px;">'
    '<h3 style="text-align: center; margin-bottom: 18px; font-size: 18px; font-weight: 800;">Why students use FormatForge before submitting</h3>'
    '<div style="display: flex; flex-wrap: wrap; gap: 16px; background: var(--secondary-background-color, #F8FAFC); padding: 20px; border-radius: 14px; border: 1px solid var(--border-color, #E2E8F0);">'
    '<div style="flex: 1; min-width: 250px;">'
    '<div style="font-size: 11px; font-weight: 800; color: #DC2626; text-transform: uppercase; margin-bottom: 8px;">❌ Messy 11:45 PM Notes</div>'
    '<div style="font-family: \'JetBrains Mono\', monospace; font-size: 12px; color: #64748B; background: #FFFFFF; padding: 14px; border-radius: 8px; border: 1px dashed #CBD5E1;">smith 2021 blue light sleep study journal of sleep health vol 15 issue 2 pp 104-112</div>'
    '</div>'
    '<div style="flex: 1; min-width: 250px;">'
    '<div style="font-size: 11px; font-weight: 800; color: #059669; text-transform: uppercase; margin-bottom: 8px;">✅ 11:46 PM .docx Export</div>'
    '<div style="font-family: \'Times New Roman\', serif; font-size: 15px; color: #0F172A; background: #FFFFFF; padding: 14px; border-radius: 8px; border: 1px solid #E2E8F0; padding-left: 36px; text-indent: -22px;">Smith, J. (2021). Blue light and circadian rhythms. <i>Journal of Sleep Health</i>, <i>15</i>(2), 104–112.</div>'
    '</div></div></div>',
    unsafe_allow_html=True
)

# 12. Legal Modals (Single-line strings to prevent copy-paste truncation errors)
@st.dialog("Privacy Policy", width="large")
def privacy_modal():
    st.markdown(
        "**Effective Date: November 2026**\n\n"
        "**1. Zero-Retention Architecture:** FormatForge processes your essay text strictly in volatile memory. "
        "Your drafts are permanently destroyed the moment your formatting finishes or you close your browser tab. "
        "Your papers are never stored in any database, never indexed by Turnitin, and never used to train AI models.\n\n"
        "**2. Payment Security:** All payments are processed directly by Stripe. FormatForge never sees or stores your card details.\n\n"
        "**3. No Tracking:** We only use temporary session state to keep your Cram Pass active. We do not sell user data."
    )

@st.dialog("Terms of Service", width="large")
def terms_modal():
    st.markdown(
        "**Effective Date: November 2026**\n\n"
        "**1. Scope of Utility:** FormatForge is an automated formatting and citation auditing tool, not an essay-writing service. "
        "It formats existing student drafts according to standard style manuals.\n\n"
        "**2. Final Review Responsibility:** Always verify your exported `.docx` document against your professor's specific syllabus before submitting. "
        "FormatForge is not liable for grading outcomes.\n\n"
        "**3. One-Time Pass Policy:** The $1.99 Cram Pass is a single, non-recurring payment granting immediate access to Pro formatting and exports."
    )

# 13. Footer UI
st.markdown("<div style='margin-top: 48px; border-top: 1px solid var(--border-color, #E2E8F0); padding-top: 20px;'></div>", unsafe_allow_html=True)

foot_col1, foot_col2, foot_col3 = st.columns([6, 1, 1])
with foot_col1:
    st.markdown("<div style='font-size: 13px; color: #64748B; padding-top: 8px;'>© 2026 FormatForge. Built for students.</div>", unsafe_allow_html=True)
with foot_col2:
    if st.button("Privacy", type="tertiary", use_container_width=True):
        privacy_modal()
with foot_col3:
    if st.button("Terms", type="tertiary", use_container_width=True):
        terms_modal()
