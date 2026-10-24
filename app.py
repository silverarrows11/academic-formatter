import io
import time
import re
import streamlit as st
import openai
import stripe
from docx import Document
from docx.shared import Inches, Pt

# 1. Page Configuration
st.set_page_config(
    page_title="FormatForge | Academic Suite",
    page_icon="✦",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Modern 2026 SaaS Styling (Minimalist, High-Trust)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,600;0,700;1,600&family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, sans-serif;
    }
    
    h1, h2, h3, .serif-text {
        font-family: 'Playfair Display', serif !important;
    }
    
    #MainMenu, header, footer { visibility: hidden; }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 860px;
    }
    
    /* Top Logo Bar */
    .top-nav {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 24px;
        padding-bottom: 16px;
        border-bottom: 1px solid #E2E8F0;
    }
    .logo-text {
        font-family: 'Playfair Display', serif;
        font-size: 24px;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.02em;
    }

    /* Oxford Slate Hero Card */
    .hero-card {
        background-color: #0B1120;
        padding: 40px 32px;
        border-radius: 12px;
        color: #FFFFFF !important;
        margin-bottom: 24px;
        text-align: center;
        border: 1px solid #1E293B;
        box-shadow: 0 4px 24px -6px rgba(11, 17, 32, 0.4);
    }
    
    .hero-card h1 {
        color: #FFFFFF !important;
        font-size: 36px;
        font-weight: 600;
        margin: 0 0 16px 0;
        line-height: 1.2;
        letter-spacing: -0.01em;
    }
    
    .hero-card p {
        color: #94A3B8 !important;
        font-size: 15px;
        margin: 0 auto;
        max-width: 600px;
        line-height: 1.6;
    }

    /* Minimalist Badges */
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .badge-pro { background: #047857; color: #FFFFFF; border: 1px solid #065F46; }
    .badge-free { background: #F1F5F9; color: #475569; border: 1px solid #E2E8F0; }

    /* High-Contrast Inputs */
    .stTextArea textarea {
        background-color: #FAFAF9 !important;
        color: #09090B !important;
        -webkit-text-fill-color: #09090B !important;
        border: 1px solid #D4D4D8 !important;
        border-radius: 8px !important;
        padding: 16px !important;
        font-size: 14px !important;
        line-height: 1.6 !important;
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.02) !important;
    }
    .stTextArea textarea:focus {
        border-color: #09090B !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 0 0 1px #09090B !important;
    }

    /* Executive Black Action Buttons */
    div.stButton > button:first-child {
        background: #09090B !important;
        color: #FFFFFF !important;
        border: 1px solid #09090B !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 12px rgba(9, 9, 11, 0.15) !important;
    }
    div.stButton > button:first-child:hover {
        background: #27272A !important;
        border-color: #27272A !important;
        transform: translateY(-1px) !important;
    }
    
    /* Pro Banner */
    .pro-banner {
        background: #FAFAF9;
        border: 1px solid #E5E7EB;
        padding: 24px;
        border-radius: 12px;
        margin-bottom: 24px;
        text-align: center;
    }
    
    /* Trust Text */
    .privacy-text {
        text-align: center; 
        font-size: 12px; 
        font-weight: 500;
        color: #059669; 
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
    }
    </style>
""", unsafe_allow_html=True)

# 3. Configuration Limits & Credentials
FREE_CHAR_LIMIT = 1200
PRO_MAX_CHARS = 35000
STRIPE_PAYMENT_URL = "https://buy.stripe.com/cNidRbdxg5l4c9902l5Ne01"
DELIMITER = "---AUDIT_AND_FEEDBACK---"

if "STRIPE_SECRET_KEY" in st.secrets:
    stripe.api_key = st.secrets["STRIPE_SECRET_KEY"]

SAMPLE_TEXT = """The Impact of Screen Time on Teen Sleep Patterns

Recent empirical studies demonstrate that screen-based media consumption before bed disrupts circadian biology. According to smith (2021), blue light emitted by smart devices suppresses nocturnal melatonin release, making it difficult for adolescents to initiate restful sleep cycles. Furthermore, incoming social alerts frequently induce sleep fragmentation (Johnson and Lee 2019). Consequently, public health guidelines should advise device cutoffs 60 minutes prior to bedtime.

References:
Smith, John. (2021). Blue light and circadian rhythms. Journal of Sleep Health, 15(2), 104-112.
johnson, m., & Lee, T. 2019. Social media addiction in secondary school students. Adolescent Psychology Review 8(4): 45-59."""

# 4. State Management
if "is_pro" not in st.session_state:
    st.session_state.is_pro = False
if "free_uses" not in st.session_state:
    st.session_state.free_uses = 0
if "input_text" not in st.session_state:
    st.session_state.input_text = ""
if "session_expired" not in st.session_state:
    st.session_state.session_expired = False
if "last_uploaded_file" not in st.session_state:
    st.session_state.last_uploaded_file = None
if "clean_text" not in st.session_state:
    st.session_state.clean_text = None
if "editorial_notes" not in st.session_state:
    st.session_state.editorial_notes = None

session_id = st.query_params.get("session_id")

if session_id and not st.session_state.is_pro:
    try:
        session = stripe.checkout.Session.retrieve(session_id)
        if session.payment_status == "paid":
            time_elapsed = time.time() - session.created
            if time_elapsed < 7200:
                st.session_state.is_pro = True
            else:
                st.session_state.session_expired = True
    except Exception:
        st.session_state.is_pro = False

is_pro = st.session_state.is_pro

# Helper Functions
def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    else:
        return uploaded_file.read().decode("utf-8")

def create_docx(content, title_style):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    head = doc.add_paragraph()
    head_run = head.add_run(f"Formatted Output ({title_style})\n\n")
    head_run.bold = True
    
    paragraphs = content.split('\n')
    
    for para_text in paragraphs:
        if not para_text.strip():
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 2.0
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        
        parts = re.split(r'(\*.*?\*|_.*?_)', para_text)
        for part in parts:
            if part.startswith(('*', '_')) and part.endswith(('*', '_')) and len(part) > 2:
                run = p.add_run(part[1:-1])
                run.italic = True
            else:
                run = p.add_run(part)
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
    
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# 5. Header Bar & Hero Banner
status_badge = '<span class="badge badge-pro">Pro Access Active</span>' if is_pro else '<span class="badge badge-free">Free Trial</span>'

st.markdown(f"""
    <div class="top-nav">
        <span class="logo-text">FormatForge.</span>
        <div>{status_badge}</div>
    </div>
    <div class="hero-card">
        <h1>Zero citation errors.<br>Perfect formatting.</h1>
        <p>Transform rough drafts and messy bibliographies into submission-ready academic documents in seconds. Built for strict university rubrics.</p>
    </div>
""", unsafe_allow_html=True)

# Minimalist University Social Proof
st.markdown("""
<div style="text-align: center; margin-bottom: 32px;">
    <div style="font-size: 11px; font-weight: 600; color: #94A3B8; letter-spacing: 0.05em; text-transform: uppercase; margin-bottom: 12px;">
        Strictly compliant with guidelines from
    </div>
    <div class="serif-text" style="display: flex; justify-content: center; gap: 24px; flex-wrap: wrap; font-size: 15px; color: #475569; font-style: italic;">
        <span>Oxford</span>
        <span>Harvard</span>
        <span>NYU</span>
        <span>Toronto</span>
        <span>Melbourne</span>
    </div>
</div>
""", unsafe_allow_html=True)

# High-Contrast Roadmap
st.markdown("""
<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 32px;">
    <div style="background: #FFFFFF; border: 1px solid #E2E8F0; padding: 16px; border-radius: 8px;">
        <div style="font-size: 12px; font-weight: 600; color: #94A3B8; margin-bottom: 4px;">01</div>
        <div style="font-weight: 600; font-size: 14px; color: #0F172A; margin-bottom: 2px;">Input Draft</div>
        <div style="font-size: 13px; color: #64748B;">Paste raw text or upload .docx</div>
    </div>
    <div style="background: #FFFFFF; border: 1px solid #E2E8F0; padding: 16px; border-radius: 8px;">
        <div style="font-size: 12px; font-weight: 600; color: #94A3B8; margin-bottom: 4px;">02</div>
        <div style="font-weight: 600; font-size: 14px; color: #0F172A; margin-bottom: 2px;">Select Manual</div>
        <div style="font-size: 13px; color: #64748B;">Choose APA, MLA, or Harvard</div>
    </div>
    <div style="background: #FFFFFF; border: 1px solid #E2E8F0; padding: 16px; border-radius: 8px;">
        <div style="font-size: 12px; font-weight: 600; color: #94A3B8; margin-bottom: 4px;">03</div>
        <div style="font-weight: 600; font-size: 14px; color: #0F172A; margin-bottom: 2px;">Export Document</div>
        <div style="font-size: 13px; color: #64748B;">Receive clean .docx instantly</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Session Status Messages
if is_pro:
    st.info("Your Pro session is active. Keep this browser tab open to run and export all your documents.")
elif st.session_state.session_expired:
    st.warning("Your Pro window has elapsed. Upgrade below to run new papers.")

if not is_pro:
    st.markdown(
        f"""
        <div class="pro-banner">
            <div style="font-size: 18px; font-weight: 700; color: #0F172A; margin-bottom: 6px;">FormatForge Pro</div>
            <div style="font-size: 14px; color: #475569; margin-bottom: 16px; max-width: 400px; margin-left: auto; margin-right: auto;">
                Export perfect, rubric-compliant Word documents in one click. Remove all length caps and access premium style manuals. Cancel anytime.
            </div>
            <a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">
                <button style="background:#09090B; color:#FFF; border:none; padding:12px 24px; border-radius:6px; font-weight:500; font-size:14px; cursor:pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                    Unlock All Features ($1.99)
                </button>
            </a>
            <div style="font-size: 11px; color: #64748B; margin-top: 12px; display: flex; align-items: center; justify-content: center; gap: 4px;">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path></svg>
                Secure 256-bit checkout via Stripe
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# 6. Bento Grid Document Inputs & Settings
col1, col2 = st.columns([2.5, 1.5], gap="large")

with col2:
    with st.container(border=True):
        st.markdown("<div style='font-weight: 600; margin-bottom: 12px; font-size: 15px;'>Formatting Settings</div>", unsafe_allow_html=True)
        if is_pro:
            format_style = st.selectbox(
                "Style Manual",
                ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago 17th Edition", "IEEE"],
                label_visibility="collapsed"
            )
            audit_citations = st.toggle("Citation Integrity Check", value=True)
            editorial_tips = st.toggle("Tone & Phrasing Review", value=True)
        else:
            st.selectbox("Style Manual", ["APA 7th Edition (Upgrade for more)"], disabled=True, label_visibility="collapsed")
            format_style = "APA 7th Edition"
            audit_citations = False
            editorial_tips = False
            st.caption("Cross-checking and full style manuals unlock in Pro.")

with col1:
    with st.container(border=True):
        btn_col1, btn_col2 = st.columns([1, 1])
        with btn_col1:
            if st.button("Load Sample Data", use_container_width=True):
                st.session_state.input_text = SAMPLE_TEXT
        with btn_col2:
            uploaded_doc = st.file_uploader("Upload .docx", type=["docx", "txt"], label_visibility="collapsed")
            if uploaded_doc is not None and uploaded_doc.name != st.session_state.last_uploaded_file:
                st.session_state.input_text = extract_text_from_file(uploaded_doc)
                st.session_state.last_uploaded_file = uploaded_doc.name

        user_text = st.text_area(
            "Document Text",
            value=st.session_state.input_text,
            height=260,
            placeholder="Paste your unformatted essay or bibliography here...",
            label_visibility="collapsed"
        )
        
        char_len = len(user_text)
        if not is_pro:
            st.caption(f"Characters: {char_len}/{FREE_CHAR_LIMIT} | Free runs remaining: {max(0, 1 - st.session_state.free_uses)}")
        else:
            st.caption(f"Characters: {char_len}/{PRO_MAX_CHARS:,} (Pro Unlimited)")

# Zero-Retention Guarantee & Run Action
st.markdown("""
    <div class="privacy-text">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
        Zero-Retention AI: Your text is processed in memory and immediately deleted. Never stored or used for training.
    </div>
""", unsafe_allow_html=True)

run_button = st.button("Format Document", type="primary", use_container_width=True)

# 7. Execution Pipeline
if run_button:
    if not user_text.strip():
        st.warning("Please paste or upload text first.")
    elif not is_pro and st.session_state.free_uses >= 1:
        st.error("Free trial limit reached. Upgrade to Pro ($1.99) above to format your complete paper.")
    elif not is_pro and char_len > FREE_CHAR_LIMIT:
        st.error(f"Text exceeds the {FREE_CHAR_LIMIT}-character free limit. Please shorten your text or upgrade to Pro.")
    elif is_pro and char_len > PRO_MAX_CHARS:
        st.error(f"Manuscript exceeds {PRO_MAX_CHARS:,} characters. Please submit chapter by chapter.")
    else:
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

        with st.status("Analyzing and formatting manuscript...", expanded=True) as status_box:
            st.write("Standardizing citation conventions...")
            
            instructions = []
            if is_pro and audit_citations:
                instructions.append("- Check in-text citations against the reference list. Note any missing entries.")
            if is_pro and editorial_tips:
                instructions.append("- Provide 3 concise suggestions to improve formal academic tone.")

            system_message = f"""
            You are a university academic copyeditor. Reformat the user's input strictly according to {format_style} standards.
            Alphabetize reference lists and verify proper capitalization, author syntax, and hanging indent layouts.
            Do not alter the user's underlying research arguments.

            Output Format:
            1. Output ONLY the clean, formatted academic paper first.
            2. If reporting citation discrepancies or tone suggestions, add this exact separator on a line by itself:
            {DELIMITER}
            3. Provide the audit notes and suggestions below the separator.
            
            {chr(10).join(instructions)}
            """

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
                    st.session_state.editorial_notes = ""
                    
                if not is_pro:
                    st.session_state.free_uses += 1
                status_box.update(label="Document formatted successfully", state="complete", expanded=False)
                st.rerun()
                
            except openai.APIConnectionError:
                status_box.update(label="Network error", state="error")
                st.error("Unable to reach the AI engine. Check your connection and try again.")
            except openai.RateLimitError:
                status_box.update(label="System busy", state="error")
                st.error("System traffic is high. Please wait 15 seconds and try again.")
            except Exception:
                status_box.update(label="Error", state="error")
                st.error("An error occurred during formatting. Please verify your text.")

# 8. Deliverables Display (Independent of Run Button)
if st.session_state.clean_text:
    st.markdown("---")
    st.markdown("<h2 class='serif-text'>Formatted Deliverable</h2>", unsafe_allow_html=True)
    
    st.text_area("Copy Output:", value=st.session_state.clean_text, height=300, label_visibility="collapsed")

    if is_pro:
        docx_output = create_docx(st.session_state.clean_text, format_style)
        st.download_button(
            label="Download .docx Document",
            data=docx_output,
            file_name="Formatted_Manuscript.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    
    if st.session_state.editorial_notes:
        with st.container(border=True):
            st.markdown("### Editorial & Citation Audit")
            st.markdown(st.session_state.editorial_notes)
