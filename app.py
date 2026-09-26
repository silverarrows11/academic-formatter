import io
import time
import streamlit as st
import openai
import stripe
from docx import Document
from docx.shared import Inches, Pt

# 1. Page Configuration
st.set_page_config(
    page_title="FormatForge | Academic Suite",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Mobile-Optimized SaaS Styling with Contrast Fixes
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, sans-serif;
    }
    
    #MainMenu, header, footer { visibility: hidden; }
    
    /* Responsive container padding */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 800px;
    }
    
    /* Top Logo Bar - Adapts to light/dark themes */
    .top-nav {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 16px;
    }
    .logo-text {
        font-size: 22px;
        font-weight: 800;
        letter-spacing: -0.02em;
        /* Gradient text to ensure visibility in both Light & Dark modes */
        background: linear-gradient(135deg, #6366F1 0%, #A855F7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }

    /* Mobile-first Hero Card */
    .hero-card {
        background: linear-gradient(135deg, #1E1B4B 0%, #312E81 50%, #4338CA 100%);
        padding: 24px 20px;
        border-radius: 16px;
        color: #FFFFFF !important;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(49, 46, 129, 0.4);
    }
    
    .hero-card h1 {
        color: #FFFFFF !important;
        font-size: 24px;
        font-weight: 800;
        margin: 10px 0 8px 0;
        line-height: 1.3;
    }
    
    .hero-card p {
        color: #E0E7FF !important;
        font-size: 13.5px;
        margin: 0;
        line-height: 1.5;
    }

    /* Feature Pills */
    .pill-container {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-bottom: 8px;
    }
    .feature-pill {
        display: inline-block;
        background: rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(6px);
        border: 1px solid rgba(255, 255, 255, 0.25);
        color: #FFFFFF !important;
        padding: 3px 9px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 600;
    }

    /* TEXT AREA VISIBILITY FIX: Solid high-contrast colors */
    .stTextArea textarea {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        border: 1.5px solid #CBD5E1 !important;
        border-radius: 12px !important;
        padding: 12px !important;
        font-size: 14px !important;
        line-height: 1.5 !important;
    }
    .stTextArea textarea:focus {
        border-color: #6366F1 !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.2) !important;
    }
    .stTextArea textarea::placeholder {
        color: #64748B !important;
        -webkit-text-fill-color: #64748B !important;
    }

    /* Mobile Buttons */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #4F46E5 0%, #6366F1 100%);
        color: #FFFFFF !important;
        border: none;
        border-radius: 12px;
        padding: 12px 20px;
        font-weight: 700;
        font-size: 15px;
        box-shadow: 0 4px 14px rgba(79, 70, 229, 0.35);
        transition: all 0.15s ease-in-out;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(79, 70, 229, 0.5);
    }
    
    .pro-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border: 1px solid #334155;
        color: #FFFFFF !important;
        padding: 18px;
        border-radius: 14px;
        margin-bottom: 1.5rem;
    }
    
    .trust-card {
        background: rgba(241, 245, 249, 0.6);
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 16px;
        margin-top: 1.5rem;
        font-size: 12px;
        color: #64748B;
        line-height: 1.5;
    }

    .badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .badge-pro { background: #10B981; color: #FFFFFF; }
    .badge-free { background: #334155; color: #E2E8F0; }

    /* Mobile fine-tuning */
    @media (max-width: 640px) {
        .hero-card h1 { font-size: 20px; }
        .hero-card { padding: 18px 16px; }
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

# 4. State Management & Dynamic Stripe Session Verification
if "is_pro" not in st.session_state:
    st.session_state.is_pro = False
if "free_uses" not in st.session_state:
    st.session_state.free_uses = 0
if "input_text" not in st.session_state:
    st.session_state.input_text = ""
if "session_expired" not in st.session_state:
    st.session_state.session_expired = False

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
    
    body = doc.add_paragraph()
    body_run = body.add_run(content)
    body_run.font.name = "Times New Roman"
    body_run.font.size = Pt(12)
    body.paragraph_format.line_spacing = 2.0
    
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# 5. Header Bar & Brand Logo
status_badge = '<span class="badge badge-pro">PRO ACTIVE</span>' if is_pro else '<span class="badge badge-free">FREE TRIAL</span>'

st.markdown(f"""
    <div class="top-nav">
        <span class="logo-text">🎓 FormatForge</span>
        <div>{status_badge}</div>
    </div>
    <div class="hero-card">
        <div class="pill-container">
            <span class="feature-pill">⚡ Auto-Hanging Indents</span>
            <span class="feature-pill">🔍 Audit Missing Sources</span>
            <span class="feature-pill">📄 Word .docx Export</span>
        </div>
        <h1>Never lose grade points to citation rules.</h1>
        <p>Paste rough notes, messy links, or full bibliographies. Transform them into submission-ready academic formats in seconds.</p>
    </div>
""", unsafe_allow_html=True)

if is_pro:
    st.info("💡 **Pro Access Active (2-Hour Window):** Bookmark this tab to keep working on your papers.")
elif st.session_state.session_expired:
    st.warning("⚠️ **Pass Expired:** Your 2-hour Pro window has ended. Upgrade below to run new papers.")

# 6. Pro Banner (Free Tier Only)
if not is_pro:
    st.markdown(
        f"""
        <div class="pro-banner">
            <div style="font-size: 16px; font-weight: 700; margin-bottom: 4px;">Unlock FormatForge Pro — $1.99</div>
            <div style="font-size: 13px; color: #CBD5E1; margin-bottom: 12px;">Remove length caps, export pre-formatted Word documents, and access MLA, Chicago, Harvard, and IEEE styles.</div>
            <a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">
                <button style="background:#4F46E5; color:#FFF; border:none; padding:8px 16px; border-radius:8px; font-weight:600; font-size:13px; cursor:pointer;">
                    Unlock All Features ($1.99)
                </button>
            </a>
        </div>
        """,
        unsafe_allow_html=True
    )

# 7. Document Inputs & Settings
tab_input, tab_settings = st.tabs(["📝 Document Editor", "⚙️ Options & Style Manuals"])

with tab_settings:
    st.subheader("Formatting Parameters")
    if is_pro:
        format_style = st.selectbox(
            "Style Manual",
            ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago 17th Edition (Author-Date)", "IEEE"]
        )
        audit_citations = st.toggle("Citation Integrity Check (Cross-reference in-text sources with references)", value=True)
        editorial_tips = st.toggle("Academic Tone & Phrasing Review", value=True)
    else:
        st.selectbox("Style Manual", ["APA 7th Edition (Upgrade for MLA, Chicago, Harvard)"], disabled=True)
        format_style = "APA 7th Edition"
        audit_citations = False
        editorial_tips = False
        st.caption("🔒 Cross-checking and full style manuals are unlocked in Pro.")

with tab_input:
    btn_col1, btn_col2 = st.columns([1, 1])
    with btn_col1:
        if st.button("📄 Load Sample Essay", use_container_width=True):
            st.session_state.input_text = SAMPLE_TEXT
    with btn_col2:
        uploaded_doc = st.file_uploader("Upload .docx / .txt", type=["docx", "txt"], label_visibility="collapsed")
        if uploaded_doc is not None:
            st.session_state.input_text = extract_text_from_file(uploaded_doc)

    user_text = st.text_area(
        "Document Text",
        value=st.session_state.input_text,
        height=220,
        placeholder="Paste your unformatted essay or bibliography here...",
        label_visibility="collapsed"
    )
    
    char_len = len(user_text)

    if not is_pro:
        st.caption(f"Characters: {char_len}/{FREE_CHAR_LIMIT} | Free runs remaining: {max(0, 1 - st.session_state.free_uses)}")
    else:
        st.caption(f"Characters: {char_len}/{PRO_MAX_CHARS:,} (Pro Unlimited)")

    run_button = st.button("✨ Format & Audit Document", type="primary", use_container_width=True)

# 8. Execution Pipeline
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

        with st.status("Analyzing and formatting...", expanded=True) as status_box:
            st.write("🔍 Standardizing citation conventions...")
            
            instructions = []
            if is_pro and audit_citations:
                instructions.append("- Check in-text citations against the reference list. Note any missing entries.")
            if is_pro and editorial_tips:
                instructions.append("- Provide 3 concise suggestions to improve formal academic tone and eliminate passive voice.")

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
                if not is_pro:
                    st.session_state.free_uses += 1
                status_box.update(label="Complete!", state="complete", expanded=False)
            except openai.APIConnectionError:
                status_box.update(label="Network error", state="error")
                st.error("Unable to reach the AI engine. Check your connection and try again.")
                st.stop()
            except openai.RateLimitError:
                status_box.update(label="System busy", state="error")
                st.error("System traffic is high. Please wait 15 seconds and try again.")
                st.stop()
            except Exception:
                status_box.update(label="Error", state="error")
                st.error("An error occurred during formatting. Please verify your text.")
                st.stop()

        # 9. Deliverables Display
        st.markdown("### Formatted Deliverable")
        
        if DELIMITER in output_content:
            clean_text, editorial_notes = output_content.split(DELIMITER, 1)
            clean_text = clean_text.strip()
            editorial_notes = editorial_notes.strip()
        else:
            clean_text = output_content.strip()
            editorial_notes = ""

        st.text_area("Copy Output:", value=clean_text, height=260)

        if is_pro:
            docx_output = create_docx(clean_text, format_style)
            st.download_button(
                label="📥 Download .docx File (Double-Spaced & 1-in Margins)",
                data=docx_output,
                file_name="formatted_paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        
        if editorial_notes:
            st.markdown("### Editorial & Citation Audit")
            st.markdown(editorial_notes)

# 10. Privacy & Integrity Notice
st.markdown(
    """
    <div class="trust-card">
        <b>Data Privacy:</b> Documents are processed in volatile memory and are never saved or trained on. Payments are handled via Stripe 256-bit encryption.<br>
        <b>Academic Integrity:</b> FormatForge formats layout and citations only. It does not generate arguments or write papers for students.
    </div>
    """,
    unsafe_allow_html=True
)
