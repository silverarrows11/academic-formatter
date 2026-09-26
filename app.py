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

# 2. Modern SaaS Styling
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, sans-serif;
        color: #0F172A;
    }
    #MainMenu, header, footer { visibility: hidden; }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 820px;
    }
    
    /* Sleek gradient hero card */
    .hero-card {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 50%, #EC4899 100%);
        padding: 28px 24px;
        border-radius: 18px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(99, 102, 241, 0.3);
    }
    
    /* Modern glowing action button */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #4F46E5 0%, #6366F1 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 12px 24px;
        font-weight: 700;
        font-size: 15px;
        letter-spacing: -0.01em;
        box-shadow: 0 4px 14px rgba(79, 70, 229, 0.4);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(79, 70, 229, 0.6);
    }
    
    /* Rounded text input */
    .stTextArea textarea {
        border-radius: 14px !important;
        border: 1.5px solid #E2E8F0 !important;
        padding: 14px !important;
        font-size: 14px !important;
        background-color: #FAFAFA !important;
    }
    .stTextArea textarea:focus {
        border-color: #6366F1 !important;
        box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15) !important;
    }
    
    /* Feature pills */
    .feature-pill {
        display: inline-flex;
        align-items: center;
        background: rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255, 255, 255, 0.3);
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 600;
        margin-right: 6px;
    }
    
    .trust-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 20px;
        margin-top: 1.5rem;
        font-size: 13px;
        color: #475569;
        line-height: 1.6;
    }
    .pro-banner {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        color: #FFFFFF;
        padding: 20px;
        border-radius: 14px;
        margin-bottom: 1.5rem;
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
    .badge-free { background: #E2E8F0; color: #475569; }
    </style>
""", unsafe_allow_html=True)

# 3. Configuration Limits & Credentials
FREE_CHAR_LIMIT = 1200
PRO_MAX_CHARS = 35000
STRIPE_PAYMENT_URL = "https://buy.stripe.com/cNidRbdxg5l4c9902l5Ne01"
DELIMITER = "---AUDIT_AND_FEEDBACK---"

# Initialize Stripe API Key safely
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

# Validate dynamic Stripe session parameter (?session_id=cs_...)
session_id = st.query_params.get("session_id")

if session_id and not st.session_state.is_pro:
    try:
        session = stripe.checkout.Session.retrieve(session_id)
        if session.payment_status == "paid":
            # 2-hour window validation (7200 seconds)
            time_elapsed = time.time() - session.created
            if time_elapsed < 7200:
                st.session_state.is_pro = True
            else:
                st.session_state.session_expired = True
    except Exception:
        st.session_state.is_pro = False

is_pro = st.session_state.is_pro

if is_pro:
    st.info("💡 **Pro Access Active (2-Hour Window):** Keep this browser tab open to run and export all your documents.")
elif st.session_state.session_expired:
    st.warning("⚠️ **Pass Expired:** Your 2-hour Pro pass has elapsed. Please purchase a new pass to process extended manuscripts.")

# Helpers for File Parsing & Word Generation
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

# 5. Header Bar & Hero Banner
status_badge = '<span class="badge badge-pro">PRO ACTIVE</span>' if is_pro else '<span class="badge badge-free">FREE TRIAL</span>'

st.markdown(f"""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
        <span style="font-size: 20px; font-weight: 800; color: #0F172A; letter-spacing: -0.02em;">🎓 FormatForge</span>
        <div>{status_badge}</div>
    </div>
    <div class="hero-card">
        <div style="display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap;">
            <span class="feature-pill">⚡ Instant Hanging Indents</span>
            <span class="feature-pill">🎯 Fixes Missed Citations</span>
            <span class="feature-pill">📑 Word .docx Ready</span>
        </div>
        <h1 style="color: white; font-size: 26px; font-weight: 800; margin: 0 0 8px 0; letter-spacing: -0.02em; line-height: 1.25;">
            Never lose grade points to citation rules again.
        </h1>
        <p style="color: rgba(255, 255, 255, 0.9); font-size: 14px; margin: 0; line-height: 1.5;">
            Paste messy text, sketchy links, or raw bibliographies. Get clean APA, MLA, or Harvard styling formatted in seconds.
        </p>
    </div>
""", unsafe_allow_html=True)

# 6. Pro Upgrade Banner (Free Tier only)
if not is_pro:
    st.markdown(
        f"""
        <div class="pro-banner">
            <div style="font-size: 17px; font-weight: 700; margin-bottom: 4px;">Upgrade to FormatForge Pro — Only $1.99</div>
            <div style="font-size: 13px; opacity: 0.85; margin-bottom: 12px;">Remove character limits, export pre-formatted Word documents, and access MLA, Chicago, and Harvard formatting.</div>
            <a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">
                <button style="background:#2563EB; color:#FFF; border:none; padding:8px 18px; border-radius:8px; font-weight:600; font-size:13px; cursor:pointer;">
                    Unlock All Features ($1.99)
                </button>
            </a>
        </div>
        """,
        unsafe_allow_html=True
    )

# 7. Layout Tabs
tab_input, tab_settings = st.tabs(["📝 Document Editor", "⚙️ Options & Citation Rules"])

with tab_settings:
    st.subheader("Formatting Parameters")
    if is_pro:
        format_style = st.selectbox(
            "Style Manual",
            ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago 17th Edition (Author-Date)", "IEEE"]
        )
        audit_citations = st.toggle("Citation Integrity Check (Cross-reference body against references)", value=True)
        editorial_tips = st.toggle("Academic Tone & Phrasing Review", value=True)
    else:
        st.selectbox("Style Manual", ["APA 7th Edition (Upgrade for MLA, Chicago, Harvard)"], disabled=True)
        format_style = "APA 7th Edition"
        audit_citations = False
        editorial_tips = False
        st.caption("🔒 Cross-checking and advanced style guides are reserved for Pro users.")

with tab_input:
    action_col1, action_col2 = st.columns([1, 1])
    with action_col1:
        if st.button("📄 Pre-fill Sample Draft", use_container_width=True):
            st.session_state.input_text = SAMPLE_TEXT
    with action_col2:
        uploaded_doc = st.file_uploader("Or upload .docx / .txt", type=["docx", "txt"], label_visibility="collapsed")
        if uploaded_doc is not None:
            st.session_state.input_text = extract_text_from_file(uploaded_doc)

    user_text = st.text_area(
        "Document Text",
        value=st.session_state.input_text,
        height=240,
        placeholder="Paste your essay, article, or references here...",
        label_visibility="collapsed"
    )
    
    char_len = len(user_text)

    # Usage Indicator
    if not is_pro:
        st.caption(f"Characters: {char_len}/{FREE_CHAR_LIMIT} | Free runs remaining: {max(0, 1 - st.session_state.free_uses)}")
    else:
        st.caption(f"Characters: {char_len}/{PRO_MAX_CHARS:,} (Pro Unlimited)")

    run_button = st.button("✨ Format & Standardize", type="primary", use_container_width=True)

# 8. Execution Logic & Guardrails
if run_button:
    if not user_text.strip():
        st.warning("Please paste or upload text first.")
    elif not is_pro and st.session_state.free_uses >= 1:
        st.error("Free trial limit reached. Upgrade to Pro ($1.99) above for unlimited usage.")
    elif not is_pro and char_len > FREE_CHAR_LIMIT:
        st.error(f"Text exceeds the {FREE_CHAR_LIMIT}-character limit. Shorten your input or upgrade to Pro.")
    elif is_pro and char_len > PRO_MAX_CHARS:
        st.error(f"Input exceeds the safety ceiling of {PRO_MAX_CHARS:,} characters (~6,000 words). Please process longer manuscripts in separate sections.")
    else:
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

        with st.status("Analyzing and formatting...", expanded=True) as status_box:
            st.write("🔍 Parsing text and checking citation references...")
            
            instructions = []
            if is_pro and audit_citations:
                instructions.append("- Cross-examine all in-text citations against the bibliography. List missing items.")
            if is_pro and editorial_tips:
                instructions.append("- Provide 3-4 bulleted suggestions to eliminate informal language or passive voice.")

            system_message = f"""
            You are a university academic copyeditor. Reformat the user's input strictly according to {format_style} standards.
            Alphabetize reference lists and verify proper capitalization and author layout.
            Do not alter the user's argument or core meaning.

            Formatting Rules:
            1. Output ONLY the clean, formatted academic paper first.
            2. If reporting citation integrity issues or editorial critiques, place this exact separator on its own line:
            {DELIMITER}
            3. Put all audit notes and suggestions below that separator under clear headings.
            
            {chr(10).join(instructions)}
            """

            st.write(f"📐 Applying {format_style} conventions...")
            
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
                status_box.update(label="Document processed successfully!", state="complete", expanded=False)
            except openai.APIConnectionError:
                status_box.update(label="Network error", state="error")
                st.error("Could not reach the AI formatting engine. Please check your connection and try again.")
                st.stop()
            except openai.RateLimitError:
                status_box.update(label="System busy", state="error")
                st.error("The system is receiving high traffic right now. Please wait 15 seconds and click 'Format' again.")
                st.stop()
            except Exception:
                status_box.update(label="Processing error", state="error")
                st.error("An unexpected error occurred while formatting. Please verify your input and try again.")
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

        st.text_area("Copy Formatted Text:", value=clean_text, height=260)

        if is_pro:
            docx_output = create_docx(clean_text, format_style)
            st.download_button(
                label="📥 Download Standardized .docx File",
                data=docx_output,
                file_name="academic_formatted_paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        
        if editorial_notes:
            st.markdown("### Editorial & Audit Reports")
            st.markdown(editorial_notes)

# 10. Trust, Security & Academic Integrity Footer
st.markdown(
    """
    <div class="trust-card">
        <b>Data Protection & Security:</b> Documents are processed in volatile memory and never stored, indexed, or shared. 
        Transactions are securely processed with 256-bit Stripe encryption.<br><br>
        <b>Academic Integrity Notice:</b> FormatForge is strictly an editorial and citation standardization tool. 
        It does not research, ghostwrite, or generate original arguments on behalf of students.
    </div>
    """,
    unsafe_allow_html=True
)
