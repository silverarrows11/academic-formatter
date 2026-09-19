import io
import streamlit as st
import openai
from docx import Document
from docx.shared import Inches, Pt

# 1. Page Configuration
st.set_page_config(
    page_title="FormatForge | Academic Citation & Formatting Suite",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Modern 2026 SaaS CSS Theme
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0F172A;
    }

    /* Hide default Streamlit chrome */
    #MainMenu, header, footer {visibility: hidden;}
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }

    /* Modern Card Layouts */
    .saas-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -2px rgba(0, 0, 0, 0.04);
        margin-bottom: 1.5rem;
    }

    /* Modern Text Areas & Inputs */
    .stTextArea textarea {
        border-radius: 12px !important;
        border: 1.5px solid #E2E8F0 !important;
        font-size: 14px !important;
        padding: 12px 16px !important;
        transition: all 0.2s ease;
    }
    .stTextArea textarea:focus {
        border-color: #3B82F6 !important;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15) !important;
    }

    /* Primary CTA Buttons */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.4rem !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2) !important;
        transition: transform 0.1s ease, box-shadow 0.1s ease !important;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.3) !important;
    }

    /* Status Badges */
    .badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 0.02em;
    }
    .badge-free { background-color: #F1F5F9; color: #475569; }
    .badge-pro { background-color: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; }

    /* Trust Footer Banner */
    .trust-footer {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 14px 20px;
        border-radius: 12px;
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        margin-top: 1rem;
        font-size: 12px;
        color: #64748B;
    }
    </style>
""", unsafe_allow_html=True)

# Configuration — ADD YOUR SECRETS & LINKS
FREE_CHAR_LIMIT = 1200
STRIPE_PAYMENT_URL = "https://buy.stripe.com/3cIaEZ64O7tc6OP16p5Ne00"  # <-- INSERT YOUR STRIPE LINK
VALID_PRO_CODE = "PRO2026"                                          # <-- INSERT YOUR CODE

# Session State Setup
if "show_passcode_input" not in st.session_state:
    st.session_state.show_passcode_input = False
if "is_pro" not in st.session_state:
    st.session_state.is_pro = False
if "free_uses" not in st.session_state:
    st.session_state.free_uses = 0

# --- SIDEBAR: ACCOUNT & SUBSCRIPTION ---
with st.sidebar:
    st.markdown("### 🎓 Account Status")

    if st.session_state.is_pro:
        st.markdown('<span class="badge badge-pro">● PRO UNLOCKED</span>', unsafe_allow_html=True)
        st.caption("All premium styles and document exports enabled.")
    else:
        st.markdown('<span class="badge badge-free">STANDARD PLAN</span>', unsafe_allow_html=True)
        st.caption(f"Trial Tier • {FREE_CHAR_LIMIT} Chars • APA 7th Only")
        
        st.markdown(
            f"""
            <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:14px; padding:18px; margin: 18px 0;">
                <div style="font-weight:700; font-size:16px; color:#0F172A; margin-bottom:4px;">Full Access Pass</div>
                <div style="font-size:22px; font-weight:800; color:#1E293B; margin-bottom:12px;">€1.00 <span style="font-size:12px; font-weight:400; color:#64748B;">/ one-off</span></div>
                <ul style="font-size:13px; color:#475569; padding-left:18px; line-height:1.6; margin:0 0 16px 0;">
                    <li>Unlimited runs & document length</li>
                    <li>Download standardized <b>.docx</b></li>
                    <li>MLA 9, Harvard, Chicago, IEEE</li>
                    <li>Missing Citation Cross-Check</li>
                    <li>Writing & Tone Critique</li>
                </ul>
                <a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">
                    <button style="width:100%; background:#0F172A; color:#FFFFFF; border:none; padding:10px 0; border-radius:8px; font-weight:600; font-size:13px; cursor:pointer;">
                        Unlock Pro Access
                    </button>
                </a>
            </div>
            """,
            unsafe_allow_html=True
        )

        if not st.session_state.show_passcode_input:
            if st.button("Already purchased? Enter passcode", use_container_width=True):
                st.session_state.show_passcode_input = True
                st.rerun()
        else:
            code_input = st.text_input("Enter Passcode:", type="password", placeholder="Enter your key...")
            if st.button("Confirm Code", use_container_width=True):
                if code_input == VALID_PRO_CODE:
                    st.session_state.is_pro = True
                    st.success("Pro pass validated.")
                    st.rerun()
                else:
                    st.error("Invalid passcode.")

is_pro = st.session_state.is_pro

# --- MAIN APP HEADER ---
st.title("FormatForge Academic")
st.markdown("Automated publication formatting, bibliography audit, and stylistic refinement.")

# --- OPTIONS / CONFIGURATION BAR ---
opt_col1, opt_col2, opt_col3 = st.columns([2, 1.5, 1.5])

with opt_col1:
    if is_pro:
        format_style = st.selectbox(
            "Target Style Guide",
            ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago 17th Edition", "IEEE"]
        )
    else:
        st.selectbox("Target Style Guide", ["APA 7th Edition (Free Standard)"], disabled=True)
        format_style = "APA 7th Edition"

with opt_col2:
    audit_citations = st.checkbox("Audit Citations", value=is_pro, disabled=not is_pro)

with opt_col3:
    provide_critique = st.checkbox("Writing Critique & Tips", value=is_pro, disabled=not is_pro)

# --- WORKSPACE: TWO-COLUMN LAYOUT ---
col_input, col_output = st.columns(2, gap="large")

with col_input:
    st.subheader("Source Input")
    user_text = st.text_area(
        "Paste unformatted draft or references:",
        height=320,
        placeholder="Paste your unformatted essay text, citations, or references here..."
    )
    char_count = len(user_text)

    # Usage Indicator
    if not is_pro:
        pct = min(1.0, char_count / FREE_CHAR_LIMIT)
        st.progress(pct)
        runs_left = max(0, 1 - st.session_state.free_uses)
        st.caption(f"Chars: {char_count}/{FREE_CHAR_LIMIT} | Trial runs remaining: {runs_left}")
    else:
        st.caption(f"Length: {char_count} characters (Pro Unlimited)")

    process_btn = st.button("Process Document", type="primary", use_container_width=True)

# Helper function to generate Word document
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

with col_output:
    st.subheader("Formatted Deliverable")

    if process_btn:
        if not user_text.strip():
            st.warning("Please input document content before running.")
        elif not is_pro and st.session_state.free_uses >= 1:
            st.error("Free trial completed. Unlock Pro (€1.00) in the sidebar for unlimited usage.")
        elif not is_pro and char_count > FREE_CHAR_LIMIT:
            st.error(f"Character limit reached ({char_count}/{FREE_CHAR_LIMIT}). Upgrade to Pro to continue.")
        else:
            client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

            extra_prompts = []
            if is_pro and audit_citations:
                extra_prompts.append("- Check in-text citations against the reference list. Note discrepancies under a clear '=== CITATION AUDIT REPORT ===' header.")
            if is_pro and provide_critique:
                extra_prompts.append("- Provide 3-5 specific, bulleted recommendations to strengthen academic tone, clarity, and conciseness under an '=== EDITORIAL FEEDBACK & REVISION TIPS ===' header.")

            prompt_additions = "\n".join(extra_prompts)

            system_prompt = f"""
            You are a senior academic copyeditor.
            Reformat the provided text strictly according to {format_style} standards.
            Preserve the author's core thesis and voice while ensuring publication-grade adherence to citation rules.
            {prompt_additions}
            Structure your output cleanly with Markdown headings.
            """

            with st.spinner("Processing document to standard..."):
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_text}
                    ]
                )

                if not is_pro:
                    st.session_state.free_uses += 1

                output = response.choices[0].message.content

                if is_pro and (provide_critique or audit_citations):
                    tab1, tab2 = st.tabs(["📄 Clean Paper", "🔍 Audit & Feedback"])
                    with tab1:
                        formatted_only = output.split("===")[0].strip()
                        st.text_area("Final Text", value=formatted_only, height=280)
                        docx_file = create_docx(formatted_only, format_style)
                        st.download_button(
                            label="📥 Download Submission-Ready .docx",
                            data=docx_file,
                            file_name="academic_formatted_paper.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    with tab2:
                        st.markdown(output)
                else:
                    st.text_area("Formatted Text", value=output, height=320)
    else:
        st.info("Your formatted document and citation reports will appear here once processed.")

# --- TRUST & SECURITY FOOTER ---
st.markdown("---")
st.markdown(
    """
    <div class="trust-footer">
        <div>🔒 <b>Data Protection:</b> Submissions are processed in memory and never stored or used to train models.</div>
        <div>💳 <b>Payments:</b> Encrypted 256-bit checkout via Stripe.</div>
        <div>🎓 <b>Compliance:</b> Conforms to official APA 7th, MLA 9th, and Chicago 17th manuals.</div>
    </div>
    """,
    unsafe_allow_html=True
)
