import io
import streamlit as st
import openai
from docx import Document
from docx.shared import Inches, Pt

# 1. Page Configuration
st.set_page_config(
    page_title="FormatForge | Academic Suite",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Modern 2026 Mobile-Optimized SaaS Styling
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
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
    .trust-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 18px;
        margin-top: 1.5rem;
        font-size: 13px;
        color: #475569;
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

# 3. Configuration — INSERT YOUR STRIPE PAYMENT LINK ON LINE 62
FREE_CHAR_LIMIT = 1200
STRIPE_PAYMENT_URL = "https://buy.stripe.com/3cIaEZ64O7tc6OP16p5Ne00"  # <-- LINE 62

SAMPLE_TEXT = """The Impact of Screen Time on Teen Sleep Patterns

Recent empirical studies demonstrate that screen-based media consumption before bed disrupts circadian biology. According to smith (2021), blue light emitted by smart devices suppresses nocturnal melatonin release, making it difficult for adolescents to initiate restful sleep cycles. Furthermore, incoming social alerts frequently induce sleep fragmentation (Johnson and Lee 2019). Consequently, public health guidelines should advise device cutoffs 60 minutes prior to bedtime.

References:
Smith, John. (2021). Blue light and circadian rhythms. Journal of Sleep Health, 15(2), 104-112.
johnson, m., & Lee, T. 2019. Social media addiction in secondary school students. Adolescent Psychology Review 8(4): 45-59."""

# 4. State Management & Auto-Unlock via Stripe Redirect
if "is_pro" not in st.session_state:
    st.session_state.is_pro = False
if "free_uses" not in st.session_state:
    st.session_state.free_uses = 0
if "input_text" not in st.session_state:
    st.session_state.input_text = ""

# Check for automatic return from Stripe (?pro=true)
if st.query_params.get("pro") == "true":
    st.session_state.is_pro = True

is_pro = st.session_state.is_pro

# Helper: Extract text from uploaded files (.txt or .docx)
def extract_text_from_file(uploaded_file):
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    else:
        return uploaded_file.read().decode("utf-8")

# Helper: Build Microsoft Word output
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

# 5. Header Section
head_col1, head_col2 = st.columns([3, 1])
with head_col1:
    st.title("FormatForge")
    st.caption("AI Academic Styling, Citation Auditing & Paper Verification")
with head_col2:
    if is_pro:
        st.markdown('<div style="text-align:right; margin-top:20px;"><span class="badge badge-pro">PRO ACTIVE</span></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="text-align:right; margin-top:20px;"><span class="badge badge-free">FREE TRIAL</span></div>', unsafe_allow_html=True)

# 6. Pro Upgrade Banner (Only shows if user is on the Free tier)
if not is_pro:
    st.markdown(
        f"""
        <div class="pro-banner">
            <div style="font-size: 17px; font-weight: 700; margin-bottom: 4px;">Upgrade to FormatForge Pro — Only €1.00</div>
            <div style="font-size: 13px; opacity: 0.85; margin-bottom: 12px;">Remove character limits, export pre-formatted Word documents, and access MLA, Chicago, and Harvard formatting.</div>
            <a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">
                <button style="background:#2563EB; color:#FFF; border:none; padding:8px 18px; border-radius:8px; font-weight:600; font-size:13px; cursor:pointer;">
                    Unlock All Features (€1.00)
                </button>
            </a>
        </div>
        """,
        unsafe_allow_html=True
    )

# 7. Main Tabs (Clean layout on both phone and laptop)
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
    # Action row for quick testing and file upload
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
        st.caption(f"Characters: {char_len} (Unlimited Pro Access)")

    run_button = st.button("✨ Format & Standardize", type="primary", use_container_width=True)

# 8. Execution & Stepped Progress
if run_button:
    if not user_text.strip():
        st.warning("Please paste or upload text first.")
    elif not is_pro and st.session_state.free_uses >= 1:
        st.error("Free trial limit reached. Upgrade to Pro (€1.00) above for unlimited usage.")
    elif not is_pro and char_len > FREE_CHAR_LIMIT:
        st.error(f"Text exceeds the {FREE_CHAR_LIMIT}-character limit. Shorten your input or upgrade to Pro.")
    else:
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

        # Dynamic Status Container
        with st.status("Analyzing and formatting...", expanded=True) as status_box:
            st.write("🔍 Parsing text and checking citation references...")
            
            instructions = []
            if is_pro and audit_citations:
                instructions.append("- Cross-examine all in-text citations against the bibliography. List missing items under '=== CITATION INTEGRITY REPORT ==='.")
            if is_pro and editorial_tips:
                instructions.append("- Provide 3-4 bulleted suggestions to eliminate informal language or passive voice under '=== ACADEMIC WRITING CRITIQUE ==='.")

            system_message = f"""
            You are a university academic copyeditor. Reformat the user's input strictly according to {format_style} standards.
            Alphabetize reference lists and verify proper capitalization and author layout.
            {chr(10).join(instructions)}
            Do not alter the user's argument or core meaning.
            """

            st.write(f"📐 Applying {format_style} conventions...")
            
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": user_text}
                ]
            )
            
            output_content = response.choices[0].message.content
            
            if not is_pro:
                st.session_state.free_uses += 1
                
            status_box.update(label="Document processed successfully!", state="complete", expanded=False)

        # 9. Deliverables Section
        st.markdown("### Formatted Deliverable")
        clean_text = output_content.split("===")[0].strip()
        st.text_area("Copy Formatted Text:", value=clean_text, height=260)

        # Pro Download Button
        if is_pro:
            docx_output = create_docx(clean_text, format_style)
            st.download_button(
                label="📥 Download Standardized .docx File",
                data=docx_output,
                file_name="formatted_paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        
        # Display Audit or Critique tabs if Pro
        if "===" in output_content:
            st.markdown("### Editorial Reports")
            st.markdown(output_content)

# 10. Trust and Security Footer
st.markdown(
    """
    <div class="trust-card">
        <b>Data Protection & Security:</b> Documents are processed in volatile memory and never stored, indexed, or shared. 
        Transactions are securely handled by Stripe with 256-bit encryption.
    </div>
    """,
    unsafe_allow_html=True
)
