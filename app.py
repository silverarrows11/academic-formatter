import io
import streamlit as st
import openai
from docx import Document
from docx.shared import Inches, Pt

st.set_page_config(
    page_title="FormatForge | Academic Suite",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Clean CSS Styling
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    .stTextArea textarea {
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        font-size: 14px;
    }
    .stButton button {
        border-radius: 10px;
        font-weight: 600;
        transition: all 0.2s ease-in-out;
    }
    .pricing-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 18px;
        border-radius: 14px;
        margin-bottom: 20px;
    }
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .badge-free { background-color: #E2E8F0; color: #475569; }
    .badge-pro { background-color: #DCFCE7; color: #15803D; }
    </style>
""", unsafe_allow_html=True)

# Configuration
FREE_CHAR_LIMIT = 1200
STRIPE_PAYMENT_URL = "https://buy.stripe.com/3cIaEZ64O7tc6OP16p5Ne00"
VALID_PRO_CODE = "PRO2026"

# --- SIDEBAR: MODERN PLAN CARDS ---
with st.sidebar:
    st.markdown("### 🎓 Account Status")
    user_code = st.text_input("Enter Pro Passcode:", type="password", placeholder="Enter key...")
    is_pro = (user_code == VALID_PRO_CODE)

    if is_pro:
        st.markdown('<span class="status-badge badge-pro">PRO ACTIVE</span>', unsafe_allow_html=True)
        st.caption("All premium features unlocked.")
    else:
        st.markdown('<span class="status-badge badge-free">FREE TIER</span>', unsafe_allow_html=True)
        st.caption(f"Max {FREE_CHAR_LIMIT} characters. APA 7th style only.")
        st.divider()
        st.markdown(
            f"""
            <div class="pricing-card">
                <h4 style="margin:0 0 8px 0; color:#1E293B;">Upgrade to Pro (€1.00)</h4>
                <ul style="font-size:13px; color:#64748B; padding-left:18px; margin:0 0 12px 0;">
                    <li>Unlimited character length</li>
                    <li>Download formatted .docx files</li>
                    <li>MLA, Harvard, Chicago, IEEE</li>
                    <li>Citations Auditor</li>
                    <li><b>Writing Critique & Tips</b></li>
                </ul>
                <a href="{STRIPE_PAYMENT_URL}" target="_blank" style="text-decoration:none;">
                    <button style="width:100%; background:#2563EB; color:white; border:none; padding:8px 0; border-radius:8px; font-weight:600; cursor:pointer;">
                        Unlock Pro (€1.00)
                    </button>
                </a>
            </div>
            """,
            unsafe_allow_html=True
        )

# --- MAIN PAGE HEADER ---
st.title("Academic Document Suite")
st.caption("Standardized formatting, source auditing, and editorial feedback.")

# Format Selections
col1, col2 = st.columns([2, 1])

with col1:
    if is_pro:
        format_style = st.selectbox(
            "Target Citation Style",
            ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago Manual of Style", "IEEE"]
        )
    else:
        st.selectbox("Target Citation Style", ["APA 7th Edition (Locked to Free)"], disabled=True)
        format_style = "APA 7th Edition"

with col2:
    if is_pro:
        audit_citations = st.checkbox("Audit Citations", value=True)
        provide_critique = st.checkbox("Include Writing Feedback & Tips", value=True)
    else:
        st.checkbox("Audit Citations (Pro)", disabled=True)
        st.checkbox("Writing Feedback & Tips (Pro)", disabled=True)
        audit_citations = False
        provide_critique = False

user_text = st.text_area("Source Text / Draft", height=220, placeholder="Paste your draft or bibliography here...")
char_count = len(user_text)

# Character Counter Display
if not is_pro:
    pct = min(1.0, char_count / FREE_CHAR_LIMIT)
    st.progress(pct)
    st.caption(f"Usage: {char_count} / {FREE_CHAR_LIMIT} characters")
else:
    st.caption(f"Usage: {char_count} characters (Unlimited)")

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

# Run Processing
if st.button("Process Document", type="primary"):
    if not user_text.strip():
        st.warning("Please enter text first.")
    elif not is_pro and char_count > FREE_CHAR_LIMIT:
        st.error(f"Text exceeds the {FREE_CHAR_LIMIT}-character free limit. Please upgrade to continue.")
    else:
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
        
        extra_prompts = []
        if is_pro and audit_citations:
            extra_prompts.append("- Check in-text citations against the reference list. Note discrepancies under a clear '=== CITATION AUDIT REPORT ===' header.")
        if is_pro and provide_critique:
            extra_prompts.append("- Provide 3-5 specific, bulleted recommendations to strengthen academic tone, clarity, and conciseness under an '=== EDITORIAL FEEDBACK & REVISION TIPS ===' header.")

        prompt_additions = "\n".join(extra_prompts)

        system_prompt = f"""
        You are an expert academic editor and writing consultant.
        Reformat the provided text strictly according to {format_style} standards (capitalization, references, layout).
        Keep the author's core thesis intact.
        {prompt_additions}
        
        Structure your answer cleanly with Markdown.
        """
        
        with st.spinner("Processing document and generating suggestions..."):
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ]
            )
            
            output = response.choices[0].message.content
            st.success("Completed successfully!")
            
            # Display output in tabs if Pro options are used
            if is_pro and (provide_critique or audit_citations):
                tab1, tab2 = st.tabs(["📄 Formatted Document", "💡 Feedback & Reports"])
                with tab1:
                    st.text_area("Formatted Text:", value=output.split("===")[0].strip(), height=300)
                    docx_file = create_docx(output, format_style)
                    st.download_button(
                        label="📥 Download .docx",
                        data=docx_file,
                        file_name="academic_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                with tab2:
                    st.markdown(output)
            else:
                st.text_area("Formatted Output:", value=output, height=300)
