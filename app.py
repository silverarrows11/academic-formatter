import io
import streamlit as st
import openai
from docx import Document
from docx.shared import Inches, Pt

st.set_page_config(page_title="Academic Document Formatter", page_icon="📝", layout="wide")

# Configuration
FREE_CHAR_LIMIT = 1200
STRIPE_PAYMENT_URL = "https://buy.stripe.com/3cIaEZ64O7tc6OP16p5Ne00"
VALID_PRO_CODE = "PRO2026"  # Replace with your custom passkey or dynamic system

# --- SIDEBAR: PRICING TIERS ---
st.sidebar.title("Plan Comparison")
st.sidebar.markdown(
    """
| Feature | Free Tier | Pro Pass (€1.00) |
| :--- | :---: | :---: |
| **Character Limit** | 1,200 chars | Unlimited |
| **Styles** | APA 7th only | APA, MLA, Harvard, Chicago, IEEE |
| **Download .docx** | ❌ | ✅ |
| **Citation Auditor**| ❌ | ✅ |
| **Auto-Alphabetize**| ❌ | ✅ |
"""
)

st.sidebar.markdown(f"[👉 **Upgrade to Pro**]({STRIPE_PAYMENT_URL})")
user_code = st.sidebar.text_input("Enter Pro Access Code:", type="password")
is_pro = (user_code == VALID_PRO_CODE)

if is_pro:
    st.sidebar.success("Pro Tier Active! All features unlocked.")
else:
    st.sidebar.info("Using Free Tier.")

# --- MAIN INTERFACE ---
st.title("📝 Academic Formatter & Citation Engine")
st.write("Convert raw drafts and bibliographies into clean, publication-ready academic formats.")

# Tier-dependent inputs
if is_pro:
    format_style = st.selectbox(
        "Select Formatting Style:",
        ["APA 7th Edition", "MLA 9th Edition", "Harvard Style", "Chicago Manual of Style", "IEEE"]
    )
    audit_citations = st.checkbox("Run Citation Auditor (Flags uncited sources or missing references)")
else:
    st.caption("🔒 *Free tier uses APA 7th Edition. Upgrade to access MLA, Harvard, Chicago, and IEEE.*")
    format_style = "APA 7th Edition"
    audit_citations = False

user_text = st.text_area("Paste your text or references below:", height=250)
char_count = len(user_text)

# Character limit display
if not is_pro:
    st.caption(f"Character Count: {char_count}/{FREE_CHAR_LIMIT}")
    if char_count > FREE_CHAR_LIMIT:
        st.warning(f"Your input exceeds the free {FREE_CHAR_LIMIT}-character limit. Please shorten your text or unlock Pro.")
else:
    st.caption(f"Character Count: {char_count} (Unlimited)")

# --- HELPER: DOCX GENERATOR ---
def create_docx(content, title_style):
    doc = Document()
    
    # 1-inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    p = doc.add_paragraph()
    run = p.add_run(f"Formatted Output ({title_style})\n\n")
    run.bold = True
    
    # Body text formatting (Times New Roman, 12pt, double-spaced)
    body = doc.add_paragraph()
    body_run = body.add_run(content)
    body_run.font.name = "Times New Roman"
    body_run.font.size = Pt(12)
    body.paragraph_format.line_spacing = 2.0
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- PROCESS BUTTON ---
if st.button("Format Document"):
    if not user_text.strip():
        st.warning("Please enter text to format.")
    elif not is_pro and char_count > FREE_CHAR_LIMIT:
        st.error(f"Cannot process: Text exceeds the free tier limit of {FREE_CHAR_LIMIT} characters.")
    else:
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
        
        # Build prompt based on tier
        extra_instructions = ""
        if is_pro and audit_citations:
            extra_instructions += "\n- Cross-check body citations against the reference list and append an 'Auditor Report' at the end noting any missing matches."
        if is_pro:
            extra_instructions += "\n- Alphabetize the bibliography strictly according to author last name/title guidelines."

        system_prompt = f"""
        You are an expert academic editor. Reformat the user's text strictly according to {format_style} guidelines.
        Ensure correct capitalization, italics, punctuation, and structure.{extra_instructions}
        Return only the formatted document ready for submission.
        """
        
        with st.spinner("Processing document..."):
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ]
            )
            
            output_text = response.choices[0].message.content
            st.success("Formatting Complete!")
            st.text_area("Formatted Result:", value=output_text, height=300)
            
            # Export Option (Pro Only)
            if is_pro:
                docx_file = create_docx(output_text, format_style)
                st.download_button(
                    label="📥 Download as .docx",
                    data=docx_file,
                    file_name="formatted_paper.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.info("💡 Want a pre-formatted Word (.docx) file with 1-inch margins and double spacing? Upgrade to Pro.")
