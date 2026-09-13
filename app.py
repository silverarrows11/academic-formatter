import streamlit as st
import openai

# 1. Page Configuration
st.set_page_config(
    page_title="Instant APA/MLA/Harvard Document Formatter", 
    page_icon="📝",
    layout="centered"
)

st.title("📝 Instant Document & Citation Formatter")
st.write("Convert your unformatted drafts, reference lists, and inline citations into perfectly styled academic documents in seconds.")

# 2. Check Stripe Payment Verification
query_params = st.query_params
has_paid = query_params.get("paid") == "true"

if not has_paid:
    st.info("🔒 Complete a one-time payment of €1.00 to unlock instant automated formatting.")
    st.link_button("Pay €1.00 to Format Document", "https://buy.stripe.com/3cIaEZ64O7tc6OP16p5Ne00")
else:
    st.success("✅ Payment verified! You now have full access to the formatting engine.")
    
    # 3. User Inputs
    user_text = st.text_area(
        "Paste your unformatted draft or reference list below:", 
        height=250,
        placeholder="Paste your essay, citations, or paper notes here..."
    )
    
    col1, col2 = st.columns(2)
    with col1:
        format_style = st.selectbox(
            "Select Citation Style", 
            ["APA 7th Edition", "MLA 9th Edition", "Chicago Manual of Style", "Harvard Style"]
        )
    with col2:
        output_focus = st.selectbox(
            "Primary Task",
            ["Full Paper Formatting", "Reference List Only", "Inline Citations Only"]
        )

    custom_notes = st.text_input(
        "Optional: Any specific rules? (e.g., 'Include hanging indents', 'Double space')", 
        placeholder="e.g., Fix missing author names if detectable"
    )

    # 4. Processing Action
    if st.button("Generate Formatted Document"):
        if not user_text.strip():
            st.warning("Please paste your text into the box above before formatting.")
        else:
            client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
            
            system_prompt = f"""
            You are a rigorous academic editor specializing in citation standard compliance.
            Reformat the user's input strictly according to {format_style} standards.
            Focus area requested: {output_focus}.
            Additional user instructions: {custom_notes if custom_notes else 'None'}.

            Rules:
            1. Fix all inline parenthetical citations and reference entries.
            2. Standardize headings, capitalization, italics, and structure per standard formatting guidelines.
            3. Do not alter the underlying thesis, facts, or writing voice.
            4. Output ONLY the polished, formatted document text ready for submission.
            """

            with st.spinner("Analyzing and formatting your document..."):
                try:
                    response = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_text}
                        ]
                    )
                    
                    formatted_result = response.choices[0].message.content
                    
                    st.subheader("🎉 Your Formatted Document")
                    st.text_area("Result:", value=formatted_result, height=350)
                    
                    # File Download Option
                    st.download_button(
                        label="Download Text File (.txt)",
                        data=formatted_result,
                        file_name="formatted_document.txt",
                        mime="text/plain"
                    )
                except Exception as e:
                    st.error("An error occurred during formatting. Please verify your settings or try again.")
