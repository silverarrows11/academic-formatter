import streamlit as st
import openai

st.set_page_config(page_title="Instant APA/MLA Document Formatter", page_icon="📝")

st.title("📝 Instant Document & Citation Formatter")
st.write("Paste your raw text below to convert it into perfect APA 7th Edition format instantly.")

# User Input
user_text = st.text_area("Paste your unformatted draft or reference list here:", height=250)
format_style = st.selectbox("Select Formatting Style", ["APA 7th Edition", "MLA 9th Edition", "Harvard Style"])

if st.button("Format Document"):
    if not user_text:
        st.warning("Please paste some text first.")
    else:
        # Access secret key securely
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
        
        system_prompt = f"""
        You are an expert academic editor. Reformat the user's input strictly according to {format_style} guidelines.
        Fix inline citations, reference lists, headings, and overall structure.
        Do not change the core meaning or writing voice of the user.
        Return only the formatted output ready to copy/paste.
        """
        
        with st.spinner("Formatting your document..."):
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ]
            )
            
            formatted_result = response.choices[0].message.content
            st.success("Formatting Complete!")
            st.text_area("Formatted Output:", value=formatted_result, height=300)
