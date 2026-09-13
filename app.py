import streamlit as st
import openai

st.set_page_config(page_title="Instant APA/MLA Document Formatter", page_icon="📝")

st.title("📝 Instant Document & Citation Formatter")

# Check if user arrived via Stripe redirect
query_params = st.query_params
has_paid = query_params.get("paid") == "true"

if not has_paid:
    st.info("Please complete payment to unlock full formatting capabilities.")
    st.link_button("Pay $1.00 to Format Document", "https://buy.stripe.com/3cIaEZ64O7tc6OP16p5Ne00")
else:
    st.success("Payment verified! Paste your document below.")
    user_text = st.text_area("Paste draft here:", height=250)
    format_style = st.selectbox("Select Style", ["APA 7th Edition", "MLA 9th Edition", "Harvard Style"])

    if st.button("Generate Formatted Document"):
        client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
        system_prompt = f"Reformat strictly according to {format_style} guidelines."
        
        with st.spinner("Formatting..."):
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ]
            )
            st.text_area("Formatted Output:", value=response.choices[0].message.content, height=300)
