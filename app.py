
import streamlit as st
from langchain_openai import ChatOpenAI
from langchain.prompts import PromptTemplate
import pandas as pd
import io
import tempfile
import os

# 🔑 OpenAI API Key (Ali's key)
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page Config
st.set_page_config(page_title="🧠 GenAI Excel Cleaner", layout="wide")

# 🧩 Header Section
st.title("🧠 GenAI Excel Cleaner")
st.markdown("Clean Excel files intelligently with AI — handles multiple sheets, removes duplicates, and explains what was done.")
st.markdown("<span style='font-size:14px; color:gray;'>Developed by <b>Ali Al Shamsi</b> — Data Engineer</span>", unsafe_allow_html=True)
st.divider()

# 📂 Upload Excel File
uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.3)
    
    if st.button("🚀 Clean My Excel File"):
        st.info("Processing your Excel file... please wait ⏳")
        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets = {}
        explanations = []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            original_shape = df.shape
            df = df.drop_duplicates().reset_index(drop=True)
            df = df.dropna(how='all')
            df.columns = [col.strip().replace(' ', '_').lower() for col in df.columns]
            cleaned_shape = df.shape

            # AI explanation
            template = PromptTemplate(
                input_variables=["sheet_name", "original_shape", "cleaned_shape"],
                template="""
                You are a professional data cleaning assistant.
                Explain what was cleaned on the sheet '{sheet_name}'.
                Original shape: {original_shape}, Cleaned shape: {cleaned_shape}.
                Mention duplicate removal, empty rows, or header formatting if applicable.
                """
            )
            prompt = template.format(sheet_name=sheet_name, original_shape=original_shape, cleaned_shape=cleaned_shape)
            explanation = llm.invoke(prompt).content

            explanations.append(f"### 🧾 Sheet: {sheet_name}\n{explanation}\n")
            cleaned_sheets[sheet_name] = df

        # Save cleaned Excel
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, df in cleaned_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        st.success("✅ Cleaning Completed!")
        st.markdown("\n".join(explanations))

        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning("Please upload an Excel file to begin.")

