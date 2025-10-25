import streamlit as st
from langchain_openai import ChatOpenAI
from langchain.prompts import PromptTemplate
import pandas as pd
import tempfile
import os
import re

# 🔐 Secure API key from Streamlit Secrets
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page Config
st.set_page_config(page_title="🧠 GenAI Excel Cleaner v3.0", layout="wide")

# ─────────────────────────────────────────────────────────────
# 🧩 HEADER
st.title("🧠 GenAI Excel Cleaner v3.0")
st.markdown(
    "A smart, AI-powered assistant that cleans messy Excel files automatically — "
    "handles duplicates, spacing, missing values, formatting, and provides AI explanations per sheet."
)
st.markdown(
    "<span style='font-size:14px;color:gray;'>Developed by <b>Group 3 – PwC Data & AI Mastery Program</b></span>",
    unsafe_allow_html=True,
)
st.divider()

# ─────────────────────────────────────────────────────────────
# 🧰 SIDEBAR OPTIONS
st.sidebar.header("⚙️ Cleaning Options")
remove_duplicates = st.sidebar.checkbox("🧩 Remove Duplicates", True)
remove_empty_rows = st.sidebar.checkbox("🗑️ Remove Empty Rows", True)
trim_whitespace = st.sidebar.checkbox("✂️ Trim & Collapse Spaces", True)
normalize_case = st.sidebar.checkbox("🔠 Normalize Capitalization", True)
remove_symbols = st.sidebar.checkbox("💬 Remove Special Characters", True)
handle_missing = st.sidebar.checkbox("❓ Replace Missing with 'Unknown'", True)
convert_types = st.sidebar.checkbox("🔢 Auto-Detect Numbers/Dates", True)
ai_explanation = st.sidebar.checkbox("🤖 Generate AI Cleaning Explanation", True)
st.sidebar.divider()
st.sidebar.info("Select which cleaning steps to apply before running the process.")

# ─────────────────────────────────────────────────────────────
# 📂 Upload Excel File
uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"])

def clean_dataframe(df):
    """Perform rule-based cleaning on a DataFrame based on sidebar selections."""
    if remove_duplicates:
        df = df.drop_duplicates().reset_index(drop=True)
    if remove_empty_rows:
        df = df.dropna(how="all")

    # Normalize headers
    df.columns = [re.sub(r"\s+", "_", col.strip().lower()) for col in df.columns]

    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str)
        if trim_whitespace:
            df[col] = df[col].str.strip().str.replace(r"\s+", " ", regex=True)
        if remove_symbols:
            df[col] = df[col].str.replace(r"[^\w\s\-./]", "", regex=True)
        if normalize_case:
            df[col] = df[col].str.title()
        if handle_missing:
            df[col] = df[col].replace(["Nan", "None", "Na", ""], "Unknown")

    if convert_types:
        for col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col])
            except Exception:
                try:
                    df[col] = pd.to_numeric(df[col])
                except Exception:
                    pass
    return df

# ─────────────────────────────────────────────────────────────
if uploaded_file:
    llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.3)

    if st.button("🚀 Clean My Excel File"):
        st.info("Processing your Excel file... Please wait ⏳")
        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets, explanations = {}, []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            original_shape = df.shape
            cleaned_df = clean_dataframe(df)
            cleaned_shape = cleaned_df.shape

            # 🤖 Optional AI Explanation
            if ai_explanation:
                template = PromptTemplate(
                    input_variables=["sheet_name", "original_shape", "cleaned_shape"],
                    template="""
You are a professional data-cleaning assistant.
Explain clearly what cleaning actions were applied to the Excel sheet '{sheet_name}'.
Compare the original vs. cleaned shapes ({original_shape} → {cleaned_shape}) and mention:
- Duplicate removal
- Empty-row deletion
- Column normalization
- Text spacing/capitalization fixes
- Type conversions
Conclude with one concise summary sentence.
""",
                )
                prompt = template.format(
                    sheet_name=sheet_name,
                    original_shape=original_shape,
                    cleaned_shape=cleaned_shape,
                )
                explanation = llm.invoke(prompt).content
                explanations.append(f"### 🧾 Sheet: {sheet_name}\n{explanation}\n")

            cleaned_sheets[sheet_name] = cleaned_df

        # 💾 Save Cleaned Excel
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for s, d in cleaned_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)

        st.success("✅ Cleaning Completed!")
        if ai_explanation:
            st.markdown("\n".join(explanations))

        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.warning("Please upload an Excel file to begin.")
