import streamlit as st
from langchain_openai import ChatOpenAI
from langchain.prompts import PromptTemplate
import pandas as pd
import tempfile
import os
import re
import numpy as np

# 🔐 Secure API key (from Streamlit Secrets)
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page Config
st.set_page_config(page_title="🧠 GenAI Excel Cleaner", layout="wide")

# 🧩 Header
st.title("🧠 GenAI Excel Cleaner")
st.markdown(
    "An **autonomous AI-powered assistant** that analyzes and cleans messy Excel files intelligently — "
    "fixing duplicates, formatting, missing values, inconsistent types, and providing AI explanations per sheet."
)
st.markdown(
    "<span style='font-size:14px;color:gray;'>Developed by <b>Group 3 – PwC Data & AI Mastery Program</b></span>",
    unsafe_allow_html=True,
)
st.divider()

# 📂 Upload Excel File
uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"])


def clean_and_infer_types(df):
    """Automatic rule-based data cleaning + type inference"""
    df = df.drop_duplicates().reset_index(drop=True)
    df = df.dropna(how="all")
    df.columns = [re.sub(r"\s+", "_", col.strip().lower()) for col in df.columns]

    # Basic text cleaning for object columns
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str)
        df[col] = df[col].str.strip()
        df[col] = df[col].str.replace(r"\s+", " ", regex=True)
        df[col] = df[col].str.replace(r"[^\w\s\-./%]", "", regex=True)
        df[col] = df[col].replace(["Nan", "None", "Na", ""], np.nan)

        # Try to detect and convert numbers (even with commas or symbols)
        try:
            df[col] = (
                df[col]
                .str.replace(",", "")
                .str.replace("AED", "", case=False)
                .str.replace("%", "")
                .astype(float)
            )
            continue
        except Exception:
            pass

        # Try converting to dates
        try:
            df[col] = pd.to_datetime(df[col], errors="raise", infer_datetime_format=True)
            continue
        except Exception:
            pass

        # Otherwise, standardize text casing
        df[col] = df[col].str.title()

    return df


def ai_clean_dataframe(df, sheet_name, llm):
    """AI-guided analysis + cleaning explanation"""
    preview = df.head(5).to_string(index=False)

    # 🧠 Ask AI to analyze and decide what to clean
    analysis_template = PromptTemplate(
        input_variables=["sheet_name", "preview"],
        template="""
You are a professional data cleaning assistant.
Analyze the following data sample from the sheet '{sheet_name}':
{preview}

Identify issues such as:
- Duplicates or blank rows
- Extra spaces or inconsistent capitalization
- Symbols or special characters
- Mixed data types (e.g., numbers stored as text, dates stored as strings)
- Missing values

Then describe the main cleaning actions that should be performed.
Respond in clear bullet points.
""",
    )
    ai_plan = llm.invoke(
        analysis_template.format(sheet_name=sheet_name, preview=preview)
    ).content

    # Apply local cleaning & type inference
    original_shape = df.shape
    cleaned_df = clean_and_infer_types(df)
    cleaned_shape = cleaned_df.shape

    # 🧠 Ask AI to summarize the improvements
    summary_template = PromptTemplate(
        input_variables=["sheet_name", "ai_plan", "original_shape", "cleaned_shape"],
        template="""
You are summarizing data cleaning results for the sheet '{sheet_name}'.

AI identified and addressed the following issues:
{ai_plan}

Original shape: {original_shape} → Cleaned shape: {cleaned_shape}

Summarize the cleaning improvements in clear sentences.
End with a single summary sentence about the data quality after cleaning.
""",
    )
    ai_summary = llm.invoke(
        summary_template.format(
            sheet_name=sheet_name,
            ai_plan=ai_plan,
            original_shape=original_shape,
            cleaned_shape=cleaned_shape,
        )
    ).content

    return cleaned_df, ai_summary


if uploaded_file:
    llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.3)

    if st.button("🚀 Let AI Clean My File"):
        st.info("AI is analyzing and cleaning your Excel file... Please wait ⏳")

        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets, explanations = {}, []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            cleaned_df, explanation = ai_clean_dataframe(df, sheet_name, llm)
            cleaned_sheets[sheet_name] = cleaned_df
            explanations.append(f"### 🧾 Sheet: {sheet_name}\n{explanation}\n")

        # 💾 Save Cleaned Excel
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for s, d in cleaned_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)

        st.success("✅ AI Cleaning Completed!")
        st.markdown("\n".join(explanations))

        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_data_ai.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.warning("Please upload an Excel file to begin.")
