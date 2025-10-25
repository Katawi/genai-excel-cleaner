import streamlit as st
from langchain_openai import ChatOpenAI
from langchain.prompts import PromptTemplate
import pandas as pd
import tempfile
import os
import re
import numpy as np

# 🔐 Secure API key
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page Config
st.set_page_config(page_title="🧠 GenAI Excel Cleaner", layout="wide")

# 🧩 Header
st.title("🧠 GenAI Excel Cleaner")
st.markdown(
    "An **autonomous AI-powered data cleaning assistant** that analyzes, cleans, and explains changes in Excel files intelligently."
)
st.markdown(
    "<span style='font-size:14px;color:gray;'>Developed by <b>Group 3 – PwC Data & AI Mastery Program</b></span>",
    unsafe_allow_html=True,
)
st.divider()

# 📂 Upload Excel File
uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"])


# Helper: compute differences
def get_change_log(original_df, cleaned_df):
    """Generate a summary of what changed between original and cleaned DataFrames."""
    changes = []

    # Structural changes
    row_diff = len(original_df) - len(cleaned_df)
    if row_diff > 0:
        changes.append(f"🗑️ Removed {row_diff} duplicate or empty row(s).")
    elif row_diff < 0:
        changes.append(f"⚠️ Added {abs(row_diff)} new row(s) (unexpected).")

    # Column changes
    orig_cols = set(original_df.columns)
    clean_cols = set(cleaned_df.columns)
    added_cols = clean_cols - orig_cols
    removed_cols = orig_cols - clean_cols

    if added_cols:
        changes.append(f"➕ Added columns: {', '.join(added_cols)}")
    if removed_cols:
        changes.append(f"➖ Removed columns: {', '.join(removed_cols)}")

    # Column name normalization
    for col in original_df.columns:
        normalized = re.sub(r"\s+", "_", col.strip().lower())
        if col != normalized and normalized in cleaned_df.columns:
            changes.append(f"🔤 Renamed '{col}' → '{normalized}'")

    # Type changes
    for col in cleaned_df.columns:
        if col in original_df.columns:
            orig_type = str(original_df[col].dtype)
            clean_type = str(cleaned_df[col].dtype)
            if orig_type != clean_type:
                changes.append(f"🔢 Converted '{col}' type: {orig_type} → {clean_type}")

    # Value cleaning checks
    for col in cleaned_df.select_dtypes(include=["object"]).columns:
        if col in original_df.columns:
            before_nulls = original_df[col].isna().sum()
            after_nulls = cleaned_df[col].isna().sum()
            if after_nulls < before_nulls:
                changes.append(f"✨ Filled missing values in '{col}' ({before_nulls - after_nulls} fixed).")
            if any(cleaned_df[col].str.contains("Unknown", case=False, na=False)):
                changes.append(f"❓ Replaced empty or invalid values with 'Unknown' in '{col}'.")
    return changes


def clean_and_infer_types(df):
    """Automatic rule-based data cleaning + type inference."""
    df = df.drop_duplicates().reset_index(drop=True)
    df = df.dropna(how="all")
    df.columns = [re.sub(r"\s+", "_", col.strip().lower()) for col in df.columns]

    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str)
        df[col] = df[col].str.strip()
        df[col] = df[col].str.replace(r"\s+", " ", regex=True)
        df[col] = df[col].str.replace(r"[^\w\s\-./%]", "", regex=True)
        df[col] = df[col].replace(["Nan", "None", "Na", ""], np.nan)

        # Try to detect numeric with commas/symbols
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

        # Try converting to datetime
        try:
            df[col] = pd.to_datetime(df[col], errors="raise", infer_datetime_format=True)
            continue
        except Exception:
            pass

        df[col] = df[col].str.title()

    return df


def ai_clean_dataframe(df, sheet_name, llm):
    """AI-guided analysis + cleaning explanation + change log."""
    preview = df.head(5).to_string(index=False)

    # 🧠 Ask AI to analyze what needs cleaning
    analysis_template = PromptTemplate(
        input_variables=["sheet_name", "preview"],
        template="""
You are a professional data cleaning assistant.
Analyze this sample from the Excel sheet '{sheet_name}':
{preview}

Identify the main data quality problems and suggest how to clean them.
List your suggestions as bullet points.
""",
    )
    ai_plan = llm.invoke(analysis_template.format(sheet_name=sheet_name, preview=preview)).content

    # Apply local cleaning & detect changes
    original_df = df.copy()
    cleaned_df = clean_and_infer_types(df)
    changes = get_change_log(original_df, cleaned_df)

    # 🧠 Ask AI to summarize
    summary_template = PromptTemplate(
        input_variables=["sheet_name", "ai_plan", "changes"],
        template="""
You are summarizing the data cleaning actions for the sheet '{sheet_name}'.

AI's cleaning plan:
{ai_plan}

Detected changes during cleaning:
{changes}

Summarize these improvements in a concise, professional way suitable for a data quality report.
""",
    )
    ai_summary = llm.invoke(
        summary_template.format(sheet_name=sheet_name, ai_plan=ai_plan, changes="\n".join(changes))
    ).content

    return cleaned_df, ai_summary, changes


if uploaded_file:
    llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.3)

    if st.button("🚀 Let AI Clean My File"):
        st.info("AI is analyzing and cleaning your Excel file... Please wait ⏳")

        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets, explanations = {}, []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            cleaned_df, explanation, changes = ai_clean_dataframe(df, sheet_name, llm)
            cleaned_sheets[sheet_name] = cleaned_df

            st.markdown(f"### 🧾 Sheet: {sheet_name}")
            st.markdown(f"**AI Explanation:**\n{explanation}")
            st.markdown("**🔍 Detected Changes:**")
            for change in changes:
                st.markdown(f"- {change}")
            st.divider()

        # 💾 Save Cleaned Excel
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for s, d in cleaned_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)

        st.success("✅ AI Cleaning Completed! All detected changes have been listed above.")
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_data_ai.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.warning("Please upload an Excel file to begin.")
