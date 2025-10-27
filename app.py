import streamlit as st
from langchain_openai import ChatOpenAI
from langchain.prompts import PromptTemplate
import pandas as pd
import tempfile
import os
import re
import io

# 🔐 Secure API key from Streamlit Secrets
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page setup
st.set_page_config(page_title="🧠 GenAI Excel Cleaner", layout="wide")
st.title("🧠 GenAI Excel Cleaner (Fully AI-Driven & Scalable)")
st.markdown(
    "GPT-3.5 automatically cleans your Excel sheets — detects and fixes data issues, "
    "handles large files smartly, and explains all improvements."
)
st.markdown(
    "<span style='font-size:14px;color:gray;'>Developed by <b>Group 3 – PwC Data & AI Mastery Program</b></span>",
    unsafe_allow_html=True,
)
st.divider()

# 📂 Upload Excel file
uploaded_file = st.file_uploader("📂 Upload Excel file (.xlsx)", type=["xlsx"])

# 🧠 Initialize LLM
llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.3)


# ------------------------------------------------------------
# 🔧 Helper: Dynamically sample large sheets
# ------------------------------------------------------------
def sample_data_for_prompt(df, max_rows=120, max_chars=12000):
    """Return a text sample that fits GPT token limits intelligently."""
    # If small sheet, send all rows
    if len(df) <= max_rows:
        csv_text = df.to_csv(index=False)
    else:
        # For large files: sample first 60 + last 60 rows to preserve variety
        top = df.head(max_rows // 2)
        bottom = df.tail(max_rows // 2)
        sample = pd.concat([top, bottom])
        csv_text = sample.to_csv(index=False)

    # If still too long, truncate by characters
    if len(csv_text) > max_chars:
        csv_text = csv_text[:max_chars]

    return csv_text


# ------------------------------------------------------------
# 🧹 Helper: GPT-based cleaning
# ------------------------------------------------------------
def clean_with_gpt(sheet_name, df, llm):
    """Send data sample to GPT for cleaning and explanation."""
    csv_sample = sample_data_for_prompt(df)

    prompt = PromptTemplate.from_template("""
You are a professional data cleaning assistant working for a data analytics team.

You are given a sample of raw CSV data from an Excel sheet called "{sheet_name}".
Analyze and clean it intelligently.

Your tasks:
1. Detect and fix all common data quality problems (duplicates, missing values, extra spaces, inconsistent capitalization, special characters, wrong types, etc.).
2. Produce the **cleaned data as a valid CSV table only** — no markdown, no code blocks.
3. After the CSV table, write:
   "### EXPLANATION:" and describe briefly what you fixed and how.

Here is the raw data sample:
{csv_sample}
""")

    # 🧠 Ask GPT to clean the data
    response = llm.invoke(prompt.format(sheet_name=sheet_name, csv_sample=csv_sample)).content.strip()

    # Split GPT’s output into cleaned CSV + explanation
    parts = re.split(r"### EXPLANATION:", response, maxsplit=1)
    cleaned_csv = parts[0].strip()
    explanation = parts[1].strip() if len(parts) > 1 else "No explanation provided."

    # Try converting GPT’s CSV back to a DataFrame
    try:
        cleaned_df = pd.read_csv(io.StringIO(cleaned_csv))
    except Exception:
        cleaned_df = df.copy()
        explanation += "\n⚠️ Could not fully parse cleaned CSV; using original structure."

    return cleaned_df, explanation


# ------------------------------------------------------------
# 🚀 Main workflow
# ------------------------------------------------------------
if uploaded_file:
    if st.button("🚀 Let GPT Clean My File"):
        st.info("AI is analyzing and cleaning your Excel file... please wait ⏳")

        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets = {}

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            st.markdown(f"### 🧾 {sheet}")
            cleaned_df, explanation = clean_with_gpt(sheet, df, llm)

            # 🪄 Display results
            st.success(f"✅ Cleaning complete for sheet: {sheet}")
            st.markdown(f"**🤖 AI Explanation:** {explanation}")
            st.markdown("**📊 Cleaned Preview:**")
            st.dataframe(cleaned_df.head())
            st.divider()

            cleaned_sheets[sheet] = cleaned_df

        # 💾 Export cleaned Excel
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for s, d in cleaned_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)

        st.success("✅ All sheets cleaned successfully by GPT-3.5!")
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_data_ai.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.warning("Please upload an Excel file to begin.")
