import streamlit as st
from langchain_openai import ChatOpenAI
from langchain.prompts import PromptTemplate
import pandas as pd
import tempfile
import os
import json
import re

# 🔐 API key (Streamlit Secrets)
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page setup
st.set_page_config(page_title="🧠 GenAI Excel Cleaner", layout="wide")
st.title("🧠 GenAI Excel Cleaner")
st.markdown(
    "AI analyzes your Excel file, decides cleaning actions, and automatically executes them."
)
st.markdown(
    "<span style='font-size:14px;color:gray;'>Developed by <b>Group 3 – PwC Data & AI Mastery Program</b></span>",
    unsafe_allow_html=True,
)
st.divider()

uploaded_file = st.file_uploader("📂 Upload Excel file (.xlsx)", type=["xlsx"])

# 🧩 LLM setup
llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.2)

# --- Helper: Apply cleaning actions in Python -------------------------
def apply_cleaning_actions(df, actions):
    change_log = []
    for act in actions:
        a_type = act.get("type")

        if a_type == "remove_duplicates":
            before = len(df)
            df = df.drop_duplicates().reset_index(drop=True)
            change_log.append(f"🗑️ Removed {before - len(df)} duplicate rows.")

        elif a_type == "trim_whitespace":
            for c in df.select_dtypes(include=["object"]).columns:
                df[c] = df[c].astype(str).str.strip()
            change_log.append("✂️ Trimmed whitespace in text columns.")

        elif a_type == "standardize_case":
            cols = act.get("columns", df.select_dtypes(include=["object"]).columns)
            for c in cols:
                df[c] = df[c].str.title()
            change_log.append(f"🔠 Standardized capitalization in {len(cols)} column(s).")

        elif a_type == "fill_missing":
            val = act.get("value", "Unknown")
            df = df.fillna(val)
            change_log.append(f"❓ Filled all missing values with '{val}'.")

        elif a_type == "convert_to_numeric":
            cols = act.get("columns", [])
            for c in cols:
                try:
                    df[c] = (
                        df[c].astype(str)
                        .str.replace(",", "")
                        .str.replace("AED", "", case=False)
                        .str.replace("%", "")
                        .astype(float)
                    )
                    change_log.append(f"🔢 Converted '{c}' to numeric.")
                except Exception:
                    change_log.append(f"⚠️ Could not convert '{c}' to numeric.")

        elif a_type == "remove_special_chars":
            for c in df.select_dtypes(include=["object"]).columns:
                df[c] = df[c].str.replace(r"[^\w\s\-./]", "", regex=True)
            change_log.append("💬 Removed special characters in text columns.")

        elif a_type == "drop_empty_rows":
            before = len(df)
            df = df.dropna(how="all")
            change_log.append(f"🧹 Dropped {before - len(df)} empty rows.")

    return df, change_log

# --- Main process -----------------------------------------------------
if uploaded_file:
    if st.button("🚀 Clean with AI"):
        st.info("AI analyzing and cleaning... please wait ⏳")
        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets = {}
        all_logs = []

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            preview = df.head(10).to_csv(index=False)

            # --- Step 1: Ask GPT what to do ---
            prompt = PromptTemplate(
                input_variables=["sheet", "preview"],
                template="""
You are a data cleaning planner.
You will receive a small CSV sample from the sheet '{sheet}':
{preview}

Your job:
1. Identify all data quality issues.
2. Output a JSON plan describing cleaning actions to apply.
3. Each action must follow this schema:
   {"type": "action_name", "columns": [optional list], "value": [optional default]}

Valid actions:
- remove_duplicates
- drop_empty_rows
- trim_whitespace
- remove_special_chars
- standardize_case
- fill_missing
- convert_to_numeric

Output ONLY the JSON (no explanations).
""",
            )
            plan_text = llm.invoke(prompt.format(sheet=sheet, preview=preview)).content.strip()

            # Clean possible code fences
            plan_text = re.sub(r"```json|```", "", plan_text).strip()
            try:
                actions = json.loads(plan_text)["actions"]
            except Exception:
                st.error(f"⚠️ Could not parse AI plan for sheet '{sheet}'. Using defaults.")
                actions = [{"type": "remove_duplicates"}, {"type": "trim_whitespace"}]

            # --- Step 2: Apply actions in Python ---
            cleaned_df, change_log = apply_cleaning_actions(df, actions)
            cleaned_sheets[sheet] = cleaned_df
            all_logs.append(f"### 🧾 {sheet}\n" + "\n".join(change_log))

        # --- Step 3: Save cleaned Excel ---
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for s, d in cleaned_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)

        st.success("✅ Cleaning completed and applied by AI.")
        st.markdown("\n\n".join(all_logs))
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_data_ai.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.warning("Please upload an Excel file to begin.")
