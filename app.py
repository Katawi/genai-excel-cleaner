import streamlit as st
from langchain_openai import ChatOpenAI
import pandas as pd
import tempfile
import os
import json
import re
import matplotlib.pyplot as plt

# 🔐 API key (from Streamlit Secrets)
os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]

# 🎨 Page setup
st.set_page_config(page_title="🧠 GenAI Excel Cleaner", layout="wide")
st.title("🧠 GenAI Excel Cleaner")
st.markdown(
    "AI analyzes your Excel file, decides cleaning actions, executes them automatically, "
    "and visualizes before-and-after data quality metrics."
)
st.markdown(
    "<span style='font-size:14px;color:gray;'>Developed by <b>Group 3 – PwC Data & AI Mastery Program</b></span>",
    unsafe_allow_html=True,
)
st.divider()

uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"])
llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.2)

# ─────────────────────────────────────────────────────────────
# 🧩 Helper functions
def profile_dataframe(df):
    """Return a summary dict of key metrics."""
    return {
        "rows": len(df),
        "nulls": int(df.isna().sum().sum()),
        "object_cols": len(df.select_dtypes(include=["object"]).columns),
        "numeric_cols": len(df.select_dtypes(include=["number"]).columns),
        "datetime_cols": len(df.select_dtypes(include=["datetime64[ns]"]).columns),
    }


def plot_before_after(before_stats, after_stats, sheet):
    """Draw before/after comparison charts."""
    fig, axes = plt.subplots(1, 2, figsize=(10, 4))

    # Chart 1 – Rows & Nulls
    axes[0].bar(["Rows", "Nulls"], [before_stats["rows"], before_stats["nulls"]], label="Before")
    axes[0].bar(["Rows", "Nulls"], [after_stats["rows"], after_stats["nulls"]], label="After", alpha=0.7)
    axes[0].set_title("Rows & Nulls")
    axes[0].legend()

    # Chart 2 – Column Types
    labels = ["Text", "Numeric", "Datetime"]
    axes[1].bar(labels,
                [before_stats["object_cols"], before_stats["numeric_cols"], before_stats["datetime_cols"]],
                label="Before")
    axes[1].bar(labels,
                [after_stats["object_cols"], after_stats["numeric_cols"], after_stats["datetime_cols"]],
                label="After", alpha=0.7)
    axes[1].set_title("Column Types")
    axes[1].legend()

    fig.suptitle(f"🧾 Sheet: {sheet}", fontsize=12)
    st.pyplot(fig)


def apply_cleaning_actions(df, actions):
    """Apply cleaning steps based on GPT plan."""
    change_log = []
    for act in actions:
        a_type = act.get("type")

        if a_type == "remove_duplicates":
            before = len(df)
            df = df.drop_duplicates().reset_index(drop=True)
            change_log.append(f"🗑️ Removed {before - len(df)} duplicate rows.")

        elif a_type == "drop_empty_rows":
            before = len(df)
            df = df.dropna(how="all")
            change_log.append(f"🧹 Dropped {before - len(df)} empty rows.")

        elif a_type == "trim_whitespace":
            for c in df.select_dtypes(include=["object"]).columns:
                df[c] = df[c].astype(str).str.strip()
            change_log.append("✂️ Trimmed whitespace in text columns.")

        elif a_type == "remove_special_chars":
            for c in df.select_dtypes(include=["object"]).columns:
                df[c] = df[c].str.replace(r"[^\w\s\-./]", "", regex=True)
            change_log.append("💬 Removed special characters from text columns.")

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
    return df, change_log


# ─────────────────────────────────────────────────────────────
if uploaded_file:
    if st.button("🚀 Clean with AI"):
        st.info("AI analyzing and cleaning... please wait ⏳")

        xls = pd.ExcelFile(uploaded_file)
        cleaned_sheets = {}
        all_logs = []

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            preview = df.head(10).to_csv(index=False)
            before_stats = profile_dataframe(df)

            # --- GPT decides cleaning plan (f-string fix) ---
            prompt_text = f"""
You are a data cleaning planner.
You will receive a small CSV sample from the sheet '{sheet}':
{preview}

Your job:
1. Identify data quality issues.
2. Output a JSON plan describing cleaning actions to apply.
3. Each action must follow this format:
   {{"type": "action_name", "columns": [optional], "value": [optional]}}

Valid actions:
- remove_duplicates
- drop_empty_rows
- trim_whitespace
- remove_special_chars
- standardize_case
- fill_missing
- convert_to_numeric

Output ONLY the JSON (no explanations).
"""
            plan_text = llm.invoke(prompt_text).content.strip()
            plan_text = re.sub(r"```json|```", "", plan_text).strip()

            try:
                actions = json.loads(plan_text)["actions"]
            except Exception:
                st.error(f"⚠️ Could not parse AI plan for sheet '{sheet}'. Using defaults.")
                actions = [{"type": "remove_duplicates"}, {"type": "trim_whitespace"}]

            # --- Execute plan ---
            cleaned_df, change_log = apply_cleaning_actions(df, actions)
            after_stats = profile_dataframe(cleaned_df)

            # --- Show visualization ---
            plot_before_after(before_stats, after_stats, sheet)

            cleaned_sheets[sheet] = cleaned_df
            all_logs.append(f"### 🧾 {sheet}\n" + "\n".join(change_log))

        # --- Save cleaned Excel ---
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for s, d in cleaned_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)

        st.success("✅ Cleaning completed and visualized by AI.")
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
