import streamlit as st
import pandas as pd
from io import BytesIO
import uuid  # Used to refresh uploader key on restart

# -------------- Session init & state flags --------------
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = str(uuid.uuid4())

# -------------- App Header --------------
st.set_page_config(page_title="Sheet-wise Excel Merger", layout="centered")
st.title("📚 Merge Excel Files (Sheets Supported)")
st.write("Upload Excel files where either:")
st.markdown("- Files have **multiple same-named sheets**, or")
st.markdown("- Files have **a single sheet each** (same structure)")

# -------------- File Uploader with dynamic key --------------
uploaded_files = st.file_uploader(
    "📂 Upload Excel Files",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    key=st.session_state.uploader_key
)

# -------------- Merge Logic --------------
if uploaded_files:
    sheet_dict = {}
    single_sheet_files = []

    for file in uploaded_files:
        try:
            xls = pd.ExcelFile(file)
            if len(xls.sheet_names) == 1:
                df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
                df["SourceFile"] = file.name
                single_sheet_files.append(df)
            else:
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    if sheet_name not in sheet_dict:
                        sheet_dict[sheet_name] = []
                    sheet_dict[sheet_name].append(df)
        except Exception as e:
            st.error(f"❌ Error reading {file.name}: {e}")

    if sheet_dict:
        try:
            merged_sheets = {}
            for sheet_name, dfs in sheet_dict.items():
                merged_df = pd.concat(dfs, ignore_index=True)
                merged_sheets[sheet_name] = merged_df

            st.success("✅ Multiple sheets merged by sheet name!")
            st.subheader("📑 Preview of Merged Sheets")
            for name, df in merged_sheets.items():
                st.markdown(f"**🔹 {name}**")
                st.dataframe(df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for name, df in merged_sheets.items():
                    df.to_excel(writer, index=False, sheet_name=name[:31])
            output.seek(0)

            st.download_button(
                label="📥 Download Merged Excel File",
                data=output,
                file_name="merged_sheets.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Failed to merge sheets: {e}")

    elif single_sheet_files:
        try:
            merged_df = pd.concat(single_sheet_files, ignore_index=True)
            st.success("✅ Single-sheet Excel files merged into one sheet!")

            st.subheader("📑 Preview of Merged Sheet")
            st.dataframe(merged_df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False, sheet_name="MergedSheet")
            output.seek(0)

            st.download_button(
                label="📥 Download Merged Excel File",
                data=output,
                file_name="merged_single_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Failed to merge single-sheet files: {e}")
    else:
        st.warning("⚠️ No valid sheets found to merge.")
else:
    st.info("ℹ️ Please upload at least 2 Excel files with similar sheet structure.")

# -------------- Restart Button --------------
st.markdown("---")
if st.button("🔁 Restart"):
    # Change uploader key to refresh file uploader
    st.session_state.uploader_key = str(uuid.uuid4())
    st.rerun()
