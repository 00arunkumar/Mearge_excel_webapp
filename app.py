import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Sheet-wise Excel Merger", layout="centered")

st.title("📚 Merge Excel Sheets Across Multiple Files")
st.write("Upload Excel files where sheets have the same names & structure. This app will merge each **same-named sheet** across files.")

uploaded_files = st.file_uploader("📂 Upload Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    sheet_dict = {}  # To store sheet-wise data

    for file in uploaded_files:
        try:
            xls = pd.ExcelFile(file)
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

            st.success("✅ Sheets merged successfully by sheet name!")

            st.subheader("📑 Preview of Merged Sheets")
            for name, df in merged_sheets.items():
                st.markdown(f"**🔹 {name}**")
                st.dataframe(df.head())

            # Prepare merged Excel file with all merged sheets
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for name, df in merged_sheets.items():
                    df.to_excel(writer, index=False, sheet_name=name[:31])  # Excel sheet name limit
            output.seek(0)

            st.download_button(
                label="📥 Download Merged Excel File",
                data=output,
                file_name="merged_sheets.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Failed to merge sheets: {e}")
    else:
        st.warning("⚠️ No valid data found in sheets.")
else:
    st.info("ℹ️ Please upload at least 2 Excel files with matching sheet names.")
