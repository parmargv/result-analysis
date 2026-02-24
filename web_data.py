import streamlit as st
import pandas as pd
import result
import os
from openpyxl import load_workbook

def main():
    st.set_page_config(page_title="Excel Processor", layout="centered")
    st.title("📊 Result Processing App")
    st.markdown(
        '<h2 style="color:#ffd700;font-size:15px;">Upload only excel .xlsx file.....</h2>',
        unsafe_allow_html=True
    )
    process()

def process():

    def process_data(df, branch):
        file_path, visitor_count = result.result_ana(df, branch)
        st.metric("Total Visitors", int(visitor_count/2))
        return file_path

    absolute_path = os.path.dirname(__file__)
    file_path = os.path.join(absolute_path, 'BRANCH_CODE.xlsx')

    df1 = pd.read_excel(file_path)
    b_code = df1["Branch_code"].tolist()

    br = st.selectbox("Select Branch", b_code)
    confirm = st.selectbox('Are you sure?', ('N', 'Y'))

    branch = br if confirm == "Y" else None

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        st.write("Uploaded file name:", uploaded_file.name)

        try:
            df = pd.read_excel(uploaded_file)

            if "BR_CODE" not in df.columns:
                st.error("❌ Missing column: BR_CODE")
                return

            if st.button("Process Data"):

                st.write("Processing Data...")
                file_path = process_data(df, branch)

                with open(file_path, "rb") as f:
                    st.success("✅ File Ready")
                    st.download_button(
                        label="📥 Download Excel",
                        data=f,
                        file_name="processed_file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"❌ Error: {e}")

    else:
        st.info("Please upload Excel file.")

    st.markdown(
        '<h2 style="color:#ffd700;font-size:18px;">Prepared by SHRI G.V.PARMAR AVPTI RAJKOT</h2>',
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()