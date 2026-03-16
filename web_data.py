import streamlit as st
import pandas as pd
import result
import os
import base64
from openpyxl import load_workbook


def main():
    st.set_page_config(page_title="GTU Result Analysis - AVPTI", layout="centered")

    # ── Logo + Title header (single HTML block for perfect alignment) ──
    absolute_path = os.path.dirname(__file__)
    logo_path = os.path.join(absolute_path, "logo.jpg")

    logo_b64 = ""
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()

    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:20px;padding:10px 0 8px 0;">
            <img src="data:image/jpeg;base64,{logo_b64}" width="110"
                 style="flex-shrink:0;border-radius:6px;" />
            <div>
                <h2 style="margin:0;color:#F54927;font-size:26px;font-weight:700;line-height:1.3;">
                    A. V. PAREKH TECHNICAL INSTITUTE RAJKOT
                </h2>
                <p style="margin:2px 0 0 0;color:#57A9C7;font-size:16px;">
                    GTU Affiliated
                </p>
                <h3 style="margin:6px 0 0 0;color:#c8960c;font-size:30px;font-weight:700;">
                    &#128202; GTU Result Analysis
                </h3>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()
    process()


def process():

    def process_data(df, branch):
        file_path, visitor_count = result.result_ana(df, branch)
        st.metric("Total Visitors", int(visitor_count / 2))
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

    st.divider()
    st.markdown(
        '<p style="color:#888;font-size:13px;text-align:center;">'
        'Prepared by <strong>SHRI G.V. PARMAR</strong> &nbsp;|&nbsp; AVPTI, Rajkot'
        '</p>',
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()