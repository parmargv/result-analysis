import streamlit as st
import pandas as pd
import result
import os

# Your existing processing function
def main():
    def process_data(df,branch):
        excel_file = result.result_ana(df, branch)
        # Example: Add a new column (replace this with your logic)
        return excel_file

    # Streamlit app
    st.set_page_config(page_title="Excel Processor", layout="centered")
    st.markdown(f'<h1 style="color:#319AA2 ;font-size:35px;">Welcome to GTU result analysis</h1>',unsafe_allow_html=True)
    st.markdown(f'<h2 style="color:#ffd700, ;font-size:20px;">Prepared by SHRI G.V.PARMAR AVPTI RAJKOT...</h2>',unsafe_allow_html=True)
    st.title("📊 Excel Processing App")
    df1 = pd.read_excel("BRANCH_CODE.xlsx")
    # code = df['Branch_code']
    b_code = df1["Branch_code"].tolist()
    # branch = df['Branch_name']
    br = st.selectbox("Select Branch", b_code)
    option1 = st.selectbox('Are you sure?', ('N', 'Y'))
    if option1 == "Y":
        branch = br
    else:
        branch =1
    # Upload file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        st.write("Uploaded file name:", uploaded_file.name)
        with open("temp.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        df = pd.read_excel("temp.xlsx")
        #df = pd.read_excel(uploaded_file)
        #st.subheader("Preview of Uploaded Data")
        # st.dataframe(df.head())
        # Process Data
        if st.button("Process Data"):
            file = process_data(df,branch)
            # st.subheader("Preview of Processed Data")
            # st.dataframe(processed_df.head())

            # Convert DataFrame to Excel in memory
            absolute_path = os.path.dirname(__file__)
            file_path = os.path.join(absolute_path, 'gtu_result_analysis.xlsx')
            # file_path = 'gtu_result_analysis.xlsx'
            with open(file_path, "rb") as f:
                st.success("✅ Your file is ready! Click below to download.")
                st.balloons()
                st.download_button(
                    label="📥 Download Excel",
                    data=f,
                    file_name="processed_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        os.remove("temp.xlsx")

if __name__ == "__main__":
    main()