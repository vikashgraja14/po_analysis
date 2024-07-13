# from tab2 import *
from tab3 import *
# from tab1 import *
# from tab4 import *
# from tab5 import *
# from tab6 import *


def on_upload_click():
    """
    Callback function for the upload button click event.
    """
    st.session_state.upload = True


def on_analyse_click():
    """
    Callback function for the analyse button click event.
    """
    st.session_state.analyse = True


# Initialize session state keys
if "upload" not in st.session_state:
    st.session_state.upload = False
    st.session_state.upload_disabled = True

if "analyse" not in st.session_state:
    st.session_state.analyse = False

st.set_page_config(layout="wide")
st.markdown(
    "<h1 style='text-align: center; color: black;'><b>NON PO PAYMENT ANALYSIS </b></h1>",
    unsafe_allow_html=True,
)

with st.expander("Upload Excel Files", expanded=False):
    with st.form("my_form"):
        # File Uploader
        files = st.file_uploader(
            "Non PO Payments",
            type="xlsx",
            accept_multiple_files=True,
            key="file_uploader",
        )
        if not files:
            st.session_state.upload_disabled = True
            st.session_state.upload = False
        else:
            st.session_state.upload_disabled = False

        # Add the submit button
        submit_button = st.form_submit_button("Analyze")

# Automatically trigger upload when files are selected
if files:
    on_upload_click()

if st.session_state.upload:
    # Process only the first file from the list
    if len(files) == 1:
        grouped_data = process_data(files[0])
    else:
        # Process each file and concatenate the dataframes
        all_data = []
        for file in files:
            all_data.append(process_data(file))
        grouped_data = pd.concat(all_data)

    # Create copies of the data
    radio1 = grouped_data.copy()
    data = grouped_data.copy()
    specialdf = grouped_data.copy()
    inv_specialdf = grouped_data.copy()
    specialexp = grouped_data.copy()

    exceptions = grouped_data.copy()
    exceptions2 = grouped_data.copy()
    ApprovalExceptions = grouped_data.copy()
    RejactionRemarks = grouped_data.copy()

    # Create tabs for different analyses
    t1, t2, t4, t3, t5, t6 = st.tabs(
        [
            "Overall Analysis",
            "Yearly Analysis",
            "Exceptions",
            "Specific Exceptions",
            "Approval Exceptions",
            "Approved -Rejection Remarks",
        ]
    )

    # with t2:
    #     @st.experimental_fragment
    #     def call_t2():
    #         display_dashboard(grouped_data)
    #     call_t2()
    #
    # with t1:
    #
    #     @st.experimental_fragment
    #     def call_t1():
    #         datatab1(data)
    #     call_t1()
    #
    # with t4:
    #
    #     @st.experimental_fragment
    #     def call_t4():
    #         checkbox(radio1, dfholiday, dfholiday2)
    #     call_t4()

    with t3:

        @st.experimental_fragment
        def call_t3():
            Special_exceptions(specialexp)
        call_t3()
              
    # with t5:
    #     uploaded_files = st.file_uploader(
    #         "Upload Attendance files", type=["xlsx"], accept_multiple_files=True
    #     )
    #     # Initialize concatenated_df outside the if block
    #     concatenated_df = None
    #     # Button to trigger analysis
    #     if st.button("Analyze Files"):
    #         if uploaded_files:
    #             dfs = []
    #             for file in uploaded_files:
    #                 df = pd.read_excel(file)
    #                 dfs.append(df)
    #
    #             # Concatenate DataFrames
    #             concatenated_df = pd.concat(dfs, ignore_index=True)
    #
    #             concatenated_df = concatenated_df[
    #                 (concatenated_df["IN Time"] == "00:00:00")
    #                 & (concatenated_df["OUT Time"] == "00:00:00")
    #                 ]
    #             concatenated_df["Date"] = pd.to_datetime(
    #                 concatenated_df["Date"], errors="coerce"
    #             )
    #             concatenated_df["Date"] = concatenated_df["Date"].dt.date
    #             concatenated_df = concatenated_df.sort_values(
    #                 by=["Empl./appl.name", "Date"]
    #             )
    #             concatenated_df.reset_index(drop=True, inplace=True)
    #             concatenated_df.index += 1  # Start index from 1
    #     ApprovalExceptions["HOG Approval by"] = ApprovalExceptions[
    #         "HOG Approval by"
    #     ].astype(str)
    #     ApprovalExceptions["HOG Approval by"] = ApprovalExceptions[
    #         "HOG Approval by"
    #     ].apply(lambda x: str(x) if isinstance(x, str) else "")
    #     ApprovalExceptions["HOG Approval by"] = ApprovalExceptions[
    #         "HOG Approval by"
    #     ].apply(lambda x: re.sub(r"\..*", "", x))
    #     ApprovalExceptions["HOG Approval on"] = pd.to_datetime(
    #         ApprovalExceptions["HOG Approval on"], errors="coerce"
    #     )
    #     ApprovalExceptions["HOG Approval on"] = ApprovalExceptions[
    #         "HOG Approval on"
    #     ].dt.date
    #     ApprovalExceptions["HOD Apr/Rej by"] = ApprovalExceptions[
    #         "HOD Apr/Rej by"
    #     ].astype(str)
    #     ApprovalExceptions["HOD Apr/Rej by"] = ApprovalExceptions[
    #         "HOD Apr/Rej by"
    #     ].apply(lambda x: str(x) if isinstance(x, str) else "")
    #     ApprovalExceptions["HOD Apr/Rej by"] = ApprovalExceptions[
    #         "HOD Apr/Rej by"
    #     ].apply(lambda x: re.sub(r"\..*", "", x))
    #     ApprovalExceptions["HOD Apr/Rej on"] = pd.to_datetime(
    #         ApprovalExceptions["HOD Apr/Rej on"], errors="coerce"
    #     )
    #     ApprovalExceptions["HOD Apr/Rej on"] = ApprovalExceptions[
    #         "HOD Apr/Rej on"
    #     ].dt.date
    #     @st.experimental_fragment
    #     def call_t5():
    #         tab5(ApprovalExceptions,exceptions2,concatenated_df)
    #     call_t5()
    #
    # with t6:
    #
    #     @st.experimental_fragment
    #     def call_t6():
    #         tab6(RejactionRemarks)
    #     call_t6()
