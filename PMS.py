
import streamlit as st
import pandas as pd
import re
from io import BytesIO


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Approval Category Processor",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("📊 Approval Category Processor")

st.markdown(
    """
    Upload your **Input Excel file** and **Reference Excel file**,
    then the application will generate the final Output Excel file.
    """
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("📋 Required Columns")

st.sidebar.markdown(
    """
    **Input file:**
    - CustomerName
    - Comment

    **Reference file:**
    - FunctionName
    - Comment1 / Comment2 / Comment3 ...
    """
)


# ============================================================
# UPLOAD INPUT FILE
# ============================================================

st.header("1️⃣ Upload Input File")

input_file = st.file_uploader(
    "Choose Input Excel file",
    type=["xlsx", "xls"],
    key="input_file"
)


# ============================================================
# UPLOAD REFERENCE FILE
# ============================================================

st.header("2️⃣ Upload Reference File")

reference_file = st.file_uploader(
    "Choose Reference Excel file",
    type=["xlsx", "xls"],
    key="reference_file"
)


# ============================================================
# PROCESS BUTTON
# ============================================================

process_button = st.button(
    "🚀 Process Files",
    type="primary",
    use_container_width=True
)


# ============================================================
# PROCESS
# ============================================================

if process_button:

    # --------------------------------------------------------
    # CHECK FILES
    # --------------------------------------------------------

    if input_file is None:

        st.error("❌ Please upload the Input file.")

        st.stop()


    if reference_file is None:

        st.error("❌ Please upload the Reference file.")

        st.stop()


    # ========================================================
    # READ INPUT
    # ========================================================

    try:

        df_input = pd.read_excel(input_file)

    except Exception as e:

        st.error(
            f"❌ Error reading Input file:\n\n{e}"
        )

        st.stop()


    # ========================================================
    # CLEAN INPUT COLUMN NAMES
    # ========================================================

    df_input.columns = (
        df_input.columns
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # CHECK INPUT COLUMNS
    # ========================================================

    required_input_columns = [
        "CustomerName",
        "Comment"
    ]


    missing_input = [
        col
        for col in required_input_columns
        if col not in df_input.columns
    ]


    if missing_input:

        st.error(
            f"""
            ❌ Missing columns in INPUT file:

            {missing_input}

            Available columns:

            {df_input.columns.tolist()}
            """
        )

        st.stop()


    # ========================================================
    # READ REFERENCE
    # ========================================================

    try:

        df_reference = pd.read_excel(reference_file)

    except Exception as e:

        st.error(
            f"❌ Error reading Reference file:\n\n{e}"
        )

        st.stop()


    # ========================================================
    # CLEAN REFERENCE COLUMN NAMES
    # ========================================================

    df_reference.columns = (
        df_reference.columns
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # CHECK FUNCTION NAME
    # ========================================================

    if "FunctionName" not in df_reference.columns:

        st.error(
            "❌ Reference file must contain 'FunctionName' column."
        )

        st.stop()


    # ========================================================
    # FIND COMMENT COLUMNS AUTOMATICALLY
    # ========================================================

    comment_columns = [
        col
        for col in df_reference.columns
        if col.lower().startswith("comment")
    ]


    if not comment_columns:

        st.error(
            """
            ❌ No Comment columns found in Reference file.

            Expected columns such as:

            Comment1
            Comment2
            Comment3
            """
        )

        st.stop()


    # ========================================================
    # SHOW FILE INFORMATION
    # ========================================================

    st.success("✅ Files loaded successfully.")


    col1, col2 = st.columns(2)


    with col1:

        st.metric(
            "Input Rows",
            f"{len(df_input):,}"
        )


    with col2:

        st.metric(
            "Reference Rows",
            f"{len(df_reference):,}"
        )


    # ========================================================
    # NORMALIZE TEXT
    # ========================================================

    def normalize_text(value):

        if pd.isna(value):

            return ""

        value = str(value).strip()

        # Replace multiple spaces with one
        value = re.sub(
            r"\s+",
            " ",
            value
        )

        return value


    # ========================================================
    # BUILD REFERENCE LOOKUP
    # ========================================================

    reference_lookup = {}


    for _, row in df_reference.iterrows():

        function_name = normalize_text(
            row["FunctionName"]
        )


        if not function_name:

            continue


        for comment_col in comment_columns:

            reference_comment = normalize_text(
                row[comment_col]
            )


            # Ignore blank comments
            if not reference_comment:

                continue


            # Ignore placeholder comments
            if (
                reference_comment.upper()
                == "PLEASE GENERAT COMMENT"
            ):

                continue


            if reference_comment not in reference_lookup:

                reference_lookup[reference_comment] = []


            reference_lookup[
                reference_comment
            ].append(
                function_name
            )


    # ========================================================
    # REMOVE DUPLICATE FUNCTIONS
    # ========================================================

    for comment in reference_lookup:

        reference_lookup[comment] = list(
            dict.fromkeys(
                reference_lookup[comment]
            )
        )


    # ========================================================
    # APPROVAL CATEGORY RULES
    # ========================================================

    QA_CATEGORIES = [

        "Approved - B",
        "Blank Approval",
        "Pending",
        "Reject",
        "Verification",
        "Approved - C"

    ]


    UNK_CATEGORIES = [

        "Partial Unknown",
        "Unknown"

    ]


    DATA_CATEGORY = "Blank Data"


    # ========================================================
    # DETERMINE APPROVAL CATEGORY
    # ========================================================

    def get_categories(comment):

        comment = normalize_text(comment)


        if not comment:

            return []


        comment_upper = comment.upper()


        # ----------------------------------------------------
        # UNK
        # ----------------------------------------------------

        if "UNK" in comment_upper:

            return UNK_CATEGORIES


        # ----------------------------------------------------
        # QA
        # ----------------------------------------------------

        if "QA" in comment_upper:

            return QA_CATEGORIES


        # ----------------------------------------------------
        # DATA
        # ----------------------------------------------------

        if "DATA" in comment_upper:

            return [DATA_CATEGORY]


        # ----------------------------------------------------
        # Unknown
        # ----------------------------------------------------

        return []


    # ========================================================
    # FIND FUNCTION NAME FROM REFERENCE
    # ========================================================

    def find_functions(comment):

        comment = normalize_text(comment)


        if not comment:

            return []


        # ----------------------------------------------------
        # Exact match
        # ----------------------------------------------------

        if comment in reference_lookup:

            return reference_lookup[comment]


        # ----------------------------------------------------
        # Case-insensitive match
        # ----------------------------------------------------

        comment_upper = comment.upper()


        matches = []


        for ref_comment, functions in reference_lookup.items():

            if ref_comment.upper() == comment_upper:

                matches.extend(functions)


        if matches:

            return list(
                dict.fromkeys(matches)
            )


        # ----------------------------------------------------
        # Match after removing spaces
        # ----------------------------------------------------

        normalized_input = re.sub(
            r"\s+",
            "",
            comment_upper
        )


        matches = []


        for ref_comment, functions in reference_lookup.items():

            normalized_ref = re.sub(
                r"\s+",
                "",
                ref_comment.upper()
            )


            if normalized_ref == normalized_input:

                matches.extend(functions)


        return list(
            dict.fromkeys(matches)
        )


    # ========================================================
    # PROCESS INPUT
    # ========================================================

    output_rows = []

    not_match_rows = []


    progress_bar = st.progress(0)

    total_rows = len(df_input)


    for index, (_, row) in enumerate(
        df_input.iterrows()
    ):

        customer_name = row["CustomerName"]


        input_comment = normalize_text(
            row["Comment"]
        )


        # Ignore blank comments
        if not input_comment:

            progress_bar.progress(
                min(
                    (index + 1) / max(total_rows, 1),
                    1.0
                )
            )

            continue


        # ----------------------------------------------------
        # Find FunctionName
        # ----------------------------------------------------

        functions = find_functions(
            input_comment
        )


        # ----------------------------------------------------
        # Determine ApprovalCategory
        # ----------------------------------------------------

        categories = get_categories(
            input_comment
        )


        # ====================================================
        # NOT MATCH
        # ====================================================

        if not functions:

            not_match_rows.append({

                "CustomerName":
                    customer_name,

                "Comment":
                    input_comment

            })


            progress_bar.progress(
                min(
                    (index + 1) / max(total_rows, 1),
                    1.0
                )
            )

            continue


        # ====================================================
        # NO CATEGORY
        # ====================================================

        if not categories:

            output_rows.append({

                "CustomerName":
                    customer_name,

                "FunctionName":
                    functions[0],

                "ApprovalCategory":
                    "NOT FOUND",

                "Comment":
                    input_comment

            })


            progress_bar.progress(
                min(
                    (index + 1) / max(total_rows, 1),
                    1.0
                )
            )

            continue


        # ====================================================
        # CREATE OUTPUT ROWS
        # ====================================================

        for function_name in functions:

            for category in categories:

                output_rows.append({

                    "CustomerName":
                        customer_name,

                    "FunctionName":
                        function_name,

                    "ApprovalCategory":
                        category,

                    "Comment":
                        input_comment

                })


        progress_bar.progress(
            min(
                (index + 1) / max(total_rows, 1),
                1.0
            )
        )


    progress_bar.empty()


    # ========================================================
    # CREATE OUTPUT DATAFRAME
    # ========================================================

    df_output = pd.DataFrame(

        output_rows,

        columns=[

            "CustomerName",
            "FunctionName",
            "ApprovalCategory",
            "Comment"

        ]

    )


    # ========================================================
    # REMOVE DUPLICATES
    # ========================================================

    if not df_output.empty:

        df_output = (
            df_output
            .drop_duplicates(
                subset=[
                    "CustomerName",
                    "FunctionName",
                    "ApprovalCategory",
                    "Comment"
                ]
            )
            .reset_index(drop=True)
        )


    # ========================================================
    # CREATE NOT MATCH DATAFRAME
    # ========================================================

    df_not_match = pd.DataFrame(

        not_match_rows,

        columns=[

            "CustomerName",
            "Comment"

        ]

    )


    if not df_not_match.empty:

        df_not_match = (
            df_not_match
            .drop_duplicates()
            .reset_index(drop=True)
        )


    # ========================================================
    # CREATE CRITERIA SHEET
    # ========================================================

    criteria_data = [

        ["Approved - B", "include QA"],
        ["Blank Approval", "include QA"],
        ["Pending", "include QA"],
        ["Reject", "include QA"],
        ["Verification", "include QA"],
        ["Approved - C", "include QA"],

        ["Blank Data", "not include QA"],

        ["Partial Unknown", "include UNK"],
        ["Unknown", "include UNK"]

    ]


    df_criteria = pd.DataFrame(

        criteria_data,

        columns=[

            "ApprovalCategory",
            "Criteria"

        ]

    )


    # ========================================================
    # CREATE EXCEL FILE IN MEMORY
    # ========================================================

    output_buffer = BytesIO()


    with pd.ExcelWriter(
        output_buffer,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # OUTPUT SHEET
        # ----------------------------------------------------

        df_output.to_excel(
            writer,
            sheet_name="Output",
            index=False
        )


        # ----------------------------------------------------
        # NOT MATCH SHEET
        # ----------------------------------------------------

        df_not_match.to_excel(
            writer,
            sheet_name="Not Match",
            index=False
        )


        # ----------------------------------------------------
        # REFERENCE SHEET
        # ----------------------------------------------------

        df_reference.to_excel(
            writer,
            sheet_name="Reference",
            index=False
        )


        # ----------------------------------------------------
        # CRITERIA SHEET
        # ----------------------------------------------------

        df_criteria.to_excel(
            writer,
            sheet_name="Criteria",
            index=False
        )


    output_buffer.seek(0)


    # ========================================================
    # SUMMARY
    # ========================================================

    st.success(
        "🎉 Processing completed successfully!"
    )


    st.header("📊 Summary")


    summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)


    with summary_col1:

        st.metric(
            "Input Rows",
            f"{len(df_input):,}"
        )


    with summary_col2:

        st.metric(
            "Output Rows",
            f"{len(df_output):,}"
        )


    with summary_col3:

        st.metric(
            "Not Match Rows",
            f"{len(df_not_match):,}"
        )


    with summary_col4:

        st.metric(
            "Functions",
            f"{df_output['FunctionName'].nunique() if not df_output.empty else 0:,}"
        )


    # ========================================================
    # OUTPUT PREVIEW
    # ========================================================

    st.header("📄 Output Preview")


    if not df_output.empty:

        st.dataframe(
            df_output.head(100),
            use_container_width=True
        )

    else:

        st.info(
            "No output rows generated."
        )


    # ========================================================
    # NOT MATCH PREVIEW
    # ========================================================

    st.header("⚠️ Not Match")


    if not df_not_match.empty:

        st.warning(
            f"{len(df_not_match):,} comments were not found in the Reference file."
        )


        st.dataframe(
            df_not_match.head(100),
            use_container_width=True
        )

    else:

        st.success(
            "✅ No unmatched comments found."
        )


    # ========================================================
    # DOWNLOAD OUTPUT
    # ========================================================

    st.header("⬇️ Download Result")


    st.download_button(

        label="📥 Download Approval_Category_Output.xlsx",

        data=output_buffer.getvalue(),

        file_name="Approval_Category_Output.xlsx",

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        use_container_width=True

    )

