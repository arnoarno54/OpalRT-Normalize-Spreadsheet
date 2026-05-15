import streamlit as st
    validation_errors = []

    for index, row in df.iterrows():

        # REQUIRED FIELDS
        for field in REQUIRED_COLUMNS:
            if str(row[field]).strip() == "":
                validation_errors.append(
                    f"Row {index + 2}: Missing required field -> {field}"
                )

        # EMAIL VALIDATION
        if not is_valid_email(row["Email"]):
            validation_errors.append(
                f"Row {index + 2}: Invalid email -> {row['Email']}"
            )

        # FIELD LENGTHS
        if len(str(row["First Name"])) > 50:
            validation_errors.append(
                f"Row {index + 2}: First Name exceeds 50 characters"
            )

        if len(str(row["Last Name"])) > 50:
            validation_errors.append(
                f"Row {index + 2}: Last Name exceeds 50 characters"
            )

        if len(str(row["Job Title"])) > 100:
            validation_errors.append(
                f"Row {index + 2}: Job Title exceeds 100 characters"
            )

        if len(str(row["Company Name"])) > 100:
            validation_errors.append(
                f"Row {index + 2}: Company Name exceeds 100 characters"
            )

    # REMOVE DUPLICATES
    if "Email" in df.columns:
        df = df.drop_duplicates(subset=["Email"])

    # FINAL COLUMN ORDER
    df = df[FINAL_COLUMNS]

    # RESULTS
    st.subheader("Dynamics-Ready Lead File")
    st.dataframe(df)

    # VALIDATION REPORT
    st.subheader("Validation Report")

    if len(validation_errors) == 0:
        st.success("No validation issues detected.")
    else:
        for error in validation_errors:
            st.warning(error)

    # EXPORT
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Dynamics-Ready CSV",
        data=csv,
        file_name="opalrt_dynamics_import.csv",
        mime="text/csv"
    )
