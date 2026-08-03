import tempfile
import time
from pathlib import Path

import streamlit as st

from chdn_report import run_report

st.set_page_config(page_title="CHDN Report Step", layout="wide")
st.title("📊 CHDN Report Step")

st.info("This app runs only the report transformation stage and writes the long-format workbook from a cleaned file.")

if "report_data" not in st.session_state:
    st.session_state.report_data = None
if "report_status" not in st.session_state:
    st.session_state.report_status = None
if "report_error" not in st.session_state:
    st.session_state.report_error = None
if "report_elapsed" not in st.session_state:
    st.session_state.report_elapsed = None

uploaded_file = st.file_uploader("Upload the cleaned CHDN workbook", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    st.caption(f"Selected file: {uploaded_file.name}")

    run_clicked = st.button("Run Report", type="primary")

    if run_clicked:
        st.session_state.report_data = None
        st.session_state.report_status = None
        st.session_state.report_error = None
        st.session_state.report_elapsed = None

        output_path = Path(__file__).resolve().with_name("report_output.xlsx")
        if output_path.exists():
            output_path.unlink()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_input = Path(tmp_file.name)

        with st.spinner("Generating the report workbook..."):
            started_at = time.perf_counter()
            try:
                status = run_report(temp_input, output_path)
                st.session_state.report_status = status
                if status == 0 and output_path.exists():
                    with open(output_path, "rb") as fh:
                        st.session_state.report_data = fh.read()
            except Exception as exc:
                st.session_state.report_status = -1
                st.session_state.report_error = exc
            finally:
                st.session_state.report_elapsed = round(time.perf_counter() - started_at, 2)

    if st.session_state.report_status == 0 and st.session_state.report_data is not None:
        st.success("Report step completed successfully.")
        if st.session_state.report_elapsed is not None:
            st.caption(f"Runtime: {st.session_state.report_elapsed} seconds")
        st.download_button(
            label="Download report workbook",
            data=st.session_state.report_data,
            file_name="CHDN_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif st.session_state.report_status is not None:
        st.error("The report step did not finish successfully.")
        if st.session_state.report_elapsed is not None:
            st.caption(f"Runtime before failure: {st.session_state.report_elapsed} seconds")
        if st.session_state.report_error is not None:
            st.exception(st.session_state.report_error)
else:
    st.warning("Please upload a cleaned workbook to begin.")
