import os
import io
import tempfile
import zipfile

import pandas as pd
import streamlit as st


def process_excel(input_path):
    """Same cleaning logic as before, but returns a DataFrame instead of writing a file."""
    sheets = pd.read_excel(input_path, sheet_name=None, header=None)
    final_df = pd.DataFrame()

    for sheet_index, (sheet_name, df) in enumerate(sheets.items()):
        # ---- SHEET 1 RULES ----
        if sheet_index == 0:
            header_row = df.iloc[6]        # Row 7 = header
            df = df.iloc[7:]               # Data starts from row 8
            df.columns = header_row
        # ---- SHEET 2+ RULES ----
        else:
            header_row = sheets[list(sheets.keys())[0]].iloc[6]  # use sheet1 header
            df = df.iloc[7:]
            df.columns = header_row

        # Remove completely empty rows
        df = df.dropna(how="all")

        # Drop trailing summary rows (same as your original)
        df = df.iloc[:-8]

        final_df = pd.concat([final_df, df], ignore_index=True)

        # Blank separator row between sheets (except after the last one)
        if sheet_index < len(sheets) - 1:
            empty_row = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
            final_df = pd.concat([final_df, empty_row], ignore_index=True)

    return final_df


def show_mailbox():
    st.title("Excel Sheet Processor (Bulk)")

    uploaded = st.file_uploader(
        "Upload Excel files (select multiple) OR a single .zip folder containing them",
        type=["xlsx", "zip"],
        accept_multiple_files=True,
    )

    uploaded_files = []
    if not uploaded:
        # Uploader is empty (cleared or nothing selected yet) - drop any
        # previously processed results so stale download buttons don't linger
        st.session_state.pop("individual_outputs", None)
        st.session_state.pop("bulk_output", None)

    if uploaded:
        for f in uploaded:
            if f.name.lower().endswith(".zip"):
                # Unpack the zip in memory and treat each xlsx inside as an uploaded file
                with zipfile.ZipFile(f) as zf:
                    for name in zf.namelist():
                        if name.lower().endswith(".xlsx") and not os.path.basename(name).startswith("~$") and not name.endswith("/"):
                            data = zf.read(name)
                            bio = io.BytesIO(data)
                            bio.name = os.path.basename(name)
                            uploaded_files.append(bio)
            else:
                uploaded_files.append(f)

    if uploaded_files:
        st.write(f"{len(uploaded_files)} file(s) ready to process.")

        if st.button("Process Files"):
            individual_outputs = {}   # filename -> bytes
            bulk_df = pd.DataFrame()
            skipped = []
            failed = []

            # Filter out Excel lock/temp files (e.g. "~$Mailbox....xlsx") created
            # when the source file is open, and any file that isn't real xlsx content
            files_to_process = []
            for uploaded in uploaded_files:
                if uploaded.name.startswith("~$"):
                    skipped.append(uploaded.name)
                else:
                    files_to_process.append(uploaded)

            progress = st.progress(0)
            status = st.empty()

            for i, uploaded in enumerate(files_to_process):
                status.text(f"Processing {uploaded.name} ({i + 1}/{len(files_to_process)})")

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
                    tmp_input.write(uploaded.read())
                    tmp_input_path = tmp_input.name

                try:
                    df = process_excel(tmp_input_path)
                except Exception as e:
                    failed.append((uploaded.name, str(e)))
                    os.remove(tmp_input_path)
                    continue
                finally:
                    if os.path.exists(tmp_input_path):
                        os.remove(tmp_input_path)

                # ---- Individual output for this file ----
                base_name, ext = os.path.splitext(uploaded.name)
                out_name = f"{base_name}-processed{ext}"

                buf = io.BytesIO()
                df.to_excel(buf, index=False)
                buf.seek(0)
                individual_outputs[out_name] = buf.getvalue()

                # ---- Append to the combined bulk dataframe ----
                if not bulk_df.empty:
                    empty_row = pd.DataFrame([[""] * len(bulk_df.columns)], columns=bulk_df.columns)
                    bulk_df = pd.concat([bulk_df, empty_row], ignore_index=True)
                bulk_df = pd.concat([bulk_df, df], ignore_index=True)

                progress.progress((i + 1) / len(files_to_process))

            status.text("Done!")

            if skipped:
                st.info(f"Skipped {len(skipped)} temp/lock file(s): {', '.join(skipped)}")
            if failed:
                st.warning("Some files failed to process:")
                for name, err in failed:
                    st.write(f"- {name}: {err}")

            bulk_buf = io.BytesIO()
            bulk_df.to_excel(bulk_buf, index=False)
            bulk_buf.seek(0)

            # Keep results around across reruns (e.g. when clicking download buttons)
            st.session_state["individual_outputs"] = individual_outputs
            st.session_state["bulk_output"] = bulk_buf.getvalue()

    if "individual_outputs" in st.session_state:
        individual_outputs = st.session_state["individual_outputs"]
        bulk_output = st.session_state["bulk_output"]

        st.subheader("Downloads")
        st.caption(f"{len(individual_outputs)} individual file(s) + 1 bulk combined file = {len(individual_outputs) + 1} total")

        # Everything zipped together — the practical way to hand back 91 files
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("bulk-combined.xlsx", bulk_output)
            for name, data in individual_outputs.items():
                zf.writestr(f"individual/{name}", data)
        zip_buf.seek(0)

        st.download_button(
            label=f"Download All ({len(individual_outputs) + 1} files, zipped)",
            data=zip_buf,
            file_name="processed_files.zip",
            mime="application/zip",
        )

        st.download_button(
            label="Download Bulk Combined File Only",
            data=bulk_output,
            file_name="bulk-combined-processed.xlsx",
        )

        with st.expander("Download individual files separately"):
            for name, data in individual_outputs.items():
                st.download_button(
                    label=f"Download {name}",
                    data=data,
                    file_name=name,
                    key=f"dl_{name}",
                )