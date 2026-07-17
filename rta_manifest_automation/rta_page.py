import streamlit as st
import tempfile
import os
import gc
from rta_manifest_automation.processor.pipeline import run_processing_pipeline


def rta_page():
    # --- Custom CSS for a cleaner card look ---
    st.markdown("""
        <style>
        .upload-section {
            background: #f8f9fb;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            padding: 18px 20px;
            margin-bottom: 16px;
        }
        .required-badge {
            display: inline-block;
            background: #dc2626;
            color: white;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-left: 8px;
            vertical-align: middle;
        }
        .optional-badge {
            display: inline-block;
            background: #6b7280;
            color: white;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-left: 8px;
            vertical-align: middle;
        }
        .section-title {
            font-size: 1.15rem;
            font-weight: 700;
            color: #111827;
            margin-bottom: 4px;
        }
        .section-desc {
            font-size: 0.9rem;
            color: #6b7280;
            margin-bottom: 12px;
        }
        .status-ok {
            color: #059669;
            font-weight: 600;
            font-size: 0.9rem;
        }
        .status-missing {
            color: #9ca3af;
            font-size: 0.9rem;
        }
        </style>
    """, unsafe_allow_html=True)

    st.header("📄 RTA Excel File Processor")
    st.caption("Supports files up to 1 year of transaction data")
    st.markdown("---")

    # ============================================================
    # SECTION 1: RTA FILE (REQUIRED)
    # ============================================================
    st.markdown(
        '<div class="section-title">📊 RTA Transaction File'
        '<span class="required-badge">REQUIRED</span></div>'
        '<div class="section-desc">The main Register Transaction Activity Excel export from your POS system.</div>',
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "Drop your RTA file here",
        type=["xlsx"],
        key="rta_upload",
        label_visibility="collapsed",
    )
    if uploaded_file:
        st.markdown(f'<span class="status-ok">✅ {uploaded_file.name}</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-missing">⏳ No file uploaded yet</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # SECTION 2: ITEM MASTER (OPTIONAL)
    # ============================================================
    st.markdown(
        '<div class="section-title">🏷️ Item Master File'
        '<span class="optional-badge">OPTIONAL</span></div>'
        '<div class="section-desc">Adds <b>Item NetSuite ID</b> and <b>Item NetSuite Name</b> columns from Step 4 onwards.</div>',
        unsafe_allow_html=True,
    )
    item_master_file = st.file_uploader(
        "Drop your Item Master file here",
        type=["xlsx", "csv"],
        key="item_master_upload",
        label_visibility="collapsed",
    )
    if item_master_file:
        st.markdown(f'<span class="status-ok">✅ {item_master_file.name}</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-missing">⏳ Not uploaded (NetSuite Item columns will be blank)</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # SECTION 3: CUSTOMER FILE (OPTIONAL)
    # ============================================================
    st.markdown(
        '<div class="section-title">👤 Customer List (Silk File)'
        '<span class="optional-badge">OPTIONAL</span></div>'
        '<div class="section-desc">Adds <b>Customer NetSuite ID</b> column. Rows with multiple matching IDs are highlighted <b style="background:#fff59d;padding:1px 6px;border-radius:3px;">yellow</b>.</div>',
        unsafe_allow_html=True,
    )
    customer_file = st.file_uploader(
        "Drop your Customer file here",
        type=["xlsx", "csv"],
        key="customer_upload",
        label_visibility="collapsed",
    )
    if customer_file:
        st.markdown(f'<span class="status-ok">✅ {customer_file.name}</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-missing">⏳ Not uploaded (Customer NetSuite ID column will be blank)</span>', unsafe_allow_html=True)

    st.markdown("---")

    # ============================================================
    # PROCESS BUTTON
    # ============================================================
    if not uploaded_file:
        st.warning("⚠️ Please upload the RTA file above to enable processing.")
        return

    original_name = os.path.splitext(uploaded_file.name)[0]
    processed_name = f"{original_name} - Processed.xlsx"

    # Show summary of what will be processed
    with st.expander("📋 Processing summary", expanded=True):
        cols = st.columns(3)
        with cols[0]:
            st.markdown("**RTA File**")
            st.markdown(f"✅ `{uploaded_file.name}`")
        with cols[1]:
            st.markdown("**Item Master**")
            if item_master_file:
                st.markdown(f"✅ `{item_master_file.name}`")
            else:
                st.markdown("⚪ Skipped")
        with cols[2]:
            st.markdown("**Customer File**")
            if customer_file:
                st.markdown(f"✅ `{customer_file.name}`")
            else:
                st.markdown("⚪ Skipped")

    st.markdown("<br>", unsafe_allow_html=True)

    process_clicked = st.button(
        "🚀 Process File",
        type="primary",
        use_container_width=True,
    )

    if not process_clicked:
        return

    tmp_path = None
    item_master_path = None
    customer_path = None
    output_path = None

    try:
        # --- Save RTA ---
        with st.spinner("💾 Saving RTA file..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                while True:
                    chunk = uploaded_file.read(8 * 1024 * 1024)
                    if not chunk:
                        break
                    tmp.write(chunk)
                tmp_path = tmp.name

        # --- Save Item Master ---
        if item_master_file is not None:
            with st.spinner("💾 Saving Item Master file..."):
                suffix = os.path.splitext(item_master_file.name)[1].lower() or ".xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_im:
                    while True:
                        chunk = item_master_file.read(8 * 1024 * 1024)
                        if not chunk:
                            break
                        tmp_im.write(chunk)
                    item_master_path = tmp_im.name

        # --- Save Customer file ---
        if customer_file is not None:
            with st.spinner("💾 Saving Customer file..."):
                suffix = os.path.splitext(customer_file.name)[1].lower() or ".csv"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_c:
                    while True:
                        chunk = customer_file.read(8 * 1024 * 1024)
                        if not chunk:
                            break
                        tmp_c.write(chunk)
                    customer_path = tmp_c.name

        # --- Run pipeline ---
        with st.spinner("⚙️ Processing... This may take a few minutes for large files."):
            output_path, error_message = run_processing_pipeline(
                tmp_path,
                return_output_path=True,
                item_master_path=item_master_path,
                customer_path=customer_path,
            )

        if error_message:
            st.error(f"⚠️ {error_message}")
            if output_path and os.path.exists(output_path):
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Download File (processed until error)",
                        data=f.read(),
                        file_name=processed_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
        else:
            st.success("🎉 Processing complete! Your file is ready to download.")
            with open(output_path, "rb") as f:
                st.download_button(
                    label="⬇️ Download Processed File",
                    data=f.read(),
                    file_name=processed_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
    except Exception as e:
        st.error("❌ Something went wrong during processing.")
        st.exception(e)
    finally:
        for p in (tmp_path, item_master_path, customer_path):
            if p and os.path.exists(p):
                try:
                    os.remove(p)
                except:
                    pass
        gc.collect()