#!/usr/bin/env python3
"""
Streamlit front end for mpp_to_excel.py.

Lets a user upload a Microsoft Project (.mpp) file, converts it to the
Project for the Web "Excel Import Template" (Project / Resources / Tasks
sheets, Tasks wrapped in a table named "Estimates"), and offers the result
as a downloadable .xlsx.

This file is the Streamlit ENTRY POINT. In your Streamlit Cloud app
settings, "Main file path" should point to this file (streamlit_app.py),
not mpp_to_excel.py — mpp_to_excel.py is just the conversion library and
has no Streamlit UI code, which is why the app was hanging before.
"""

import os
import tempfile

import streamlit as st

from mpp_to_excel import (
    read_project,
    build_rows,
    project_estimated_start,
    write_excel,
)

st.set_page_config(page_title="MPP to Excel Import Template", page_icon="📅")

st.title("📅 MPP → Excel Import Template")
st.write(
    "Upload a Microsoft Project (`.mpp`) file and download it as the "
    "Project / Resources / Tasks Excel import template."
)

uploaded_file = st.file_uploader("Choose an .mpp file", type=["mpp"])

if uploaded_file is not None:
    project_name = os.path.splitext(uploaded_file.name)[0]

    if st.button("Convert", type="primary"):
        with st.spinner("Reading project and converting…"):
            try:
                # mpxj's UniversalProjectReader needs a real file path, so
                # write the upload to a temp file first.
                with tempfile.NamedTemporaryFile(
                    suffix=".mpp", delete=False
                ) as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    tmp_in_path = tmp_in.name

                project = read_project(tmp_in_path)
                rows, resources = build_rows(project)
                project_start = project_estimated_start(project, rows)

                out_path = tempfile.NamedTemporaryFile(
                    suffix=".xlsx", delete=False
                ).name
                write_excel(
                    rows,
                    resources,
                    out_path,
                    project_name=project_name,
                    project_start=project_start,
                )

                with open(out_path, "rb") as f:
                    xlsx_bytes = f.read()

            except Exception as e:
                st.error(f"Conversion failed: {e}")
                st.exception(e)
            else:
                st.success(
                    f"Converted {len(rows)} task rows and "
                    f"{len(resources)} resources."
                )
                if project_start is not None:
                    st.write(f"**Estimated Start Date:** {project_start}")

                st.download_button(
                    label="⬇️ Download Excel Import Template",
                    data=xlsx_bytes,
                    file_name=f"{project_name}_Import.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".spreadsheetml.sheet"
                    ),
                )
            finally:
                # Clean up temp files
                for p in (tmp_in_path, out_path):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
else:
    st.info("Upload a .mpp file to get started.")
