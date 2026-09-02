import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import os
import subprocess
import tempfile

# ============================================================
# PAGE SETUP
# ============================================================

st.set_page_config(
    page_title="WNW Template Engine",
    layout="wide"
)

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx + .pdf)")


# ============================================================
# SIDEBAR - BOOKING PROFILE
# ============================================================

with st.sidebar:
    st.header("📋 Booking Profile")

    school_name = st.text_input(
        "School Name:",
        "AA School"
    )

    start_city = st.text_input(
        "Starting From:",
        "Jaipur"
    )

    destination_name = st.text_input(
        "Destination Label:",
        "CHANDIGARH – MANALI"
    )

    tour_duration = st.text_input(
        "Duration Frame:",
        "6 Nights / 7 Days"
    )

    st.header("💰 Pricing Matrix")

    student_cost = st.text_input(
        "Cost Per Student:",
        "Rs. 13,500/-"
    )

    group_strength = st.text_input(
        "Group Strength:",
        "45 (Minimum)"
    )

    teacher_ratio = st.text_input(
        "Teacher Ratio:",
        "15:01"
    )


# ============================================================
# ITINERARY INPUT
# ============================================================

pasted_itinerary = st.text_area(
    "Pasted Itinerary Body Text:",
    height=550,
    placeholder="Paste the TABLE_START / INCLUSIONS_START / EXCLUSIONS_START / DAY itinerary here..."
)


# ============================================================
# HELPER - ADD STYLED ITINERARY LINE
# ============================================================

def append_styled_line(doc, curr_p, d_line):

    stripped = d_line.strip()

    if not stripped:
        return curr_p

    new_p = doc.add_paragraph()

    # --------------------------------------------------------
    # DAY HEADING
    # --------------------------------------------------------

    if stripped.upper().startswith("DAY ") and ":" in stripped:

        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        numbers = re.findall(r"\d+", stripped)
        day_num = numbers[0] if numbers else "1"

        run_day = new_p.add_run(f"DAY {day_num}")

        run_day.font.name = "Arial"
        run_day.font.size = Pt(14)
        run_day.font.bold = True
        run_day.font.color.rgb = RGBColor(0, 163, 224)

    # --------------------------------------------------------
    # SUB ROUTE
    # --------------------------------------------------------

    elif stripped.startswith("[SUB_ROUTE]"):

        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        route_text = stripped.replace(
            "[SUB_ROUTE]",
            ""
        ).strip()

        r_sub = new_p.add_run(route_text)

        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(11)
        r_sub.font.bold = True
        r_sub.font.color.rgb = RGBColor(0, 86, 179)

    # --------------------------------------------------------
    # OVERNIGHT DIVIDER
    # --------------------------------------------------------

    elif "Overnight Journey" in stripped or "---" in stripped:

        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r_div = new_p.add_run(stripped)

        r_div.font.name = "Arial"
        r_div.font.size = Pt(10)
        r_div.font.color.rgb = RGBColor(100, 100, 100)

    # --------------------------------------------------------
    # TIMELINE EVENTS
    # --------------------------------------------------------

    else:

        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Detect time such as:
        # 08:00 AM
        # 10:30 PM
        # 21:00 PM

        time_match = re.match(
            r"^(\d{1,2}:\d{2}\s?(?:AM|PM))\s+(.*)$",
            stripped,
            re.IGNORECASE
        )

        if time_match:

            time_text = time_match.group(1)
            event_text = time_match.group(2)

            r_time = new_p.add_run(time_text + "\t")

            r_time.font.name = "Arial"
            r_time.font.size = Pt(11)
            r_time.font.bold = True

            r_text = new_p.add_run(event_text)

            r_text.font.name = "Arial"
            r_text.font.size = Pt(11)

        else:

            r_txt = new_p.add_run(stripped)

            r_txt.font.name = "Arial"
            r_txt.font.size = Pt(11)

    # Insert directly after current paragraph
    curr_p._p.addnext(new_p._p)

    return new_p


# ============================================================
# PDF CONVERSION FUNCTION
# ============================================================

def convert_docx_to_pdf(docx_bytes):

    with tempfile.TemporaryDirectory() as temp_dir:

        docx_path = os.path.join(
            temp_dir,
            "WNW_Itinerary.docx"
        )

        pdf_path = os.path.join(
            temp_dir,
            "WNW_Itinerary.pdf"
        )

        # Save temporary Word file
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        # Convert Word → PDF using LibreOffice
        process = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                temp_dir,
                docx_path
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        # Check conversion
        if process.returncode != 0:
            raise RuntimeError(
                process.stderr or
                "LibreOffice could not convert the document."
            )

        if not os.path.exists(pdf_path):
            raise RuntimeError(
                "PDF file was not created."
            )

        # Read PDF
        with open(pdf_path, "rb") as f:
            return f.read()


# ============================================================
# COMPILE BUTTON
# ============================================================

if st.button(
    "🚀 Compile Official Word + PDF Proposal",
    type="primary"
):

    if not pasted_itinerary.strip():

        st.error(
            "Please paste the formatted itinerary text block first!"
        )

    else:

        try:

            # =================================================
            # LOAD MASTER TEMPLATE
            # =================================================

            doc = Document("template.docx")

            lines = pasted_itinerary.splitlines()

            table_rows_data = []
            detailed_lines = []
            inclusions = []
            exclusions = []

            current_mode = "detailed"


            # =================================================
            # PARSE INPUT BLOCKS
            # =================================================

            for line in lines:

                stripped_line = line.strip()

                if "TABLE_START" in stripped_line:

                    current_mode = "table"

                    continue

                elif "TABLE_END" in stripped_line:

                    current_mode = "detailed"

                    continue

                elif "INCLUSIONS_START" in stripped_line:

                    current_mode = "inclusions"

                    continue

                elif "INCLUSIONS_END" in stripped_line:

                    current_mode = "detailed"

                    continue

                elif "EXCLUSIONS_START" in stripped_line:

                    current_mode = "exclusions"

                    continue

                elif "EXCLUSIONS_END" in stripped_line:

                    current_mode = "detailed"

                    continue


                # ---------------------------------------------
                # SUMMARY TABLE
                # ---------------------------------------------

                if current_mode == "table":

                    if "|" in stripped_line:

                        splits = [
                            c.strip()
                            for c in stripped_line.split("|")
                            if c.strip()
                        ]

                        # Ignore accidental header row
                        if (
                            len(splits) >= 2
                            and splits[0].lower() != "day"
                        ):
                            table_rows_data.append(splits)


                # ---------------------------------------------
                # INCLUSIONS
                # ---------------------------------------------

                elif current_mode == "inclusions":

                    if stripped_line:
                        inclusions.append(stripped_line)


                # ---------------------------------------------
                # EXCLUSIONS
                # ---------------------------------------------

                elif current_mode == "exclusions":

                    if stripped_line:
                        exclusions.append(stripped_line)


                # ---------------------------------------------
                # DETAILED ITINERARY
                # ---------------------------------------------

                else:

                    if stripped_line:
                        detailed_lines.append(stripped_line)


            # =================================================
            # CALCULATE TOUR ROUTE
            # =================================================

            calculated_tour_route = ""

            if len(table_rows_data) > 0:

                route_cities = []

                for r_line in table_rows_data:

                    if len(r_line) > 1:

                        city_field = str(
                            r_line[1]
                        ).upper()

                        cleaned_cell = (
                            city_field
                            .replace("[", "")
                            .replace("]", "")
                            .replace("'", "")
                            .replace('"', "")
                        )

                        sub_cities = [
                            c.strip()
                            for c in cleaned_cell.split("→")
                            if c.strip()
                        ]

                        for sc in sub_cities:

                            if sc not in route_cities:

                                route_cities.append(sc)


                # Add starting city at end if required
                if (
                    len(route_cities) > 0
                    and start_city.upper() not in route_cities[-1]
                ):

                    route_cities.append(
                        start_city.upper()
                    )


                calculated_tour_route = " → ".join(
                    route_cities
                )


            # =================================================
            # PROCESS SUMMARY TABLE
            # =================================================

            processed_table_rows = []

            for r_line in table_rows_data:

                if len(r_line) >= 5:

                    # Everything after first 4 columns
                    # belongs to meals
                    combined_meals = " | ".join(
                        r_line[4:]
                    )

                    processed_table_rows.append(
                        r_line[:4] + [combined_meals]
                    )

                else:

                    processed_table_rows.append(
                        r_line
                    )


            # =================================================
            # REPLACE TEMPLATE PARAGRAPH VARIABLES
            # =================================================

            for p in doc.paragraphs:

                # ---------------------------------------------
                # SCHOOL NAME
                # ---------------------------------------------

                if "{{SCHOOL_NAME}}" in p.text:

                    p.text = p.text.replace(
                        "{{SCHOOL_NAME}}",
                        school_name
                    )


                # ---------------------------------------------
                # DESTINATION
                # ---------------------------------------------

                if "{{DESTINATION_NAME}}" in p.text:

                    p.text = p.text.replace(
                        "{{DESTINATION_NAME}}",
                        destination_name.upper()
                    )


                # ---------------------------------------------
                # DURATION
                # ---------------------------------------------

                if "{{TOUR_DURATION}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_DURATION}}",
                        tour_duration
                    )


                # ---------------------------------------------
                # ROUTE
                # ---------------------------------------------

                if "{{TOUR_ROUTE}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_ROUTE}}",
                        calculated_tour_route
                    )


                # ---------------------------------------------
                # INCLUSIONS
                # ---------------------------------------------

                if "{{TOUR_INCLUSIONS}}" in p.text:

                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    p.text = p.text.replace(
                        "{{TOUR_INCLUSIONS}}",
                        ""
                    )

                    curr_inc_p = p

                    for inc_line in inclusions:

                        if not inc_line:
                            continue

                        # Remove existing bullet if supplied
                        clean_inc = re.sub(
                            r"^•\s*",
                            "",
                            inc_line
                        )

                        new_inc = doc.add_paragraph(
                            clean_inc,
                            style="List Bullet"
                        )

                        new_inc.alignment = (
                            WD_ALIGN_PARAGRAPH.LEFT
                        )

                        for run in new_inc.runs:

                            run.font.name = "Arial"
                            run.font.size = Pt(10.5)

                        curr_inc_p._p.addnext(
                            new_inc._p
                        )

                        curr_inc_p = new_inc


                # ---------------------------------------------
                # EXCLUSIONS
                # ---------------------------------------------

                if "{{TOUR_EXCLUSIONS}}" in p.text:

                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    p.text = p.text.replace(
                        "{{TOUR_EXCLUSIONS}}",
                        ""
                    )

                    curr_exc_p = p

                    for exc_line in exclusions:

                        if not exc_line:
                            continue

                        clean_exc = re.sub(
                            r"^•\s*",
                            "",
                            exc_line
                        )

                        new_exc = doc.add_paragraph(
                            clean_exc,
                            style="List Bullet"
                        )

                        new_exc.alignment = (
                            WD_ALIGN_PARAGRAPH.LEFT
                        )

                        for run in new_exc.runs:

                            run.font.name = "Arial"
                            run.font.size = Pt(10.5)

                        curr_exc_p._p.addnext(
                            new_exc._p
                        )

                        curr_exc_p = new_exc


                # ---------------------------------------------
                # DETAILED ITINERARY
                # ---------------------------------------------

                if "{{DETAILED_ITINERARY}}" in p.text:

                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    p.text = p.text.replace(
                        "{{DETAILED_ITINERARY}}",
                        ""
                    )

                    curr_p = p

                    for d_line in detailed_lines:

                        curr_p = append_styled_line(
                            doc,
                            curr_p,
                            d_line
                        )


            # =================================================
            # PROCESS TABLES
            # =================================================

            for table in doc.tables:

                is_target = False
                target_row_idx = -1


                # ---------------------------------------------
                # FIND ITINERARY SUMMARY TABLE
                # ---------------------------------------------

                for r_idx, row in enumerate(
                    table.rows
                ):

                    for cell in row.cells:

                        if "{{DAY_NUM}}" in cell.text:

                            is_target = True
                            target_row_idx = r_idx

                            break

                    if is_target:
                        break


                # ---------------------------------------------
                # INSERT SUMMARY ROWS
                # ---------------------------------------------

                if (
                    is_target
                    and target_row_idx != -1
                ):

                    for idx, row_data in enumerate(
                        processed_table_rows
                    ):

                        if idx == 0:

                            new_row = table.rows[
                                target_row_idx
                            ]

                        else:

                            new_row = table.add_row()


                        for i in range(
                            min(
                                len(row_data),
                                len(new_row.cells)
                            )
                        ):

                            cell_text = row_data[i]

                            # Stack meals vertically
                            if i == 4:

                                cell_text = (
                                    cell_text
                                    .replace(
                                        "L:",
                                        "\nL:"
                                    )
                                    .replace(
                                        "D:",
                                        "\nD:"
                                    )
                                )

                            new_row.cells[i].text = (
                                cell_text
                            )


                            # ---------------------------------
                            # FORMAT TABLE CELL
                            # ---------------------------------

                            for paragraph in (
                                new_row.cells[i]
                                .paragraphs
                            ):

                                for run in paragraph.runs:

                                    run.font.name = "Arial"
                                    run.font.size = Pt(9)


                # ---------------------------------------------
                # OTHER TEMPLATE TABLE VARIABLES
                # ---------------------------------------------

                for row in table.rows:

                    for cell in row.cells:

                        # -------------------------------------
                        # STUDENT COST
                        # -------------------------------------

                        if "{{STUDENT_COST}}" in cell.text:

                            cell.text = cell.text.replace(
                                "{{STUDENT_COST}}",
                                ""
                            )

                            p_run = (
                                cell.paragraphs[0]
                                .add_run(student_cost)
                            )

                            p_run.font.name = "Arial"
                            p_run.font.size = Pt(14)
                            p_run.font.bold = True


                        # -------------------------------------
                        # GROUP STRENGTH
                        # -------------------------------------

                        if "{{GROUP_STRENGTH}}" in cell.text:

                            cell.text = cell.text.replace(
                                "{{GROUP_STRENGTH}}",
                                group_strength
                            )


                        # -------------------------------------
                        # TEACHER RATIO
                        # -------------------------------------

                        if "{{TEACHER_RATIO}}" in cell.text:

                            cell.text = cell.text.replace(
                                "{{TEACHER_RATIO}}",
                                teacher_ratio
                            )


            # =================================================
            # SAVE FINAL WORD DOCUMENT IN MEMORY
            # =================================================

            bio = io.BytesIO()

            doc.save(bio)

            docx_data = bio.getvalue()


            # =================================================
            # SUCCESS MESSAGE
            # =================================================

            st.success(
                "🎉 Final Word document compiled successfully!"
            )


            # =================================================
            # PDF CONVERSION
            # =================================================

            try:

                pdf_data = convert_docx_to_pdf(
                    docx_data
                )

                st.success(
                    "📄 PDF version created successfully!"
                )


                # ---------------------------------------------
                # DOWNLOAD BUTTONS
                # ---------------------------------------------

                col1, col2 = st.columns(2)


                with col1:

                    st.download_button(
                        label="💾 Download Word Document (.docx)",
                        data=docx_data,
                        file_name=(
                            f"WNW_Itinerary_"
                            f"{school_name.replace(' ', '_')}.docx"
                        ),
                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        )
                    )


                with col2:

                    st.download_button(
                        label="📄 Download PDF Document (.pdf)",
                        data=pdf_data,
                        file_name=(
                            f"WNW_Itinerary_"
                            f"{school_name.replace(' ', '_')}.pdf"
                        ),
                        mime="application/pdf"
                    )


            # =================================================
            # WORD SUCCESS BUT PDF FAILURE
            # =================================================

            except Exception as pdf_error:

                st.warning(
                    "Word document was created successfully, "
                    "but PDF conversion failed."
                )

                st.error(
                    f"PDF conversion error: {str(pdf_error)}"
                )


                # Still allow Word download

                st.download_button(
                    label="💾 Download Word Document (.docx)",
                    data=docx_data,
                    file_name=(
                        f"WNW_Itinerary_"
                        f"{school_name.replace(' ', '_')}.docx"
                    ),
                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    )
                )


        # =====================================================
        # GENERAL ERROR
        # =====================================================

        except Exception as e:

            st.error(
                f"Error merging template data strings: {str(e)}"
            )
