import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

st.set_page_config(
    page_title="WNW Template Engine",
    layout="wide"
)

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption(
    "Official Production Studio - Final Master Template Merger Engine (.docx)"
)


# ============================================================
# SIDEBAR
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
# PASTE ITINERARY
# ============================================================

pasted_itinerary = st.text_area(
    "Pasted Itinerary Body Text:",
    height=450
)


# ============================================================
# REMOVE ALL STATIC DAY HEADINGS FROM TEMPLATE
# ============================================================
#
# THIS IS THE IMPORTANT FIX.
#
# The previous code only checked doc.paragraphs.
# Word can store paragraphs inside:
#
# - tables
# - text boxes
# - shapes
# - other XML containers
#
# This function searches the complete Word XML and removes
# ANY paragraph whose text is a DAY heading.
#
# Example:
#
# DAY 1 : DAY 1
# DAY 2 : DAY 2
# DAY 3 : NOV 14
#
# ALL are removed from the template before we insert
# the generated itinerary.
# ============================================================

def remove_all_template_day_headings(doc):

    day_heading_pattern = re.compile(
        r"^\s*DAY\s+\d+\s*(?::.*)?\s*$",
        re.IGNORECASE
    )

    # Search the complete document XML.
    # This includes paragraphs inside tables and text boxes.

    root = doc.element

    paragraphs_to_remove = []

    for paragraph_xml in root.iter():

        # Word paragraph XML tag
        if paragraph_xml.tag.endswith("}p"):

            # Collect all text inside this paragraph
            text_parts = []

            for text_node in paragraph_xml.iter():

                if text_node.tag.endswith("}t"):
                    if text_node.text:
                        text_parts.append(text_node.text)

            paragraph_text = "".join(text_parts).strip()

            if day_heading_pattern.match(
                paragraph_text
            ):

                paragraphs_to_remove.append(
                    paragraph_xml
                )

    # Remove all matching paragraphs
    for paragraph_xml in paragraphs_to_remove:

        parent = paragraph_xml.getparent()

        if parent is not None:

            parent.remove(
                paragraph_xml
            )


# ============================================================
# CREATE DAY HEADING
# ============================================================

def create_day_heading(
    doc,
    curr_p,
    day_number,
    date_text=""
):

    new_p = doc.add_paragraph()

    # CENTER ALIGN
    new_p.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    # --------------------------------------------------------
    # FINAL HEADING TEXT
    #
    # If there is NO date:
    # DAY 1
    #
    # If there IS a date:
    # DAY 1 : OCT 02
    # --------------------------------------------------------

    if date_text:

        heading_text = (
            f"DAY {day_number} : "
            f"{date_text.upper()}"
        )

    else:

        heading_text = (
            f"DAY {day_number}"
        )

    # --------------------------------------------------------
    # DAY HEADING RUN
    # --------------------------------------------------------

    run_day = new_p.add_run(
        heading_text
    )

    # FONT
    run_day.font.name = "Arial"

    # SIZE
    run_day.font.size = Pt(14)

    # BOLD
    run_day.font.bold = True

    # BLUE
    run_day.font.color.rgb = RGBColor(
        0,
        86,
        179
    )

    # --------------------------------------------------------
    # INSERT AFTER PREVIOUS PARAGRAPH
    # --------------------------------------------------------

    curr_p._p.addnext(
        new_p._p
    )

    return new_p


# ============================================================
# FORMAT ITINERARY LINE
# ============================================================

def append_styled_line(
    doc,
    curr_p,
    d_line
):

    stripped = d_line.strip()

    if not stripped:
        return curr_p


    # ========================================================
    # DAY HEADING
    # ========================================================

    day_match = re.match(
        r"^DAY\s+(\d+)",
        stripped,
        re.IGNORECASE
    )

    if day_match:

        day_number = day_match.group(1)

        date_text = ""

        # ----------------------------------------------------
        # Look for text after :
        # ----------------------------------------------------

        if ":" in stripped:

            after_colon = (
                stripped
                .split(":", 1)[1]
                .strip()
            )

            # If after colon is "DAY 1",
            # there is NO date.
            if not re.match(
                r"^DAY\s+\d+$",
                after_colon,
                re.IGNORECASE
            ):

                date_text = after_colon

        return create_day_heading(
            doc,
            curr_p,
            day_number,
            date_text
        )


    # ========================================================
    # CREATE NORMAL PARAGRAPH
    # ========================================================

    new_p = doc.add_paragraph()


    # ========================================================
    # SUB ROUTE
    # ========================================================

    if stripped.startswith(
        "[SUB_ROUTE]"
    ):

        new_p.alignment = (
            WD_ALIGN_PARAGRAPH.LEFT
        )

        route_text = (
            stripped
            .replace(
                "[SUB_ROUTE]",
                ""
            )
            .strip()
        )

        r_sub = new_p.add_run(
            route_text
        )

        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(11)
        r_sub.font.bold = True

        r_sub.font.color.rgb = RGBColor(
            0,
            86,
            179
        )


    # ========================================================
    # OVERNIGHT DIVIDER
    # ========================================================

    elif "Overnight Journey" in stripped:

        new_p.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER
        )

        r_div = new_p.add_run(
            stripped
        )

        r_div.font.name = "Arial"
        r_div.font.size = Pt(10)

        r_div.font.color.rgb = RGBColor(
            100,
            100,
            100
        )


    # ========================================================
    # TIMELINE
    # ========================================================

    else:

        new_p.alignment = (
            WD_ALIGN_PARAGRAPH.LEFT
        )

        time_match = re.match(
            r"^(\d{1,2}:\d{2}\s?(?:AM|PM))\s+(.*)$",
            stripped,
            re.IGNORECASE
        )

        if time_match:

            time_text = (
                time_match
                .group(1)
                .upper()
            )

            event_text = (
                time_match
                .group(2)
                .strip()
            )

            # TIME
            r_time = new_p.add_run(
                time_text + "\t"
            )

            r_time.font.name = "Arial"
            r_time.font.size = Pt(11)
            r_time.bold = True

            # EVENT TEXT
            r_text = new_p.add_run(
                event_text
            )

            r_text.font.name = "Arial"
            r_text.font.size = Pt(11)

        else:

            r_txt = new_p.add_run(
                stripped
            )

            r_txt.font.name = "Arial"
            r_txt.font.size = Pt(11)


    # ========================================================
    # INSERT AFTER PREVIOUS PARAGRAPH
    # ========================================================

    curr_p._p.addnext(
        new_p._p
    )

    return new_p


# ============================================================
# COMPILE
# ============================================================

if st.button(
    "Compile Official Word Proposal"
):

    if not pasted_itinerary:

        st.error(
            "Please paste the formatted text block from our chat first!"
        )

    else:

        try:

            # ==================================================
            # LOAD TEMPLATE
            # ==================================================

            doc = Document(
                "template.docx"
            )


            # ==================================================
            # CRITICAL FIX
            #
            # REMOVE EVERY STATIC DAY HEADING FROM TEMPLATE
            #
            # This removes:
            #
            # DAY 1 : DAY 1
            # DAY 2 : DAY 2
            # DAY 3 : DAY 3
            #
            # even if they are inside tables/text boxes.
            # ==================================================

            remove_all_template_day_headings(
                doc
            )


            # ==================================================
            # PARSE INPUT
            # ==================================================

            lines = pasted_itinerary.split(
                "\n"
            )

            table_rows_data = []
            detailed_lines = []
            inclusions = []
            exclusions = []

            current_mode = "detailed"


            # ==================================================
            # READ STRUCTURAL BLOCKS
            # ==================================================

            for line in lines:

                stripped_line = line.strip()


                if "TABLE_START" in stripped_line:

                    current_mode = "table"


                elif "TABLE_END" in stripped_line:

                    current_mode = "detailed"


                elif "INCLUSIONS_START" in stripped_line:

                    current_mode = "inclusions"


                elif "INCLUSIONS_END" in stripped_line:

                    current_mode = "detailed"


                elif "EXCLUSIONS_START" in stripped_line:

                    current_mode = "exclusions"


                elif "EXCLUSIONS_END" in stripped_line:

                    current_mode = "detailed"


                else:

                    # TABLE
                    if (
                        current_mode == "table"
                        and "|" in line
                    ):

                        splits = [
                            c.strip()
                            for c in line.split("|")
                            if c.strip()
                        ]

                        if (
                            len(splits) >= 2
                            and splits[0].lower()
                            != "day"
                        ):

                            table_rows_data.append(
                                splits
                            )


                    # INCLUSIONS
                    elif (
                        current_mode
                        == "inclusions"
                    ):

                        if stripped_line:

                            inclusions.append(
                                stripped_line
                            )


                    # EXCLUSIONS
                    elif (
                        current_mode
                        == "exclusions"
                    ):

                        if stripped_line:

                            exclusions.append(
                                stripped_line
                            )


                    # DETAILED
                    else:

                        if stripped_line:

                            detailed_lines.append(
                                stripped_line
                            )


            # ==================================================
            # CALCULATE ROUTE
            # ==================================================

            calculated_tour_route = ""

            if table_rows_data:

                route_cities = []

                for r_line in table_rows_data:

                    if len(r_line) > 1:

                        city_field = (
                            str(r_line[1])
                            .upper()
                        )

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

                                route_cities.append(
                                    sc
                                )

                if (
                    route_cities
                    and start_city.upper()
                    not in route_cities[-1]
                ):

                    route_cities.append(
                        start_city.upper()
                    )

                calculated_tour_route = (
                    " → ".join(
                        route_cities
                    )
                )


            # ==================================================
            # PROCESS MEALS
            # ==================================================

            processed_table_rows = []

            for r_line in table_rows_data:

                if len(r_line) >= 5:

                    combined_meals = (
                        " ".join(
                            r_line[4:]
                        )
                    )

                    processed_table_rows.append(
                        r_line[:4]
                        + [combined_meals]
                    )

                else:

                    processed_table_rows.append(
                        r_line
                    )


            # ==================================================
            # TEMPLATE PARAGRAPHS
            # ==================================================

            for p in doc.paragraphs:


                # SCHOOL
                if "{{SCHOOL_NAME}}" in p.text:

                    p.text = p.text.replace(
                        "{{SCHOOL_NAME}}",
                        school_name
                    )


                # DESTINATION
                if "{{DESTINATION_NAME}}" in p.text:

                    p.text = p.text.replace(
                        "{{DESTINATION_NAME}}",
                        destination_name.upper()
                    )


                # DURATION
                if "{{TOUR_DURATION}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_DURATION}}",
                        tour_duration
                    )


                # ROUTE
                if "{{TOUR_ROUTE}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_ROUTE}}",
                        calculated_tour_route
                    )


                # ==================================================
                # INCLUSIONS
                # ==================================================

                if "{{TOUR_INCLUSIONS}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_INCLUSIONS}}",
                        ""
                    )

                    p.alignment = (
                        WD_ALIGN_PARAGRAPH.LEFT
                    )

                    curr_inc_p = p

                    for inc_line in inclusions:

                        new_inc = (
                            doc.add_paragraph()
                        )

                        new_inc.alignment = (
                            WD_ALIGN_PARAGRAPH.LEFT
                        )

                        r_inc = new_inc.add_run(
                            inc_line
                        )

                        r_inc.font.name = "Arial"
                        r_inc.font.size = Pt(10)

                        curr_inc_p._p.addnext(
                            new_inc._p
                        )

                        curr_inc_p = new_inc


                # ==================================================
                # EXCLUSIONS
                # ==================================================

                if "{{TOUR_EXCLUSIONS}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_EXCLUSIONS}}",
                        ""
                    )

                    p.alignment = (
                        WD_ALIGN_PARAGRAPH.LEFT
                    )

                    curr_exc_p = p

                    for exc_line in exclusions:

                        new_exc = (
                            doc.add_paragraph()
                        )

                        new_exc.alignment = (
                            WD_ALIGN_PARAGRAPH.LEFT
                        )

                        r_exc = new_exc.add_run(
                            exc_line
                        )

                        r_exc.font.name = "Arial"
                        r_exc.font.size = Pt(10)

                        curr_exc_p._p.addnext(
                            new_exc._p
                        )

                        curr_exc_p = new_exc


                # ==================================================
                # DETAILED ITINERARY
                # ==================================================

                if "{{DETAILED_ITINERARY}}" in p.text:

                    # Clear placeholder
                    p.text = ""

                    p.alignment = (
                        WD_ALIGN_PARAGRAPH.LEFT
                    )

                    curr_p = p

                    for d_line in detailed_lines:

                        curr_p = append_styled_line(
                            doc,
                            curr_p,
                            d_line
                        )


            # ==================================================
            # TABLE PROCESSING
            # ==================================================

            for table in doc.tables:

                is_target = False
                target_row_idx = -1


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


                if (
                    is_target
                    and target_row_idx != -1
                ):

                    for idx, row_data in enumerate(
                        processed_table_rows
                    ):

                        if idx == 0:

                            new_row = (
                                table.rows[
                                    target_row_idx
                                ]
                            )

                        else:

                            new_row = (
                                table.add_row()
                            )


                        for i in range(
                            min(
                                len(row_data),
                                len(new_row.cells)
                            )
                        ):

                            cell_text = row_data[i]

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

                            new_row.cells[
                                i
                            ].text = cell_text


                # PRICING
                for row in table.rows:

                    for cell in row.cells:

                        if (
                            "{{STUDENT_COST}}"
                            in cell.text
                        ):

                            cell.text = (
                                cell.text.replace(
                                    "{{STUDENT_COST}}",
                                    ""
                                )
                            )

                            p_run = (
                                cell.paragraphs[
                                    0
                                ].add_run(
                                    student_cost
                                )
                            )

                            p_run.font.name = "Arial"
                            p_run.font.size = Pt(14)
                            p_run.font.bold = True


                        if (
                            "{{GROUP_STRENGTH}}"
                            in cell.text
                        ):

                            cell.text = (
                                cell.text.replace(
                                    "{{GROUP_STRENGTH}}",
                                    group_strength
                                )
                            )


                        if (
                            "{{TEACHER_RATIO}}"
                            in cell.text
                        ):

                            cell.text = (
                                cell.text.replace(
                                    "{{TEACHER_RATIO}}",
                                    teacher_ratio
                                )
                            )


            # ==================================================
            # SAVE DOCUMENT
            # ==================================================

            bio = io.BytesIO()

            doc.save(
                bio
            )

            st.success(
                "🎉 Final Document compiled perfectly!"
            )

            st.download_button(
                label=(
                    "💾 Download Client Word "
                    "Document (.docx)"
                ),
                data=bio.getvalue(),
                file_name=(
                    "WNW_Itinerary_"
                    + school_name.replace(
                        " ",
                        "_"
                    )
                    + ".docx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml.document"
                )
            )


        except Exception as e:

            st.error(
                "Error merging template data strings: "
                + str(e)
            )
