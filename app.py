import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

st.set_page_config(page_title="WNW Template Engine", layout="wide")
st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx)")

with st.sidebar:
    st.header("📋 Booking Profile")
    school_name = st.text_input("School Name:", "AA School")
    start_city = st.text_input("Starting From:", "Jaipur")
    destination_name = st.text_input("Destination Label:", "CHANDIGARH – MANALI")
    tour_duration = st.text_input("Duration Frame:", "6 Nights / 7 Days")

    st.header("💰 Pricing Matrix")
    student_cost = st.text_input("Cost Per Student:", "Rs. 13,500/-")
    group_strength = st.text_input("Group Strength:", "45 (Minimum)")
    teacher_ratio = st.text_input("Teacher Ratio:", "15:01")

pasted_itinerary = st.text_area(
    "Pasted Itinerary Body Text:",
    height=450
)


# ============================================================
# DETAILED ITINERARY LINE FORMATTER
# ============================================================

def append_styled_line(doc, curr_p, d_line):

    stripped = d_line.strip()

    if not stripped:
        return curr_p

    new_p = doc.add_paragraph()

    # --------------------------------------------------------
    # DAY HEADING
    # Example input:
    # DAY 1 : DAY 1
    # DAY 1 : NOV 12
    # DAY 2 : NOV 13
    #
    # Output:
    # DAY 1
    # DAY 1 : NOV 12
    # --------------------------------------------------------
    if re.match(r"^DAY\s+\d+", stripped, re.IGNORECASE):

        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Extract day number
        day_match = re.search(r"DAY\s+(\d+)", stripped, re.IGNORECASE)

        day_num = day_match.group(1) if day_match else "1"

        # Try to extract actual date after colon
        date_part = ""

        if ":" in stripped:
            after_colon = stripped.split(":", 1)[1].strip()

            # If the text after colon is NOT another "DAY X",
            # treat it as the actual date.
            if not re.match(r"^DAY\s+\d+$", after_colon, re.IGNORECASE):
                date_part = after_colon

        # Build final heading
        if date_part:
            heading_text = f"DAY {day_num} : {date_part.upper()}"
        else:
            heading_text = f"DAY {day_num}"

        run_day = new_p.add_run(heading_text)

        run_day.font.name = "Arial"
        run_day.font.size = Pt(14)
        run_day.font.bold = True
        run_day.font.color.rgb = RGBColor(0, 86, 179)

    # --------------------------------------------------------
    # SUB ROUTE
    # --------------------------------------------------------
    elif stripped.startswith("[SUB_ROUTE]"):

        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        route_text = stripped.replace("[SUB_ROUTE]", "").strip()

        r_sub = new_p.add_run(route_text)

        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(11)
        r_sub.font.bold = True
        r_sub.font.color.rgb = RGBColor(0, 86, 179)

    # --------------------------------------------------------
    # OVERNIGHT DIVIDER
    # --------------------------------------------------------
    elif "---" in stripped:

        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r_div = new_p.add_run(stripped)

        r_div.font.name = "Arial"
        r_div.font.size = Pt(10)
        r_div.font.color.rgb = RGBColor(100, 100, 100)

    # --------------------------------------------------------
    # TIMELINE EVENT
    # Example:
    # 08:30 AM Assemble and depart...
    # --------------------------------------------------------
    else:

        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Detect 12-hour timestamp
        time_match = re.match(
            r"^(\d{1,2}:\d{2}\s?(?:AM|PM))\s+(.*)$",
            stripped,
            re.IGNORECASE
        )

        if time_match:

            time_text = time_match.group(1).upper()
            event_text = time_match.group(2).strip()

            r_time = new_p.add_run(time_text + "\t")

            r_time.bold = True
            r_time.font.name = "Arial"
            r_time.font.size = Pt(11)

            r_text = new_p.add_run(event_text)

            r_text.font.name = "Arial"
            r_text.font.size = Pt(11)

        else:

            r_txt = new_p.add_run(stripped)

            r_txt.font.name = "Arial"
            r_txt.font.size = Pt(11)

    # Insert directly after previous paragraph
    curr_p._p.addnext(new_p._p)

    return new_p


# ============================================================
# COMPILE DOCUMENT
# ============================================================

if st.button("Compile Official Word Proposal"):

    if not pasted_itinerary:

        st.error("Please paste the formatted text block from our chat first!")

    else:

        try:

            doc = Document("template.docx")

            lines = pasted_itinerary.split("\n")

            table_rows_data = []
            detailed_lines = []
            inclusions = []
            exclusions = []

            current_mode = "detailed"

            # ====================================================
            # PARSE INPUT BLOCKS
            # ====================================================

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

                    if current_mode == "table" and "|" in line:

                        splits = [
                            c.strip()
                            for c in line.split("|")
                            if c.strip()
                        ]

                        # Ignore the table header
                        if (
                            len(splits) >= 2
                            and splits[0].lower() != "day"
                        ):
                            table_rows_data.append(splits)

                    elif current_mode == "inclusions":

                        if stripped_line:
                            inclusions.append(stripped_line)

                    elif current_mode == "exclusions":

                        if stripped_line:
                            exclusions.append(stripped_line)

                    else:

                        if stripped_line:
                            detailed_lines.append(stripped_line)


            # ====================================================
            # CALCULATE TOUR ROUTE
            # ====================================================

            calculated_tour_route = ""

            if len(table_rows_data) > 0:

                route_cities = []

                for r_line in table_rows_data:

                    if len(r_line) > 1:

                        city_field = str(r_line[1]).upper()

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

                if (
                    len(route_cities) > 0
                    and start_city.upper() not in route_cities[-1]
                ):

                    route_cities.append(start_city.upper())

                calculated_tour_route = " → ".join(route_cities)


            # ====================================================
            # PROCESS TABLE ROWS
            # ====================================================

            processed_table_rows = []

            for r_line in table_rows_data:

                if len(r_line) >= 5:

                    combined_meals = " ".join(r_line[4:])

                    processed_table_rows.append(
                        r_line[:4] + [combined_meals]
                    )

                else:

                    processed_table_rows.append(r_line)


            # ====================================================
            # REPLACE NORMAL TEMPLATE PLACEHOLDERS
            # ====================================================

            for p in doc.paragraphs:

                if "{{SCHOOL_NAME}}" in p.text:

                    p.text = p.text.replace(
                        "{{SCHOOL_NAME}}",
                        school_name
                    )


                if "{{DESTINATION_NAME}}" in p.text:

                    p.text = p.text.replace(
                        "{{DESTINATION_NAME}}",
                        destination_name.upper()
                    )


                if "{{TOUR_DURATION}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_DURATION}}",
                        tour_duration
                    )


                if "{{TOUR_ROUTE}}" in p.text:

                    p.text = p.text.replace(
                        "{{TOUR_ROUTE}}",
                        calculated_tour_route
                    )


                # ====================================================
                # INCLUSIONS
                # ====================================================

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

                        new_inc = doc.add_paragraph()

                        new_inc.alignment = WD_ALIGN_PARAGRAPH.LEFT

                        r_inc = new_inc.add_run(inc_line)

                        r_inc.font.name = "Arial"
                        r_inc.font.size = Pt(10)

                        curr_inc_p._p.addnext(new_inc._p)

                        curr_inc_p = new_inc


                # ====================================================
                # EXCLUSIONS
                # ====================================================

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

                        new_exc = doc.add_paragraph()

                        new_exc.alignment = WD_ALIGN_PARAGRAPH.LEFT

                        r_exc = new_exc.add_run(exc_line)

                        r_exc.font.name = "Arial"
                        r_exc.font.size = Pt(10)

                        curr_exc_p._p.addnext(new_exc._p)

                        curr_exc_p = new_exc


                # ====================================================
                # DETAILED ITINERARY
                # ====================================================

                if "{{DETAILED_ITINERARY}}" in p.text:

                    # IMPORTANT:
                    # Completely clear the existing placeholder
                    # paragraph so template text such as
                    # "DAY 1 :" cannot remain and cause:
                    #
                    # DAY 1 : DAY 1
                    #
                    p.text = ""

                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    curr_p = p

                    for d_line in detailed_lines:

                        curr_p = append_styled_line(
                            doc,
                            curr_p,
                            d_line
                        )


            # ====================================================
            # PROCESS TABLES
            # ====================================================

            for table in doc.tables:

                is_target = False
                target_row_idx = -1

                for r_idx, row in enumerate(table.rows):

                    for cell in row.cells:

                        if "{{DAY_NUM}}" in cell.text:

                            is_target = True
                            target_row_idx = r_idx

                            break

                    if is_target:
                        break


                if is_target and target_row_idx != -1:

                    for idx, row_data in enumerate(
                        processed_table_rows
                    ):

                        if idx == 0:

                            new_row = table.rows[target_row_idx]

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
                                    .replace("L:", "\nL:")
                                    .replace("D:", "\nD:")
                                )

                            new_row.cells[i].text = cell_text


                # ====================================================
                # PRICING PLACEHOLDERS
                # ====================================================

                for row in table.rows:

                    for cell in row.cells:

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


                        if "{{GROUP_STRENGTH}}" in cell.text:

                            cell.text = cell.text.replace(
                                "{{GROUP_STRENGTH}}",
                                group_strength
                            )


                        if "{{TEACHER_RATIO}}" in cell.text:

                            cell.text = cell.text.replace(
                                "{{TEACHER_RATIO}}",
                                teacher_ratio
                            )


            # ====================================================
            # SAVE DOCUMENT
            # ====================================================

            bio = io.BytesIO()

            doc.save(bio)

            st.success(
                "🎉 Final Document compiled perfectly!"
            )

            st.download_button(
                label="💾 Download Client Word Document (.docx)",
                data=bio.getvalue(),
                file_name=(
                    f"WNW_Itinerary_"
                    f"{school_name.replace(' ', '_')}.docx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            )


        except Exception as e:

            st.error(
                f"Error merging template data strings: {str(e)}"
            )
