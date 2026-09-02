import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# Page Configuration with Official Blue Identity
st.set_page_config(page_title="WNW Master Template Engine", layout="wide")

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx)")

# Sidebar inputs
with st.sidebar:
    st.header("📋 Booking Profile")
    school_name = st.text_input("School/College Name:", "AA School")
    start_city = st.text_input("Starting From City:", "Jaipur")
    destination_name = st.text_input("Destination Label:", "CHANDIGARH – MANALI")
    tour_duration = st.text_input("Tour Duration Frame:", "6 Nights / 7 Days")
    
    st.header("💰 Pricing Grid Parameters")
    student_cost = st.text_input("Cost Per Student:", "Rs. 13,500/-")
    group_strength = st.text_input("Group Strength:", "45 (Minimum)")
    teacher_ratio = st.text_input("Teacher Ratio:", "15:01")

st.markdown("### 📋 Paste the AI Text Output Block Below:")
pasted_itinerary = st.text_area("Pasted Itinerary Body Text:", height=450)

if st.button("Compile Official Word Proposal"):
    if not pasted_itinerary:
        st.error("Please paste the formatted text block from our chat first!")
    else:
        try:
            doc = Document("template.docx")
            lines = pasted_itinerary.split('\n')
            
            table_rows_data = []
            detailed_itinerary_lines = []
            inclusions_data = []
            exclusions_data = []
            current_mode = "detailed"
            
            for line in lines:
                if "TABLE_START" in line:
                    current_mode = "table"
                    continue
                elif "TABLE_END" in line:
                    current_mode = "detailed"
                    continue
                elif "INCLUSIONS_START" in line:
                    current_mode = "inclusions"
                    continue
                elif "INCLUSIONS_END" in line:
                    current_mode = "detailed"
                    continue
                elif "EXCLUSIONS_START" in line:
                    current_mode = "exclusions"
                    continue
                elif "EXCLUSIONS_END" in line:
                    current_mode = "detailed"
                    continue
                
                if current_mode == "table":
                    if "|" in line:
                        raw_splits = [cell.strip() for cell in line.split('|') if cell.strip()]
                        if len(raw_splits) >= 5:
                            fixed_row = raw_splits[:4] + ["  ".join(raw_splits[4:])]
                            table_rows_data.append(fixed_row)
                        else:
                            table_rows_data.append(raw_splits)
                elif current_mode == "inclusions":
                    inclusions_data.append(line)
                elif current_mode == "exclusions":
                    exclusions_data.append(line)
                else:
                    detailed_itinerary_lines.append(line)

            # SWAP PARAGRAPH LEVEL TAGS WITH ADVANCED COLOR AND DESIGN LOGIC
            for paragraph in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in paragraph.text:
                    for line in table_rows_data:
                        if "Day 1" in line or "Nov 12" in line:
                            paragraph.text = paragraph.text.replace("{{TOUR_ROUTE}}", line.upper())
                
                if "{{TOUR_INCLUSIONS}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_INCLUSIONS}}", "\n".join(inclusions_data))
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if "{{TOUR_EXCLUSIONS}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_EXCLUSIONS}}", "\n".join(exclusions_data))
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                if "{{DETAILED_ITINERARY}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{DETAILED_ITINERARY}}", "")
                    current_para = paragraph
                    
                    for idx, d_line in enumerate(detailed_itinerary_lines):
                        stripped_line = d_line.strip()
                        if not stripped_line:
                            continue
                            
                        new_p = doc.add_paragraph()
                        
                        # FORMAT ANCHOR 1: STYLING THE DAY TITLES IN BLUE
                        if stripped_line.upper().startswith("DAY ") and ":" in stripped_line:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            day_part, route_part = stripped_line.split(":", 1)
                            
                            run_day = new_p.add_run(day_part.upper() + "\n")
                            run_day.font.name = 'Arial'
                            run_day.font.size = Pt(14)
                            run_day.font.bold = True
                            run_day.font.color.rgb = RGBColor(0, 163, 224) # Exact WNW Sky Blue
                            
                            run_route = new_p.add_run(route_part.strip().upper())
                            run_route.font.name = 'Arial'
                            run_route.font.size = Pt(14)
                            run_route.font.bold = True
                            run_route.font.color.rgb = RGBColor(0, 0, 0)
                            
                        # FORMAT ANCHOR 2: STYLING THE DRIVING SUB-ROUTES IN ITALIC BLUE
                        elif stripped_line.startswith("[SUB_ROUTE]"):
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_sub = new_p.add_run(stripped_line.replace("[SUB_ROUTE]", "").strip())
                            run_sub.font.name = 'Arial'
                            run_sub.font.size = Pt(10.5)
                            run_sub.font.bold = True
                            run_sub.font.italic = True
                            run_sub.font.color.rgb = RGBColor(0, 163, 224)
                            
                        # FORMAT ANCHOR 3: STYLING DIVIDER LINES
                        elif "---" in stripped_line:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_div = new_p.add_run(stripped_line)
                            run_div.font.name = 'Arial'
                            run_div.font.size = Pt(10)
                            run_div.font.color.rgb = RGBColor(100, 100, 100)
                            
                        # FORMAT ANCHOR 4: REGULAR TIMELINE ENTRIES
                        else:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            if len(stripped_line) > 5 and stripped_line[0:2].isdigit() and ":" in stripped_line:
                                space_idx = stripped_line.find(" ")
                                time_part = stripped_line[:space_idx]
                                text_part = stripped_line[space_idx:]
                                
                                run_time = new_p.add_run(time_part + "\t")
                                run_time.bold = True
                                run_time.font.name = 'Arial'
                                run_time.font.size = Pt(11)
                                
                                run_text = new_p.add_run(text_part.strip())
                                run_text.font.name = 'Arial'
                                run_text.font.size = Pt(11)
                            else:
                                run_txt = new_p.add_run(stripped_line)
                                run_txt.font.name = 'Arial'
                                run_txt.font.size = Pt(11)
                        
                        current_para._p.addnext(new_p._p)
                        current_para = new_p

            # SWAP TABLES GRID FOR HIGHLIGHTS AND PRICING
            for table in doc.tables:
                is_target_table = False
                if len(table.rows) > 1:
                    # Clean look across cells inside row 1 to see if the tag exists safely
                    for cell in table.rows[1].cells:
                        if "{{DAY_NUM}}" in cell.text:
                            is_target_table = True
                            break
                
                if is_target_table:
                    base_row = table.rows[1]
                    for index, row_data in enumerate(table_rows_data):
                        if index == 0:
                            new_row = base_row
                        else:
                            new_row = table.add_row()
                        for i in range(min(len(row_data), len(new_row.cells))):
                            new_row.cells[i].text = row_data[i]
                
                for row in table.rows:
                    for cell in row.cells:
                        if "{{STUDENT_COST}}" in cell.text:
                            cell.text = cell.text.replace("{{STUDENT_COST}}", student_cost)
                        if "{{GROUP_STRENGTH}}" in cell.text:
                            cell.text = cell.text.replace("{{GROUP_STRENGTH}}", group_strength)
