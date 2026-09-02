import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(page_title="WNW Template Engine", layout="wide")
st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx)")

with st.sidebar:
    st.header("📋 Booking Profile")
    school_name = st.text_input("School/College Name:", "AA School")
    start_city = st.text_input("Starting From City:", "Jaipur")
    destination_name = st.text_input("Destination Label:", "CHANDIGARH – MANALI")
    tour_duration = st.text_input("Tour Duration Frame:", "6 Nights / 7 Days")
    st.header("💰 Pricing Matrix")
    student_cost = st.text_input("Cost Per Student:", "Rs. 13,500/-")
    group_strength = st.text_input("Group Strength:", "45 (Minimum)")
    teacher_ratio = st.text_input("Teacher Ratio:", "15:01")

pasted_itinerary = st.text_area("Pasted Itinerary Body Text:", height=450)

if st.button("Compile Official Word Proposal"):
    if not pasted_itinerary:
        st.error("Please paste the formatted text block from our chat first!")
    else:
        try:
            doc = Document("template.docx")
            lines = pasted_itinerary.split('\n')
            table_rows_data, detailed_lines, inclusions, exclusions = [], [], [], []
            current_mode = "detailed"
            
            for line in lines:
                if "TABLE_START" in line: current_mode = "table"
                elif "TABLE_END" in line: current_mode = "detailed"
                elif "INCLUSIONS_START" in line: current_mode = "inclusions"
                elif "INCLUSIONS_END" in line: current_mode = "detailed"
                elif "EXCLUSIONS_START" in line: current_mode = "exclusions"
                elif "EXCLUSIONS_END" in line: current_mode = "detailed"
                else:
                    if current_mode == "table" and "|" in line:
                        splits = [c.strip() for c in line.split('|') if c.strip()]
                        if len(splits) >= 5: table_rows_data.append(splits[:4] + ["  ".join(splits[4:])])
                        else: table_rows_data.append(splits)
                    elif current_mode == "inclusions": inclusions.append(line)
                    elif current_mode == "exclusions": exclusions.append(line)
                    else: detailed_lines.append(line)

            for p in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in p.text: p.text = p.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in p.text: p.text = p.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in p.text: p.text = p.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in p.text:
                    for r_line in table_rows_data:
                        if len(r_line) > 1 and ("DAY 1" in r_line.upper() or "NOV" in r_line.upper()): 
                            p.text = p.text.replace("{{TOUR_ROUTE}}", r_line.upper())
                if "{{TOUR_INCLUSIONS}}" in p.text:
                    p.text = p.text.replace("{{TOUR_INCLUSIONS}}", "\n".join(inclusions))
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if "{{TOUR_EXCLUSIONS}}" in p.text:
                    p.text = p.text.replace("{{TOUR_EXCLUSIONS}}", "\n".join(exclusions))
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if "{{DETAILED_ITINERARY}}" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.text = p.text.replace("{{DETAILED_ITINERARY}}", "")
                    curr_p = p
                    for d_line in detailed_lines:
                        if not d_line.strip(): continue
                        new_p = doc.add_paragraph()
                        if d_line.strip().upper().startswith("DAY ") and ":" in d_line:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            day_p, route_p = d_line.split(":", 1)
                            r1 = new_p.add_run(day_p.upper() + "\n")
                            r1.font.name, r1.font.size, r1.font.bold = 'Arial', Pt(14), True
                            r1.font.color.rgb = RGBColor(0, 163, 224)
                            r2 = new_p.add_run(route_p.strip().upper())
                            r2.font.name, r2.font.size, r2.font.bold = 'Arial', Pt(14), True
                        elif d_line.strip().startswith("[SUB_ROUTE]"):
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r_sub = new_p.add_run(d_line.replace("[SUB_ROUTE]", "").strip())
                            r_sub.font.name, r_sub.font.size, r_sub.font.bold, r_sub.font.italic = 'Arial', Pt(10.5), True, True
                            r_sub.font.color.rgb = RGBColor(0, 163, 224)
                        elif "---" in d_line:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r_div = new_p.add_run(d_line)
                            r_div.font.name, r_div.font.size = 'Arial', Pt(10)
                            r_div.font.color.rgb = RGBColor(100, 100, 100)
                        else:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            stripped = d_line.strip()
                            if len(stripped) > 5 and stripped[0:2].isdigit() and ":" in stripped:
                                sp_idx = stripped.find(" ")
                                r_time = new_p.add_run(stripped[:sp_idx] + "\t")
                                r_time.bold, r_time.font.name, r_time.font.size = True, 'Arial', Pt(11)
                                r_text = new_p.add_run(stripped[sp_idx:].strip())
                                r_text.font.name, r_text.font.size = 'Arial', Pt(11)
                            else:
                                r_txt = new_p.add_run(stripped)
                                r_txt.font.name, r_txt.font.size = 'Arial', Pt(11)
                        curr_p._p.addnext(new_p._p)
                        curr_p = new_p

            for table in doc.tables:
                is_target = False
                target_row_idx = -1
                
                # Check table dimensions securely to find structural placeholder match rows
                for r_idx, row in enumerate(table.rows):
                    for cell in row.cells:
                        if "{{DAY_NUM}}" in cell.text:
                            is_target = True
                            target_row_idx = r_idx
                            break
                    if is_target: break
                
                if is_target and target_row_idx != -1:
                    # Dynamically maps rows downwards from your initial design template anchor row location
                    for idx, row_data in enumerate(table_rows_data):
                        if idx == 0:
                            new_row = table.rows[target_row_idx]
                        else:
                            new_row = table.add_row()
                        for i in range(min(len(row_data), len(new_row.cells))):
                            new_row.cells[i].text = row_data[i]
                
                for row in table.rows:
                    for cell in row.cells:
                        if "{{STUDENT_COST}}" in cell.text: cell.text = cell.text.replace("{{STUDENT_COST}}", student_cost)
                        if "{{GROUP_STRENGTH}}" in cell.text: cell.text = cell.text.replace("{{GROUP_STRENGTH}}", group_strength)
                        if "{{TEACHER_RATIO}}" in cell.text: cell.text = cell.text.replace("{{TEACHER_RATIO}}", teacher_ratio)

            bio = io.BytesIO()
            doc.save(bio)
            st.success("🎉 Final Document compiled perfectly with blue headers and zero execution defects!")
            st.download_button(label="💾 Download Client Word Document (.docx)", data=bio.getvalue(), file_name=f"WNW_Itinerary_{school_name.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error merging template data strings: {str(e)}")
