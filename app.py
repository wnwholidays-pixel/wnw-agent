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
    school_name = st.text_input("School Name:", "AA School")
    start_city = st.text_input("Starting From:", "Jaipur")
    destination_name = st.text_input("Destination Label:", "CHANDIGARH – MANALI")
    tour_duration = st.text_input("Duration Frame:", "6 Nights / 7 Days")
    st.header("💰 Pricing Matrix")
    student_cost = st.text_input("Cost Per Student:", "Rs. 13,500/-")
    group_strength = st.text_input("Group Strength:", "45 (Minimum)")
    teacher_ratio = st.text_input("Teacher Ratio:", "15:01")

pasted_itinerary = st.text_area("Pasted Itinerary Body Text:", height=450)

def append_styled_line(doc, curr_p, d_line):
    stripped = d_line.strip()
    if not stripped: return curr_p
    new_p = doc.add_paragraph()
    
    if stripped.upper().startswith("DAY ") and ":" in stripped:
        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        day_part, _ = stripped.split(":", 1)
        r1 = new_p.add_run(day_part.upper().strip())
        r1.font.name, r1.font.size, r1.font.bold = 'Arial', Pt(16), True
        r1.font.color.rgb = RGBColor(0, 86, 179)
        
    elif stripped.startswith("[SUB_ROUTE]"):
        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_sub = new_p.add_run(stripped.replace("[SUB_ROUTE]", "").strip())
        r_sub.font.name, r_sub.font.size, r_sub.font.bold = 'Arial', Pt(11), True
        r_sub.font.italic = False 
        r_sub.font.color.rgb = RGBColor(0, 86, 179)
        
    elif "---" in stripped:
        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_div = new_p.add_run(stripped)
        r_div.font.name, r_div.font.size = 'Arial', Pt(10)
        r_div.font.color.rgb = RGBColor(100, 100, 100)
    else:
        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if len(stripped) > 5 and stripped[0:2].isdigit() and ":" in stripped:
            space_idx = stripped.find(" ")
            r_time = new_p.add_run(stripped[:space_idx] + "\t")
            r_time.bold, r_time.font.name, r_time.font.size = True, 'Arial', Pt(11)
            r_text = new_p.add_run(stripped[space_idx:].strip())
            r_text.font.name, r_text.font.size = 'Arial', Pt(11)
        else:
            r_txt = new_p.add_run(stripped)
            r_txt.font.name, r_txt.font.size = 'Arial', Pt(11)
    curr_p._p.addnext(new_p._p)
    return new_p

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
                    elif current_mode == "inclusions": inclusions.append(line.strip())
                    elif current_mode == "exclusions": exclusions.append(line.strip())
                    else:
                        if line.strip(): detailed_lines.append(line)

            calculated_tour_route = ""
            if len(table_rows_data) > 0:
                route_cities = []
                for r_line in table_rows_data:
                    if len(r_line) > 1:
                        sub_cities = [c.strip() for c in str(r_line).upper().split('→') if c.strip()]
                        for sc in sub_cities:
                            if sc not in route_cities: route_cities.append(sc)
                if len(route_cities) > 0 and start_city.upper() not in route_cities[-1]:
                    route_cities.append(start_city.upper())
                calculated_tour_route = " → ".join(route_cities)

            for p in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in p.text: p.text = p.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in p.text: p.text = p.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in p.text: p.text = p.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in p.text: p.text = p.text.replace("{{TOUR_ROUTE}}", calculated_tour_route)
                
                if "{{TOUR_INCLUSIONS}}" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.text = p.text.replace("{{TOUR_INCLUSIONS}}", "")
                    curr_inc_p = p
                    for inc_line in inclusions:
                        if not inc_line: continue
                        new_inc = doc.add_paragraph(inc_line, style='List Bullet' if not inc_line.startswith('•') else 'Normal')
                        new_inc.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        curr_inc_p._p.addnext(new_inc._p)
                        curr_inc_p = new_inc

                if "{{TOUR_EXCLUSIONS}}" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.text = p.text.replace("{{TOUR_EXCLUSIONS}}", "")
                    curr_exc_p = p
                    for exc_line in exclusions:
                        if not exc_line: continue
                        new_exc = doc.add_paragraph(exc_line, style='List Bullet' if not exc_line.startswith('•') else 'Normal')
                        new_exc.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        curr_exc_p._p.addnext(new_exc._p)
                        curr_exc_p = new_exc

                if "{{DETAILED_ITINERARY}}" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.text = p.text.replace("{{DETAILED_ITINERARY}}", "")
                    curr_p = p
                    for d_line in detailed_lines:
                        curr_p = append_styled_line(doc, curr_p, d_line)

            for table in doc.tables:
                is_target, target_row_idx = False, -1
                for r_idx, row in enumerate(table.rows):
                    for cell in row.cells:
                        if "{{DAY_NUM}}" in cell.text:
                            is_target, target_row_idx = True, r_idx
                            break
                    if is_target: break
                
                if is_target and target_row_idx != -1:
                    for idx, row_data in enumerate(table_rows_data):
                        new_row = table.rows[target_row_idx] if idx == 0 else table.add_row()
                        for i in range(min(len(row_data), len(new_row.cells))): 
                            new_row.cells[i].text = row_data[i]
                
                for row in table.rows:
                    for cell in row.cells:
                        # FIXED COSTING DESIGN ENGINE: Replaces the tag and applies Size 14 Bold to the text cell
                        if "{{STUDENT_COST}}" in cell.text: 
                            cell.text = cell.text.replace("{{STUDENT_COST}}", "")
                            p_run = cell.paragraphs[0].add_run(student_cost)
                            p_run.font.name, p_run.font.size, p_run.font.bold = 'Arial', Pt(14), True
                        if "{{GROUP_STRENGTH}}" in cell.text: cell.text = cell.text.replace("{{GROUP_STRENGTH}}", group_strength)
                        if "{{TEACHER_RATIO}}" in cell.text: cell.text = cell.text.replace("{{TEACHER_RATIO}}", teacher_ratio)

            bio = io.BytesIO()
            doc.save(bio)
            st.success("🎉 Final Document compiled perfectly with size 14 bold pricing!")
            st.download_button(label="💾 Download Client Word Document (.docx)", data=bio.getvalue(), file_name=f"WNW_Itinerary_{school_name.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error merging template data strings: {str(e)}")
