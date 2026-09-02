import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# Page Configuration with Official Wings 'N' Wheels Blue Identity
st.set_page_config(page_title="WNW Template Engine", layout="wide")
st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx)")

# 1. SIDEBAR INPUT CONTROLS
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

st.markdown("### 📋 Paste the AI Text Output Block Below:")
pasted_itinerary = st.text_area("Pasted Itinerary Body Text:", height=450)

if st.button("Compile Official Word Proposal"):
    if not pasted_itinerary:
        st.error("Please paste the formatted text block from our chat first!")
    else:
        try:
            # Load the untouchable master template from your GitHub project directory
            doc = Document("template.docx")
            lines = pasted_itinerary.split('\n')
            
            table_rows_data = []
            detailed_lines = []
            inclusions = []
            exclusions = []
            current_mode = "detailed"
            
            # 2. PARSE COPIED DATA SECTIONS ACCORDING TO SYSTEM HIGHLIGHT TAGS
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
                        splits = [c.strip() for c in line.split('|') if c.strip()]
                        if len(splits) >= 5: 
                            # Combines trailing blocks cleanly inside your Meal cell
                            fixed_row = splits[:4] + ["  ".join(splits[4:])]
                            table_rows_data.append(fixed_row)
                        else: 
                            table_rows_data.append(splits)
                elif current_mode == "inclusions": 
                    inclusions.append(line.strip())
                elif current_mode == "exclusions": 
                    exclusions.append(line.strip())
                else: 
                    detailed_lines.append(line)

            # 3. SWAP PARAGRAPH PLACEHOLDERS AND APPLY VISUAL TYPOGRAPHY DESIGN
            for p in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in p.text: 
                    p.text = p.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in p.text: 
                    p.text = p.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in p.text: 
                    p.text = p.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in p.text:
                    for r_line in table_rows_data:
                        if len(r_line) > 1 and ("DAY 1" in r_line.upper() or "NOV" in r_line.upper()): 
                            p.text = p.text.replace("{{TOUR_ROUTE}}", r_line.upper())
                
                # Dynamic list injection handler for inclusions
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

                # Dynamic list injection handler for exclusions
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

                # TIMELINE TRANSLATION LOOPS: Builds Centered Blue Headers & Aligned Timestamps
                if "{{DETAILED_ITINERARY}}" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.text = p.text.replace("{{DETAILED_ITINERARY}}", "")
                    curr_p = p
                    
                    for d_line in detailed_lines:
                        if not d_line.strip(): continue
                        new_p = doc.add_paragraph()
                        
                        # Style Day Headers (Centered, Bold, WNW Sky Blue Color)
                        if d_line.strip().upper().startswith("DAY ") and ":" in d_line:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            day_p, route_p = d_line.split(":", 1)
                            r1 = new_p.add_run(day_p.upper() + "\n")
                            r1.font.name, r1.font.size, r1.font.bold = 'Arial', Pt(14), True
                            r1.font.color.rgb = RGBColor(0, 163, 224)
                            r2 = new_p.add_run(route_p.strip().upper())
                            r2.font.name, r2.font.size, r2.font.bold = 'Arial', Pt(14), True
                        
                        # Style Sub-routes (Centered, Italic Blue)
                        elif d_line.strip().startswith("[SUB_ROUTE]"):
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r_sub = new_p.add_run(d_line.replace("[SUB_ROUTE]", "").strip())
                            r_sub.font.name, r_sub.font.size, r_sub.font.bold, r_sub.font.italic = 'Arial', Pt(10.5), True, True
                            r_sub.font.color.rgb = RGBColor(0, 163, 224)
                        
                        # Style Day Divider Lines
                        elif "---" in d_line:
                            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r_div = new_p.add_run(d_line)
                            r_div.font.name, r_div.font.size = 'Arial', Pt(10)
                            r_div.font.color.rgb = RGBColor(100, 100, 100)
                        
                        # Style Normal Left-Aligned Timestamps
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

            # 4. DATA MATRIX TABLE MULTIPLICATION LOOPS
            for table in doc.tables:
                is_target = False
                target_row_idx = -1
                
                # Locate your dynamic row index location without reference errors
                for r_idx, row in enumerate(table.rows):
                    for cell in row.cells:
                        if "{{DAY_NUM}}" in cell.text:
                            is_target = True
                            target_row_idx = r_idx
                            break
                    if is_target: break
                
                # Dynamically append rows cleanly beneath your table headers configuration
                if is_target and target_row_idx != -1:
                    for idx, row_data in enumerate(table_rows_data):
                        if idx == 0:
                            new_row = table.rows[target_row_idx]
                        else:
                            new_row = table.add_row()
                        for i in range(min(len(row_data), len(new_row.cells))):
                            new_row.cells[i].text = row_data[i]
                
