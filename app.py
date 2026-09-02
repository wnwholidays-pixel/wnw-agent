import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# Page Configuration with Official Wings 'N' Wheels Corporate Identity
st.set_page_config(page_title="WNW Template Engine", layout="wide")
st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx)")

# 1. SIDEBAR INPUT CONTROLS
with st.sidebar:
    st.header("📋 Booking Profile")
    school_name = st.text_input("School Name:", "AA School")
    start_city = st.text_input("Starting From:", "Jaipur")
    destination_name = st.text_input("Destination Label:", "CHANDIGARH – MANALI")
    tour_duration = st.text_input("Duration Frame:", "6 Nights / 7 Days")
    st.header("💰 Pricing Matrix")
    student_cost = st.text_input("Cost Per Student:", "**Rs. 13,500/-**")
    group_strength = st.text_input("Group Strength:", "45 (Minimum)")
    teacher_ratio = st.text_input("Teacher Ratio:", "15:01")

st.markdown("### 📋 Paste the AI Text Output Block Below:")
pasted_itinerary = st.text_area("Pasted Itinerary Body Text:", height=450)

def append_styled_line(doc, curr_p, d_line):
    stripped = d_line.strip()
    if not stripped: 
        return curr_p
    new_p = doc.add_paragraph()
    
    # FORMAT ANCHOR 1: CENTERED DAY HEADINGS (e.g. DAY 1)
    if stripped.upper().startswith("DAY ") and ":" in stripped:
        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        day_part, _ = stripped.split(":", 1)
        r1 = new_p.add_run(day_part.upper().strip())
        r1.font.name, r1.font.size, r1.font.bold = 'Arial', Pt(12), True
        r1.font.color.rgb = RGBColor(0, 86, 179) # WNW Deep Corporate Theme Blue
        
    # FORMAT ANCHOR 2: LEFT-ALIGNED SUB-ROUTE BANNER
    elif stripped.startswith("[SUB_ROUTE]"):
        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_sub = new_p.add_run(stripped.replace("[SUB_ROUTE]", "").strip())
        r_sub.font.name, r_sub.font.size, r_sub.font.bold, r_sub.font.italic = 'Arial', Pt(11), True, True
        r_sub.font.color.rgb = RGBColor(0, 86, 179)
        
    # FORMAT ANCHOR 3: PERFECTLY CENTERED TIMELINE DIVIDERS
    elif "---" in stripped:
        new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_div = new_p.add_run(stripped)
        r_div.font.name, r_div.font.size = 'Arial', Pt(10)
        r_div.font.color.rgb = RGBColor(100, 100, 100)
        
    # FORMAT ANCHOR 4: REGULAR LEFT-ALIGNED TIMESTAMPS & ENTRIES
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
            
            # 2. PARSE COPIED DATA SECTIONS ACCORDING TO BLOCK MARKERS
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
                else:
                    if current_mode == "table" and "|" in line:
                        splits = [c.strip() for c in line.split('|') if c.strip()]
                        if len(splits) >= 5: 
                            fixed_row = splits[:4] + ["  ".join(splits[4:])]
                            table_rows_data.append(fixed_row)
                        else: 
                            table_rows_data.append(splits)
                    elif current_mode == "inclusions": 
                        inclusions.append(line.strip())
                    elif current_mode == "exclusions": 
                        exclusions.append(line.strip())
                    else:
                        if line.strip(): 
                            detailed_lines.append(line)

            # 3. BUILD CLEAN COMPACT LOOPS ROUTE STRING FOR PAGE 1 BANNER
            calculated_tour_route = ""
            if len(table_rows_data) > 0:
                route_cities = []
                for r_line in table_rows_data:
                    if len(r_line) > 1:
                        # Safely read cell data index 1 which contains the route loops map
                        sub_cities = [c.strip() for c in r_line[1].upper().split('→') if c.strip()]
                        for sc in sub_cities:
                            if sc not in route_cities: 
                                route_cities.append(sc)
                if len(route_cities) > 0 and start_city.upper() in route_cities:
                    # Append the starting point back to the tail index to close the loop blueprint natively
                    route_cities.append(start_city.upper())
                calculated_tour_route = " → ".join(route_cities)

            # 4. SWAP TEXT PARAGRAPH LEVEL PLACEHOLDERS
            for p in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in p.text: 
                    p.text = p.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in p.text: 
                    p.text = p.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in p.text: 
                    p.text = p.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in p.text: 
                    p.text = p.text.replace("{{TOUR_ROUTE}}", calculated_tour_route)
                
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

                # Dynamic detailed itinerary handler line processing loops
                if "{{DETAILED_ITINERARY}}" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.text = p.text.replace("{{DETAILED_ITINERARY}}", "")
                    curr_p = p
                    for d_line in detailed_lines:
                        curr_p = append_styled_line(doc, curr_p, d_line)

            # 5. DATA MATRIX GRID TABLES SYNC CHANNELS
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
                        if "{{STUDENT_COST}}" in cell.text: 
                            cell.text = cell.text.replace("{{STUDENT_COST}}", student_cost)
                        if "{{GROUP_STRENGTH}}" in cell.text: 
                            cell.text = cell.text.replace("{{GROUP_STRENGTH}}", group_strength)
                        if "{{TEACHER_RATIO}}" in cell.text: 
                            cell.text = cell.text.replace("{{TEACHER_RATIO}}", teacher_ratio)

            # 6. EXPORT OUT BUFFER CHANNELS
