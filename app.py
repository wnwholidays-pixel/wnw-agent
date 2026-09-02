import streamlit as st
from docx import Document
import io

# Page Configuration with Official Blue Identity
st.set_page_config(page_title="WNW Master Template Engine", layout="wide")

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Production Studio - Final Master Template Merger Engine (.docx)")

# 1. SIDEBAR INPUT CONTROLS FOR DIRECT BALANCING
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
            
            # Internal routing matrix tags
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
                        table_rows_data.append([cell.strip() for cell in line.split('|')])
                elif current_mode == "inclusions":
                    inclusions_data.append(line)
                elif current_mode == "exclusions":
                    exclusions_data.append(line)
                else:
                    detailed_itinerary_lines.append(line)
            
            clean_body_text = "\n".join(detailed_itinerary_lines)
            clean_inclusions = "\n".join(inclusions_data)
            clean_exclusions = "\n".join(exclusions_data)

            # SWAP PARAGRAPH LEVEL TAGS
            for paragraph in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in paragraph.text:
                    for line in table_rows_data:
                        if "Day 1" in line:
                            paragraph.text = paragraph.text.replace("{{TOUR_ROUTE}}", line[1].upper())
                if "{{DETAILED_ITINERARY}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{DETAILED_ITINERARY}}", clean_body_text)
                if "{{TOUR_INCLUSIONS}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_INCLUSIONS}}", clean_inclusions)
                if "{{TOUR_EXCLUSIONS}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_EXCLUSIONS}}", clean_exclusions)

            # SWAP TABLES GRID FOR HIGHLIGHTS AND PRICING
            for table in doc.tables:
                if len(table.rows) > 1 and "{{DAY_NUM}}" in table.rows[1].cells[0].text:
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
                        if "{{TEACHER_RATIO}}" in cell.text:
                            cell.text = cell.text.replace("{{TEACHER_RATIO}}", teacher_ratio)

            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("🎉 Final Document compiled perfectly with locked styling dimensions!")
            st.download_button(
                label="💾 Download Client Word Document (.docx)",
                data=bio.getvalue(),
                file_name=f"WNW_Itinerary_{school_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Error merging template data strings: {str(e)}")
