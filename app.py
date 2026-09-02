import streamlit as st
from docx import Document
import io

# Page Configuration with Official Blue Identity
st.set_page_config(page_title="WNW Table Router Engine", layout="wide")

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Dynamic Row Multiplication Template Engine (.docx)")

# Interactive User Parameters
school_name = st.text_input("School/College Name:", "AA School")
start_city = st.text_input("Starting From City:", "Jaipur")
destination_name = st.text_input("Destination Label:", "CHANDIGARH – MANALI")
tour_duration = st.text_input("Tour Duration Frame:", "6 Nights / 7 Days")
student_cost = st.text_input("Cost Per Student Row:", "Rs. 13,500/-")

st.markdown("### 📋 Paste the AI Text Output Block Below:")
pasted_itinerary = st.text_area("Pasted Itinerary Body Text:", height=400)

if st.button("Compile Custom Fixed-Format Proposal"):
    if not pasted_itinerary:
        st.error("Please paste the formatted text block from our chat first!")
    else:
        try:
            # 1. Open your master template from your GitHub project folder
            doc = Document("template.docx")
            
            # 2. Extract and separate the dynamic overview table rows from the main body text
            lines = pasted_itinerary.split('\n')
            table_rows_data = []
            detailed_itinerary_lines = []
            in_table_mode = False
            
            for line in lines:
                if "TABLE_START" in line:
                    in_table_mode = True
                    continue
                if "TABLE_END" in line:
                    in_table_mode = False
                    continue
                if in_table_mode:
                    if "|" in line:
                        table_rows_data.append([cell.strip() for cell in line.split('|')])
                else:
                    detailed_itinerary_lines.append(line)
            
            clean_body_text = "\n".join(detailed_itinerary_lines)

            # 3. Swap out your Page 1 top header placeholder tags seamlessly
            for paragraph in doc.paragraphs:
                if "{{SCHOOL_NAME}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{SCHOOL_NAME}}", school_name)
                if "{{DESTINATION_NAME}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{DESTINATION_NAME}}", destination_name.upper())
                if "{{TOUR_DURATION}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{TOUR_DURATION}}", tour_duration)
                if "{{TOUR_ROUTE}}" in paragraph.text:
                    # Finds the loop line generated from our chat
                    for line in table_rows_data:
                        if "Day 1" in line[0]:
                            paragraph.text = paragraph.text.replace("{{TOUR_ROUTE}}", line[1].upper())
                if "{{DETAILED_ITINERARY}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{DETAILED_ITINERARY}}", clean_body_text)

            # 4. DYNAMIC MULTIPLICATION ENGINE: Find your template table on Page 2
            for table in doc.tables:
                # Target the grid that contains our target anchor tags
                if len(table.rows) > 1 and "{{DAY_NUM}}" in table.rows[1].cells[0].text:
                    base_row = table.rows[1]
                    
                    # Cycle through each data split line provided by the chat text box
                    for index, row_data in enumerate(table_rows_data):
                        if index == 0:
                            # Fill the very first template row
                            new_row = base_row
                        else:
                            # Dynamically clone a fresh blank row keeping identical column widths & colors
                            new_row = table.add_row()
                        
                        # Populate cells securely with the arrow symbols preserved perfectly
                        for i in range(min(len(row_data), len(new_row.cells))):
                            new_row.cells[i].text = row_data[i]
                
                # Check your pricing grid structure at the bottom for replacement values
                for row in table.rows:
                    for cell in row.cells:
                        if "{{STUDENT_COST}}" in cell.text:
                            cell.text = cell.text.replace("{{STUDENT_COST}}", student_cost)

            # 5. Export out to browser download buffer channel
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("🎉 Word Document filled successfully with your custom arrow routes!")
            st.download_button(
                label="💾 Download Client Word Document (.docx)",
                data=bio.getvalue(),
                file_name=f"WNW_Proposal_{school_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Execution Error processing file text layers: {str(e)}")
