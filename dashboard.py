import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import io
import base64

# --- Core Marksheet Generation Logic ---

def calculate_single_student_grades(student_info, mid_term_weight, final_term_weight, passing_percentage):
    """Calculates total marks, GPA, and grades for a single student from form data."""
    total_marks_dict = {}
    subjects = list(student_info['marks'].keys())

    for subject, marks in student_info['marks'].items():
        mid_mark = marks['Mid-term']
        final_mark = marks['Final']
        total_marks_dict[subject] = (mid_term_weight * mid_mark) + (final_term_weight * final_mark)

    student_info['total_marks'] = total_marks_dict
    
    total_percentage = sum(total_marks_dict.values()) / len(subjects) if subjects else 0
    student_info['percentage'] = total_percentage

    gpa = (total_percentage / 100) * 4.0
    student_info['gpa'] = f"{gpa:.2f}"

    if total_percentage >= passing_percentage:
        status = 'Pass'
    else:
        status = 'Fail'
    
    if total_percentage >= 90:
        letter_grade = 'A+'
    elif 80 <= total_percentage < 90:
        letter_grade = 'A'
    elif 70 <= total_percentage < 80:
        letter_grade = 'B'
    elif 60 <= total_percentage < 70:
        letter_grade = 'C'
    elif 50 <= total_percentage < 60:
        letter_grade = 'D'
    else:
        letter_grade = 'F'
        
    student_info['grade'] = letter_grade
    student_info['status'] = status

    return student_info

def get_image_base64(image_bytes):
    """Converts image bytes to a base64 string for embedding in HTML."""
    return base64.b64encode(image_bytes).decode()

def generate_word_marksheet(student_data, school_name, school_logo_bytes):
    """Generates a Word document from a byte stream."""
    document = Document()
    
    # Header
    if school_logo_bytes:
        header_section = document.sections[0].header
        header_paragraph = header_section.paragraphs[0]
        logo_run = header_paragraph.add_run()
        logo_run.add_picture(io.BytesIO(school_logo_bytes), width=Inches(1.0))
        header_paragraph.add_run(f'\t{school_name}').bold = True
        header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        document.add_heading(school_name, level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_heading('Marksheet', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Student Profile
    profile_table = document.add_table(rows=1, cols=2)
    profile_cell = profile_table.cell(0, 0)
    profile_cell.text = f"Name: {student_data['name']}\n" \
                       f"Roll No: {student_data['roll_no']}\n" \
                       f"Department: {student_data['department']}"

    photo_cell = profile_table.cell(0, 1)
    if student_data.get('photo_bytes'):
        photo_run = photo_cell.paragraphs[0].add_run()
        photo_run.add_picture(io.BytesIO(student_data['photo_bytes']), width=Inches(1.25))
    else:
        photo_cell.text = "Photo not available"

    # Marks table
    document.add_heading('Marks Details', level=3)
    marks_table = document.add_table(rows=1, cols=4)
    marks_table.style = 'Table Grid'
    hdr_cells = marks_table.rows[0].cells
    hdr_cells[0].text = 'Subject'
    hdr_cells[1].text = 'Mid-term Marks'
    hdr_cells[2].text = 'Final Marks'
    hdr_cells[3].text = 'Total Marks (Weighted)'

    for subject, marks in student_data['marks'].items():
        row_cells = marks_table.add_row().cells
        row_cells[0].text = subject
        row_cells[1].text = str(marks['Mid-term'])
        row_cells[2].text = str(marks['Final'])
        row_cells[3].text = f"{student_data['total_marks'].get(subject, 0):.2f}"

    # Summary
    document.add_heading('Summary', level=3)
    document.add_paragraph(f"Percentage: {student_data['percentage']:.2f}%\n"
                           f"GPA: {student_data['gpa']}\n"
                           f"Grade: {student_data['grade']}\n"
                           f"Status: {student_data['status']}")

    doc_io = io.BytesIO()
    document.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_excel_marksheet(student_data, school_name):
    """Generates an Excel marksheet for a single student."""
    wb = Workbook()
    ws = wb.active
    ws.title = f"{student_data['name']} Marksheet"

    ws.append([f"{school_name} - Marksheet"])
    ws.merge_cells('A1:D1')
    ws['A1'].font = Font(bold=True, size=16)
    
    ws.append([f"Name: {student_data['name']}", f"Roll No: {student_data['roll_no']}", f"Department: {student_data['department']}"])
    ws.append([]) # Spacer

    # Marks
    ws.append(['Subject', 'Mid-term Marks', 'Final Marks', 'Total Marks (Weighted)'])
    for subject, marks in student_data['marks'].items():
        ws.append([
            subject,
            marks['Mid-term'],
            marks['Final'],
            f"{student_data['total_marks'].get(subject, 0):.2f}"
        ])
    
    # Summary
    ws.append([]) # Spacer
    ws.append(['Summary'])
    ws['A' + str(ws.max_row)].font = Font(bold=True)
    ws.append([f"Percentage:", f"{student_data['percentage']:.2f}%"])
    ws.append([f"GPA:", student_data['gpa']])
    ws.append([f"Grade:", student_data['grade']])
    ws.append([f"Status:", student_data['status']])
    
    if student_data['status'] == 'Fail':
        ws['B' + str(ws.max_row)].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

# --- Streamlit Dashboard UI ---

st.set_page_config(page_title="Marksheet Input Form", layout="wide")
st.title("Marksheet Generation System")

# Initialize session state
if 'num_subjects' not in st.session_state:
    st.session_state.num_subjects = 3 # Default value
if 'subjects' not in st.session_state:
    st.session_state.subjects = ["Math", "Science", "History"]

def update_subjects():
    st.session_state.subjects = [st.session_state[f"subject_{i}"] for i in range(st.session_state.num_subjects)]

# --- Sidebar for Configuration ---
with st.sidebar:
    st.header("1. School & Academic Criteria")
    school_name = st.text_input("School Name", "Global Tech Academy")
    logo_upload = st.file_uploader("Upload School Logo", type=['png', 'jpg'])
    passing_percentage = st.slider("Passing Percentage (%)", 0, 100, 50)
    mid_term_weight = st.slider("Mid-term Weight", 0.0, 1.0, 0.3, 0.05)
    st.info(f"Final-term Weight: **{1.0 - mid_term_weight:.2f}**")

    st.header("2. Define Subjects")
    num_subjects_input = st.number_input(
        "Number of Subjects", 
        min_value=1, 
        value=st.session_state.num_subjects, 
        step=1,
        key='num_subjects'
    )

    # Dynamically create subject name inputs
    subject_names = []
    for i in range(st.session_state.num_subjects):
        default_name = st.session_state.subjects[i] if i < len(st.session_state.subjects) else f"Subject {i+1}"
        subject_name = st.text_input(f"Subject {i+1} Name", value=default_name, key=f"subject_input_{i}")
        subject_names.append(subject_name)
    
    st.session_state.subjects = subject_names


# --- Main Panel for Input Form ---
st.header("3. Enter Student Details")

if not st.session_state.subjects or not all(st.session_state.subjects):
    st.warning("Please define all subject names in the sidebar to create the input form.")
else:
    with st.form("marksheet_form"):
        st.subheader("Student Information")
        student_name = st.text_input("Student Name", "John Doe")
        roll_no = st.text_input("Roll No", "101")
        department = st.text_input("Department", "Computer Science")
        student_photo = st.file_uploader("Upload Student Photo", type=['png', 'jpg'])

        st.subheader("Enter Marks")
        marks_data = {}
        for subject in st.session_state.subjects:
            cols = st.columns(2)
            with cols[0]:
                mid_marks = st.number_input(f"Mid-term: {subject}", min_value=0, max_value=100, value=75, key=f"mid_{subject}")
            with cols[1]:
                final_marks = st.number_input(f"Final: {subject}", min_value=0, max_value=100, value=80, key=f"final_{subject}")
            marks_data[subject] = {"Mid-term": mid_marks, "Final": final_marks}

        submitted = st.form_submit_button("Generate Marksheet Preview", type="primary")

    if submitted:
        # Process data on submission
        student_info = {
            "name": student_name,
            "roll_no": roll_no,
            "department": department,
            "marks": marks_data,
            "photo_bytes": student_photo.getvalue() if student_photo else None
        }
        
        school_logo_bytes = logo_upload.getvalue() if logo_upload else None

        # Calculate grades
        processed_data = calculate_single_student_grades(student_info, mid_term_weight, (1.0 - mid_term_weight), passing_percentage)

        # --- Display Marksheet Preview ---
        st.header("4. Marksheet Preview & Download")
        
        # Header
        cols = st.columns([1, 4])
        with cols[0]:
            if school_logo_bytes:
                st.image(school_logo_bytes, width=100)
        with cols[1]:
            st.title(school_name)
            st.subheader("Student Marksheet")
        
        st.markdown("---")

        # Student Info
        cols = st.columns([4, 1])
        with cols[0]:
            st.write(f"**Name:** {processed_data['name']}")
            st.write(f"**Roll No:** {processed_data['roll_no']}")
            st.write(f"**Department:** {processed_data['department']}")
        with cols[1]:
            if processed_data['photo_bytes']:
                st.image(processed_data['photo_bytes'], width=120)

        # Marks Table
        st.subheader("Marks Details")
        marks_df_data = []
        for subject, marks in processed_data['marks'].items():
            marks_df_data.append({
                "Subject": subject,
                "Mid-term Marks": marks['Mid-term'],
                "Final Marks": marks['Final'],
                "Total (Weighted)": f"{processed_data['total_marks'][subject]:.2f}"
            })
        st.dataframe(pd.DataFrame(marks_df_data), use_container_width=True)

        # Summary
        st.subheader("Overall Performance")
        cols = st.columns(4)
        with cols[0]:
            st.metric("Final Percentage", f"{processed_data['percentage']:.2f}%")
        with cols[1]:
            st.metric("GPA", processed_data['gpa'])
        with cols[2]:
            st.metric("Grade", processed_data['grade'])
        with cols[3]:
            st.metric("Status", processed_data['status'])

        st.markdown("---")

        # --- Download Buttons ---
        st.subheader("Download Files")
        
        # Generate files in memory
        word_file = generate_word_marksheet(processed_data, school_name, school_logo_bytes)
        excel_file = generate_excel_marksheet(processed_data, school_name)

        cols = st.columns(2)
        with cols[0]:
            st.download_button(
                label="Download as Word (.docx)",
                data=word_file,
                file_name=f"{processed_data['name']}_marksheet.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        with cols[1]:
            st.download_button(
                label="Download as Excel (.xlsx)",
                data=excel_file,
                file_name=f"{processed_data['name']}_marksheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
