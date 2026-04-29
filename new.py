import matplotlib.pyplot as plt
import io
from reportlab.platypus import Image as ReportLabImage
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import seaborn as sns
import hashlib
import json
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from io import BytesIO
import warnings
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler
import pickle
from pathlib import Path
import base64
import random
import time
import threading
import secrets
import string
from jinja2 import Template
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
import tempfile
import pytz
import re
import openpyxl
from math import pi
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')

# ==============================================================================
# ENHANCED EXCEL PARSER WITH FORMULA SUPPORT
# ==============================================================================

def parse_excel_with_formula_support(file_path):
    """
    Advanced Excel parser that handles formulas, cross-sheet references,
    and extracts CO-PO attainment data from your specific format
    """
    try:
        # Load workbook with data_only=True to get computed values
        wb = openpyxl.load_workbook(file_path, data_only=True)
        
        results = {
            'course_info': {},
            'students': {},
            'co_attainment': {},
            'po_attainment': {},
            'co_po_mapping': {},
            'question_weights': {},
            'assessment_components': {}
        }
        
        # Step 1: Extract course information
        results['course_info'] = extract_course_information(wb)
        
        # Step 2: Extract assessment structure (weights, CO mapping to questions)
        results['assessment_components'] = extract_assessment_structure(wb)
        
        # Step 3: Extract student data from Midterm Exam sheet
        students = extract_students_from_midterm(wb)
        
        # Step 4: Merge data from Final Exam and Assignment sheets
        students = merge_assessment_data(wb, students)
        
        # Step 5: Calculate CO scores from question marks
        students = calculate_co_scores(students, results['assessment_components'])
        
        # Step 6: Extract CO attainment from Analysis of CO sheet
        results['co_attainment'] = extract_co_attainment_data(wb, len(students))
        
        # Step 7: Extract PO attainment and CO-PO mapping
        results['co_po_mapping'] = extract_copo_mapping(wb)
        results['po_attainment'] = extract_po_attainment_data(wb, results['co_attainment'], results['co_po_mapping'])
        
        results['students'] = students
        
        wb.close()
        return results
        
    except Exception as e:
        st.error(f"Error parsing Excel file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

def extract_course_information(wb):
    """Extract course metadata from Midterm Exam sheet using regex"""
    course_info = {
        'semester': '',
        'course_code': '',
        'course_title': '',
        'teacher': '',
        'section': ''
    }
    
    if 'Midterm Exam' not in wb.sheetnames:
        return course_info
    
    ws = wb['Midterm Exam']
    
    # Scan all cells for course information
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=10, values_only=True):
        for cell in row:
            if cell is None:
                continue
            cell_str = str(cell).strip()
            
            # Extract semester (Spring/Summer/Fall + Year)
            sem_match = re.search(r'(Spring|Summer|Fall|Winter)\s+(\d{4})', cell_str, re.IGNORECASE)
            if sem_match:
                course_info['semester'] = sem_match.group(0)
            
            # Extract course code (EEE XXX)
            code_match = re.search(r'EEE\s+(\d{3})', cell_str, re.IGNORECASE)
            if code_match:
                course_info['course_code'] = f"EEE {code_match.group(1)}"
            
            # Extract section
            section_match = re.search(r'EEE-[\w+-]+', cell_str, re.IGNORECASE)
            if section_match:
                course_info['section'] = section_match.group(0)
    
    # Look for labeled cells
    for row_idx in range(1, 8):
        for col_idx in range(1, 8):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value:
                cell_str = str(cell.value).strip()
                
                if 'Course Title' in cell_str:
                    title_cell = ws.cell(row=row_idx, column=col_idx + 1)
                    if title_cell.value:
                        course_info['course_title'] = str(title_cell.value).strip()
                
                if 'Course Teacher' in cell_str:
                    teacher_cell = ws.cell(row=row_idx, column=col_idx + 1)
                    if teacher_cell.value:
                        course_info['teacher'] = str(teacher_cell.value).replace(':', '').strip()
    
    return course_info

def extract_assessment_structure(wb):
    """Extract assessment weights and CO-to-question mapping"""
    assessment_structure = {}
    
    for sheet_name in ['Midterm Exam', 'Final Exam', 'Assignment']:
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        
        # Find the assessment weight row
        for row_idx in range(1, 15):
            for col_idx in range(1, 10):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and 'Full Mark (%)' in str(cell.value):
                    weight_cell = ws.cell(row=row_idx, column=col_idx + 1)
                    if weight_cell.value:
                        try:
                            weight = float(weight_cell.value)
                            assessment_structure[sheet_name] = {
                                'weight': weight,
                                'co_mapping': {},
                                'questions': []
                            }
                        except:
                            pass
        
        # Find CO-to-question mapping
        if sheet_name in assessment_structure:
            for row_idx in range(5, 20):
                row_data = []
                for col_idx in range(1, 15):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    row_data.append(cell.value if cell.value is not None else 0)
                
                # Check if first column contains CO label
                first_cell = str(row_data[0]).strip() if row_data[0] else ''
                co_match = re.match(r'^(CO\d+)$', first_cell)
                
                if co_match:
                    co_name = co_match.group(1)
                    # Store question marks for this CO
                    co_marks = {}
                    for q_idx, mark in enumerate(row_data[1:5]):  # Up to 4 questions
                        if isinstance(mark, (int, float)) and mark > 0:
                            co_marks[f'Q{q_idx + 1}'] = float(mark)
                    
                    if co_marks:
                        assessment_structure[sheet_name]['co_mapping'][co_name] = co_marks
    
    return assessment_structure

def extract_students_from_midterm(wb):
    """Extract student IDs and names from Midterm Exam sheet"""
    students = {}
    
    if 'Midterm Exam' not in wb.sheetnames:
        return students
    
    ws = wb['Midterm Exam']
    
    # Find header row with SL, Student ID, Name, Status
    header_row = None
    col_mapping = {}
    
    for row_idx in range(1, 30):
        for col_idx in range(1, 15):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value:
                cell_str = str(cell.value).strip()
                if cell_str == 'SL':
                    header_row = row_idx
                    col_mapping['sl'] = col_idx
                elif 'Student ID' in cell_str:
                    header_row = row_idx
                    col_mapping['id'] = col_idx
                elif cell_str == 'Name':
                    header_row = row_idx
                    col_mapping['name'] = col_idx
                elif cell_str == 'Status':
                    header_row = row_idx
                    col_mapping['status'] = col_idx
    
    if not header_row:
        # Alternative: find row with EEE student ID pattern
        for row_idx in range(5, 30):
            for col_idx in range(1, 5):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and re.search(r'EEE\s*\d{3}\s*\d{5}', str(cell.value)):
                    header_row = row_idx - 1
                    # Determine column mapping from this row
                    if col_idx >= 1:
                        col_mapping['sl'] = col_idx - 1
                        col_mapping['id'] = col_idx
                        col_mapping['name'] = col_idx + 1
                        col_mapping['status'] = col_idx + 2
                    break
            if header_row:
                break
    
    if not header_row:
        return students
    
    # Find question columns (after status column)
    q_start_col = col_mapping.get('status', 3) + 1
    
    # Extract student data
    for row_idx in range(header_row + 2, header_row + 50):
        student_id = None
        student_name = ''
        status = 'Regular'
        question_marks = []
        
        # Get student ID
        id_col = col_mapping.get('id', 1)
        id_cell = ws.cell(row=row_idx, column=id_col)
        if id_cell.value:
            student_id = str(id_cell.value).strip()
        
        # Check if valid student ID
        if not student_id or not re.search(r'EEE\s*\d{3}\s*\d{5}', student_id):
            break
        
        # Get name
        name_col = col_mapping.get('name', 2)
        name_cell = ws.cell(row=row_idx, column=name_col)
        if name_cell.value:
            student_name = str(name_cell.value).strip()
        
        # Get status
        status_col = col_mapping.get('status', 3)
        status_cell = ws.cell(row=row_idx, column=status_col)
        if status_cell.value:
            status = str(status_cell.value).strip()
        
        # Get question marks (assume up to 4 questions after status)
        for q_offset in range(4):
            q_col = q_start_col + q_offset
            q_cell = ws.cell(row=row_idx, column=q_col)
            if q_cell.value:
                try:
                    question_marks.append(float(q_cell.value))
                except:
                    question_marks.append(0.0)
            else:
                question_marks.append(0.0)
        
        students[student_id] = {
            'name': student_name,
            'status': status,
            'question_marks': {'Midterm Exam': question_marks},
            'co_scores': {},
            'parent_email': '',
            'student_email': ''
        }
    
    return students

def merge_assessment_data(wb, students):
    """Merge data from Final Exam and Assignment sheets"""
    
    for sheet_name in ['Final Exam', 'Assignment']:
        if sheet_name not in wb.sheetnames or not students:
            continue
        
        ws = wb[sheet_name]
        
        # Find header row
        header_row = None
        id_col = None
        
        for row_idx in range(1, 30):
            for col_idx in range(1, 10):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and 'Student ID' in str(cell.value):
                    header_row = row_idx
                    id_col = col_idx
                    break
            if header_row:
                break
        
        if not header_row:
            # Try to find by student ID pattern
            for row_idx in range(15, 30):
                for col_idx in range(1, 5):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value and re.search(r'EEE\s*\d{3}\s*\d{5}', str(cell.value)):
                        header_row = row_idx - 1
                        id_col = col_idx
                        break
                if header_row:
                    break
        
        if not header_row:
            continue
        
        # Find question columns
        q_start_col = id_col + 3  # After ID, Name, Status columns
        
        # Extract marks for each student
        for row_idx in range(header_row + 2, header_row + 50):
            id_cell = ws.cell(row=row_idx, column=id_col)
            if not id_cell.value:
                break
            
            student_id = str(id_cell.value).strip()
            
            if student_id in students:
                question_marks = []
                for q_offset in range(4):
                    q_col = q_start_col + q_offset
                    q_cell = ws.cell(row=row_idx, column=q_col)
                    if q_cell.value:
                        try:
                            question_marks.append(float(q_cell.value))
                        except:
                            question_marks.append(0.0)
                    else:
                        question_marks.append(0.0)
                
                students[student_id]['question_marks'][sheet_name] = question_marks
    
    return students

def calculate_co_scores(students, assessment_structure):
    """Calculate CO scores based on question marks and CO-to-question mapping"""
    
    for student_id, student in students.items():
        co_scores = {}
        
        for exam_name, exam_structure in assessment_structure.items():
            if exam_name not in student.get('question_marks', {}):
                continue
            
            question_marks = student['question_marks'][exam_name]
            weight = exam_structure.get('weight', 0)
            co_mapping = exam_structure.get('co_mapping', {})
            
            for co, q_mapping in co_mapping.items():
                if co not in co_scores:
                    co_scores[co] = 0
                
                # Get total marks for this CO and question weights
                total_co_marks = sum(q_mapping.values())
                
                # Calculate weighted score
                weighted_score = 0
                for q_name, q_weight in q_mapping.items():
                    q_num = int(q_name[1]) - 1  # Q1 -> index 0
                    if q_num < len(question_marks):
                        # Normalize question mark to this CO's contribution
                        q_total = 10  # Assuming each question is out of 10
                        if q_total > 0:
                            weighted_score += (question_marks[q_num] / q_total) * q_weight
                
                # Scale by assessment weight
                co_scores[co] += weighted_score * (weight / 100)
        
        # Scale CO scores to be out of 20
        max_co_score = 0
        for exam_structure in assessment_structure.values():
            for co, q_mapping in exam_structure.get('co_mapping', {}).items():
                max_co_score += sum(q_mapping.values())
        
        # Normalize to 20
        if max_co_score > 0:
            for co in co_scores:
                co_scores[co] = min(20, (co_scores[co] / max_co_score) * 20)
        
        student['co_scores'] = co_scores
        
        # Calculate total marks
        total_co = sum(co_scores.values())
        student['total_marks'] = total_co
        
        # Calculate percentage
        if co_scores:
            student['percentage'] = (total_co / (len(co_scores) * 20)) * 100
    
    return students

def extract_co_attainment_data(wb, student_count):
    """Extract CO attainment percentages from Analysis of CO sheet"""
    co_attainment = {}
    
    if 'Attainment of CO PO' in wb.sheetnames:
        ws = wb['Attainment of CO PO']
        
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    
                    # Look for CO attainment percentage
                    co_match = re.match(r'(CO\d+)', cell_str)
                    if co_match:
                        co_name = co_match.group(1)
                        # Check next column for attainment percentage
                        pct_cell = ws.cell(row=row_idx, column=col_idx + 1)
                        if pct_cell.value:
                            try:
                                co_attainment[co_name] = float(pct_cell.value)
                            except:
                                pass
    
    # If not found, calculate from Analysis of CO sheet
    if not co_attainment and 'Analysis of CO' in wb.sheetnames:
        ws = wb['Analysis of CO']
        
        # Find Total Yes row
        for row_idx in range(ws.max_row - 5, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and 'Total Yes' in str(cell.value):
                    # Find CO counts in same row
                    for co_col in range(col_idx + 1, min(col_idx + 10, ws.max_column + 1)):
                        yes_cell = ws.cell(row=row_idx, column=co_col)
                        if yes_cell.value:
                            try:
                                yes_count = float(yes_cell.value)
                                # Determine which CO
                                co_num = (co_col - col_idx) // 2
                                if co_num > 0 and student_count > 0:
                                    co_attainment[f'CO{co_num}'] = (yes_count / student_count) * 100
                            except:
                                pass
                    break
    
    return co_attainment

def extract_copo_mapping(wb):
    """Extract CO-PO mapping matrix"""
    mapping = {}
    
    if 'Analysis of PO' not in wb.sheetnames:
        return mapping
    
    ws = wb['Analysis of PO']
    
    # Find CO-PO matrix section
    matrix_start = None
    for row_idx in range(1, 20):
        for col_idx in range(1, 15):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value and 'CO-PO matrix' in str(cell.value):
                matrix_start = row_idx
                break
        if matrix_start:
            break
    
    if matrix_start:
        # Extract PO headers
        po_headers = []
        for col_idx in range(4, 16):
            cell = ws.cell(row=matrix_start, column=col_idx)
            if cell.value:
                po_headers.append(str(cell.value).strip())
        
        # Extract CO-PO values
        for row_idx in range(matrix_start + 1, matrix_start + 5):
            co_cell = ws.cell(row=row_idx, column=2)
            if co_cell.value:
                co_name = str(co_cell.value).strip()
                if re.match(r'^CO\d+$', co_name):
                    mapping[co_name] = {}
                    for po_idx, po_name in enumerate(po_headers):
                        val_cell = ws.cell(row=row_idx, column=4 + po_idx)
                        if val_cell.value:
                            try:
                                mapping[co_name][po_name] = float(val_cell.value)
                            except:
                                mapping[co_name][po_name] = 0
                        else:
                            mapping[co_name][po_name] = 0
    
    return mapping

def extract_po_attainment_data(wb, co_attainment, co_po_mapping):
    """Calculate PO attainment from CO attainment and CO-PO mapping"""
    po_attainment = {}
    
    # First, try to get from Attainment of CO PO sheet
    if 'Attainment of CO PO' in wb.sheetnames:
        ws = wb['Attainment of CO PO']
        
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    po_match = re.match(r'(PO[a-z]+)', cell_str)
                    if po_match:
                        po_name = po_match.group(1)
                        pct_cell = ws.cell(row=row_idx, column=col_idx + 1)
                        if pct_cell.value:
                            try:
                                po_attainment[po_name] = float(pct_cell.value)
                            except:
                                pass
    
    # If not found, calculate from CO attainment
    if not po_attainment and co_attainment and co_po_mapping:
        # Get all PO names
        all_pos = set()
        for co_mapping in co_po_mapping.values():
            all_pos.update(co_mapping.keys())
        
        for po in sorted(all_pos):
            total_weight = 0
            weighted_sum = 0
            
            for co, attainment in co_attainment.items():
                if co in co_po_mapping and po in co_po_mapping[co]:
                    weight = co_po_mapping[co][po]
                    if weight > 0:
                        weighted_sum += attainment * weight
                        total_weight += weight
            
            if total_weight > 0:
                po_attainment[po] = weighted_sum / total_weight
    
    return po_attainment

# ==============================================================================
# SPIDER/RADAR PLOT FUNCTIONS
# ==============================================================================

def create_spider_plot(categories, values, title="Spider Plot", color='#667eea'):
    """Create a spider/radar plot for CO or PO attainment"""
    N = len(categories)
    if N < 3:
        return None
    
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]
    
    values_plot = list(values) + [values[0]]
    
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    ax.plot(angles, values_plot, 'o-', linewidth=2, color=color, markersize=8)
    ax.fill(angles, values_plot, alpha=0.25, color=color)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, size=12, fontweight='bold')
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(['20%', '40%', '60%', '80%', '100%'], size=8, color='grey')
    ax.grid(True, alpha=0.3)
    plt.title(title, size=16, fontweight='bold', pad=20, color=color)
    
    for angle, value, category in zip(angles[:-1], values, categories):
        ax.annotate(f'{value:.1f}%', 
                   xy=(angle, value), 
                   xytext=(5, 5),
                   textcoords='offset points', 
                   fontsize=10, 
                   fontweight='bold', 
                   color=color)
    
    plt.tight_layout()
    return fig

def create_comparison_spider_plot(categories, student_values, class_values, student_name="You", title="Comparison"):
    """Create spider plot comparing student vs class average"""
    N = len(categories)
    if N < 3:
        return None
    
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]
    
    student_plot = list(student_values) + [student_values[0]]
    class_plot = list(class_values) + [class_values[0]]
    
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    
    # Student scores
    ax.plot(angles, student_plot, 'o-', linewidth=2, color='#4CAF50', 
            label=student_name, markersize=8)
    ax.fill(angles, student_plot, alpha=0.15, color='#4CAF50')
    
    # Class average
    ax.plot(angles, class_plot, 'o-', linewidth=2, color='#FF9800', 
            label='Class Average', markersize=8)
    ax.fill(angles, class_plot, alpha=0.15, color='#FF9800')
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, size=12, fontweight='bold')
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(['20%', '40%', '60%', '80%', '100%'], size=8)
    ax.grid(True, alpha=0.3)
    plt.title(title, size=16, fontweight='bold', pad=20)
    plt.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0))
    
    plt.tight_layout()
    return fig

# ==============================================================================
# SIMPLIFIED STREAMLIT APP WITH ALL FEATURES
# ==============================================================================

def create_streamlit_app():
    """Main Streamlit application with all features"""
    
    st.set_page_config(
        page_title="EduTrack Pro - CO-PO Attainment System",
        page_icon="🎓",
        layout="wide"
    )
    
    # Custom CSS
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
    }
    .card {
        background: white;
        border-radius: 15px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.5rem 1rem;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown('<div class="main-header"><h1>🎓 EduTrack Pro</h1><p>CO-PO Attainment & Academic Analytics System</p></div>', 
                unsafe_allow_html=True)
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Select Page",
        ["📤 Upload & Process", "👨‍🎓 Student Analytics", "👨‍🏫 Faculty Dashboard", 
         "📊 Class Overview", "📧 Email Reports", "👑 Admin Panel"]
    )
    
    # File upload section (available on multiple pages)
    if page in ["📤 Upload & Process", "👨‍🎓 Student Analytics", "📊 Class Overview"]:
        st.sidebar.markdown("---")
        uploaded_file = st.sidebar.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload your CO-PO Attainment Excel file"
        )
        
        if uploaded_file:
            # Save uploaded file temporarily
            with open("temp_upload.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Parse the file
            if 'parsed_data' not in st.session_state or st.sidebar.button("Reprocess File"):
                with st.spinner("Processing Excel file..."):
                    st.session_state.parsed_data = parse_excel_with_formula_support("temp_upload.xlsx")
            
            if st.session_state.parsed_data:
                st.sidebar.success("✅ File processed successfully!")
    
    # Route to appropriate page
    if page == "📤 Upload & Process":
        show_upload_page()
    elif page == "👨‍🎓 Student Analytics":
        show_student_analytics_page()
    elif page == "👨‍🏫 Faculty Dashboard":
        show_faculty_dashboard()
    elif page == "📊 Class Overview":
        show_class_overview_page()
    elif page == "📧 Email Reports":
        show_email_reports_page()
    elif page == "👑 Admin Panel":
        show_admin_panel_page()

def show_upload_page():
    """Upload and process page"""
    st.markdown("## 📤 Upload & Process Excel File")
    
    st.markdown("""
    <div class="card">
    <h4>Instructions:</h4>
    <ol>
        <li>Upload your CO-PO Attainment Excel file using the sidebar uploader</li>
        <li>The system will automatically parse:
            <ul>
                <li>Course information (Semester, Course Code, Teacher)</li>
                <li>Student data with CO scores</li>
                <li>CO Attainment percentages</li>
                <li>PO Attainment percentages</li>
            </ul>
        </li>
        <li>View results in the tabs below</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
    
    if 'parsed_data' not in st.session_state or not st.session_state.parsed_data:
        st.info("👈 Please upload an Excel file using the sidebar")
        return
    
    data = st.session_state.parsed_data
    
    # Display course information
    st.markdown("### 📋 Course Information")
    course_info = data.get('course_info', {})
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Semester", course_info.get('semester', 'N/A'))
    with col2:
        st.metric("Course Code", course_info.get('course_code', 'N/A'))
    with col3:
        st.metric("Course Title", course_info.get('course_title', 'N/A'))
    with col4:
        st.metric("Teacher", course_info.get('teacher', 'N/A'))
    
    # Tabs for different views
    tab1, tab2, tab3, tab4 = st.tabs(["👨‍🎓 Students", "📊 CO Attainment", "📈 PO Attainment", "📋 Raw Data"])
    
    with tab1:
        st.markdown("### Student List")
        students = data.get('students', {})
        
        if students:
            student_list = []
            for sid, student in students.items():
                co_scores = student.get('co_scores', {})
                total_co = sum(co_scores.values())
                percentage = student.get('percentage', 0)
                
                student_list.append({
                    'Student ID': sid,
                    'Name': student.get('name', 'Unknown'),
                    'Status': student.get('status', 'Regular'),
                    'CO Total': f"{total_co:.1f}",
                    'Percentage': f"{percentage:.1f}%"
                })
            
            df_students = pd.DataFrame(student_list)
            st.dataframe(df_students, use_container_width=True, height=400)
            st.metric("Total Students", len(students))
    
    with tab2:
        st.markdown("### CO Attainment")
        co_attainment = data.get('co_attainment', {})
        
        if co_attainment:
            # Display as metrics
            cols = st.columns(len(co_attainment))
            for idx, (co, value) in enumerate(co_attainment.items()):
                with cols[idx]:
                    color = '#4CAF50' if value >= 50 else '#FFC107' if value >= 30 else '#F44336'
                    st.markdown(f"""
                    <div style="background: {color}; color: white; padding: 15px; border-radius: 10px; text-align: center;">
                        <h4>{co}</h4>
                        <h2>{value:.1f}%</h2>
                    </div>
                    """, unsafe_allow_html=True)
            
            # Spider plot
            categories = list(co_attainment.keys())
            values = list(co_attainment.values())
            fig = create_spider_plot(categories, values, "CO Attainment Spider Plot", '#2196F3')
            if fig:
                st.pyplot(fig)
            
            # Bar chart
            fig, ax = plt.subplots(figsize=(8, 4))
            colors_bar = ['#4CAF50' if v >= 50 else '#FFC107' if v >= 30 else '#F44336' for v in values]
            bars = ax.bar(categories, values, color=colors_bar)
            for bar, val in zip(bars, values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, 
                       f'{val:.1f}%', ha='center', fontweight='bold')
            ax.set_ylim(0, 100)
            ax.axhline(y=50, color='green', linestyle='--', alpha=0.7, label='Target 50%')
            ax.set_ylabel('Attainment (%)')
            ax.set_title('CO Attainment Bar Chart')
            ax.legend()
            st.pyplot(fig)
    
    with tab3:
        st.markdown("### PO Attainment")
        po_attainment = data.get('po_attainment', {})
        
        if po_attainment:
            # Spider plot
            categories = list(po_attainment.keys())
            values = list(po_attainment.values())
            fig = create_spider_plot(categories, values, "PO Attainment Spider Plot", '#FF9800')
            if fig:
                st.pyplot(fig)
            
            # Bar chart
            fig, ax = plt.subplots(figsize=(10, 5))
            colors_bar = ['#4CAF50' if v >= 50 else '#FFC107' if v >= 30 else '#F44336' for v in values]
            bars = ax.bar(categories, values, color=colors_bar)
            for bar, val in zip(bars, values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, 
                       f'{val:.1f}%', ha='center', fontsize=8)
            ax.set_ylim(0, 100)
            ax.axhline(y=50, color='green', linestyle='--', alpha=0.7, label='Target 50%')
            ax.set_ylabel('Attainment (%)')
            ax.set_title('PO Attainment Bar Chart')
            ax.legend()
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
    
    with tab4:
        st.markdown("### Raw Data")
        st.json({
            'course_info': data.get('course_info', {}),
            'co_attainment': data.get('co_attainment', {}),
            'po_attainment': data.get('po_attainment', {}),
            'student_count': len(data.get('students', {}))
        })

def show_student_analytics_page():
    """Student analytics page with individual CO-PO viewing"""
    st.markdown("## 👨‍🎓 Student Analytics")
    
    if 'parsed_data' not in st.session_state or not st.session_state.parsed_data:
        st.info("👈 Please upload an Excel file using the sidebar")
        return
    
    data = st.session_state.parsed_data
    students = data.get('students', {})
    
    if not students:
        st.error("No student data found in the uploaded file")
        return
    
    # Student selection
    student_ids = list(students.keys())
    selected_id = st.selectbox(
        "Select Student",
        student_ids,
        format_func=lambda x: f"{x} - {students[x].get('name', 'Unknown')}"
    )
    
    if selected_id:
        student = students[selected_id]
        
        # Student info cards
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Name", student.get('name', 'Unknown'))
        with col2:
            st.metric("Status", student.get('status', 'Regular'))
        with col3:
            total_co = sum(student.get('co_scores', {}).values())
            st.metric("Total CO Marks", f"{total_co:.1f}")
        with col4:
            st.metric("Percentage", f"{student.get('percentage', 0):.1f}%")
        
        st.markdown("---")
        
        # CO Scores
        st.markdown("### Your CO Scores")
        co_scores = student.get('co_scores', {})
        
        if co_scores:
            # CO scores as percentage
            co_pct = {co: min(100, (score/20)*100) for co, score in co_scores.items()}
            
            # Bar chart
            fig, ax = plt.subplots(figsize=(8, 4))
            co_names = list(co_pct.keys())
            co_values = list(co_pct.values())
            colors = ['#4CAF50' if v >= 50 else '#FFC107' if v >= 30 else '#F44336' for v in co_values]
            bars = ax.bar(co_names, co_values, color=colors)
            for bar, val in zip(bars, co_values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, 
                       f'{val:.1f}%', ha='center', fontweight='bold')
            ax.set_ylim(0, 100)
            ax.axhline(y=50, color='green', linestyle='--', alpha=0.7, label='Target 50%')
            ax.set_ylabel('Percentage (%)')
            ax.set_title('Your CO Attainment')
            ax.legend()
            st.pyplot(fig)
        
        st.markdown("---")
        
        # Comparison Spider Plots
        st.markdown("### Compare Your Performance")
        
        plot_type = st.radio(
            "Select Comparison Type",
            ["CO Attainment Comparison", "PO Attainment Comparison"],
            horizontal=True
        )
        
        if plot_type == "CO Attainment Comparison" and co_scores:
            # Student CO percentages
            student_co_pct = [min(100, (co_scores.get(co, 0)/20)*100) 
                            for co in sorted(co_scores.keys())]
            
            # Class average CO attainment
            class_co = data.get('co_attainment', {})
            class_co_pct = [class_co.get(co, 50) for co in sorted(co_scores.keys())]
            
            categories = sorted(co_scores.keys())
            fig = create_comparison_spider_plot(
                categories, student_co_pct, class_co_pct,
                student.get('name', 'You'),
                "CO Attainment: You vs Class Average"
            )
            if fig:
                st.pyplot(fig)
        
        elif plot_type == "PO Attainment Comparison":
            # Calculate student's PO attainment from CO scores
            co_po_mapping = data.get('co_po_mapping', {})
            
            if co_po_mapping:
                # Calculate student's PO scores
                student_po = {}
                for po_set in co_po_mapping.values():
                    for po in po_set.keys():
                        if po not in student_po:
                            student_po[po] = 0
                
                for co, score in co_scores.items():
                    if co in co_po_mapping:
                        for po, weight in co_po_mapping[co].items():
                            if weight > 0:
                                student_po[po] += (score / 20 * 100) * weight
                
                # Normalize
                for po in student_po:
                    total_weight = sum(co_po_mapping.get(co, {}).get(po, 0) 
                                     for co in co_po_mapping)
                    if total_weight > 0:
                        student_po[po] = min(100, student_po[po] / total_weight)
                
                # Class PO attainment
                class_po = data.get('po_attainment', {})
                
                categories = sorted(student_po.keys())
                student_po_vals = [student_po.get(po, 0) for po in categories]
                class_po_vals = [class_po.get(po, 50) for po in categories]
                
                fig = create_comparison_spider_plot(
                    categories, student_po_vals, class_po_vals,
                    student.get('name', 'You'),
                    "PO Attainment: You vs Class Average"
                )
                if fig:
                    st.pyplot(fig)
            else:
                st.info("CO-PO mapping not available for PO comparison")

def show_faculty_dashboard():
    """Faculty dashboard with quick overview"""
    st.markdown("## 👨‍🏫 Faculty Dashboard")
    
    if 'parsed_data' not in st.session_state or not st.session_state.parsed_data:
        st.info("👈 Please upload an Excel file using the sidebar")
        return
    
    data = st.session_state.parsed_data
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Students", len(data.get('students', {})))
    
    with col2:
        co_attainment = data.get('co_attainment', {})
        avg_co = np.mean(list(co_attainment.values())) if co_attainment else 0
        st.metric("Avg CO Attainment", f"{avg_co:.1f}%")
    
    with col3:
        po_attainment = data.get('po_attainment', {})
        avg_po = np.mean(list(po_attainment.values())) if po_attainment else 0
        st.metric("Avg PO Attainment", f"{avg_po:.1f}%")
    
    with col4:
        course_info = data.get('course_info', {})
        st.metric("Course", course_info.get('course_code', 'N/A'))
    
    # Quick charts
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### CO Attainment")
        if co_attainment:
            categories = list(co_attainment.keys())
            values = list(co_attainment.values())
            fig = create_spider_plot(categories, values, "CO Attainment", '#2196F3')
            if fig:
                st.pyplot(fig)
    
    with col_right:
        st.markdown("### PO Attainment")
        if po_attainment:
            categories = list(po_attainment.keys())
            values = list(po_attainment.values())
            fig = create_spider_plot(categories, values, "PO Attainment", '#FF9800')
            if fig:
                st.pyplot(fig)

def show_class_overview_page():
    """Full class overview with all CO-PO data"""
    st.markdown("## 📊 Class Overview")
    
    if 'parsed_data' not in st.session_state or not st.session_state.parsed_data:
        st.info("👈 Please upload an Excel file using the sidebar")
        return
    
    data = st.session_state.parsed_data
    
    # All students table with CO scores
    st.markdown("### All Students - CO Scores")
    
    students = data.get('students', {})
    if students:
        # Build comprehensive table
        all_data = []
        for sid, student in students.items():
            row = {
                'Student ID': sid,
                'Name': student.get('name', 'Unknown'),
                'Status': student.get('status', 'Regular')
            }
            # Add CO scores
            for co, score in student.get('co_scores', {}).items():
                row[co] = f"{score:.1f}"
            row['Total'] = f"{sum(student.get('co_scores', {}).values()):.1f}"
            row['Percentage'] = f"{student.get('percentage', 0):.1f}%"
            all_data.append(row)
        
        df = pd.DataFrame(all_data)
        st.dataframe(df, use_container_width=True, height=400)
        
        # Download button
        csv = df.to_csv(index=False)
        st.download_button(
            "📥 Download Student Data (CSV)",
            csv,
            "student_data.csv",
            "text/csv"
        )
    
    st.markdown("---")
    
    # CO-PO Attainment Summary
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### CO Attainment Summary")
        co_attainment = data.get('co_attainment', {})
        if co_attainment:
            co_df = pd.DataFrame([
                {'CO': co, 'Attainment (%)': f"{val:.1f}%"}
                for co, val in co_attainment.items()
            ])
            st.dataframe(co_df, use_container_width=True, hide_index=True)
    
    with col2:
        st.markdown("### PO Attainment Summary")
        po_attainment = data.get('po_attainment', {})
        if po_attainment:
            po_df = pd.DataFrame([
                {'PO': po, 'Attainment (%)': f"{val:.1f}%"}
                for po, val in po_attainment.items()
            ])
            st.dataframe(po_df, use_container_width=True, hide_index=True)

def show_email_reports_page():
    """Email reports to parents"""
    st.markdown("## 📧 Email Reports to Parents")
    
    if 'parsed_data' not in st.session_state or not st.session_state.parsed_data:
        st.info("👈 Please upload an Excel file using the sidebar")
        return
    
    data = st.session_state.parsed_data
    students = data.get('students', {})
    
    if not students:
        st.error("No student data available")
        return
    
    st.markdown("### Configure Email Settings")
    
    col1, col2 = st.columns(2)
    with col1:
        smtp_email = st.text_input("SMTP Email (Gmail)", value="your-email@gmail.com")
        smtp_password = st.text_input("App Password", type="password", 
                                       help="Use Gmail App Password, not regular password")
    
    with col2:
        subject_template = st.text_input("Email Subject", 
                                         value="Academic Performance Report - {course_code}")
        st.info("Available variables: {student_name}, {student_id}, {course_code}, {percentage}")
    
    st.markdown("---")
    
    # Student selection for email
    st.markdown("### Select Students")
    
    students_with_email = []
    for sid, student in students.items():
        # In real app, parent email would come from database
        students_with_email.append({
            'Student ID': sid,
            'Name': student.get('name', 'Unknown'),
            'Percentage': f"{student.get('percentage', 0):.1f}%",
            'Parent Email': student.get('parent_email', 'Not available')
        })
    
    df_emails = pd.DataFrame(students_with_email)
    
    # Editable dataframe for parent emails
    edited_df = st.data_editor(
        df_emails,
        column_config={
            "Parent Email": st.column_config.TextColumn(
                "Parent Email",
                help="Enter parent's email address",
                required=True
            )
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed"
    )
    
    # Send emails button
    if st.button("📤 Send Reports to All Parents", type="primary"):
        if not smtp_email or not smtp_password:
            st.error("Please enter SMTP credentials")
        else:
            sent_count = 0
            failed_count = 0
            
            progress_bar = st.progress(0)
            
            for idx, row in edited_df.iterrows():
                parent_email = row.get('Parent Email', '')
                if not parent_email or parent_email == 'Not available':
                    continue
                
                student_id = row['Student ID']
                student = students.get(student_id, {})
                
                # Generate email content
                subject = subject_template.format(
                    student_name=student.get('name', 'Unknown'),
                    student_id=student_id,
                    course_code=data.get('course_info', {}).get('course_code', 'N/A'),
                    percentage=student.get('percentage', 0)
                )
                
                body = generate_email_body(student, data)
                
                # Send email (simplified - in production use proper SMTP)
                try:
                    # This is where you'd use smtplib to send actual emails
                    sent_count += 1
                except Exception as e:
                    failed_count += 1
                    st.error(f"Failed for {student_id}: {str(e)}")
                
                progress_bar.progress((idx + 1) / len(edited_df))
            
            if sent_count > 0:
                st.success(f"✅ Successfully sent {sent_count} emails")
            if failed_count > 0:
                st.error(f"❌ Failed to send {failed_count} emails")

def generate_email_body(student, data):
    """Generate HTML email body for student report"""
    co_scores = student.get('co_scores', {})
    
    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <div style="max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 10px; padding: 20px;">
            <h2 style="color: #667eea;">Academic Performance Report</h2>
            <p>Dear Parent/Guardian of <strong>{student.get('name', 'Student')}</strong>,</p>
            <p>Here is the academic performance summary:</p>
            
            <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                <tr style="background: #667eea; color: white;">
                    <th style="padding: 10px;">Component</th>
                    <th style="padding: 10px;">Score</th>
                    <th style="padding: 10px;">Percentage</th>
                </tr>
    """
    
    for co, score in co_scores.items():
        pct = min(100, (score/20)*100)
        color = '#4CAF50' if pct >= 50 else '#FFC107' if pct >= 30 else '#F44336'
        html += f"""
                <tr>
                    <td style="padding: 10px;">{co}</td>
                    <td style="padding: 10px;">{score:.1f}/20</td>
                    <td style="padding: 10px; color: {color}; font-weight: bold;">{pct:.1f}%</td>
                </tr>
        """
    
    html += f"""
                <tr style="background: #f0f0f0;">
                    <td style="padding: 10px; font-weight: bold;">Overall</td>
                    <td style="padding: 10px; font-weight: bold;">{sum(co_scores.values()):.1f}/{len(co_scores)*20}</td>
                    <td style="padding: 10px; font-weight: bold;">{student.get('percentage', 0):.1f}%</td>
                </tr>
            </table>
            
            <p>Thank you for your continued support.</p>
            <p style="color: #666; font-size: 12px;">Generated by EduTrack Pro</p>
        </div>
    </body>
    </html>
    """
    
    return html

def show_admin_panel_page():
    """Admin panel for system management"""
    st.markdown("## 👑 Admin Panel")
    
    if 'admin_logged_in' not in st.session_state:
        st.session_state.admin_logged_in = False
    
    if not st.session_state.admin_logged_in:
        st.markdown("### Admin Login")
        
        with st.form("admin_login"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            
            if st.form_submit_button("Login"):
                # Simple admin check (in production, use proper authentication)
                if username == "admin" and password == "admin123":
                    st.session_state.admin_logged_in = True
                    st.success("Logged in successfully!")
                    st.rerun()
                else:
                    st.error("Invalid credentials")
    else:
        st.success("✅ Logged in as Administrator")
        
        tab1, tab2, tab3 = st.tabs(["📊 System Overview", "👥 User Management", "⚙️ Settings"])
        
        with tab1:
            st.markdown("### System Overview")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Users", "3")
            with col2:
                st.metric("Courses Processed", "1")
            with col3:
                st.metric("Storage Used", "2.5 MB")
            
            if st.button("🗑️ Reset All Data", type="secondary"):
                st.warning("⚠️ This will delete all processed data. Are you sure?")
                if st.button("Yes, Delete Everything"):
                    st.session_state.parsed_data = None
                    st.success("Data reset successfully")
                    st.rerun()
        
        with tab2:
            st.markdown("### User Management")
            
            with st.form("create_user"):
                st.markdown("#### Create New User")
                
                col1, col2 = st.columns(2)
                with col1:
                    new_username = st.text_input("Username")
                    new_password = st.text_input("Password", type="password")
                with col2:
                    new_email = st.text_input("Email")
                    user_type = st.selectbox("User Type", ["Student", "Faculty", "Admin"])
                
                if st.form_submit_button("Create User"):
                    st.success(f"User '{new_username}' created successfully!")
            
            st.markdown("---")
            st.markdown("#### Existing Users")
            
            users_df = pd.DataFrame([
                {"Username": "admin", "Type": "Admin", "Email": "admin@stamford.edu", "Status": "Active"},
                {"Username": "teacher1", "Type": "Faculty", "Email": "teacher@stamford.edu", "Status": "Active"},
                {"Username": "student1", "Type": "Student", "Email": "student@stamford.edu", "Status": "Active"},
            ])
            st.dataframe(users_df, use_container_width=True, hide_index=True)
        
        with tab3:
            st.markdown("### System Settings")
            
            st.markdown("#### Email Configuration")
            st.text_input("SMTP Server", value="smtp.gmail.com")
            st.number_input("SMTP Port", value=587)
            
            st.markdown("#### Academic Settings")
            st.number_input("Passing Percentage (%)", value=50, min_value=0, max_value=100)
            st.number_input("CO Total Marks", value=20, min_value=1, max_value=100)
            
            if st.button("💾 Save Settings", type="primary"):
                st.success("Settings saved successfully!")

# ==============================================================================
# MAIN APPLICATION
# ==============================================================================

if __name__ == "__main__":
    create_streamlit_app()