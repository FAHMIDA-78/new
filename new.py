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
from datetime import datetime, timedelta
from io import BytesIO
import warnings
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestClassifier
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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
import tempfile
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
import pytz
import re
from openpyxl import load_workbook
warnings.filterwarnings('ignore')

# ==============================================================================
# EMAIL CONFIGURATION
# ==============================================================================
SMTP_CONFIG = {
    'server': 'smtp.gmail.com',
    'port': 587,
    'email': 'fahmidafaiza918@gmail.com',
    'password': 'karzuqmdxkbnauuw'
}

# ==============================================================================
# CHECK AND INSTALL MISSING DEPENDENCIES
# ==============================================================================
try:
    import xlsxwriter
except ImportError:
    st.error("Missing dependency: xlsxwriter")
    st.info("Please install it using: pip install xlsxwriter")
    if st.button("Install xlsxwriter (requires internet)"):
        import subprocess
        import sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter"])
        st.success("xlsxwriter installed successfully! Please restart the app.")
        st.stop()

try:
    from reportlab.lib import colors
except ImportError:
    st.error("Missing dependency: reportlab")
    st.info("Please install it using: pip install reportlab")
    if st.button("Install reportlab (requires internet)"):
        import subprocess
        import sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
        st.success("reportlab installed successfully! Please restart the app.")
        st.stop()

# ==============================================================================
# APPLICATION CONFIGURATION
# ==============================================================================
st.set_page_config(
    page_title="EduTrack Pro 2025",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# PROFESSIONAL THEME (Kept from original)
# ==============================================================================
def apply_professional_theme():
    # [Theme code remains the same as original]
    pass

# ==============================================================================
# PARSE EXCEL FILE DYNAMICALLY
# ==============================================================================
def parse_excel_file(uploaded_file):
    """Dynamically parse Excel file with any structure"""
    try:
        wb = load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames
        
        # Find sheets with student data (look for SL, Student ID columns)
        student_data_sheets = []
        
        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            
            # Check if this sheet has student data
            has_student_data = False
            for row in ws.iter_rows(min_row=1, max_row=min(30, ws.max_row), values_only=False):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        if 'student' in cell.value.lower() or 'name' in cell.value.lower():
                            has_student_data = True
                            break
                if has_student_data:
                    break
            
            if has_student_data:
                student_data_sheets.append(sheet_name)
        
        if not student_data_sheets:
            # Try to read all sheets and find any with student-like data
            for sheet_name in sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                for col in df.columns:
                    if df[col].astype(str).str.contains('EEE|Student|Name|ID', case=False).any():
                        student_data_sheets.append(sheet_name)
                        break
        
        # Parse each sheet that might contain student data
        all_student_data = {}
        co_po_info = {}
        
        for sheet_name in sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            # Extract metadata (Trimester, Course Code, etc.)
            metadata = {}
            for idx, row in df.iterrows():
                if idx > 10:  # Only check first 10 rows for metadata
                    break
                for col_idx, value in enumerate(row):
                    if pd.notna(value) and isinstance(value, str):
                        value_str = str(value).strip().lower()
                        if 'trimester' in value_str or 'semester' in value_str:
                            next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                            metadata['semester'] = str(next_val).strip() if pd.notna(next_val) else ''
                        elif 'course code' in value_str:
                            next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                            metadata['course_code'] = str(next_val).strip() if pd.notna(next_val) else ''
                        elif 'course title' in value_str:
                            next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                            metadata['course_title'] = str(next_val).strip() if pd.notna(next_val) else ''
                        elif 'teacher' in value_str:
                            next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                            metadata['teacher'] = str(next_val).strip() if pd.notna(next_val) else ''
            
            # Find the header row (look for SL, Student ID, Name patterns)
            header_row = None
            for idx, row in df.iterrows():
                row_str = ' '.join([str(v) for v in row if pd.notna(v)])
                if any(pattern in row_str.lower() for pattern in ['student id', 'sl', 'name']):
                    header_row = idx
                    break
            
            if header_row is not None:
                # Read data from header row onwards
                df_data = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
                
                # Find CO and PO columns dynamically
                co_columns = [col for col in df_data.columns if 'co' in str(col).lower()]
                po_columns = [col for col in df_data.columns if 'po' in str(col).lower()]
                
                # Store CO/PO info
                if co_columns or po_columns:
                    co_po_info[sheet_name] = {
                        'co_columns': co_columns,
                        'po_columns': po_columns,
                        'metadata': metadata
                    }
                
                all_student_data[sheet_name] = {
                    'data': df_data,
                    'metadata': metadata,
                    'co_columns': co_columns,
                    'po_columns': po_columns
                }
        
        return all_student_data, co_po_info, sheet_names
    
    except Exception as e:
        st.error(f"Error parsing Excel file: {str(e)}")
        return {}, {}, []

def extract_co_po_attainment(excel_data, co_po_info):
    """Extract CO and PO attainment values from parsed Excel data"""
    attainment_data = {
        'co_attainment': {},
        'po_attainment': {},
        'student_co_scores': {},
        'student_po_scores': {}
    }
    
    for sheet_name, info in co_po_info.items():
        df = excel_data.get(sheet_name, {}).get('data')
        if df is None:
            continue
        
        # Extract CO scores for each student
        co_columns = info.get('co_columns', [])
        po_columns = info.get('po_columns', [])
        
        for idx, row in df.iterrows():
            student_id = None
            student_name = None
            
            # Find student ID and name
            for col in df.columns:
                if 'student' in str(col).lower() and 'id' in str(col).lower():
                    student_id = str(row[col]) if pd.notna(row[col]) else None
                if 'name' in str(col).lower():
                    student_name = str(row[col]) if pd.notna(row[col]) else None
            
            if student_id and student_name:
                if student_id not in attainment_data['student_co_scores']:
                    attainment_data['student_co_scores'][student_id] = {
                        'name': student_name,
                        'courses': {}
                    }
                
                # Extract CO scores
                co_scores = {}
                for co_col in co_columns:
                    if co_col in row.index:
                        val = row[co_col]
                        co_scores[str(co_col)] = float(val) if pd.notna(val) and val != '' else 0.0
                
                # Extract PO scores
                po_scores = {}
                for po_col in po_columns:
                    if po_col in row.index:
                        val = row[po_col]
                        po_scores[str(po_col)] = float(val) if pd.notna(val) and val != '' else 0.0
                
                course_code = info.get('metadata', {}).get('course_code', sheet_name)
                attainment_data['student_co_scores'][student_id]['courses'][course_code] = {
                    'co_scores': co_scores,
                    'po_scores': po_scores
                }
    
    # Calculate overall CO and PO attainment
    # Calculate averages across all courses for each student
    for student_id, data in attainment_data['student_co_scores'].items():
        all_co = {}
        all_po = {}
        course_count = len(data['courses'])
        
        for course_code, course_data in data['courses'].items():
            for co, score in course_data['co_scores'].items():
                if co not in all_co:
                    all_co[co] = []
                all_co[co].append(score)
            
            for po, score in course_data['po_scores'].items():
                if po not in all_po:
                    all_po[po] = []
                all_po[po].append(score)
        
        # Calculate averages
        avg_co = {co: np.mean(scores) for co, scores in all_co.items() if scores}
        avg_po = {po: np.mean(scores) for po, scores in all_po.items() if scores}
        
        attainment_data['co_attainment'] = avg_co
        attainment_data['po_attainment'] = avg_po
    
    return attainment_data

# ==============================================================================
# COLOR UTILITY FUNCTIONS (Kept from original)
# ==============================================================================
def get_color_by_value(value, metric_type):
    # [Same as original]
    pass

def get_color_class(value, metric_type):
    # [Same as original]
    pass

def get_metric_description(value, metric_type):
    # [Same as original]
    pass

def create_colored_metric_card(title, value, metric_type="percentage", suffix="", prefix=""):
    # [Same as original]
    pass

# ==============================================================================
# PDF REPORT GENERATION (Kept from original)
# ==============================================================================
def generate_course_pdf_report(course_data, selected_course):
    # [Same as original]
    pass

# ==============================================================================
# EMAIL FUNCTIONS (Kept from original)
# ==============================================================================
def send_email(to_email, subject, body, pdf_buffer=None):
    # [Same as original]
    pass

def generate_individual_student_pdf(student, semester, course_code, course_name):
    # [Same as original]
    pass

def send_bulk_emails_to_parents(course_data, course_name):
    # [Same as original]
    pass

# ==============================================================================
# CHART GENERATION FUNCTIONS (Kept from original)
# ==============================================================================
def generate_marks_distribution_chart(course_data):
    # [Same as original]
    pass

def generate_grade_distribution_chart(course_data):
    # [Same as original]
    pass

def generate_co_attainment_chart(co_attainment):
    # [Same as original]
    pass

def generate_po_attainment_chart(po_attainment):
    # [Same as original]
    pass

# ==============================================================================
# SPIDER/RADAR PLOT FOR PO ATTAINMENT
# ==============================================================================
def create_spider_plot(po_attainment):
    """Create a spider/radar plot for PO attainment"""
    if not po_attainment:
        return None
    
    # Extract PO names and values
    po_names = list(po_attainment.keys())
    po_values = list(po_attainment.values())
    
    # Number of variables
    num_vars = len(po_names)
    
    # Compute angle for each axis
    angles = [n / float(num_vars) * 2 * np.pi for n in range(num_vars)]
    angles += angles[:1]  # Complete the circle
    
    # Add the first value to close the polygon
    values = po_values + [po_values[0]]
    
    # Create the plot
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
    
    # Draw one line per variable
    ax.plot(angles, values, 'o-', linewidth=2, color='#667eea')
    ax.fill(angles, values, alpha=0.25, color='#667eea')
    
    # Set labels
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(po_names, fontsize=10)
    
    # Set y-axis limits
    max_val = max(po_values) if po_values else 100
    ax.set_ylim(0, max(max_val * 1.2, 100))
    
    # Add value labels
    for angle, value, name in zip(angles[:-1], values[:-1], po_names):
        ax.annotate(f'{value:.1f}%', 
                   xy=(angle, value),
                   xytext=(5, 5),
                   textcoords='offset points',
                   fontsize=9,
                   fontweight='bold')
    
    ax.set_title('PO Attainment (Spider Plot)', fontsize=14, fontweight='bold', pad=20)
    ax.grid(True)
    
    return fig

# ==============================================================================
# STUDENT ANALYTICS PAGE (MODIFIED)
# ==============================================================================
def show_student_analytics():
    """Show student analytics with CO-PO attainment and spider plots"""
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Student Performance Analytics", unsafe_allow_html=True)
    
    # Help button
    col_title, col_help = st.columns([4, 1])
    with col_help:
        if st.button("Help", key="student_analytics_help"):
            st.info("""
            **Student Analytics Help:**
            
            1. **Course-wise Results**: View your performance in each course
            2. **CO Attainment**: Check Course Outcomes attainment for each course
            3. **PO Attainment**: View Program Outcomes attainment with spider plot
            4. **Class Position**: See your position in class
            5. **Career Prediction**: AI-based career recommendations
            """)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Get student ID based on user type
    student_id = ""
    student_name = ""
    
    if st.session_state.user_type == "student":
        student_id = st.session_state.user_data.get('student_id', '')
        student_name = st.session_state.user_data.get('full_name', 'Student')
        
        if not student_id:
            st.error("Student ID not found in your account.")
            return
    
    elif st.session_state.user_type == "parent":
        linked_student = get_linked_student_for_parent(st.session_state.username)
        if linked_student:
            student_id = linked_student.get('student_id', '')
            student_name = linked_student.get('full_name', 'Your Child')
        else:
            st.error("No student linked to your account.")
            return
    
    elif st.session_state.user_type in ["teacher", "admin"]:
        # Load all student data
        all_courses = load_all_courses()
        all_students = {}
        
        if all_courses:
            for course_key, course_data in all_courses.items():
                for stu_id, stu_data in course_data.get('students', {}).items():
                    if stu_id not in all_students:
                        all_students[stu_id] = stu_data.get('name', stu_id)
        
        if all_students:
            selected_student = st.selectbox(
                "Select Student:",
                list(all_students.keys()),
                format_func=lambda x: f"{all_students[x]} ({x})"
            )
            student_id = selected_student
            student_name = all_students[selected_student]
    
    if not student_id:
        st.info("No student data available.")
        return
    
    st.info(f"Viewing academic performance for: {student_name} ({student_id})")
    
    # Load student data
    student_data = load_student_data(student_id)
    
    if not student_data:
        st.info("No academic data found for this student.")
        return
    
    # Tabs for different views
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Course-wise Results",
        "CO Attainment",
        "PO Attainment",
        "Class Position",
        "Career Prediction"
    ])
    
    with tab1:
        show_course_wise_results(student_data, student_name)
    
    with tab2:
        show_co_attainment_analysis(student_data, student_name)
    
    with tab3:
        show_po_attainment_analysis(student_data, student_name)
    
    with tab4:
        show_class_position(student_data, student_id, student_name)
    
    with tab5:
        show_ai_career_prediction(student_data, student_id, student_name)

def show_course_wise_results(student_data, student_name):
    """Show course-wise results for a student"""
    st.markdown(f"#### Course-wise Results for {student_name}")
    
    course_results = []
    
    for course_key, course_info in student_data.items():
        student_course_data = course_info.get('student_data', {})
        course_results.append({
            'Course': course_info.get('course_code', course_key),
            'Semester': course_info.get('semester', 'N/A'),
            'Total Marks': student_course_data.get('total_marks', 0),
            'Grade': student_course_data.get('grade', 'N/A'),
            'SGPA': student_course_data.get('sgpa', 0),
            'Status': student_course_data.get('status', 'N/A')
        })
    
    if course_results:
        df_results = pd.DataFrame(course_results)
        
        # Summary metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            avg_marks = df_results['Total Marks'].mean()
            st.metric("Average Marks", f"{avg_marks:.1f}")
        with col2:
            avg_sgpa = df_results['SGPA'].mean()
            st.metric("Average SGPA", f"{avg_sgpa:.2f}")
        with col3:
            passed = len(df_results[df_results['Status'] == 'Pass'])
            total = len(df_results)
            st.metric("Pass Rate", f"{passed}/{total}")
        
        # Results table
        st.dataframe(df_results, use_container_width=True, hide_index=True)
        
        # Marks bar chart
        fig = go.Figure(data=[
            go.Bar(
                x=df_results['Course'],
                y=df_results['Total Marks'],
                marker_color='#667eea',
                text=df_results['Total Marks'].round(1),
                textposition='auto'
            )
        ])
        fig.update_layout(
            title="Marks by Course",
            xaxis_title="Course",
            yaxis_title="Total Marks",
            yaxis_range=[0, 100],
            height=400
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No course results available.")

def show_co_attainment_analysis(student_data, student_name):
    """Show CO attainment analysis for each course"""
    st.markdown(f"#### CO Attainment Analysis for {student_name}")
    
    all_co_data = {}
    
    for course_key, course_info in student_data.items():
        student_course_data = course_info.get('student_data', {})
        co_scores = student_course_data.get('co_scores', {})
        
        if co_scores:
            course_code = course_info.get('course_code', course_key)
            all_co_data[course_code] = co_scores
    
    if not all_co_data:
        st.info("No CO attainment data available.")
        return
    
    # Course selection for CO view
    selected_course = st.selectbox(
        "Select Course for CO Analysis:",
        list(all_co_data.keys())
    )
    
    if selected_course:
        co_scores = all_co_data[selected_course]
        
        # CO attainment bar chart
        cos = list(co_scores.keys())
        scores = list(co_scores.values())
        
        # Calculate percentages (assuming max 20 per CO)
        percentages = [(score / 20) * 100 for score in scores]
        
        fig = go.Figure(data=[
            go.Bar(
                x=cos,
                y=percentages,
                marker_color=['#4CAF50' if p >= 50 else '#FFC107' if p >= 40 else '#F44336' for p in percentages],
                text=[f'{p:.1f}%' for p in percentages],
                textposition='auto'
            )
        ])
        fig.update_layout(
            title=f"CO Attainment - {selected_course}",
            xaxis_title="Course Outcomes",
            yaxis_title="Attainment (%)",
            yaxis_range=[0, 100],
            height=400
        )
        fig.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="50% Threshold")
        st.plotly_chart(fig, use_container_width=True)
        
        # CO details table
        co_details = []
        for co, score in co_scores.items():
            percentage = (score / 20) * 100
            if percentage >= 50:
                status = "Achieved"
                color = "green"
            elif percentage >= 40:
                status = "Marginally Achieved"
                color = "orange"
            else:
                status = "Not Achieved"
                color = "red"
            
            co_details.append({
                'CO': co,
                'Score': f"{score:.1f}/20",
                'Percentage': f"{percentage:.1f}%",
                'Status': status
            })
        
        df_co_details = pd.DataFrame(co_details)
        st.dataframe(df_co_details, use_container_width=True, hide_index=True)
    
    # Overall CO attainment across all courses
    st.markdown("---")
    st.markdown("#### Overall CO Attainment (All Courses)")
    
    # Calculate average CO scores across courses
    avg_co_scores = {}
    for course_code, co_scores in all_co_data.items():
        for co, score in co_scores.items():
            if co not in avg_co_scores:
                avg_co_scores[co] = []
            avg_co_scores[co].append(score)
    
    avg_co = {co: np.mean(scores) for co, scores in avg_co_scores.items()}
    
    if avg_co:
        avg_percentages = [(score / 20) * 100 for score in avg_co.values()]
        
        fig_avg = go.Figure(data=[
            go.Bar(
                x=list(avg_co.keys()),
                y=avg_percentages,
                marker_color='#667eea',
                text=[f'{p:.1f}%' for p in avg_percentages],
                textposition='auto'
            )
        ])
        fig_avg.update_layout(
            title="Average CO Attainment Across All Courses",
            xaxis_title="Course Outcomes",
            yaxis_title="Attainment (%)",
            yaxis_range=[0, 100],
            height=400
        )
        fig_avg.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="50% Threshold")
        st.plotly_chart(fig_avg, use_container_width=True)

def show_po_attainment_analysis(student_data, student_name):
    """Show PO attainment analysis with spider plot"""
    st.markdown(f"#### PO Attainment Analysis for {student_name}")
    st.markdown("Program Outcomes (PO) attainment based on all completed courses")
    
    # Calculate PO attainment from CO scores and CO-PO mapping
    all_po_data = {}
    
    for course_key, course_info in student_data.items():
        student_course_data = course_info.get('student_data', {})
        co_scores = student_course_data.get('co_scores', {})
        
        if co_scores:
            # Get CO-PO mapping for this course
            course_data = load_all_courses().get(course_key, {})
            co_po_mapping = course_data.get('co_po_mapping')
            
            if co_po_mapping is None:
                # Use default mapping if not available
                co_po_mapping = create_default_copo_mapping()
            
            # Calculate PO attainment
            po_attainment = calculate_po_attainment(
                {co: score / 20 * 100 for co, score in co_scores.items()},
                co_po_mapping
            )
            
            if po_attainment:
                course_code = course_info.get('course_code', course_key)
                all_po_data[course_code] = po_attainment
    
    if not all_po_data:
        st.info("No PO attainment data available.")
        return
    
    # Calculate average PO attainment across courses
    avg_po = {}
    for course_code, po_scores in all_po_data.items():
        for po, score in po_scores.items():
            if po not in avg_po:
                avg_po[po] = []
            avg_po[po].append(score)
    
    overall_po = {po: np.mean(scores) for po, scores in avg_po.items() if scores}
    
    if overall_po:
        # Create spider plot
        st.markdown("##### PO Attainment Spider Plot")
        spider_fig = create_spider_plot(overall_po)
        if spider_fig:
            st.pyplot(spider_fig)
        
        # PO details table
        st.markdown("##### PO Attainment Details")
        po_details = []
        for po, score in overall_po.items():
            if score >= 70:
                status = "Excellent"
                color = "#4CAF50"
            elif score >= 50:
                status = "Satisfactory"
                color = "#FFC107"
            else:
                status = "Needs Improvement"
                color = "#F44336"
            
            po_details.append({
                'PO': po,
                'Attainment': f"{score:.1f}%",
                'Status': status
            })
        
        df_po = pd.DataFrame(po_details)
        st.dataframe(df_po, use_container_width=True, hide_index=True)
        
        # Course-wise PO comparison
        st.markdown("##### Course-wise PO Attainment")
        course_for_po = st.selectbox(
            "Select Course for PO View:",
            list(all_po_data.keys())
        )
        
        if course_for_po:
            course_po = all_po_data[course_for_po]
            
            pos = list(course_po.keys())
            values = list(course_po.values())
            
            fig_course_po = go.Figure(data=[
                go.Bar(
                    x=pos,
                    y=values,
                    marker_color=['#4CAF50' if v >= 70 else '#FFC107' if v >= 50 else '#F44336' for v in values],
                    text=[f'{v:.1f}%' for v in values],
                    textposition='auto'
                )
            ])
            fig_course_po.update_layout(
                title=f"PO Attainment - {course_for_po}",
                xaxis_title="Program Outcomes",
                yaxis_title="Attainment (%)",
                yaxis_range=[0, 100],
                height=400
            )
            fig_course_po.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="50% Threshold")
            st.plotly_chart(fig_course_po, use_container_width=True)

def show_class_position(student_data, student_id, student_name):
    """Show student's class position for each course"""
    st.markdown(f"#### Class Position for {student_name}")
    
    for course_key, course_info in student_data.items():
        student_course_data = course_info.get('student_data', {})
        course_code = course_info.get('course_code', course_key)
        
        # Get all students for this course
        all_courses = load_all_courses()
        course_full_data = all_courses.get(course_key, {})
        all_students = course_full_data.get('students', {})
        
        if all_students and student_id in all_students:
            # Sort students by marks
            sorted_students = sorted(all_students.items(), 
                                    key=lambda x: x[1].get('total_marks', 0), 
                                    reverse=True)
            
            # Find student's position
            position = next((i + 1 for i, (sid, _) in enumerate(sorted_students) 
                           if sid == student_id), len(sorted_students))
            
            total_students = len(sorted_students)
            student_marks = student_course_data.get('total_marks', 0)
            highest_marks = sorted_students[0][1].get('total_marks', 0) if sorted_students else 0
            avg_marks = np.mean([s[1].get('total_marks', 0) for s in sorted_students])
            
            with st.expander(f"{course_code} - Position: {position}/{total_students}", expanded=True):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Your Position", f"{position}/{total_students}")
                with col2:
                    st.metric("Your Marks", f"{student_marks:.1f}")
                with col3:
                    st.metric("Highest Marks", f"{highest_marks:.1f}")
                with col4:
                    st.metric("Class Average", f"{avg_marks:.1f}")
                
                # Position visualization
                fig = go.Figure()
                
                # All students as bars
                marks_list = [s[1].get('total_marks', 0) for s in sorted_students]
                names_list = [f"Student {i+1}" for i in range(len(sorted_students))]
                
                colors = ['#667eea'] * len(sorted_students)
                if position <= len(sorted_students):
                    colors[position - 1] = '#4CAF50'  # Highlight student
                
                fig.add_trace(go.Bar(
                    x=list(range(1, len(sorted_students) + 1)),
                    y=marks_list,
                    marker_color=colors,
                    text=[f'{m:.1f}' for m in marks_list],
                    textposition='auto'
                ))
                
                fig.update_layout(
                    title=f"Class Ranking - {course_code}",
                    xaxis_title="Student Rank",
                    yaxis_title="Total Marks",
                    height=300,
                    showlegend=False
                )
                
                st.plotly_chart(fig, use_container_width=True)

def show_ai_career_prediction(student_data, student_id, student_name):
    """Show AI-powered career predictions"""
    st.markdown(f"#### AI Career Prediction for {student_name}")
    
    st.warning("""
    **DISCLAIMER**
    
    This prediction is AI-generated and should be considered as guidance only.
    Discuss career choices with your academic supervisor and career counselor.
    """)
    
    # Get predictions from student data
    predictions = None
    for course_key, course_info in student_data.items():
        if 'predictions' in course_info and course_info['predictions']:
            predictions = course_info['predictions']
            break
    
    if not predictions:
        # Generate rule-based prediction if no ML prediction available
        if student_data:
            latest_course_key = sorted(student_data.keys())[-1]
            latest_student_data = student_data[latest_course_key].get('student_data', {})
            predictions = generate_rule_based_prediction(latest_student_data)
    
    if predictions:
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### Current Assessment")
            st.info(f"**Performance:** {predictions.get('current_performance', 'Not available')}")
            
            st.markdown("##### Key Strengths")
            strengths = predictions.get('key_strengths', [])
            if strengths:
                for strength in strengths:
                    st.markdown(f"- {strength}")
            else:
                st.info("Developing core competencies")
        
        with col2:
            st.markdown("##### Recommended Career Path")
            career_sector = predictions.get('recommended_career_sector', 'Not available')
            
            career_colors = {
                "Research & Academia": "#4CAF50",
                "Power Systems & Energy": "#2196F3",
                "Electronics & Embedded Systems": "#FF9800",
                "Control & Automation": "#9C27B0",
                "Telecommunications": "#00BCD4",
                "Renewable Energy": "#8BC34A",
                "AI & Machine Learning in EEE": "#FF5722"
            }
            
            color = career_colors.get(career_sector, "#667eea")
            st.markdown(f"""
            <div style="text-align: center; padding: 1.5rem; background: {color}20; border-radius: 15px; border: 2px solid {color};">
                <h3 style="color: {color};">{career_sector}</h3>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("##### Recommendations")
            recommendation = predictions.get('recommendation', 'Continue developing your skills')
            st.success(recommendation)
        
        # CO-PO based career insight
        st.markdown("---")
        st.markdown("##### Career Insights Based on CO-PO Attainment")
        
        # Analyze CO and PO strengths
        co_strengths = []
        po_strengths = []
        
        for course_key, course_info in student_data.items():
            student_course_data = course_info.get('student_data', {})
            co_scores = student_course_data.get('co_scores', {})
            
            for co, score in co_scores.items():
                if score >= 15:
                    co_strengths.append(co)
        
        if co_strengths:
            st.markdown("**Strong Course Outcomes:**")
            for co in set(co_strengths):
                st.markdown(f"- {co}: Above 75% attainment")
            
            # Career recommendations based on CO strengths
            st.markdown("**Career Sectors Aligned with Your Strengths:**")
            career_recommendations = []
            
            if 'CO1' in co_strengths:
                career_recommendations.append("Power Systems & Energy Sector (Strong theoretical foundation)")
            if 'CO2' in co_strengths:
                career_recommendations.append("Research & Development (Strong problem-solving skills)")
            if 'CO3' in co_strengths:
                career_recommendations.append("Control & Automation (Strong analytical skills)")
            if 'CO4' in co_strengths:
                career_recommendations.append("Project Management & Consulting (Strong professional skills)")
            
            if career_recommendations:
                for rec in career_recommendations:
                    st.markdown(f"- {rec}")
            else:
                st.info("Build stronger CO attainment for specialized career paths")
    else:
        st.info("No prediction data available yet. More academic data is needed for accurate predictions.")

# ==============================================================================
# TEACHER COURSE REPORT PAGE (MODIFIED)
# ==============================================================================
def show_course_reports():
    """Show course reports for teachers - can only see their own courses"""
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("<h3>Course Reports & Analytics</h3>", unsafe_allow_html=True)
    
    # Load all courses
    all_courses = load_all_courses()
    
    if not all_courses:
        st.info("No course data available yet. Please upload data first.")
        if st.button("Go to Upload Page"):
            st.session_state.current_page = "upload"
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        return
    
    # Filter courses for teacher - only show their courses
    if st.session_state.user_type == "teacher":
        teacher_courses = {}
        for course_key, course_data in all_courses.items():
            if course_data.get('teacher') == st.session_state.username:
                teacher_courses[course_key] = course_data
        all_courses = teacher_courses
        
        if not teacher_courses:
            st.info("You haven't uploaded any course data yet.")
            if st.button("Upload Course Data"):
                st.session_state.current_page = "upload"
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            return
    
    # Course selection
    selected_course = st.selectbox(
        "Select Course Report",
        list(all_courses.keys())
    )
    
    if selected_course:
        course_data = all_courses[selected_course]
        
        # Display course information
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"**Semester:** {course_data.get('semester', 'N/A')}")
        with col2:
            st.markdown(f"**Course Code:** {course_data.get('course_code', 'N/A')}")
        with col3:
            stats = course_data.get('course_stats', {})
            st.markdown(f"**Total Students:** {stats.get('total_students', 0)}")
        
        # PDF Download
        st.markdown("---")
        if st.button("Download PDF Report", type="primary", use_container_width=True):
            with st.spinner("Generating PDF report..."):
                pdf_buffer = generate_course_pdf_report(course_data, selected_course)
                
                st.download_button(
                    label="Download PDF Report",
                    data=pdf_buffer,
                    file_name=f"Course_Report_{course_data.get('course_code', 'Unknown')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("PDF report generated successfully!")
        
        # KPI Metrics
        if stats:
            st.markdown("#### Performance Metrics")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(create_colored_metric_card(
                    "Avg Marks", stats.get("average_marks", 0), "marks", ""
                ), unsafe_allow_html=True)
            with col2:
                st.markdown(create_colored_metric_card(
                    "Pass Rate", stats.get("pass_percentage", 0), "pass_rate", "%"
                ), unsafe_allow_html=True)
            with col3:
                st.markdown(create_colored_metric_card(
                    "Avg SGPA", stats.get("average_sgpa", 0), "cgpa", ""
                ), unsafe_allow_html=True)
            with col4:
                st.markdown(create_colored_metric_card(
                    "Highest", stats.get("highest_marks", 0), "marks", ""
                ), unsafe_allow_html=True)
        
        # Charts
        st.markdown("#### Visual Analytics")
        
        tab1, tab2, tab3, tab4 = st.tabs([
            "Marks Distribution",
            "Grade Distribution",
            "CO-PO Attainment",
            "Students"
        ])
        
        with tab1:
            # Marks distribution
            students = course_data.get('students', {})
            marks = [student['total_marks'] for student in students.values()]
            
            fig1 = go.Figure()
            fig1.add_trace(go.Histogram(
                x=marks, nbinsx=10,
                marker_color='#667eea', opacity=0.7
            ))
            fig1.update_layout(
                title="Marks Distribution",
                xaxis_title="Total Marks",
                yaxis_title="Number of Students",
                height=400
            )
            st.plotly_chart(fig1, use_container_width=True)
            
            col_stats1, col_stats2, col_stats3 = st.columns(3)
            with col_stats1:
                st.metric("Mean", f"{np.mean(marks):.1f}")
            with col_stats2:
                st.metric("Median", f"{np.median(marks):.1f}")
            with col_stats3:
                st.metric("Std Dev", f"{np.std(marks):.1f}")
        
        with tab2:
            # Grade distribution
            grades = [student['grade'] for student in students.values()]
            grade_counts = pd.Series(grades).value_counts()
            
            fig2 = go.Figure(data=[go.Pie(
                labels=grade_counts.index,
                values=grade_counts.values,
                hole=.3,
                marker_colors=['#4CAF50', '#8BC34A', '#FFC107', '#FF9800', '#FF5722', '#F44336'],
                textinfo='label+percent'
            )])
            fig2.update_layout(title="Grade Distribution", height=500)
            st.plotly_chart(fig2, use_container_width=True)
        
        with tab3:
            # CO-PO Attainment for this course
            co_attainment = course_data.get('co_attainment', {})
            po_attainment = course_data.get('po_attainment', {})
            
            if co_attainment:
                cos = list(co_attainment.keys())
                co_values = list(co_attainment.values())
                
                fig3 = go.Figure(data=[go.Bar(
                    x=cos, y=co_values,
                    marker_color=[get_color_by_value(v, "attainment") for v in co_values],
                    text=[f'{v:.1f}%' for v in co_values],
                    textposition='auto'
                )])
                fig3.update_layout(
                    title="CO Attainment",
                    yaxis_title="Attainment (%)",
                    yaxis_range=[0, 100],
                    height=400
                )
                st.plotly_chart(fig3, use_container_width=True)
            
            if po_attainment:
                pos = list(po_attainment.keys())
                po_values = list(po_attainment.values())
                
                fig4 = go.Figure(data=[go.Bar(
                    x=pos, y=po_values,
                    marker_color=[get_color_by_value(v, "attainment") for v in po_values],
                    text=[f'{v:.1f}%' for v in po_values],
                    textposition='auto'
                )])
                fig4.update_layout(
                    title="PO Attainment",
                    yaxis_title="Attainment (%)",
                    yaxis_range=[0, 100],
                    height=400
                )
                st.plotly_chart(fig4, use_container_width=True)
                
                # Spider plot for PO
                st.markdown("##### PO Attainment Spider Plot")
                spider_fig = create_spider_plot(po_attainment)
                if spider_fig:
                    st.pyplot(spider_fig)
            
            if not co_attainment and not po_attainment:
                st.info("CO-PO attainment data not available for this course.")
        
        with tab4:
            # Student details
            student_data = []
            for student_id, student in students.items():
                student_data.append({
                    'ID': student_id,
                    'Name': student['name'],
                    'Total Marks': student['total_marks'],
                    'Grade': student['grade'],
                    'SGPA': student['sgpa'],
                    'Status': student['status']
                })
            
            df_students = pd.DataFrame(student_data)
            st.dataframe(
                df_students.sort_values('Total Marks', ascending=False),
                use_container_width=True,
                height=400
            )
        
        # Email section for teachers
        if st.session_state.user_type == "teacher":
            st.markdown("---")
            st.markdown("#### Send Reports to Parents")
            
            valid_emails = len([s for s in students.values() 
                              if s.get('parent_email') and s.get('parent_email') not in ['', 'nan', 'NaN']])
            st.info(f"{valid_emails} parent emails found out of {len(students)} students")
            
            if st.button("Send Bulk Emails", type="primary", use_container_width=True):
                with st.spinner("Sending emails to parents..."):
                    success, fail, log, _ = send_bulk_emails_to_parents(course_data, selected_course)
                    
                    if success > 0:
                        st.success(f"Successfully sent: {success} emails")
                    if fail > 0:
                        st.error(f"Failed: {fail} emails")
                    
                    with st.expander("Email Log", expanded=True):
                        for entry in log:
                            st.text(entry)
    
    # Back button
    if st.button("Back to Dashboard", use_container_width=True):
        st.session_state.current_page = "dashboard"
        st.rerun()
    
    st.markdown("</div>", unsafe_allow_html=True)

# ==============================================================================
# DATA PROCESSING (MODIFIED FOR DYNAMIC CO-PO)
# ==============================================================================
def process_student_data(df, semester, course_code, teacher_username, filename):
    """Process student data with dynamic CO and PO detection"""
    results = {
        'students': {},
        'course_stats': {},
        'co_attainment': {},
        'po_attainment': {},
        'semester': semester,
        'course_code': course_code,
        'teacher': teacher_username,
        'filename': filename,
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # Detect CO and PO columns dynamically
    co_columns = [col for col in df.columns if 'co' in str(col).lower()]
    po_columns = [col for col in df.columns if 'po' in str(col).lower()]
    
    co_scores_all = []
    
    for idx, row in df.iterrows():
        try:
            student_id = str(row.get('Student_ID', f'STU{idx}'))
            student_name = str(row.get('Student_Name', f'Student {idx}'))
            
            # Extract marks from different components
            marks = {}
            
            # Mid term
            mid_cols = [col for col in df.columns if 'mid' in str(col).lower()]
            mid_total = sum([float(row[col]) for col in mid_cols if pd.notna(row.get(col, 0))])
            
            # Final
            final_cols = [col for col in df.columns if 'final' in str(col).lower()]
            final_total = sum([float(row[col]) for col in final_cols if pd.notna(row.get(col, 0))])
            
            # Class Test
            ct_cols = [col for col in df.columns if 'ct' in str(col).lower()]
            ct_total = sum([float(row[col]) for col in ct_cols if pd.notna(row.get(col, 0))]) / max(1, len(ct_cols))
            
            # Assignment
            assignment_cols = [col for col in df.columns if 'assignment' in str(col).lower()]
            assignment_total = sum([float(row[col]) for col in assignment_cols if pd.notna(row.get(col, 0))])
            
            # Attendance
            attendance = float(row.get('Attendance', 0)) if pd.notna(row.get('Attendance', 0)) else 0
            
            # Scale to expected ranges
            mid_scaled = min(30, mid_total) if mid_total > 0 else 0
            final_scaled = min(40, final_total) if final_total > 0 else 0
            ct_scaled = min(20, ct_total) if ct_total > 0 else 0
            assignment_scaled = min(5, assignment_total) if assignment_total > 0 else 0
            attendance_scaled = min(5, attendance) if attendance > 0 else 0
            
            total_marks = mid_scaled + final_scaled + ct_scaled + assignment_scaled + attendance_scaled
            
            # Calculate SGPA and Grade
            sgpa = calculate_sgpa(total_marks)
            grade = get_grade_from_marks(total_marks)
            
            # Extract CO scores dynamically
            co_scores = {}
            for co_col in co_columns:
                val = row.get(co_col, 0)
                co_scores[co_col] = float(val) if pd.notna(val) else 0
            
            # Extract PO scores if available
            po_scores = {}
            for po_col in po_columns:
                val = row.get(po_col, 0)
                po_scores[po_col] = float(val) if pd.notna(val) else 0
            
            results['students'][student_id] = {
                'id': student_id,
                'name': student_name,
                'mid': mid_scaled,
                'final': final_scaled,
                'ct': ct_scaled,
                'assignment': assignment_scaled,
                'attendance': attendance_scaled,
                'total_marks': total_marks,
                'sgpa': sgpa,
                'grade': grade,
                'co_scores': co_scores,
                'po_scores': po_scores,
                'student_email': str(row.get('Student_Email', '')),
                'parent_email': str(row.get('Parent_Email', '')),
                'status': 'Pass' if total_marks >= 40 else 'Fail'
            }
            
            co_scores_all.append(co_scores)
            
        except Exception as e:
            st.warning(f"Error processing student {idx}: {str(e)}")
    
    # Calculate course statistics
    if results['students']:
        marks_list = [s['total_marks'] for s in results['students'].values()]
        sgpas = [s['sgpa'] for s in results['students'].values()]
        passing = [m for m in marks_list if m >= 40]
        
        results['course_stats'] = {
            'average_marks': np.mean(marks_list),
            'highest_marks': max(marks_list),
            'lowest_marks': min(marks_list),
            'average_sgpa': np.mean(sgpas),
            'total_students': len(marks_list),
            'passing_students': len(passing),
            'pass_percentage': (len(passing) / len(marks_list) * 100) if marks_list else 0,
            'fail_percentage': ((len(marks_list) - len(passing)) / len(marks_list) * 100) if marks_list else 0,
            'std_deviation': np.std(marks_list)
        }
    
    # Calculate CO attainment
    if co_scores_all and co_scores_all[0]:
        df_co = pd.DataFrame(co_scores_all)
        # Convert to percentage (assumes max 20 per CO)
        co_maxes = {}
        for col in df_co.columns:
            max_val = df_co[col].max()
            co_maxes[col] = max(max_val, 1)  # Avoid division by zero
        
        results['co_attainment'] = {
            col: (df_co[col].mean() / co_maxes[col] * 100) 
            for col in df_co.columns
        }
        
        # Calculate PO attainment if mapping exists
        default_mapping = create_default_copo_mapping()
        results['po_attainment'] = calculate_po_attainment(
            {co: score / 20 * 100 for co, score in results['co_attainment'].items()},
            default_mapping
        )
    
    # Generate predictions
    results['predictions'] = generate_ai_predictions(results)
    
    # Track upload and save
    track_teacher_upload(teacher_username, semester, course_code, filename, len(results['students']))
    save_course_data(semester, course_code, results)
    
    st.session_state.processed = True
    st.session_state.results = results
    
    return results

# ==============================================================================
# AI PREDICTIONS (Kept from original with CO-PO integration)
# ==============================================================================
def generate_ai_predictions(results):
    """Generate AI predictions using ML and CO-PO attainment"""
    predictions = {}
    if not results.get('students'):
        return predictions
    
    students_data = []
    student_ids = []
    
    for student_id, student in results['students'].items():
        # Include CO scores in features
        co_scores = list(student.get('co_scores', {}).values())
        co_avg = np.mean(co_scores) if co_scores else 0
        
        features = [
            student.get('total_marks', 0),
            student.get('mid', 0),
            student.get('final', 0),
            student.get('ct', 0),
            student.get('assignment', 0),
            student.get('sgpa', 0),
            co_avg  # Add CO average as feature
        ]
        students_data.append(features)
        student_ids.append(student_id)
    
    if len(students_data) < 3:
        # Not enough data for ML, use rule-based
        for student_id, student in results['students'].items():
            predictions[student_id] = generate_rule_based_prediction(student)
        return predictions
    
    X = np.array(students_data)
    
    # Academic prediction model
    y_academic = X[:, 0]
    model_academic = LinearRegression()
    model_academic.fit(X[:, 1:], y_academic)
    
    # Career prediction with CO-PO based sectors
    career_sectors = [
        "Power Systems & Energy",
        "Electronics & Embedded Systems", 
        "Telecommunications",
        "Control & Automation",
        "Research & Academia",
        "Renewable Energy",
        "AI & Machine Learning in EEE"
    ]
    
    y_career = []
    for features in X:
        total_marks = features[0]
        sgpa = features[5]
        co_avg = features[6]
        
        # Career mapping based on performance and CO scores
        if total_marks >= 80 and sgpa >= 3.5:
            y_career.append(0)  # Research
        elif total_marks >= 75:
            y_career.append(1)  # Electronics
        elif total_marks >= 70:
            y_career.append(2)  # Telecom
        elif total_marks >= 65 and co_avg >= 15:
            y_career.append(3)  # Control
        elif total_marks >= 60:
            y_career.append(4)  # Academia
        elif total_marks >= 50:
            y_career.append(5)  # Renewable
        else:
            y_career.append(6)  # AI/ML
    
    model_career = RandomForestClassifier(n_estimators=50, random_state=42)
    model_career.fit(X[:, 1:], y_career)
    
    for idx, student_id in enumerate(student_ids):
        student = results['students'][student_id]
        features = X[idx]
        
        # Predict next semester performance
        next_sem_pred = model_academic.predict([features[1:]])[0]
        next_sem_pred = max(40, min(95, next_sem_pred))
        
        # Predict career sector
        career_idx = model_career.predict([features[1:]])[0]
        career_sector = career_sectors[career_idx]
        
        # Calculate growth
        current_marks = features[0]
        growth_percent = ((next_sem_pred - current_marks) / current_marks * 100) if current_marks > 0 else 100
        
        # Performance assessment
        if current_marks >= 80:
            performance = "Excellent"
            recommendation = "Consider graduate studies or research positions"
        elif current_marks >= 70:
            performance = "Good"
            recommendation = "Focus on specialization in strong areas"
        elif current_marks >= 60:
            performance = "Average"
            recommendation = "Improve weak areas through practice"
        elif current_marks >= 40:
            performance = "Satisfactory"
            recommendation = "Maintain consistency and seek guidance"
        else:
            performance = "Needs Improvement"
            recommendation = "Seek academic support"
        
        # Identify strengths from CO scores
        co_scores = student.get('co_scores', {})
        strengths = []
        for co, score in co_scores.items():
            if score >= 15:
                strengths.append(f"Strong in {co}")
        if not strengths:
            strengths = ["Developing core engineering skills"]
        
        predictions[student_id] = {
            'student_name': student['name'],
            'current_performance': f"{current_marks:.1f} marks ({performance})",
            'predicted_next_semester': f"{next_sem_pred:.1f} marks",
            'growth_percentage': f"{growth_percent:.1f}%",
            'recommended_career_sector': career_sector,
            'key_strengths': strengths[:3],
            'recommendation': recommendation,
            'confidence_level': "Medium" if len(students_data) >= 5 else "Low"
        }
    
    return predictions

# ==============================================================================
# RULE-BASED PREDICTION (Kept from original)
# ==============================================================================
def generate_rule_based_prediction(student):
    """Generate rule-based predictions when insufficient data for ML"""
    total_marks = student.get('total_marks', 0)
    sgpa = student.get('sgpa', 0)
    
    if total_marks >= 80:
        performance = "Excellent"
        next_sem = min(95, total_marks + np.random.uniform(0, 5))
        career = np.random.choice(["Research & Academia", "Power Systems Design", "Advanced Electronics"])
        recommendation = "Pursue graduate studies or competitive industry positions"
    elif total_marks >= 70:
        performance = "Good"
        next_sem = min(90, total_marks + np.random.uniform(-2, 8))
        career = np.random.choice(["Energy Management", "Control Systems", "Telecommunications"])
        recommendation = "Focus on specialization and internships"
    elif total_marks >= 60:
        performance = "Average"
        next_sem = min(85, total_marks + np.random.uniform(-5, 10))
        career = np.random.choice(["Renewable Energy", "Maintenance Engineering", "Technical Sales"])
        recommendation = "Improve fundamentals and seek practical experience"
    elif total_marks >= 40:
        performance = "Satisfactory"
        next_sem = max(40, total_marks + np.random.uniform(-10, 15))
        career = "General Engineering with focused skill development"
        recommendation = "Maintain consistency and seek academic guidance"
    else:
        performance = "Needs Improvement"
        next_sem = max(30, total_marks + np.random.uniform(-5, 20))
        career = "Foundation strengthening required"
        recommendation = "Seek academic support and focus on core concepts"
    
    # Identify strengths
    strengths = []
    if student.get('mid', 0) >= 20:
        strengths.append("Good exam preparation skills")
    if student.get('final', 0) >= 30:
        strengths.append("Strong comprehensive understanding")
    if student.get('ct', 0) >= 15:
        strengths.append("Consistent performance in assessments")
    if student.get('assignment', 0) >= 4:
        strengths.append("Good assignment completion")
    
    if not strengths:
        strengths = ["Developing engineering competencies"]
    
    growth = ((next_sem - total_marks) / total_marks * 100) if total_marks > 0 else 100
    
    return {
        'student_name': student['name'],
        'current_performance': f"{total_marks:.1f} marks ({performance})",
        'predicted_next_semester': f"{next_sem:.1f} marks",
        'growth_percentage': f"{growth:.1f}%",
        'recommended_career_sector': career,
        'key_strengths': strengths[:3],
        'recommendation': recommendation,
        'confidence_level': "Low (Rule-based)"
    }

# ==============================================================================
# UPLOAD PAGE (MODIFIED FOR DYNAMIC FORMAT)
# ==============================================================================
def organized_upload_page():
    """Upload page that handles dynamic Excel formats"""
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("<h3>Upload Student Data</h3>", unsafe_allow_html=True)
    
    # Help
    col_title, col_help = st.columns([4, 1])
    with col_help:
        if st.button("Help", key="upload_help"):
            st.info("""
            **Upload Help:**
            
            1. **Download Template**: Get the Excel template
            2. **Enter Details**: Provide semester, course code
            3. **Upload File**: Select your Excel file
            4. **Process Data**: Click process to analyze
            
            The system can handle various Excel formats with different CO and PO structures.
            """)
    
    st.markdown("#### Step 1: Download Template")
    excel_file = create_sample_excel()
    st.download_button(
        label="Download XLSX Template",
        data=excel_file,
        file_name="Course_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("---")
    st.markdown("#### Step 2: Enter Academic Details")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        semester = st.text_input(
            "Semester:",
            value=get_current_semester(),
            placeholder="e.g., Spring 2023"
        )
    with col2:
        course_code = st.text_input(
            "Course Code:",
            value="",
            placeholder="e.g., EEE 321"
        )
    with col3:
        course_name = st.text_input(
            "Course Name:",
            value="",
            placeholder="e.g., Power System I"
        )
    
    st.markdown("---")
    st.markdown("#### Step 3: Upload Your File")
    
    uploaded_file = st.file_uploader(
        "Choose your Excel file",
        type=['xlsx', 'xls'],
        help="Upload the Excel file containing student marks"
    )
    
    if uploaded_file is not None:
        try:
            # Parse the Excel file dynamically
            all_data, co_po_info, sheet_names = parse_excel_file(uploaded_file)
            
            st.success(f"File uploaded successfully! Found {len(sheet_names)} sheets with student data.")
            
            # Show preview of sheets found
            with st.expander("File Analysis", expanded=True):
                st.markdown("**Sheets detected:**")
                for sheet in sheet_names:
                    st.markdown(f"- {sheet}")
                
                if co_po_info:
                    st.markdown("**CO-PO columns detected:**")
                    for sheet, info in co_po_info.items():
                        st.markdown(f"- {sheet}: {len(info['co_columns'])} COs, {len(info['po_columns'])} POs")
            
            # Process the data
            st.markdown("---")
            st.markdown("#### Step 4: Process Data")
            
            if st.button("Process & Analyze Data", use_container_width=True, type="primary"):
                if not semester or not course_code:
                    st.error("Please enter semester and course code")
                else:
                    with st.spinner("Processing data..."):
                        # Extract CO-PO attainment
                        attainment_data = extract_co_po_attainment(all_data, co_po_info)
                        
                        # Process student data
                        if sheet_names:
                            # Use the first sheet with student data
                            main_sheet = sheet_names[0]
                            df = all_data[main_sheet]['data']
                            
                            # Handle the dataframe based on its structure
                            results = process_student_data(
                                df, semester, course_code,
                                st.session_state.username,
                                uploaded_file.name
                            )
                            
                            # Add CO-PO attainment data
                            results['co_attainment'] = attainment_data.get('co_attainment', {})
                            results['po_attainment'] = attainment_data.get('po_attainment', {})
                            
                            st.session_state.results = results
                            key = f"{semester} - {course_code}"
                            st.session_state.all_semester_data[key] = results
                            
                            st.success("Data processing complete!")
                            
                            # Quick summary
                            st.markdown("##### Quick Summary")
                            stats = results['course_stats']
                            col_sum1, col_sum2, col_sum3 = st.columns(3)
                            with col_sum1:
                                st.metric("Total Students", stats['total_students'])
                            with col_sum2:
                                st.metric("Average Marks", f"{stats['average_marks']:.1f}")
                            with col_sum3:
                                st.metric("Pass Rate", f"{stats['pass_percentage']:.1f}%")
                        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.info("Please ensure your file follows the expected format")
    
    st.markdown("</div>", unsafe_allow_html=True)

# ==============================================================================
# MAIN APPLICATION (MODIFIED NAVIGATION)
# ==============================================================================
def main():
    """Main application controller"""
    try:
        apply_professional_theme()
        
        if not st.session_state.logged_in:
            if st.session_state.current_page == "register":
                show_registration_page()
            elif st.session_state.current_page == "forgot_password":
                show_forgot_password_page()
            else:
                show_login_page()
            return
        
        show_sidebar()
        
        if st.session_state.get('show_about', False):
            show_about_page()
            show_footer()
            return
        
        # Navigation based on user type and current page
        if st.session_state.current_page == "dashboard":
            if st.session_state.user_type in ["teacher", "admin"]:
                show_dashboard()
            else:
                st.session_state.current_page = "student_analytics"
                st.rerun()
        
        elif st.session_state.current_page == "upload":
            if st.session_state.user_type in ["teacher", "admin"]:
                organized_upload_page()
            else:
                st.error("Access Denied: You don't have permission to access this page.")
                st.session_state.current_page = "student_analytics"
                st.rerun()
        
        elif st.session_state.current_page == "course_reports":
            if st.session_state.user_type in ["teacher", "admin"]:
                show_course_reports()
            else:
                st.error("Access Denied: You don't have permission to access this page.")
                st.session_state.current_page = "student_analytics"
                st.rerun()
        
        elif st.session_state.current_page == "student_analytics":
            show_student_analytics()
        
        elif st.session_state.current_page == "admin_panel":
            if st.session_state.user_type == "admin":
                show_admin_panel()
            else:
                st.error("Access Denied: Admin privileges required.")
                st.session_state.current_page = "student_analytics"
                st.rerun()
        
        else:
            if st.session_state.user_type in ["parent", "student"]:
                st.session_state.current_page = "student_analytics"
                st.rerun()
            else:
                show_dashboard()
        
        show_footer()
    
    except Exception as e:
        st.error("An unexpected error occurred.")
        with st.expander("Technical Details"):
            st.error(f"Error: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

# ==============================================================================
# RUN APPLICATION
# ==============================================================================
if __name__ == "__main__":
    Path("course_data").mkdir(exist_ok=True)
    
    if not os.path.exists("users_enhanced.json"):
        users = load_users()
        save_users(users)
    
    if not os.path.exists("teacher_uploads.json"):
        save_teacher_uploads({})
    
    main()
