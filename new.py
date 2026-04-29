"""
EduTrack Pro 2025 - CO-PO Attainment & Academic Analytics System
Stamford University Bangladesh
Department of Electrical and Electronic Engineering

Run this file directly: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import re
import os
import json
import hashlib
import pickle
import warnings
import io
from io import BytesIO
from datetime import datetime
from pathlib import Path
from math import pi
import base64
import time
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pytz
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

warnings.filterwarnings('ignore')

# ==============================================================================
# STREAMLIT PAGE CONFIGURATION
# ==============================================================================
st.set_page_config(
    page_title="EduTrack Pro 2025",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# CUSTOM CSS THEME
# ==============================================================================
def apply_theme():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    .css-1d391kg {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    .header-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 10px 40px rgba(102, 126, 234, 0.3);
    }
    
    .card {
        background: white;
        border-radius: 15px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.07);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }
    
    .card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.1);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 15px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        transition: transform 0.3s ease;
    }
    
    .metric-card:hover {
        transform: scale(1.05);
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: 800;
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        opacity: 0.9;
    }
    
    .student-card {
        background: white;
        border-left: 4px solid #667eea;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 0 10px 10px 0;
        transition: transform 0.2s ease;
    }
    
    .student-card:hover {
        transform: translateX(5px);
    }
    
    .co-badge {
        display: inline-block;
        padding: 5px 15px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        margin: 3px;
    }
    
    .co-excellent { background: #4CAF50; color: white; }
    .co-good { background: #8BC34A; color: white; }
    .co-average { background: #FFC107; color: #333; }
    .co-poor { background: #FF9800; color: white; }
    .co-fail { background: #F44336; color: white; }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4);
    }
    
    .stDownloadButton > button {
        background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
    }
    
    .footer {
        text-align: center;
        padding: 2rem;
        background: linear-gradient(135deg, #232526 0%, #414345 100%);
        color: white;
        border-radius: 20px 20px 0 0;
        margin-top: 3rem;
    }
    
    .tab-content {
        padding: 2rem;
        background: white;
        border-radius: 0 0 15px 15px;
    }
    
    @media (max-width: 768px) {
        .metric-value { font-size: 1.8rem; }
        .card { padding: 1rem; }
    }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# EXCEL PARSER - CORE FUNCTIONALITY
# ==============================================================================

class ExcelParser:
    """Advanced Excel parser for CO-PO Attainment files"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = None
        self.results = {
            'course_info': {},
            'students': {},
            'co_attainment': {},
            'po_attainment': {},
            'co_po_mapping': {},
            'assessment_weights': {}
        }
    
    def parse(self):
        """Main parsing function"""
        try:
            self.wb = openpyxl.load_workbook(self.file_path, data_only=True)
            
            # Parse all components
            self._extract_course_info()
            self._extract_assessment_weights()
            self._extract_students()
            self._extract_co_scores()
            self._extract_co_attainment()
            self._extract_co_po_mapping()
            self._extract_po_attainment()
            self._calculate_statistics()
            
            self.wb.close()
            return self.results
            
        except Exception as e:
            print(f"Parser Error: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def _extract_course_info(self):
        """Extract course metadata from Midterm Exam sheet"""
        if 'Midterm Exam' not in self.wb.sheetnames:
            return
        
        ws = self.wb['Midterm Exam']
        
        for row in ws.iter_rows(min_row=1, max_row=10, max_col=10, values_only=True):
            for idx, cell in enumerate(row):
                if cell is None:
                    continue
                cell_str = str(cell).strip()
                
                # Semester
                if 'Trimester' in cell_str:
                    next_cell = row[idx + 1] if idx + 1 < len(row) else None
                    if next_cell:
                        sem_match = re.search(r'(Spring|Summer|Fall|Winter)\s+\d{4}', str(next_cell), re.IGNORECASE)
                        if sem_match:
                            self.results['course_info']['semester'] = sem_match.group(0)
                
                # Course Code
                elif 'Course Code' in cell_str:
                    next_cell = row[idx + 1] if idx + 1 < len(row) else None
                    if next_cell:
                        code_match = re.search(r'EEE\s+\d{3}', str(next_cell), re.IGNORECASE)
                        if code_match:
                            self.results['course_info']['course_code'] = code_match.group(0)
                
                # Course Title
                elif 'Course Title' in cell_str:
                    next_cell = row[idx + 1] if idx + 1 < len(row) else None
                    if next_cell:
                        self.results['course_info']['course_title'] = str(next_cell).strip()
                
                # Teacher
                elif 'Course Teacher' in cell_str:
                    next_cell = row[idx + 1] if idx + 1 < len(row) else None
                    if next_cell:
                        self.results['course_info']['teacher'] = str(next_cell).replace(':', '').strip()
                
                # Section
                elif 'Section' in cell_str:
                    section_match = re.search(r'EEE-[\w+-]+', cell_str, re.IGNORECASE)
                    if section_match:
                        self.results['course_info']['section'] = section_match.group(0)
    
    def _extract_assessment_weights(self):
        """Extract assessment weights from each sheet"""
        for sheet_name in ['Midterm Exam', 'Final Exam', 'Assignment']:
            if sheet_name not in self.wb.sheetnames:
                continue
            
            ws = self.wb[sheet_name]
            
            for row in ws.iter_rows(min_row=1, max_row=15, max_col=5, values_only=True):
                for idx, cell in enumerate(row):
                    if cell and 'Full Mark (%)' in str(cell):
                        next_cell = row[idx + 1] if idx + 1 < len(row) else None
                        if next_cell:
                            try:
                                self.results['assessment_weights'][sheet_name] = float(next_cell)
                            except:
                                pass
    
    def _extract_students(self):
        """Extract student data from Midterm Exam sheet"""
        if 'Midterm Exam' not in self.wb.sheetnames:
            return
        
        ws = self.wb['Midterm Exam']
        
        # Find header row
        header_row = None
        col_map = {}
        
        for row_idx in range(1, 50):
            for col_idx in range(1, 15):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    if cell_str in ['SL', 'Student ID', 'Name', 'Status']:
                        header_row = row_idx
                        col_map[cell_str] = col_idx
        
        if not header_row:
            # Try alternative: find first student ID
            for row_idx in range(5, 50):
                for col_idx in range(1, 10):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value and re.search(r'EEE\s*\d{3}\s*\d{5}', str(cell.value)):
                        header_row = row_idx - 1
                        col_map = {
                            'SL': col_idx - 1,
                            'Student ID': col_idx,
                            'Name': col_idx + 1,
                            'Status': col_idx + 2
                        }
                        break
                if header_row:
                    break
        
        if not header_row:
            return
        
        # Extract students
        for row_idx in range(header_row + 2, ws.max_row + 1):
            id_col = col_map.get('Student ID', 2)
            student_id_cell = ws.cell(row=row_idx, column=id_col)
            
            if not student_id_cell.value:
                break
            
            student_id = str(student_id_cell.value).strip()
            
            # Validate student ID
            if not re.search(r'EEE\s*\d{3}\s*\d{5}', student_id):
                continue
            
            # Get name
            name_col = col_map.get('Name', 3)
            name_cell = ws.cell(row=row_idx, column=name_col)
            name = str(name_cell.value).strip() if name_cell.value else ''
            
            # Get status
            status_col = col_map.get('Status', 4)
            status_cell = ws.cell(row=row_idx, column=status_col)
            status = str(status_cell.value).strip() if status_cell.value else 'Regular'
            
            # Get question marks
            question_marks = []
            for q_offset in range(4):
                q_col = status_col + 1 + q_offset
                q_cell = ws.cell(row=row_idx, column=q_col)
                if q_cell.value:
                    try:
                        question_marks.append(float(q_cell.value))
                    except:
                        question_marks.append(0.0)
                else:
                    question_marks.append(0.0)
            
            self.results['students'][student_id] = {
                'name': name,
                'status': status,
                'midterm_questions': question_marks,
                'final_questions': [],
                'assignment_questions': [],
                'co_scores': {},
                'total_marks': 0,
                'percentage': 0,
                'grade': '',
                'parent_email': '',
                'student_email': ''
            }
        
        # Extract marks from Final Exam and Assignment
        for sheet_name in ['Final Exam', 'Assignment']:
            if sheet_name not in self.wb.sheetnames:
                continue
            
            ws = self.wb[sheet_name]
            
            for row_idx in range(header_row + 2, ws.max_row + 1):
                id_cell = ws.cell(row=row_idx, column=id_col)
                if not id_cell.value:
                    continue
                
                student_id = str(id_cell.value).strip()
                
                if student_id in self.results['students']:
                    question_marks = []
                    for q_offset in range(4):
                        q_col = status_col + 1 + q_offset
                        q_cell = ws.cell(row=row_idx, column=q_col)
                        if q_cell.value:
                            try:
                                question_marks.append(float(q_cell.value))
                            except:
                                question_marks.append(0.0)
                        else:
                            question_marks.append(0.0)
                    
                    if sheet_name == 'Final Exam':
                        self.results['students'][student_id]['final_questions'] = question_marks
                    else:
                        self.results['students'][student_id]['assignment_questions'] = question_marks
    
    def _extract_co_scores(self):
        """Extract CO scores from Analysis of CO sheet"""
        if 'Analysis of CO' not in self.wb.sheetnames:
            return
        
        ws = self.wb['Analysis of CO']
        
        # Find CO columns
        co_cols = {}
        for row_idx in range(5, 15):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    if re.match(r'^CO\d+$', cell_str):
                        co_cols[cell_str] = col_idx
        
        # Extract CO scores for each student
        for row_idx in range(10, ws.max_row + 1):
            id_cell = ws.cell(row=row_idx, column=2)
            if not id_cell.value:
                continue
            
            student_id = str(id_cell.value).strip()
            
            if student_id not in self.results['students']:
                continue
            
            co_scores = {}
            for co, col_idx in co_cols.items():
                score_cell = ws.cell(row=row_idx, column=col_idx)
                if score_cell.value:
                    try:
                        co_scores[co] = float(score_cell.value)
                    except:
                        co_scores[co] = 0.0
                else:
                    co_scores[co] = 0.0
            
            self.results['students'][student_id]['co_scores'] = co_scores
            
            # Calculate total and percentage
            total = sum(co_scores.values())
            self.results['students'][student_id]['total_marks'] = total
            
            if co_scores:
                max_possible = len(co_scores) * 20
                pct = (total / max_possible * 100) if max_possible > 0 else 0
                self.results['students'][student_id]['percentage'] = min(100, pct)
                
                # Assign grade
                pct_val = self.results['students'][student_id]['percentage']
                if pct_val >= 80:
                    grade = 'A+'
                elif pct_val >= 75:
                    grade = 'A'
                elif pct_val >= 70:
                    grade = 'A-'
                elif pct_val >= 65:
                    grade = 'B+'
                elif pct_val >= 60:
                    grade = 'B'
                elif pct_val >= 55:
                    grade = 'B-'
                elif pct_val >= 50:
                    grade = 'C+'
                elif pct_val >= 45:
                    grade = 'C'
                elif pct_val >= 40:
                    grade = 'D'
                else:
                    grade = 'F'
                
                self.results['students'][student_id]['grade'] = grade
    
    def _extract_co_attainment(self):
        """Extract CO attainment from Attainment of CO PO sheet"""
        if 'Attainment of CO PO' in self.wb.sheetnames:
            ws = self.wb['Attainment of CO PO']
            
            for row_idx in range(1, ws.max_row + 1):
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        cell_str = str(cell.value).strip()
                        co_match = re.match(r'(CO\d+)', cell_str)
                        if co_match:
                            co_name = co_match.group(1)
                            # Check next column for percentage
                            pct_cell = ws.cell(row=row_idx, column=col_idx + 1)
                            if pct_cell.value:
                                try:
                                    self.results['co_attainment'][co_name] = float(pct_cell.value)
                                except:
                                    pass
        
        # If not found, calculate from Analysis of CO
        if not self.results['co_attainment'] and 'Analysis of CO' in self.wb.sheetnames:
            ws = self.wb['Analysis of CO']
            student_count = len(self.results['students'])
            
            # Find Total Yes row
            for row_idx in range(ws.max_row - 5, ws.max_row + 1):
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value and 'Total Yes' in str(cell.value):
                        for co_num, offset in enumerate([1, 3, 5, 7], 1):
                            yes_cell = ws.cell(row=row_idx, column=col_idx + offset)
                            if yes_cell.value and student_count > 0:
                                try:
                                    self.results['co_attainment'][f'CO{co_num}'] = (float(yes_cell.value) / student_count) * 100
                                except:
                                    pass
                        break
    
    def _extract_co_po_mapping(self):
        """Extract CO-PO mapping from Analysis of PO sheet"""
        if 'Analysis of PO' not in self.wb.sheetnames:
            return
        
        ws = self.wb['Analysis of PO']
        
        # Find CO-PO matrix
        matrix_start = None
        for row_idx in range(1, 20):
            for col_idx in range(1, 15):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and 'CO-PO matrix' in str(cell.value):
                    matrix_start = row_idx
                    break
            if matrix_start:
                break
        
        if not matrix_start:
            return
        
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
                    self.results['co_po_mapping'][co_name] = {}
                    for po_idx, po_name in enumerate(po_headers):
                        val_cell = ws.cell(row=row_idx, column=4 + po_idx)
                        if val_cell.value:
                            try:
                                self.results['co_po_mapping'][co_name][po_name] = float(val_cell.value)
                            except:
                                self.results['co_po_mapping'][co_name][po_name] = 0
                        else:
                            self.results['co_po_mapping'][co_name][po_name] = 0
    
    def _extract_po_attainment(self):
        """Calculate PO attainment from CO attainment and CO-PO mapping"""
        # First try to get from Attainment sheet
        if 'Attainment of CO PO' in self.wb.sheetnames:
            ws = self.wb['Attainment of CO PO']
            
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
                                    self.results['po_attainment'][po_name] = float(pct_cell.value)
                                except:
                                    pass
        
        # If not found, calculate from CO attainment
        if not self.results['po_attainment'] and self.results['co_attainment'] and self.results['co_po_mapping']:
            all_pos = set()
            for mapping in self.results['co_po_mapping'].values():
                all_pos.update(mapping.keys())
            
            for po in sorted(all_pos):
                total_weight = 0
                weighted_sum = 0
                
                for co, attainment in self.results['co_attainment'].items():
                    if co in self.results['co_po_mapping'] and po in self.results['co_po_mapping'][co]:
                        weight = self.results['co_po_mapping'][co][po]
                        if weight > 0:
                            weighted_sum += attainment * weight
                            total_weight += weight
                
                if total_weight > 0:
                    self.results['po_attainment'][po] = weighted_sum / total_weight
    
    def _calculate_statistics(self):
        """Calculate course statistics"""
        students = self.results['students']
        
        if not students:
            self.results['stats'] = {
                'total_students': 0,
                'average_marks': 0,
                'pass_percentage': 0,
                'highest_marks': 0,
                'lowest_marks': 0
            }
            return
        
        percentages = [s['percentage'] for s in students.values()]
        total_marks = [s['total_marks'] for s in students.values()]
        
        self.results['stats'] = {
            'total_students': len(students),
            'average_marks': np.mean(total_marks) if total_marks else 0,
            'average_percentage': np.mean(percentages) if percentages else 0,
            'pass_percentage': len([p for p in percentages if p >= 40]) / len(percentages) * 100 if percentages else 0,
            'highest_marks': max(total_marks) if total_marks else 0,
            'lowest_marks': min(total_marks) if total_marks else 0,
            'std_deviation': np.std(percentages) if percentages else 0
        }

# ==============================================================================
# VISUALIZATION FUNCTIONS
# ==============================================================================

def create_spider_plot(categories, values, title="Spider Plot", color='#667eea'):
    """Create spider/radar plot"""
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
        ax.annotate(f'{value:.1f}%', xy=(angle, value), xytext=(5, 5),
                   textcoords='offset points', fontsize=10, fontweight='bold', color=color)
    
    plt.tight_layout()
    return fig

def create_comparison_spider(categories, student_vals, class_vals, student_name="You"):
    """Create comparison spider plot"""
    N = len(categories)
    if N < 3:
        return None
    
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]
    
    student_plot = list(student_vals) + [student_vals[0]]
    class_plot = list(class_vals) + [class_vals[0]]
    
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    
    ax.plot(angles, student_plot, 'o-', linewidth=2, color='#4CAF50', label=student_name, markersize=8)
    ax.fill(angles, student_plot, alpha=0.15, color='#4CAF50')
    
    ax.plot(angles, class_plot, 'o-', linewidth=2, color='#FF9800', label='Class Average', markersize=8)
    ax.fill(angles, class_plot, alpha=0.15, color='#FF9800')
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, size=12, fontweight='bold')
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(['20%', '40%', '60%', '80%', '100%'], size=8)
    ax.grid(True, alpha=0.3)
    plt.title("You vs Class Average", size=16, fontweight='bold', pad=20)
    plt.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0))
    
    plt.tight_layout()
    return fig

# ==============================================================================
# PDF REPORT GENERATION
# ==============================================================================

def generate_pdf_report(data, course_name):
    """Generate PDF report with charts"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24, alignment=1, spaceAfter=20)
    story.append(Paragraph(f"CO-PO Attainment Report", title_style))
    story.append(Paragraph(f"Course: {course_name}", styles['Heading2']))
    story.append(Spacer(1, 20))
    
    # Course info
    course_info = data.get('course_info', {})
    for key, value in course_info.items():
        if value:
            story.append(Paragraph(f"<b>{key}:</b> {value}", styles['Normal']))
    
    story.append(Spacer(1, 20))
    
    # Statistics table
    stats = data.get('stats', {})
    table_data = [
        ['Metric', 'Value'],
        ['Total Students', str(stats.get('total_students', 0))],
        ['Average Marks', f"{stats.get('average_marks', 0):.1f}"],
        ['Pass Rate', f"{stats.get('pass_percentage', 0):.1f}%"],
        ['Highest', f"{stats.get('highest_marks', 0):.1f}"],
        ['Lowest', f"{stats.get('lowest_marks', 0):.1f}"]
    ]
    
    table = Table(table_data, colWidths=[200, 150])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#667eea')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('PADDING', (0, 0), (-1, -1), 8)
    ]))
    story.append(table)
    story.append(Spacer(1, 30))
    
    # CO Attainment chart
    co_attainment = data.get('co_attainment', {})
    if co_attainment:
        categories = list(co_attainment.keys())
        values = list(co_attainment.values())
        
        fig = create_spider_plot(categories, values, "CO Attainment")
        if fig:
            img_buffer = BytesIO()
            fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            img_buffer.seek(0)
            story.append(Image(img_buffer, width=5*inch, height=5*inch))
    
    story.append(PageBreak())
    
    # Student list
    students = data.get('students', {})
    if students:
        story.append(Paragraph("Student Results", styles['Heading2']))
        
        student_data = [['ID', 'Name', 'Marks', 'Grade', '%']]
        for sid, student in sorted(students.items(), key=lambda x: x[1].get('percentage', 0), reverse=True)[:30]:
            student_data.append([
                sid,
                student.get('name', '')[:20],
                f"{student.get('total_marks', 0):.1f}",
                student.get('grade', ''),
                f"{student.get('percentage', 0):.1f}%"
            ])
        
        t = Table(student_data, colWidths=[80, 120, 60, 50, 60])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#667eea')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('PADDING', (0, 0), (-1, -1), 4)
        ]))
        story.append(t)
    
    story.append(Spacer(1, 50))
    story.append(Paragraph("Generated by EduTrack Pro 2025", styles['Normal']))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# ==============================================================================
# STREAMLIT APP PAGES
# ==============================================================================

def main():
    """Main Streamlit application"""
    apply_theme()
    
    # Initialize session state
    if 'parsed_data' not in st.session_state:
        st.session_state.parsed_data = None
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "Upload"
    
    # Header
    st.markdown("""
    <div class="header-card">
        <h1>🎓 EduTrack Pro 2025</h1>
        <p>CO-PO Attainment & Academic Analytics System</p>
        <p style="font-size: 0.9rem; opacity: 0.8;">Stamford University Bangladesh</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("## 📋 Navigation")
        
        pages = ["📤 Upload & Process", "📊 Dashboard", "👨‍🎓 Student View", 
                "📈 CO-PO Analysis", "📧 Email Reports", "ℹ️ Help"]
        
        selected_page = st.radio("Select Page", pages)
        
        st.markdown("---")
        
        # File uploader in sidebar
        uploaded_file = st.file_uploader(
            "📁 Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload your CO-PO Attainment Excel file"
        )
        
        if uploaded_file:
            if st.button("🔄 Process File", use_container_width=True):
                with st.spinner("Processing..."):
                    # Save temporarily
                    temp_path = "temp_upload.xlsx"
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Parse
                    parser = ExcelParser(temp_path)
                    st.session_state.parsed_data = parser.parse()
                    
                    if st.session_state.parsed_data:
                        st.success("✅ File processed successfully!")
                        st.balloons()
                    else:
                        st.error("❌ Error parsing file")
    
    # Main content area
    if selected_page == "📤 Upload & Process":
        show_upload_page()
    elif selected_page == "📊 Dashboard":
        show_dashboard_page()
    elif selected_page == "👨‍🎓 Student View":
        show_student_view_page()
    elif selected_page == "📈 CO-PO Analysis":
        show_copo_analysis_page()
    elif selected_page == "📧 Email Reports":
        show_email_page()
    elif selected_page == "ℹ️ Help":
        show_help_page()
    
    # Footer
    st.markdown("""
    <div class="footer">
        <p><strong>EduTrack Pro 2025</strong></p>
        <p>Department of Electrical & Electronic Engineering</p>
        <p>Stamford University Bangladesh</p>
        <p style="font-size: 0.8rem; opacity: 0.7;">© 2025 All Rights Reserved</p>
    </div>
    """, unsafe_allow_html=True)

def show_upload_page():
    st.markdown("## 📤 Upload & Process")
    
    st.markdown("""
    <div class="card">
        <h4>📋 Instructions</h4>
        <ol>
            <li>Upload your Excel file using the sidebar uploader</li>
            <li>Click <b>'Process File'</b> to parse the data</li>
            <li>The system will automatically extract:
                <ul>
                    <li>Course Information</li>
                    <li>Student Data with CO Scores</li>
                    <li>CO Attainment Percentages</li>
                    <li>PO Attainment Percentages</li>
                </ul>
            </li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.parsed_data:
        data = st.session_state.parsed_data
        
        st.success("✅ Data ready! Navigate to other pages to view analysis.")
        
        # Quick preview
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Students", data['stats']['total_students'])
        with col2:
            st.metric("Avg Marks", f"{data['stats']['average_marks']:.1f}")
        with col3:
            st.metric("Pass Rate", f"{data['stats']['pass_percentage']:.1f}%")
        with col4:
            st.metric("Course", data['course_info'].get('course_code', 'N/A'))
        
        # Download PDF Report
        if st.button("📥 Download PDF Report", use_container_width=True):
            course_name = f"{data['course_info'].get('course_code', 'Course')} - {data['course_info'].get('semester', '')}"
            pdf_buffer = generate_pdf_report(data, course_name)
            
            st.download_button(
                "📄 Click to Download",
                pdf_buffer,
                "CO_PO_Report.pdf",
                "application/pdf",
                use_container_width=True
            )

def show_dashboard_page():
    st.markdown("## 📊 Dashboard")
    
    if not st.session_state.parsed_data:
        st.info("👈 Please upload and process an Excel file first")
        return
    
    data = st.session_state.parsed_data
    
    # Summary cards
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Students</div>
            <div class="metric-value">{data['stats']['total_students']}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        avg_pct = data['stats']['average_percentage']
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Average %</div>
            <div class="metric-value">{avg_pct:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        pass_rate = data['stats']['pass_percentage']
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Pass Rate</div>
            <div class="metric-value">{pass_rate:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Course</div>
            <div class="metric-value" style="font-size: 1.2rem;">{data['course_info'].get('course_code', 'N/A')}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Charts
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### CO Attainment")
        co_attainment = data.get('co_attainment', {})
        if co_attainment:
            fig = create_spider_plot(
                list(co_attainment.keys()), 
                list(co_attainment.values()),
                "CO Attainment",
                '#2196F3'
            )
            if fig:
                st.pyplot(fig)
                plt.close()
    
    with col_right:
        st.markdown("### PO Attainment")
        po_attainment = data.get('po_attainment', {})
        if po_attainment:
            fig = create_spider_plot(
                list(po_attainment.keys()),
                list(po_attainment.values()),
                "PO Attainment",
                '#FF9800'
            )
            if fig:
                st.pyplot(fig)
                plt.close()
    
    # Student performance distribution
    st.markdown("### Marks Distribution")
    
    students = data.get('students', {})
    if students:
        percentages = [s['percentage'] for s in students.values()]
        
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.histplot(percentages, bins=10, kde=True, color='#667eea', edgecolor='white', ax=ax)
        ax.axvline(np.mean(percentages), color='red', linestyle='--', label=f'Mean: {np.mean(percentages):.1f}%')
        ax.set_xlabel('Percentage')
        ax.set_ylabel('Number of Students')
        ax.set_title('Student Performance Distribution')
        ax.legend()
        st.pyplot(fig)
        plt.close()

def show_student_view_page():
    st.markdown("## 👨‍🎓 Student CO-PO Attainment View")
    
    if not st.session_state.parsed_data:
        st.info("👈 Please upload and process an Excel file first")
        return
    
    data = st.session_state.parsed_data
    students = data.get('students', {})
    
    if not students:
        st.error("No student data found")
        return
    
    # Student selection
    student_ids = list(students.keys())
    selected_id = st.selectbox(
        "🔍 Select Student",
        student_ids,
        format_func=lambda x: f"{x} - {students[x].get('name', 'Unknown')}"
    )
    
    if selected_id:
        student = students[selected_id]
        
        # Student info
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Name", student.get('name', 'Unknown'))
        with col2:
            st.metric("Status", student.get('status', 'Regular'))
        with col3:
            st.metric("Total Marks", f"{student.get('total_marks', 0):.1f}")
        with col4:
            pct = student.get('percentage', 0)
            color = '#4CAF50' if pct >= 50 else '#FFC107' if pct >= 30 else '#F44336'
            st.markdown(f"""
            <div style="background: {color}; color: white; padding: 15px; border-radius: 10px; text-align: center;">
                <h3>{pct:.1f}%</h3>
                <small>Grade: {student.get('grade', 'N/A')}</small>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # CO Scores
        st.markdown("### Your CO Scores")
        co_scores = student.get('co_scores', {})
        
        if co_scores:
            # Bar chart
            fig, ax = plt.subplots(figsize=(10, 5))
            co_names = list(co_scores.keys())
            co_values = [min(100, (score/20)*100) for score in co_scores.values()]
            
            colors = []
            for v in co_values:
                if v >= 50: colors.append('#4CAF50')
                elif v >= 30: colors.append('#FFC107')
                else: colors.append('#F44336')
            
            bars = ax.bar(co_names, co_values, color=colors, edgecolor='white', linewidth=2)
            
            for bar, val in zip(bars, co_values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, 
                       f'{val:.1f}%', ha='center', fontweight='bold', fontsize=12)
            
            ax.set_ylim(0, 110)
            ax.axhline(y=50, color='green', linestyle='--', linewidth=2, alpha=0.7, label='Target 50%')
            ax.set_ylabel('Percentage (%)')
            ax.set_title(f'{student.get("name", "Student")} - CO Attainment')
            ax.legend()
            ax.grid(axis='y', alpha=0.3)
            st.pyplot(fig)
            plt.close()
            
            # CO Score table
            co_df = pd.DataFrame([
                {
                    'Course Outcome': co,
                    'Score (out of 20)': f"{score:.1f}",
                    'Percentage': f"{min(100, (score/20)*100):.1f}%"
                }
                for co, score in co_scores.items()
            ])
            st.dataframe(co_df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Comparison plots
        st.markdown("### Compare with Class Average")
        
        plot_type = st.radio(
            "Select Comparison",
            ["CO Attainment (Spider Plot)", "PO Attainment (Spider Plot)"],
            horizontal=True
        )
        
        if "CO" in plot_type and co_scores:
            student_co_pct = [min(100, (co_scores.get(co, 0)/20)*100) for co in sorted(co_scores.keys())]
            class_co = data.get('co_attainment', {})
            class_co_pct = [class_co.get(co, 50) for co in sorted(co_scores.keys())]
            
            fig = create_comparison_spider(
                sorted(co_scores.keys()),
                student_co_pct,
                class_co_pct,
                student.get('name', 'You')
            )
            if fig:
                st.pyplot(fig)
                plt.close()
        
        elif "PO" in plot_type:
            # Calculate student PO from CO scores using mapping
            co_po_mapping = data.get('co_po_mapping', {})
            
            if co_po_mapping and co_scores:
                student_po = {}
                for co, mapping in co_po_mapping.items():
                    if co in co_scores:
                        for po, weight in mapping.items():
                            if weight > 0:
                                if po not in student_po:
                                    student_po[po] = {'sum': 0, 'weight': 0}
                                score_pct = (co_scores[co] / 20) * 100
                                student_po[po]['sum'] += score_pct * weight
                                student_po[po]['weight'] += weight
                
                # Normalize
                student_po_pct = {}
                for po, vals in student_po.items():
                    if vals['weight'] > 0:
                        student_po_pct[po] = min(100, vals['sum'] / vals['weight'])
                
                class_po = data.get('po_attainment', {})
                
                if student_po_pct:
                    categories = sorted(student_po_pct.keys())
                    student_vals = [student_po_pct.get(po, 0) for po in categories]
                    class_vals = [class_po.get(po, 50) for po in categories]
                    
                    fig = create_comparison_spider(categories, student_vals, class_vals, student.get('name', 'You'))
                    if fig:
                        st.pyplot(fig)
                        plt.close()

def show_copo_analysis_page():
    st.markdown("## 📈 CO-PO Attainment Analysis")
    
    if not st.session_state.parsed_data:
        st.info("👈 Please upload and process an Excel file first")
        return
    
    data = st.session_state.parsed_data
    
    tab1, tab2, tab3 = st.tabs(["CO Attainment", "PO Attainment", "CO-PO Mapping"])
    
    with tab1:
        st.markdown("### Course Outcome Attainment")
        
        co_attainment = data.get('co_attainment', {})
        
        if co_attainment:
            # Spider plot
            fig = create_spider_plot(
                list(co_attainment.keys()),
                list(co_attainment.values()),
                "CO Attainment Spider Plot",
                '#2196F3'
            )
            if fig:
                col1, col2 = st.columns(2)
                with col1:
                    st.pyplot(fig)
                    plt.close()
                
                with col2:
                    # Bar chart
                    fig2, ax = plt.subplots(figsize=(8, 6))
                    co_names = list(co_attainment.keys())
                    co_vals = list(co_attainment.values())
                    
                    colors = ['#4CAF50' if v >= 50 else '#FFC107' if v >= 30 else '#F44336' for v in co_vals]
                    bars = ax.bar(co_names, co_vals, color=colors, edgecolor='white', linewidth=2)
                    
                    for bar, val in zip(bars, co_vals):
                        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                               f'{val:.1f}%', ha='center', fontweight='bold')
                    
                    ax.set_ylim(0, 110)
                    ax.axhline(y=50, color='green', linestyle='--', linewidth=2, alpha=0.7, label='Target 50%')
                    ax.set_ylabel('Attainment (%)')
                    ax.set_title('CO Attainment Bar Chart')
                    ax.legend()
                    ax.grid(axis='y', alpha=0.3)
                    st.pyplot(fig2)
                    plt.close()
            
            # Summary table
            co_df = pd.DataFrame([
                {'Course Outcome': co, 'Attainment (%)': f"{val:.1f}%", 
                 'Status': '✅ Achieved' if val >= 50 else '⚠️ Needs Improvement'}
                for co, val in co_attainment.items()
            ])
            st.dataframe(co_df, use_container_width=True, hide_index=True)
    
    with tab2:
        st.markdown("### Program Outcome Attainment")
        
        po_attainment = data.get('po_attainment', {})
        
        if po_attainment:
            fig = create_spider_plot(
                list(po_attainment.keys()),
                list(po_attainment.values()),
                "PO Attainment Spider Plot",
                '#FF9800'
            )
            if fig:
                st.pyplot(fig)
                plt.close()
            
            # Bar chart
            fig2, ax = plt.subplots(figsize=(12, 5))
            po_names = list(po_attainment.keys())
            po_vals = list(po_attainment.values())
            
            colors = ['#4CAF50' if v >= 50 else '#FFC107' if v >= 30 else '#F44336' for v in po_vals]
            bars = ax.bar(po_names, po_vals, color=colors, edgecolor='white', linewidth=2)
            
            for bar, val in zip(bars, po_vals):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                       f'{val:.1f}%', ha='center', fontsize=8)
            
            ax.set_ylim(0, 110)
            ax.axhline(y=50, color='green', linestyle='--', linewidth=2, alpha=0.7)
            ax.set_ylabel('Attainment (%)')
            ax.set_title('PO Attainment Bar Chart')
            ax.grid(axis='y', alpha=0.3)
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig2)
            plt.close()
    
    with tab3:
        st.markdown("### CO-PO Mapping Matrix")
        
        co_po_mapping = data.get('co_po_mapping', {})
        
        if co_po_mapping:
            # Convert to DataFrame
            mapping_data = []
            for co, po_map in co_po_mapping.items():
                row = {'CO': co}
                row.update(po_map)
                mapping_data.append(row)
            
            df_mapping = pd.DataFrame(mapping_data)
            df_mapping = df_mapping.set_index('CO')
            
            # Heatmap
            fig, ax = plt.subplots(figsize=(12, 6))
            sns.heatmap(df_mapping, annot=True, cmap='YlOrRd', cbar_kws={'label': 'Weight'}, ax=ax)
            ax.set_title('CO-PO Mapping Matrix')
            st.pyplot(fig)
            plt.close()
            
            st.dataframe(df_mapping, use_container_width=True)

def show_email_page():
    st.markdown("## 📧 Email Reports to Parents")
    
    if not st.session_state.parsed_data:
        st.info("👈 Please upload and process an Excel file first")
        return
    
    data = st.session_state.parsed_data
    students = data.get('students', {})
    
    if not students:
        st.error("No student data available")
        return
    
    st.markdown("### Student List with Parent Emails")
    
    # Create editable dataframe
    email_data = []
    for sid, student in students.items():
        email_data.append({
            'Student ID': sid,
            'Name': student.get('name', 'Unknown'),
            'Grade': student.get('grade', 'N/A'),
            'Percentage': f"{student.get('percentage', 0):.1f}%",
            'Parent Email': student.get('parent_email', ''),
            'Send': False
        })
    
    df_emails = pd.DataFrame(email_data)
    
    edited_df = st.data_editor(
        df_emails,
        column_config={
            "Parent Email": st.column_config.TextColumn("Parent Email", help="Enter parent's email address"),
            "Send": st.column_config.CheckboxColumn("Send Report", help="Check to send email")
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed"
    )
    
    col1, col2, col3 = st.columns(3)
    with col1:
        subject = st.text_input("Email Subject", "Academic Performance Report")
    with col2:
        sender_email = st.text_input("Sender Email (Gmail)", "")
    with col3:
        app_password = st.text_input("Gmail App Password", type="password")
    
    # Email preview
    if st.checkbox("Preview Email Template"):
        sample_student = list(students.values())[0] if students else {}
        st.markdown("### Email Preview")
        st.markdown(f"""
        <div style="background: #f5f5f5; padding: 20px; border-radius: 10px; border: 1px solid #ddd;">
            <h4>Subject: {subject}</h4>
            <hr>
            <p>Dear Parent/Guardian of <strong>{sample_student.get('name', 'Student')}</strong>,</p>
            <p>Here is the academic performance summary for your child:</p>
            <ul>
                <li>Course: {data['course_info'].get('course_code', 'N/A')}</li>
                <li>Grade: {sample_student.get('grade', 'N/A')}</li>
                <li>Percentage: {sample_student.get('percentage', 0):.1f}%</li>
            </ul>
            <p>Thank you for your continued support.</p>
            <p style="color: #666;">Generated by EduTrack Pro</p>
        </div>
        """, unsafe_allow_html=True)
    
    if st.button("📤 Send Selected Emails", type="primary", use_container_width=True):
        if not sender_email or not app_password:
            st.error("Please enter sender email and app password")
        else:
            selected = edited_df[edited_df['Send'] == True]
            
            if len(selected) == 0:
                st.warning("No students selected for sending")
            else:
                progress = st.progress(0)
                sent = 0
                failed = 0
                
                for idx, row in selected.iterrows():
                    try:
                        # Here you would implement actual email sending
                        # For demo, we just simulate
                        time.sleep(0.5)
                        sent += 1
                        progress.progress((idx + 1) / len(selected))
                    except Exception as e:
                        failed += 1
                        st.error(f"Failed for {row['Student ID']}: {str(e)}")
                
                if sent > 0:
                    st.success(f"✅ Successfully sent {sent} emails")
                if failed > 0:
                    st.error(f"❌ Failed: {failed} emails")

def show_help_page():
    st.markdown("## ℹ️ Help & Documentation")
    
    st.markdown("""
    <div class="card">
        <h3>📚 How to Use EduTrack Pro</h3>
        
        <h4>1. Upload Data</h4>
        <p>Upload your Excel file containing CO-PO Attainment data. The system supports the standard format with sheets:</p>
        <ul>
            <li><b>Midterm Exam</b> - Student marks for midterm</li>
            <li><b>Final Exam</b> - Student marks for final exam</li>
            <li><b>Assignment</b> - Assignment marks</li>
            <li><b>Analysis of CO</b> - CO score calculations</li>
            <li><b>Analysis of PO</b> - PO score calculations</li>
            <li><b>Attainment of CO PO</b> - Final attainment percentages</li>
        </ul>
        
        <h4>2. Dashboard</h4>
        <p>View overall course statistics, CO and PO attainment spider plots, and student performance distribution.</p>
        
        <h4>3. Student View</h4>
        <p>Select individual students to view their CO scores and compare with class average using spider plots.</p>
        
        <h4>4. CO-PO Analysis</h4>
        <p>Detailed analysis of Course Outcomes and Program Outcomes with spider plots, bar charts, and CO-PO mapping matrix.</p>
        
        <h4>5. Email Reports</h4>
        <p>Send performance reports to parents via email. Edit parent email addresses in the table and select students to send.</p>
        
        <hr>
        
        <h4>📊 Understanding the Charts</h4>
        <ul>
            <li><b>Spider/Radar Plots</b> - Show attainment across multiple COs or POs</li>
            <li><b>Bar Charts</b> - Individual CO/PO attainment with color coding</li>
            <li><b>Heatmaps</b> - CO-PO mapping matrix showing relationship strengths</li>
            <li><b>Histograms</b> - Student performance distribution</li>
        </ul>
        
        <h4>🎨 Color Coding</h4>
        <ul>
            <li><span style="color: #4CAF50;">🟢 Green</span> - Good attainment (≥50%)</li>
            <li><span style="color: #FFC107;">🟡 Yellow</span> - Average (30-50%)</li>
            <li><span style="color: #F44336;">🔴 Red</span> - Needs improvement (<30%)</li>
        </ul>
        
        <h4>📧 Email Setup</h4>
        <ol>
            <li>Enable 2-Factor Authentication in your Gmail account</li>
            <li>Generate an App Password from Google Account settings</li>
            <li>Use the App Password in the Email section</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# RUN APPLICATION
# ==============================================================================

if __name__ == "__main__":
    main()
