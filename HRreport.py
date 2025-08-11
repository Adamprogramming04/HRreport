from flask import Flask, render_template_string, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import sys
import subprocess
from datetime import datetime
import json
import socket

# Auto-install required packages
def install_package(package):
    try:
        __import__(package.split('==')[0] if '==' in package else package)
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

# Install required packages
packages = [
    'flask',
    'pandas',
    'openpyxl',
    'reportlab',
    'matplotlib',
    'seaborn',
    'Pillow',
    'numpy'
]

for package in packages:
    install_package(package)

try:
    import pandas as pd
    import numpy as np
    import matplotlib
    matplotlib.use('Agg')  # Use non-GUI backend
    import matplotlib.pyplot as plt
    import seaborn as sns
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from PIL import Image as PILImage
    import io
    import base64
except ImportError as e:
    print(f"Error importing required modules: {e}")
    sys.exit(1)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'reports'
app.config['TEMP_FOLDER'] = 'temp'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max file size

# Get the script directory for reliable paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(SCRIPT_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(SCRIPT_DIR, 'reports')
app.config['TEMP_FOLDER'] = os.path.join(SCRIPT_DIR, 'temp')

# Create necessary directories
for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], app.config['TEMP_FOLDER']]:
    os.makedirs(folder, exist_ok=True)

# Global variables
uploaded_files = []
report_data = {}

def get_local_ip():
    """Get the local IP address of this machine"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"

def detect_column_type(series, column_name):
    """Enhanced column type detection"""
    # Clean the series
    clean_series = series.dropna()
    if len(clean_series) == 0:
        return 'empty'
    
    column_name_lower = str(column_name).lower().strip()
    
    # Check for date columns first
    if any(keyword in column_name_lower for keyword in ['date', 'time', 'created', 'updated', 'birth', 'hire', 'start', 'end']):
        try:
            pd.to_datetime(clean_series.head(10), errors='raise')
            return 'date'
        except:
            pass
    
    # Check if numeric
    if pd.api.types.is_numeric_dtype(clean_series):
        # Check if it's likely an ID or code (all integers, high uniqueness)
        if pd.api.types.is_integer_dtype(clean_series):
            unique_ratio = len(clean_series.unique()) / len(clean_series)
            if unique_ratio > 0.8 and any(keyword in column_name_lower for keyword in ['id', 'code', 'number']):
                return 'identifier'
        return 'numeric'
    
    # Check for categorical data
    unique_values = len(clean_series.unique())
    total_values = len(clean_series)
    
    if unique_values <= 20 or (unique_values / total_values) < 0.5:
        return 'categorical'
    
    return 'text'

def analyze_excel_data(dataframes):
    """Enhanced data analysis with better categorization"""
    analysis = {
        'summary': {},
        'data_overview': [],
        'charts_data': {
            'numeric': {},
            'categorical': {},
            'dates': {},  # Fixed: was missing 's'
            'identifiers': {}
        },
        'detailed_analysis': {},
        'insights': []
    }
    
    total_rows = 0
    total_files = len(dataframes)
    all_columns = {}
    
    for filename, sheets in dataframes.items():
        file_summary = {
            'filename': filename,
            'sheets': {},
            'total_rows': 0,
            'total_columns': 0
        }
        
        file_analysis = {
            'numeric_stats': {},
            'categorical_counts': {},
            'data_quality': {}
        }
        
        for sheet_name, df in sheets.items():
            if df.empty:
                continue
            
            rows, cols = df.shape
            total_rows += rows
            file_summary['sheets'][sheet_name] = {
                'rows': rows,
                'columns': cols,
                'column_names': list(df.columns)
            }
            file_summary['total_rows'] += rows
            file_summary['total_columns'] = max(file_summary['total_columns'], cols)
            
            # Analyze each column
            for col in df.columns:
                col_clean = str(col).strip()
                col_type = detect_column_type(df[col], col)
                
                # Initialize column data if not exists
                if col_clean not in all_columns:
                    all_columns[col_clean] = {
                        'type': col_type,
                        'files': [],
                        'data': []
                    }
                
                all_columns[col_clean]['files'].append(f"{filename}:{sheet_name}")
                
                if col_type == 'numeric':
                    numeric_data = pd.to_numeric(df[col], errors='coerce').dropna()
                    if len(numeric_data) > 0:
                        analysis['charts_data']['numeric'][col_clean] = analysis['charts_data']['numeric'].get(col_clean, [])
                        analysis['charts_data']['numeric'][col_clean].extend(numeric_data.tolist())
                        
                        # Store statistics
                        file_analysis['numeric_stats'][col_clean] = {
                            'mean': float(numeric_data.mean()),
                            'median': float(numeric_data.median()),
                            'std': float(numeric_data.std()),
                            'min': float(numeric_data.min()),
                            'max': float(numeric_data.max()),
                            'count': len(numeric_data)
                        }
                
                elif col_type == 'categorical':
                    value_counts = df[col].value_counts()
                    if col_clean not in analysis['charts_data']['categorical']:
                        analysis['charts_data']['categorical'][col_clean] = {}
                    
                    for value, count in value_counts.items():
                        if pd.notna(value):
                            key = str(value).strip()
                            analysis['charts_data']['categorical'][col_clean][key] = \
                                analysis['charts_data']['categorical'][col_clean].get(key, 0) + count
                
                elif col_type == 'date':
                    try:
                        date_data = pd.to_datetime(df[col], errors='coerce').dropna()
                        if len(date_data) > 0:
                            analysis['charts_data']['dates'][col_clean] = analysis['charts_data']['dates'].get(col_clean, [])
                            analysis['charts_data']['dates'][col_clean].extend(date_data.tolist())
                    except:
                        pass
                
                # Data quality analysis
                null_count = df[col].isnull().sum()
                file_analysis['data_quality'][col_clean] = {
                    'null_count': int(null_count),
                    'null_percentage': float(null_count / len(df) * 100),
                    'unique_count': int(df[col].nunique())
                }
        
        analysis['detailed_analysis'][filename] = file_analysis
        analysis['data_overview'].append(file_summary)
    
    # Generate enhanced summary
    analysis['summary'] = {
        'total_files': total_files,
        'total_rows': total_rows,
        'total_columns': len(all_columns),
        'numeric_columns': len(analysis['charts_data']['numeric']),
        'categorical_columns': len(analysis['charts_data']['categorical']),
        'date_columns': len(analysis['charts_data']['dates'])
    }
    
    # Generate enhanced insights
    insights = []
    insights.append(f"📊 Analyzed {total_files} Excel files containing {total_rows:,} total data records")
    insights.append(f"📈 Found {len(analysis['charts_data']['numeric'])} numeric columns for quantitative analysis")
    insights.append(f"📋 Identified {len(analysis['charts_data']['categorical'])} categorical columns for distribution analysis")
    
    if analysis['charts_data']['dates']:
        insights.append(f"📅 Detected {len(analysis['charts_data']['dates'])} date columns for temporal analysis")
    
    # Numeric insights
    for col, values in analysis['charts_data']['numeric'].items():
        if values and len(values) > 1:
            mean_val = np.mean(values)
            std_val = np.std(values)
            insights.append(f"📊 {col}: Mean = {mean_val:,.2f}, Std Dev = {std_val:,.2f}, Range = {min(values):,.2f} - {max(values):,.2f}")
    
    # Categorical insights
    for col, categories in analysis['charts_data']['categorical'].items():
        if categories:
            total_count = sum(categories.values())
            top_category = max(categories, key=categories.get)
            percentage = (categories[top_category] / total_count) * 100
            insights.append(f"📋 {col}: Most frequent value is '{top_category}' ({percentage:.1f}% of data, {categories[top_category]:,} records)")
    
    analysis['insights'] = insights
    return analysis

def create_professional_chart(chart_type, data, title, filename, column_name="", subtitle=""):
    """Create professional charts with blue theme"""
    plt.style.use('default')
    
    # Define professional blue color palette - only blues
    blue_colors = ['#1E88E5', '#2196F3', '#42A5F5', '#64B5F6', '#90CAF9', '#BBDEFB', 
                   '#0D47A1', '#1565C0', '#1976D2', '#1E88E5', '#2196F3', '#42A5F5']
    
    # Create figure with specific styling
    fig, ax = plt.subplots(figsize=(14, 10))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F5F9FF')
    
    # Grid styling
    ax.grid(True, alpha=0.3, color='#CCCCCC', linestyle='-', linewidth=0.5)
    ax.set_axisbelow(True)
    
    if chart_type == 'categorical_pie':
        # Limit to top 12 categories for better readability
        sorted_data = dict(sorted(data.items(), key=lambda x: x[1], reverse=True)[:12])
        
        # Create pie chart
        wedges, texts, autotexts = ax.pie(
            sorted_data.values(), 
            labels=sorted_data.keys(), 
            autopct='%1.1f%%',
            colors=blue_colors[:len(sorted_data)],
            startangle=90,
            explode=[0.02] * len(sorted_data),  # Slight separation
            shadow=True,
            textprops={'fontsize': 11, 'fontweight': 'bold'}
        )
        
        # Style the text
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(10)
        
        for text in texts:
            text.set_color('#333333')
            text.set_fontweight('bold')
            text.set_fontsize(10)
        
        ax.axis('equal')
    
    elif chart_type == 'categorical_bar':
        # Limit to top 20 categories
        sorted_data = dict(sorted(data.items(), key=lambda x: x[1], reverse=True)[:20])
        
        # Create bar chart
        bars = ax.bar(
            range(len(sorted_data)), 
            sorted_data.values(),
            color=blue_colors[0],
            alpha=0.8,
            edgecolor='white',
            linewidth=1
        )
        
        # Customize axes
        ax.set_xticks(range(len(sorted_data)))
        ax.set_xticklabels(sorted_data.keys(), rotation=45, ha='right', fontsize=10, fontweight='bold')
        ax.set_ylabel('Count', fontsize=12, fontweight='bold', color='#333333')
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + max(sorted_data.values()) * 0.01,
                   f'{int(height):,}', ha='center', va='bottom', 
                   fontweight='bold', fontsize=9, color='#333333')
        
        # Style the axes
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#CCCCCC')
        ax.spines['bottom'].set_color('#CCCCCC')
    
    elif chart_type == 'numeric_histogram':
        # Create histogram
        n, bins, patches = ax.hist(
            data, 
            bins=min(30, max(10, len(set(data))//3)), 
            color=blue_colors[0], 
            alpha=0.7, 
            edgecolor='white',
            linewidth=1
        )
        
        # Color gradient for bars
        for i, patch in enumerate(patches):
            patch.set_facecolor(blue_colors[i % len(blue_colors)])
        
        ax.set_xlabel(f'{column_name} Values', fontsize=12, fontweight='bold', color='#333333')
        ax.set_ylabel('Frequency', fontsize=12, fontweight='bold', color='#333333')
        
        # Add statistics text
        mean_val = np.mean(data)
        std_val = np.std(data)
        stats_text = f'Mean: {mean_val:.2f}\nStd Dev: {std_val:.2f}\nSamples: {len(data):,}'
        ax.text(0.75, 0.95, stats_text, transform=ax.transAxes, 
                bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.8),
                verticalalignment='top', fontsize=10, fontweight='bold')
    
    # Set title and subtitle
    main_title = f"{title}"
    if subtitle:
        main_title += f"\n{subtitle}"
    
    ax.set_title(main_title, fontsize=16, fontweight='bold', color='#1565C0', pad=20)
    
    # Style tick parameters
    ax.tick_params(colors='#333333', labelsize=10)
    
    # Tight layout
    plt.tight_layout()
    
    # Save chart with high quality
    chart_path = os.path.join(app.config['TEMP_FOLDER'], filename)
    plt.savefig(chart_path, facecolor='white', dpi=300, bbox_inches='tight', 
                edgecolor='none', transparent=False)
    plt.close()
    
    return chart_path

def create_time_series_chart(date_data, column_name, filename):
    """Create time series analysis chart"""
    if not date_data:
        return None
    
    plt.style.use('default')
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
    fig.patch.set_facecolor('white')
    
    # Convert to datetime and sort
    dates = pd.to_datetime(date_data).sort_values()
    
    # Timeline plot
    ax1.hist(dates, bins=min(20, len(dates)//5 + 1), color='#1E88E5', alpha=0.7, edgecolor='white')
    ax1.set_title(f'{column_name} - Timeline Distribution', fontsize=14, fontweight='bold', color='#1565C0')
    ax1.set_ylabel('Frequency', fontsize=12, fontweight='bold')
    ax1.grid(True, alpha=0.3)
    
    # Monthly/yearly aggregation
    date_counts = dates.dt.to_period('M').value_counts().sort_index()
    ax2.plot(date_counts.index.astype(str), date_counts.values, 
             marker='o', linewidth=2, color='#1E88E5', markersize=6)
    ax2.fill_between(range(len(date_counts)), date_counts.values, alpha=0.3, color='#1E88E5')
    ax2.set_title(f'{column_name} - Monthly Trend', fontsize=14, fontweight='bold', color='#1565C0')
    ax2.set_ylabel('Count', fontsize=12, fontweight='bold')
    ax2.tick_params(axis='x', rotation=45)
    ax2.grid(True, alpha=0.3)
    
    plt.tight_layout()
    
    chart_path = os.path.join(app.config['TEMP_FOLDER'], filename)
    plt.savefig(chart_path, facecolor='white', dpi=300, bbox_inches='tight')
    plt.close()
    
    return chart_path

def generate_pdf_report(analysis, report_title, company_name):
    """Generate enhanced PDF report with better formatting"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"HR_Report_{timestamp}.pdf"
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    story = []
    
    # Enhanced styles
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=28,
        textColor=colors.HexColor('#1565C0'),
        spaceAfter=30,
        alignment=1,
        fontName='Helvetica-Bold'
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=20,
        textColor=colors.HexColor('#1E88E5'),
        spaceAfter=20,
        spaceBefore=10,
        alignment=1,
        fontName='Helvetica-Bold'
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=colors.HexColor('#1565C0'),
        spaceBefore=25,
        spaceAfter=15,
        fontName='Helvetica-Bold'
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#333333'),
        spaceAfter=12,
        fontName='Helvetica'
    )
    
    # Title page
    story.append(Paragraph(company_name, title_style))
    story.append(Paragraph(report_title, subtitle_style))
    story.append(Paragraph(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", normal_style))
    story.append(Spacer(1, 30))
    
    # Executive Summary
    story.append(Paragraph("📊 Executive Summary", heading_style))
    
    summary_data = [
        ["📋 Metric", "📈 Value", "📝 Description"],
        ["Files Analyzed", f"{analysis['summary']['total_files']}", "Total Excel files processed"],
        ["Total Records", f"{analysis['summary']['total_rows']:,}", "Combined data rows across all files"],
        ["Data Columns", f"{analysis['summary']['total_columns']}", "Unique column types identified"],
        ["Numeric Fields", f"{analysis['summary']['numeric_columns']}", "Quantitative analysis columns"],
        ["Category Fields", f"{analysis['summary']['categorical_columns']}", "Classification/grouping columns"],
        ["Date Fields", f"{analysis['summary']['date_columns']}", "Temporal analysis columns"]
    ]
    
    summary_table = Table(summary_data, colWidths=[2.2*inch, 1.5*inch, 3*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E3F2FD')),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#333333')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    ]))
    
    story.append(summary_table)
    story.append(Spacer(1, 25))
    
    # Key Insights
    story.append(Paragraph("🔍 Key Data Insights", heading_style))
    for i, insight in enumerate(analysis['insights'], 1):
        story.append(Paragraph(f"{i}. {insight}", normal_style))
    story.append(PageBreak())
    
    # Categorical Analysis Charts
    categorical_data = analysis['charts_data']['categorical']
    chart_count = 0
    
    if categorical_data:
        story.append(Paragraph("📊 Categorical Data Analysis", heading_style))
        
        for col_name, cat_data in list(categorical_data.items())[:6]:
            if len(cat_data) > 1:
                chart_count += 1
                total_records = sum(cat_data.values())
                
                # Choose chart type based on number of categories
                if len(cat_data) <= 8:
                    chart_path = create_professional_chart(
                        'categorical_pie', 
                        cat_data,
                        f'{col_name} Distribution',
                        f'cat_pie_{chart_count}.png',
                        col_name,
                        f'Total Records: {total_records:,}'
                    )
                else:
                    chart_path = create_professional_chart(
                        'categorical_bar', 
                        cat_data,
                        f'{col_name} Distribution',
                        f'cat_bar_{chart_count}.png',
                        col_name,
                        f'Total Records: {total_records:,} | Showing Top 20 Categories'
                    )
                
                story.append(Paragraph(f"📋 {col_name} Analysis", heading_style))
                
                # Add summary statistics
                top_3 = sorted(cat_data.items(), key=lambda x: x[1], reverse=True)[:3]
                summary_text = f"Most common values: "
                for i, (value, count) in enumerate(top_3):
                    percentage = (count / total_records) * 100
                    summary_text += f"{value} ({percentage:.1f}%)"
                    if i < len(top_3) - 1:
                        summary_text += ", "
                
                story.append(Paragraph(summary_text, normal_style))
                story.append(Image(chart_path, width=6.5*inch, height=4.5*inch))
                story.append(Spacer(1, 20))
    
    # Numeric Analysis Charts
    numeric_data = analysis['charts_data']['numeric']
    
    if numeric_data:
        story.append(PageBreak())
        story.append(Paragraph("📈 Numeric Data Analysis", heading_style))
        
        for col_name, num_data in list(numeric_data.items())[:5]:
            if len(num_data) > 1:
                chart_count += 1
                
                # Create histogram
                chart_path = create_professional_chart(
                    'numeric_histogram', 
                    num_data,
                    f'{col_name} Distribution Analysis',
                    f'num_hist_{chart_count}.png',
                    col_name,
                    f'Sample Size: {len(num_data):,} records'
                )
                
                story.append(Paragraph(f"📊 {col_name} Statistical Analysis", heading_style))
                
                # Add detailed statistics
                stats_text = f"Mean: {np.mean(num_data):,.2f} | Median: {np.median(num_data):,.2f} | "
                stats_text += f"Range: {np.min(num_data):,.2f} - {np.max(num_data):,.2f} | "
                stats_text += f"Std Deviation: {np.std(num_data):,.2f}"
                
                story.append(Paragraph(stats_text, normal_style))
                story.append(Image(chart_path, width=6.5*inch, height=4.5*inch))
                story.append(Spacer(1, 20))
    
    # Date Analysis
    date_data = analysis['charts_data']['dates']
    if date_data:
        story.append(PageBreak())
        story.append(Paragraph("📅 Temporal Data Analysis", heading_style))
        
        for col_name, dates in list(date_data.items())[:3]:
            if len(dates) > 1:
                chart_count += 1
                chart_path = create_time_series_chart(dates, col_name, f'date_analysis_{chart_count}.png')
                
                if chart_path:
                    story.append(Paragraph(f"📅 {col_name} Temporal Analysis", heading_style))
                    
                    # Date range info
                    min_date = min(dates).strftime('%Y-%m-%d')
                    max_date = max(dates).strftime('%Y-%m-%d')
                    date_range = f"Date Range: {min_date} to {max_date} | Total Records: {len(dates):,}"
                    
                    story.append(Paragraph(date_range, normal_style))
                    story.append(Image(chart_path, width=6.5*inch, height=4.5*inch))
                    story.append(Spacer(1, 20))
    
    # Build the PDF
    doc.build(story)
    return filename

# Clean HTML Template with Blue Theme
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HR Monthly Report Generator</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1976D2 0%, #1565C0 100%);
            color: white;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding: 30px 0;
        }
        
        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
        }
        
        .card {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 25px;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        
        .upload-section {
            border: 2px dashed rgba(255, 255, 255, 0.5);
            text-align: center;
            padding: 40px 20px;
            min-height: 200px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        
        .upload-section.dragover {
            border-color: #42A5F5;
            background: rgba(66, 165, 245, 0.1);
        }
        
        .file-input {
            display: none;
        }
        
        .btn {
            background: linear-gradient(135deg, #1976D2, #1565C0);
            color: white;
            border: none;
            padding: 12px 25px;
            font-size: 1rem;
            cursor: pointer;
            border-radius: 8px;
            transition: all 0.3s ease;
            margin: 5px;
            box-shadow: 0 4px 12px rgba(25, 118, 210, 0.3);
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(25, 118, 210, 0.4);
        }
        
        .btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }
        
        .btn-success {
            background: linear-gradient(135deg, #1976D2, #42A5F5);
        }
        
        .btn-danger {
            background: linear-gradient(135deg, #1565C0, #0D47A1);
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-label {
            display: block;
            margin-bottom: 8px;
            font-weight: bold;
        }
        
        .form-input {
            width: 100%;
            padding: 12px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            font-size: 1rem;
        }
        
        .form-input:focus {
            outline: none;
            border-color: #42A5F5;
            box-shadow: 0 0 0 3px rgba(66, 165, 245, 0.2);
        }
        
        .form-input::placeholder {
            color: rgba(255, 255, 255, 0.7);
        }
        
        .file-item {
            background: rgba(255, 255, 255, 0.1);
            padding: 15px;
            margin: 10px 0;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-left: 4px solid #42A5F5;
        }
        
        .status-message {
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
            font-weight: bold;
        }
        
        .status-success {
            background: rgba(30, 136, 229, 0.3);
            border: 2px solid #1E88E5;
        }
        
        .status-error {
            background: rgba(13, 71, 161, 0.3);
            border: 2px solid #0D47A1;
        }
        
        .status-info {
            background: rgba(25, 118, 210, 0.3);
            border: 2px solid #1976D2;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 30px;
        }
        
        .spinner {
            width: 40px;
            height: 40px;
            border: 4px solid rgba(255, 255, 255, 0.3);
            border-top: 4px solid #42A5F5;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        
        .stat-card {
            background: linear-gradient(135deg, rgba(30, 136, 229, 0.2), rgba(25, 118, 210, 0.2));
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            border: 1px solid rgba(66, 165, 245, 0.3);
        }
        
        .stat-value {
            font-size: 1.5rem;
            font-weight: bold;
            margin-bottom: 5px;
            color: #42A5F5;
        }
        
        .data-preview {
            background: rgba(0, 0, 0, 0.2);
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
            overflow-x: auto;
            border-left: 4px solid #42A5F5;
        }
        
        .data-preview table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        
        .data-preview th,
        .data-preview td {
            padding: 8px;
            text-align: left;
            border-bottom: 1px solid rgba(255, 255, 255, 0.2);
        }
        
        .data-preview th {
            background: rgba(30, 136, 229, 0.3);
            font-weight: bold;
            color: #42A5F5;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 HR Monthly Report Generator</h1>
            <p>Upload Excel files to generate comprehensive HR reports with charts and analysis</p>
        </div>
        
        <!-- File Upload Section -->
        <div class="card upload-section" id="uploadSection">
            <div>
                <h3>📁 Upload Excel Files</h3>
                <p>Select multiple .xlsx or .xls files</p>
                <input type="file" id="fileInput" class="file-input" multiple accept=".xlsx,.xls">
                <button class="btn" onclick="document.getElementById('fileInput').click()">
                    📂 Browse Files
                </button>
            </div>
            <div id="filesList"></div>
        </div>
        
        <!-- Configuration Section -->
        <div class="card" id="configSection" style="display: none;">
            <h3>⚙️ Report Configuration</h3>
            
            <div class="stats-grid" id="statsGrid"></div>
            
            <div class="form-group">
                <label class="form-label" for="reportTitle">Report Title</label>
                <input type="text" id="reportTitle" class="form-input" 
                       placeholder="HR Monthly Report - August 2024" 
                       value="HR Monthly Report">
            </div>
            
            <div class="form-group">
                <label class="form-label" for="companyName">Company Name</label>
                <input type="text" id="companyName" class="form-input" 
                       placeholder="Your Company Name" 
                       value="Company Analytics">
            </div>
            
            <button class="btn btn-success" id="generateBtn" onclick="generateReports()" disabled>
                📊 Generate Reports
            </button>
            <button class="btn btn-danger" onclick="clearFiles()">
                🗑️ Clear Files
            </button>
        </div>
        
        <!-- Data Preview -->
        <div class="card" id="previewSection" style="display: none;">
            <h3>📈 Data Preview</h3>
            <div id="dataPreview"></div>
        </div>
        
        <!-- Loading -->
        <div class="loading" id="loadingSection">
            <div class="spinner"></div>
            <p>Generating reports...</p>
        </div>
        
        <!-- Status Messages -->
        <div id="statusMessages"></div>
    </div>
    
    <script>
        let uploadedFiles = [];
        let fileData = {};
        
        const fileInput = document.getElementById('fileInput');
        const uploadSection = document.getElementById('uploadSection');
        const filesList = document.getElementById('filesList');
        const configSection = document.getElementById('configSection');
        const previewSection = document.getElementById('previewSection');
        const generateBtn = document.getElementById('generateBtn');
        const statsGrid = document.getElementById('statsGrid');
        
        fileInput.addEventListener('change', handleFileUpload);
        
        // Drag and drop
        uploadSection.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadSection.classList.add('dragover');
        });
        
        uploadSection.addEventListener('dragleave', () => {
            uploadSection.classList.remove('dragover');
        });
        
        uploadSection.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadSection.classList.remove('dragover');
            
            const files = Array.from(e.dataTransfer.files).filter(file => 
                file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')
            );
            
            if (files.length > 0) {
                const dt = new DataTransfer();
                files.forEach(file => dt.items.add(file));
                fileInput.files = dt.files;
                handleFileUpload();
            } else {
                showStatus('Please upload Excel files (.xlsx or .xls only)', 'error');
            }
        });
        
        function handleFileUpload() {
            const files = Array.from(fileInput.files);
            
            if (files.length === 0) return;
            
            showStatus(`Uploading ${files.length} Excel files...`, 'info');
            
            const formData = new FormData();
            files.forEach(file => {
                formData.append('excel_files', file);
            });
            
            fetch('/upload_excel', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
                return response.json();
            })
            .then(data => {
                if (data.success) {
                    uploadedFiles = data.files;
                    fileData = data.preview_data;
                    updateFilesList();
                    updateDataPreview();
                    updateStatsGrid(data.summary);
                    configSection.style.display = 'block';
                    previewSection.style.display = 'block';
                    generateBtn.disabled = false;
                    showStatus(`Successfully uploaded ${files.length} Excel files with ${data.summary.total_rows.toLocaleString()} records`, 'success');
                } else {
                    showStatus(`Error: ${data.error}`, 'error');
                }
            })
            .catch(error => {
                showStatus(`Upload failed: ${error.message}`, 'error');
            });
        }
        
        function updateStatsGrid(summary) {
            const stats = [
                { label: 'Files', value: summary.total_files },
                { label: 'Records', value: summary.total_rows.toLocaleString() },
                { label: 'Columns', value: summary.total_columns },
                { label: 'Numeric', value: summary.numeric_columns },
                { label: 'Categories', value: summary.categorical_columns },
                { label: 'Dates', value: summary.date_columns }
            ];
            
            statsGrid.innerHTML = stats.map(stat => `
                <div class="stat-card">
                    <div class="stat-value">${stat.value}</div>
                    <div>${stat.label}</div>
                </div>
            `).join('');
        }
        
        function updateFilesList() {
            filesList.innerHTML = '';
            
            uploadedFiles.forEach((file, index) => {
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item';
                fileItem.innerHTML = `
                    <div>
                        <h4>${file.filename}</h4>
                        <div>Sheets: ${file.sheets.join(', ')} | Size: ${file.size}</div>
                    </div>
                    <button class="btn btn-danger" onclick="removeFile(${index})">Remove</button>
                `;
                filesList.appendChild(fileItem);
            });
        }
        
        function updateDataPreview() {
            const previewDiv = document.getElementById('dataPreview');
            previewDiv.innerHTML = '';
            
            Object.keys(fileData).forEach(filename => {
                const fileDiv = document.createElement('div');
                fileDiv.className = 'data-preview';
                
                let html = `<h4>${filename}</h4>`;
                
                Object.keys(fileData[filename]).forEach(sheetName => {
                    const sheetData = fileData[filename][sheetName];
                    
                    html += `<h5>Sheet: ${sheetName} (${sheetData.rows} rows)</h5>`;
                    html += `<p><strong>Columns:</strong> ${sheetData.column_names.slice(0, 8).join(', ')}</p>`;
                    
                    if (sheetData.sample_data && sheetData.sample_data.length > 0) {
                        html += '<table><thead><tr>';
                        
                        const displayColumns = Object.keys(sheetData.sample_data[0]).slice(0, 6);
                        displayColumns.forEach(col => {
                            html += `<th>${col}</th>`;
                        });
                        html += '</tr></thead><tbody>';
                        
                        sheetData.sample_data.slice(0, 3).forEach(row => {
                            html += '<tr>';
                            displayColumns.forEach(col => {
                                const value = row[col];
                                const displayValue = value !== null && value !== undefined ? 
                                    String(value).substring(0, 30) : '-';
                                html += `<td>${displayValue}</td>`;
                            });
                            html += '</tr>';
                        });
                        
                        html += '</tbody></table>';
                    }
                });
                
                fileDiv.innerHTML = html;
                previewDiv.appendChild(fileDiv);
            });
        }
        
        function removeFile(index) {
            uploadedFiles.splice(index, 1);
            
            if (uploadedFiles.length === 0) {
                configSection.style.display = 'none';
                previewSection.style.display = 'none';
                generateBtn.disabled = true;
                fileData = {};
                statsGrid.innerHTML = '';
            }
            
            updateFilesList();
            updateDataPreview();
        }
        
        function clearFiles() {
            uploadedFiles = [];
            fileData = {};
            filesList.innerHTML = '';
            document.getElementById('dataPreview').innerHTML = '';
            configSection.style.display = 'none';
            previewSection.style.display = 'none';
            generateBtn.disabled = true;
            fileInput.value = '';
            statsGrid.innerHTML = '';
            showStatus('All files cleared', 'info');
        }
        
        function generateReports() {
            const reportTitle = document.getElementById('reportTitle').value.trim() || 'HR Monthly Report';
            const companyName = document.getElementById('companyName').value.trim() || 'Company';
            
            if (uploadedFiles.length === 0) {
                showStatus('Please upload Excel files first', 'error');
                return;
            }
            
            document.getElementById('loadingSection').style.display = 'block';
            generateBtn.disabled = true;
            
            fetch('/generate_reports', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    report_title: reportTitle,
                    company_name: companyName
                })
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
                return response.json();
            })
            .then(data => {
                document.getElementById('loadingSection').style.display = 'none';
                generateBtn.disabled = false;
                
                if (data.success) {
                    showStatus('Reports generated successfully!', 'success');
                    
                    const downloadDiv = document.createElement('div');
                    downloadDiv.className = 'status-message status-success';
                    downloadDiv.innerHTML = `
                        <h4>📊 Reports Ready for Download</h4>
                        <div style="margin-top: 15px;">
                            <a href="${data.pdf_url}" class="btn" download style="text-decoration: none; margin-right: 10px;">
                                📄 Download PDF Report
                            </a>
                        </div>
                    `;
                    
                    document.getElementById('statusMessages').appendChild(downloadDiv);
                } else {
                    showStatus(`Report generation failed: ${data.error}`, 'error');
                }
            })
            .catch(error => {
                document.getElementById('loadingSection').style.display = 'none';
                generateBtn.disabled = false;
                showStatus(`Report generation failed: ${error.message}`, 'error');
            });
        }
        
        function showStatus(message, type) {
            const statusDiv = document.createElement('div');
            statusDiv.className = `status-message status-${type}`;
            statusDiv.innerHTML = message;
            
            const statusContainer = document.getElementById('statusMessages');
            statusContainer.appendChild(statusDiv);
            
            setTimeout(() => {
                if (statusDiv.parentNode) {
                    statusDiv.parentNode.removeChild(statusDiv);
                }
            }, 5000);
        }
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    global uploaded_files, report_data
    
    try:
        if 'excel_files' not in request.files:
            return jsonify({'error': 'No files selected'}), 400
        
        files = request.files.getlist('excel_files')
        if not files or all(file.filename == '' for file in files):
            return jsonify({'error': 'No valid files selected'}), 400
        
        uploaded_files = []
        dataframes = {}
        preview_data = {}
        
        for file in files:
            if file.filename == '':
                continue
                
            if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
                continue
                
            try:
                filename = secure_filename(file.filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                unique_filename = f"{timestamp}_{filename}"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
                file.save(filepath)
                
                # Read all sheets
                excel_file = pd.ExcelFile(filepath)
                sheets = {}
                sheet_names = []
                file_size = os.path.getsize(filepath)
                
                preview_data[file.filename] = {}
                
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(filepath, sheet_name=sheet_name)
                        
                        # Clean column names
                        df.columns = df.columns.astype(str).str.strip()
                        
                        # Remove empty rows and columns
                        df = df.dropna(how='all').dropna(axis=1, how='all')
                        
                        if not df.empty and len(df) > 0:
                            sheets[sheet_name] = df
                            sheet_names.append(sheet_name)
                            
                            preview_data[file.filename][sheet_name] = {
                                'rows': len(df),
                                'columns': len(df.columns),
                                'column_names': list(df.columns),
                                'sample_data': df.head(5).fillna('').to_dict('records') if not df.empty else []
                            }
                    except Exception as sheet_error:
                        print(f"Warning: Could not read sheet '{sheet_name}': {sheet_error}")
                        continue
                
                if sheets:
                    dataframes[file.filename] = sheets
                    uploaded_files.append({
                        'filename': file.filename,
                        'filepath': filepath,
                        'sheets': sheet_names,
                        'size': f"{file_size / 1024:.1f} KB" if file_size < 1024*1024 else f"{file_size / (1024*1024):.1f} MB"
                    })
                    
            except Exception as file_error:
                print(f"Error processing file {file.filename}: {file_error}")
                return jsonify({'error': f'Failed to process {file.filename}: {str(file_error)}'}), 400
        
        if not uploaded_files:
            return jsonify({'error': 'No valid Excel files could be processed'}), 400
        
        # Analyze data
        report_data = analyze_excel_data(dataframes)
        
        return jsonify({
            'success': True,
            'files': uploaded_files,
            'preview_data': preview_data,
            'summary': report_data['summary']
        })
        
    except Exception as e:
        print(f"Upload error: {e}")
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

@app.route('/generate_reports', methods=['POST'])
def generate_reports():
    global report_data
    
    try:
        if not report_data:
            return jsonify({'error': 'No data available. Please upload Excel files first.'}), 400
        
        data = request.json
        if not data:
            return jsonify({'error': 'Invalid request data'}), 400
            
        report_title = data.get('report_title', 'HR Monthly Report').strip()
        company_name = data.get('company_name', 'Company').strip()
        
        # Generate PDF report
        pdf_filename = generate_pdf_report(report_data, report_title, company_name)
        
        return jsonify({
            'success': True,
            'pdf_filename': pdf_filename,
            'pdf_url': f'/download/{pdf_filename}'
        })
        
    except Exception as e:
        print(f"Report generation error: {e}")
        return jsonify({'error': f'Failed to generate reports: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        filename = secure_filename(filename)
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'File not found'}), 404
            
    except Exception as e:
        print(f"Download error: {e}")
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

if __name__ == '__main__':
    local_ip = get_local_ip()
    print(f"\n🚀 HR Monthly Report Generator")
    print(f"📱 Local: http://127.0.0.1:5000")
    print(f"🌐 Network: http://{local_ip}:5000")
    print(f"\nReady to process Excel files and generate HR reports!")
    
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
