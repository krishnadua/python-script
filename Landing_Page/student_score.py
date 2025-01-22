import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
import os

# Function to generate report card PDF for each student
def generate_report_card(student_id, student_name, subject_scores, total_score, average_score):
    # Define file name for PDF with the format report_card_<StudentID>.pdf
    filename = f"report_card_{student_id}.pdf"
    
    # Create a PDF document
    document = SimpleDocTemplate(filename, pagesize=letter)
    content = []

    # Add title
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    content.append(Paragraph(f"Report Card for {student_name}", title_style))
    
    # Add total and average score
    content.append(Paragraph(f"Total Score: {total_score}", styles['Normal']))
    content.append(Paragraph(f"Average Score: {average_score:.2f}", styles['Normal']))
    
    # Create the table of subject scores
    table_data = [["Subject", "Score"]]  # Table header
    for subject, score in subject_scores.items():
        table_data.append([subject, score])
    
    # Create and style the table
    table = Table(table_data, colWidths=[150, 100])
    table.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONT_SIZE', (0, 0), (-1, -1), 10),
    ]))
    
    content.append(table)
    
    # Build the document
    document.build(content)

# Function to process the Excel data and generate the report cards
def process_student_scores(excel_file):
    try:
        # Read the Excel file using pandas
        df = pd.read_excel(excel_file)
        
        # Clean up the column names (remove extra spaces)
        df.columns = df.columns.str.strip()  # Remove any leading/trailing spaces
        
        # Validate the columns
        if 'Student ID' not in df.columns or 'Name' not in df.columns or 'Subject Score' not in df.columns:
            raise ValueError("Excel file must contain 'Student ID', 'Name', and 'Subject Score' columns")
        
        # Group the data by Student ID
        grouped = df.groupby(['Student ID', 'Name'])['Subject Score'].apply(list).reset_index()
        
        # Iterate over each student group to generate report cards
        for _, row in grouped.iterrows():
            student_id = row['Student ID']
            student_name = row['Name']
            subject_scores = row['Subject Score']
            
            if not subject_scores:
                print(f"Missing scores for student {student_name} (ID: {student_id}), skipping.")
                continue
            
            # Calculate total and average score
            total_score = sum(subject_scores)
            average_score = total_score / len(subject_scores)
            
            # Create a dictionary of subjects and their scores (assuming unique subjects for simplicity)
            subject_dict = {f"Subject {i+1}": score for i, score in enumerate(subject_scores)}
            
            # Generate the PDF report card for this student
            generate_report_card(student_id, student_name, subject_dict, total_score, average_score)
    
    except Exception as e:
        print(f"Error processing the Excel file: {e}")

# Main function to run the script
if __name__ == "__main__":
    excel_file = "student_scores.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"Error: The file '{excel_file}' does not exist.")
    else:
        process_student_scores(excel_file)
