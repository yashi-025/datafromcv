# Import necessary libraries
import pandas as pd
import os
import re
import docx2txt
from pdfminer.high_level import extract_text
from pdfminer.pdfparser import PDFSyntaxError

# Function to remove illegal characters from text
def remove_illegal_characters(text):
    illegal_chars_pattern = r'[^\x09\x0A\x0D\x20-\x7E]'
    cleaned_text = re.sub(illegal_chars_pattern, '', text)
    return cleaned_text

# Function to extract text from PDF files
def extract_text_from_pdf(pdf_path):
    try:
        return extract_text(pdf_path)
    except PDFSyntaxError as e:
        print(f"Error processing {pdf_path}: {e}")
        return ''

# Function to extract text from DOCX files
def extract_text_from_docx(docx_path):
    return docx2txt.process(docx_path)

# Function to load all CV files
def loadAllCV():
    dir_list = os.listdir('./cv')
    resumes = []
    for cv in dir_list:
        file_extension = os.path.splitext(cv)[1].lower()
        if file_extension == '.pdf':
            rr = loadSinglePDF(cv)
        elif file_extension == '.docx':
            rr = loadSingleDocx(cv)
        else:
            print(f"Ignoring file: {cv} (Unsupported format)")
            continue
        resumes.append(rr)
    return resumes

# Function to load text from a PDF file
def loadSinglePDF(file):
    try:
        data = extract_text_from_pdf('./cv/' + file)
        print('Processing PDF:', file)
        cleaned_data = remove_illegal_characters(data)
        return cleaned_data
    except Exception as e:
        print(f"Error processing {file}: {e}")
        return None

# Function to load text from a DOCX file
def loadSingleDocx(file):
    try:
        data = extract_text_from_docx('./cv/' + file)
        print('Processing DOCX:', file)
        cleaned_data = remove_illegal_characters(data)
        return cleaned_data
    except Exception as e:
        print(f"Error processing {file}: {e}")
        return None

# Function to extract information from the resume text
def extract_info(resume_text):
    # Initialize empty fields
    name = ''
    mobile = ''
    email = ''
    skills = ''
    designation = ''
    education = ''
    awards = ''
    projects = ''

    # Extract name, mobile, and email
    name_match = re.search(r'^[A-Z][a-z]+ [A-Z][a-z]+$', resume_text)
    if name_match:
        name = name_match.group(0)
    
    mobile_match = re.search(r'\b\d{10}\b', resume_text)
    if mobile_match:
        mobile = mobile_match.group(0)
    
    email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', resume_text)
    if email_match:
        email = email_match.group(0)
    
    # Extract skills, designation, education, awards, and projects
    sections = re.split(r'\n\n|\.\s+', resume_text)
    for section in sections:
        if 'Skill' in section:
            skills = section
        elif 'Designation' in section:
            designation = section
        elif 'Education' in section:
            education = section
        elif 'Award' in section:
            awards = section
        elif 'Project' in section:
            projects = section
    
    return {
        'Name': name,
        'Mobile': mobile,
        'Email': email,
        'Skills': skills,
        'Designation': designation,
        'Education': education,
        'Awards': awards,
        'Projects': projects
    }

# Main function
if __name__ == '__main__':
    resumes = loadAllCV()
    data = []
    for resume in resumes:
        if resume:
            info = extract_info(resume)
            data.append(info)
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Write DataFrame to Excel file
    df.to_excel('resumes.xlsx', index=False)
