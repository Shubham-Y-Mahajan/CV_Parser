import os

import pythoncom
from pdf2docx import Converter
from docx import Document
import re
from PyPDF2 import PdfReader
import xlsxwriter
import aspose.words as aw
def identify_filetypes(folder_path):
    pdf_files = []
    docx_files = []
    doc_files = []
    # Iterate over files in the folder
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        # Check if it's a file (not a directory)
        if os.path.isfile(filepath):
            # Check file extension
            if filename.lower().endswith('.pdf'):
                pdf_files.append(filename)
            elif filename.lower().endswith('.docx'):
                docx_files.append(filename)
            elif filename.lower().endswith('.doc'):
                doc_files.append(filename)

    return [pdf_files,docx_files,doc_files]

def pdf_to_docx(files,content_path):
    for pdf_file in files:
        file_name=os.path.splitext(pdf_file)[0]
        docx_name=file_name + ".docx"

        pdf_path=f"{content_path}/{pdf_file}"
        docx_path=f"formatted/{docx_name}"
        # Initialize Converter object
        cv = Converter(pdf_path)
        # Convert the PDF to DOCX
        cv.convert(docx_path, start=0, end=None)
        # Close the Converter object
        cv.close()


def doc_to_docx(files,content_path):
    for doc_file in files:
        file_name=os.path.splitext(doc_file)[0]
        docx_name=file_name + ".docx"

        doc_path=f"{content_path}/{doc_file}"
        docx_path=f"{content_path}/{docx_name}"

        doc = aw.Document(doc_path)
        doc.save(docx_path)
        os.remove(doc_path)





# Function to extract text from DOCX
def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

def extract_text_from_doc(doc_file_path):

    try:
        doc = Document(doc_file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        print(f"An error occurred: {e}")
        return None


def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, "rb") as file:
        pdf_reader = PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text
def extract_section(text, section_heading):
    pattern = re.compile(r'(?i)' + re.escape(section_heading) + r'\s*\n')
    start_match = pattern.search(text)
    if start_match:
        start_index = start_match.end()
        next_heading_match = pattern.search(text[start_match.end():])
        end_index = next_heading_match.start() if next_heading_match else len(text)
        section_text = text[start_index:end_index].strip()
        return section_text
    else:
        return None

def extract_text_between_words(text, word1, word2):
    pattern = re.compile(r'{}(.*?){}'.format(re.escape(word1), re.escape(word2)),re.DOTALL)
    match = pattern.search(text)
    if match:
        return match.group(1).strip()
    else:
        return None

def section_extractor(text):
    headings = [
        'Personal Information',
        'Objective',
        'Work History',
        'Education',
        'Work Experience',
        'Professional Experience',
        'Skills',
        'Certifications',
        'Certificates'
        'Projects',
        'Publications',
        'Awards',
        'Employment History'
        'Professional Affiliations',
        'References',
        'Languages',
        'Achievements',
        'Academic Credentials',
        'Profile',
        'Details',
        'Personal Details',
        'Work Summary',
        'Desired Job Details'

    ]

    text_data = text.split()
    sections_present=[]

    "---------Removing Aspose watermark-----------------"
    if text_data[4] == "Aspose.Words.":
        for i in range(10):
            text_data.pop(0)
    "---------------"

    sections_present.append(text_data[0])
    for word in headings:
        capital=word.upper()
        if capital in text:
            sections_present.append(capital)
        elif word in text:
            sections_present.append(word)


    "--------------------------------------------"
    "special case handling"

    if 'Work Experience' not in sections_present and 'Work Experience'.upper() not in sections_present and \
            'Professional Experience' not in sections_present and 'Professional Experience'.upper() not in sections_present:
        if "EXPERIENCE" in text:
            sections_present.append("EXPERIENCE")
        elif "Experience" in text:
            sections_present.append("Experience")

    if "ACADEMIC CREDENTIALS" in sections_present and "Education" in sections_present:
        sections_present.remove("Education") # special case for Akash Goel

    "------------------------------------------------------------------------------------------------------"
    extracted_data=[]

    for section in sections_present:

        min_data=None

        for next_section in sections_present:
            if section != next_section:
                data = extract_text_between_words(text=text, word1=f"{section}", word2=f"{next_section}")

                if min_data:
                    if data:
                        if len(data) < len(min_data) and len(data) > 0:
                            min_data=data
                else:
                    if data:
                        min_data=data

        if not min_data:
            "now consider its the last section thus we wil match till end of doc"
            end_pattern = re.compile(r'{}(.*)'.format(re.escape(section)), re.DOTALL)
            match = end_pattern.search(text)
            if match:
                min_data = match.group(1).strip()

        extracted_data.append([section,min_data])

    for item in extracted_data:
        print(item[0])
        print(item[1])
    return extracted_data



def extract_data(folder_path,section_headings):


    # List to store data
    data = []

    # Loop through PDF files
    for filename in os.listdir(folder_path):
        print(filename)
        filepath = os.path.join(folder_path, filename)
        print(filepath)
        if os.path.isfile(filepath):
            # Check file extension
            if filename.lower().endswith('.pdf'):
                text=extract_text_from_pdf(filepath)
            elif filename.lower().endswith('.docx'):
                text=extract_text_from_docx(filepath)
            elif filename.lower().endswith('.doc'):
                text=extract_text_from_doc(filepath)
            else:
                text=""
        else:
            text = ""


        # Extract each section dynamically based on section headings
        sections = {}
        for section_heading in section_headings:
            section_text = extract_section(text, section_heading)
            sections[section_heading] = section_text

        # Append extracted sections to data list
        data.append({'Filename': filename, **sections})

    return data


def excel_writer(data):
    if os.path.exists("Report.xlsx"):
        os.remove("Report.xlsx")
    workbook = xlsxwriter.Workbook(f"Report.xlsx")



    for CV in data:
        worksheet = workbook.add_worksheet()
        col = 0
        for index,item in enumerate(CV):
            if index == 0:

                worksheet.write(0, col, f"Introduction")
                worksheet.write(1, col, f"{item[0]} {item[1]}")
                col += 1
            else:
                worksheet.write(0, col, f"{item[0]} ")
                worksheet.write(1, col, f"{item[1]}")
                col += 1


    workbook.close()
    return 1

if __name__=="__main__":
    text=extract_text_from_pdf("extracted/Sample2/AarushiRohatgi.pdf")
    #text=extract_text_from_docx("extracted/Sample2/heemSen.docx")

    section_extractor(text=text)

