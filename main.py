import zipfile
import os
import streamlit as st
from backend import identify_filetypes,pdf_to_docx,extract_data,excel_writer,doc_to_docx,section_extractor,\
    extract_text_from_pdf,extract_text_from_docx
import xlsxwriter

st.title("CV data extractor")
uploaded_zip = st.file_uploader('XML File', type="zip",accept_multiple_files=False)
if (uploaded_zip is not None):
    #print(uploaded_zip)
    zf = zipfile.ZipFile(uploaded_zip)
    zf.extractall(path="extracted")

    folder_name=(os.listdir("extracted"))[0]

    filetypes=identify_filetypes(f"extracted/{folder_name}")
    pdf=filetypes[0]
    docx=filetypes[1]
    doc=filetypes[2]
    print("-------------------------------------------------------------")

    directory=f"extracted\{folder_name}"

    doc_to_docx(files=doc,content_path=directory)


    excel_data=[]
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        # Check file extension
        if filename.lower().endswith('.pdf'):
            text=extract_text_from_pdf(filepath)
        elif filename.lower().endswith('.docx'):
            text=extract_text_from_docx(filepath)

        else:
            text=""

        data=section_extractor(text=text)
        print("Loading.....")
        excel_data.append(data)


    excel_writer(data=excel_data)
    print("complete")









