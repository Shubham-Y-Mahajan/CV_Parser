import zipfile
import os
import streamlit as st
from backend import identify_filetypes,excel_writer,doc_to_docx,section_extractor,\
    extract_text_from_pdf,extract_text_from_docx,clear_directory


st.title("CV Data Extractor")
st.info("Developed By - Shubham Mahajan | +91 8879466601 |shubhamy@iitbhilai.ac.in")
uploaded_zip = st.file_uploader('XML File', type="zip",accept_multiple_files=False)
if (uploaded_zip is not None):



    try:
        zf = zipfile.ZipFile(uploaded_zip)
        zf.extractall(path="extracted")
        folder_name=(os.listdir("extracted"))[0]
        directory = os.path.join("extracted", folder_name)

        filetypes=identify_filetypes(directory)
        pdf=filetypes[0]
        docx=filetypes[1]
        doc=filetypes[2]

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

            data=section_extractor(text=text,filename=filename)
            print("Loading.....")
            excel_data.append(data)


        excel_writer(data=excel_data)

        print("complete")

        with open("Report.xlsx", 'rb') as my_file:
            st.download_button(label='Download Output', data=my_file, file_name='Shubham_Mahajan_App_Output.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except IndexError and NotADirectoryError:
        clear_directory("extracted")

        st.info("Error : Submit a zipped folder which contains the bundle of CVs")







