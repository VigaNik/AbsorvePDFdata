import subprocess
import json #(PyMuPDF)for reading and checking json files
import os

local_dir_input_pdf = r'C:\WORK\StudiesPDF\corpus\sidra\pdf'  #PDF files local directory
local_dir_input_meta = r'C:\WORK\Corpus\corpus-metadata\sidra\meta'

def is_it_scanned_pdf(meta_file):
    """Here we check if meta file is scanned or writing, if file have text layers as False otherwise is True"""

    try:
        #Open json file
        with open(meta_file, 'r', encoding='utf-8') as file:
            new_file = json.load(file)
    except Exception as e:
        print(f"Error in {meta_file}: {e} do not have informathion ")
        return False
    if 'content' in new_file:
        content = new_file['content']
        is_page_scan = content.get("isPageScan", False)

        #Chek if "isPageScan" exthist and is it true
        print(f"'isPageScan' value in {meta_file}: {is_page_scan}")

        return is_page_scan  # Вернёт True, если isPageScan = True, иначе False
    else:
        print(f"Key 'content' not found in {meta_file}")
        return False

def extract_pdf_and_meta_from_local_folder():
    """Download all pdf and meta files from local folder, and call is_it_scanned_pdf"""
    #Chek if files is in correct formart
    pdf_files = [file for file in os.listdir(local_dir_input_pdf) if file.endswith(".pdf")]
    meta_files = [file for file in os.listdir(local_dir_input_meta) if file.endswith(".json")]

    for filename_pdf in pdf_files:
        base_name_pdf = os.path.splitext(filename_pdf)[0]  # Get file name without extension
        file_name_pdf = os.path.join(local_dir_input_pdf, filename_pdf)
        for filename_meta in meta_files:
            base_name_meta = os.path.splitext(filename_meta)[0]  # Get file name without extension
            file_name_meta = os.path.join(local_dir_input_meta, filename_meta)
            if base_name_pdf == base_name_meta:
                if is_it_scanned_pdf(file_name_meta):
                    print(f"File name:  {base_name_pdf} is a scanned PDF. Applying OCR...")
                    #text = moduleForScannedPDF(file_name_pdf)
                else:
                    print(f" File name: {base_name_pdf} is a printed PDF. Extracting text directly...")

                    subprocess.run(['python', 'PrintedPdf.py', file_name_pdf], check=True)


if __name__ == "__main__":
     extract_pdf_and_meta_from_local_folder()
