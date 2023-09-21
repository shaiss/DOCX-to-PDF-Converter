import logging
import os
import docx2pdf
import sys
import pythoncom
from concurrent.futures import ThreadPoolExecutor
from termcolor import colored, cprint

def convert_docx_to_pdf(docx_file, pdf_file, retries=3):
    logger = logging.getLogger(__name__)
    pythoncom.CoInitialize()
    for i in range(retries + 1):
        try:
            docx2pdf.convert(docx_file, pdf_file)
            return True
        except Exception as e:
            logger.error(f"Failed to convert {docx_file} to PDF on attempt {i + 1} of {retries}: {e} ({sys.exc_info()[1]})")
    return False

def convert_all_docx_to_pdf(folder_path, max_workers=4, retries=3):
    logger = logging.getLogger(__name__)

    docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".docx")]
    num_docx_files = len(docx_files)

    if num_docx_files == 0:
        cprint("No .docx files found in the specified folder.", "yellow")
        return

    executed_folder_path = os.path.join(folder_path, "executed")
    if not os.path.exists(executed_folder_path):
        os.makedirs(executed_folder_path)

    num_converted_files = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {executor.submit(convert_docx_to_pdf, docx_file, os.path.join(executed_folder_path, os.path.basename(docx_file).replace(".docx", ".pdf")), retries): docx_file for docx_file in docx_files}
        for future in future_to_file:
            if future.result():
                num_converted_files += 1

    success_msg = f"{num_converted_files} out of {num_docx_files} .docx files were successfully converted to .pdf files."
    if num_converted_files == num_docx_files:
        cprint(success_msg, "green")
    else:
        cprint(success_msg, "yellow")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = os.getcwd()
        count = len([f for f in os.listdir(folder_path) if f.endswith(".docx")])
        response = input(colored(f"There are {count} .docx files in {folder_path}. Do you want to continue? (Y/n) ", "cyan"))
        if response.strip().lower() not in ["y", "yes", ""]:
            sys.exit(0)

    logging.basicConfig(filename="docx2pdf.log", level=logging.INFO)
    convert_all_docx_to_pdf(folder_path)
