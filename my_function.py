import base64
import pandas as pd
import os
import fitz  # PyMuPDF
from datetime import datetime

def compress_pdf(input_pdf_path, output_pdf_path, target_size_kb=1):
    # Open the PDF file
    pdf_document = fitz.open(input_pdf_path)
    
    # Define compression parameters
    garbage_values = [1, 2, 3, 4]  # Possible values for the garbage parameter
    deflate_values = [True]  # Possible values for the deflate parameter
    
    # Iterate over combinations of garbage and deflate parameters
    for garbage in garbage_values:
        print(garbage)
        for deflate in deflate_values:
            print(deflate)
            # Compress the PDF with the current parameters
            pdf_document.save(output_pdf_path, garbage=garbage, deflate=deflate)
            # Get the size of the resulting file
            output_size_kb = os.path.getsize(output_pdf_path) / 1024  # Convert to KB
            # Check if the resulting size is close to the target size
            if abs(output_size_kb - target_size_kb) <= 1:
                # If close enough to the target size, exit the loop
                pdf_document.close()
                return
            # If the resulting size is larger than the target, try next combination
            elif output_size_kb > target_size_kb:
                break
    
    # Close the PDF document if no suitable compression is found
    pdf_document.close()
    return output_pdf_path

def pdf_to_base64(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        pdf_bytes = pdf_file.read()
        base64_bytes = base64.b64encode(pdf_bytes)
        base64_string = base64_bytes.decode("utf-8")
    return base64_string

def pdfs_to_base64(pdf_paths):
    # print(pdf_paths)
    base64_dict = {}
    for pdf_path in pdf_paths:
        with open(pdf_path, "rb") as pdf_file:
            pdf_bytes = pdf_file.read()
            base64_bytes = base64.b64encode(pdf_bytes)
            base64_string = base64_bytes.decode("utf-8")
        base64_dict[pdf_path] = base64_string
    return base64_dict

def pdfs_to_base64_to_excel(pdf_paths):
    base64_dict = {}
    for pdf_path in pdf_paths:
        compressed_pdf_path = 'compressed_' + os.path.basename(pdf_path)
        compressed_path = compress_pdf(pdf_path, compressed_pdf_path, target_size_kb=5)
        
        with open(compressed_path, "rb") as pdf_file:
            print(f"""=== {compressed_path}""")
            pdf_bytes = pdf_file.read()
            base64_bytes = base64.b64encode(pdf_bytes)
            base64_string = base64_bytes.decode("utf-8")
        
        # if int(len(base64_string)) >= 36000:
        #     base64_string = "FILE TOO LARGE " + str(len(base64_string))
                
        file_name = os.path.basename(pdf_path)
        base64_dict[pdf_path] = ("", 
                                 "", 
                                 file_name, 
                                 "", 
                                 base64_string)
    # print(base64_dict)
    return base64_dict