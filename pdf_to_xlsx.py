import os
import glob
import shutil
import aspose.pdf as ap  # Ensure you have installed the aspose-pdf package

def convertPDFToExcel(input_folder, output_folder, converted_folder):
    """ 
    Function that convert PDF File into Excel

    Args:
        input_folder (_type_): use this folder to input your pdf(s)
        output_folder (_type_): all the xlsx (from converted pdf) will be here
        converted_folder (_type_): pdf will be moved here if conversion success
    """
    # Ensure the output and converted directories exist
    print("========== START PDF TO XLSX ==========")
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(converted_folder, exist_ok=True)
    i = 1

    # Get all PDF files in the input directory
    pdf_files = glob.glob(os.path.join(input_folder, "*.pdf"))

    for pdf_file in pdf_files:
        print(str(i)+" ===")
        try:
            # Construct the output file path with .xlsx extension
            base_name = os.path.basename(pdf_file)
            output_file = os.path.join(output_folder, os.path.splitext(base_name)[0] + ".xlsx")

            # Open PDF document
            document = ap.Document(pdf_file)

            # Define save options for Excel format
            save_option = ap.ExcelSaveOptions()

            # Save the file into MS Excel format
            document.save(output_file, save_option)
            print(f"Converted {pdf_file} to {output_file}")

            # Move the original PDF to the converted folder
            shutil.move(pdf_file, os.path.join(converted_folder, base_name))
            print(f"Moved {pdf_file} to {os.path.join(converted_folder, base_name)}")
        except Exception as e:
            # Handle conversion errors
            print(f"Failed to convert {pdf_file}: {e}")
        i += 1
    print("========== FINISH PDF TO XLSX ==========\n")