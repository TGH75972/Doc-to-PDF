import os
import pypandoc
from docx2pdf import convert as convert_docx
from pathlib import Path
import logging
from datetime import datetime

logging.basicConfig(filename='conversion.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_docx_to_pdf(docx_path, pdf_path):
    """Convert a .docx file to PDF using docx2pdf."""
    try:
        logging.info(f"Starting conversion of {docx_path} to PDF.")
        convert_docx(docx_path, pdf_path)
        logging.info(f"Successfully converted {docx_path} to {pdf_path}.")
    except Exception as e:
        logging.error(f"Error converting {docx_path} to PDF: {e}")

def convert_doc_to_pdf(doc_path, pdf_path):
    """Convert a .doc file to PDF using Pandoc."""
    try:
        logging.info(f"Starting conversion of {doc_path} to PDF.")
        output = pypandoc.convert_file(doc_path, 'pdf', outputfile=pdf_path)
        logging.info(f"Successfully converted {doc_path} to {pdf_path}.")
        return output
    except Exception as e:
        logging.error(f"Error converting {doc_path} to PDF: {e}")

def get_files(directory, extension):
    """Get a list of files in a directory with the given extension."""
    try:
        files = [f for f in Path(directory).glob(f'*{extension}')]
        logging.info(f"Found {len(files)} files with extension {extension}.")
        return files
    except Exception as e:
        logging.error(f"Error retrieving files with extension {extension}: {e}")
        return []

def convert_files(directory, extension):
    """Convert all files in the directory with the given extension to PDF."""
    files = get_files(directory, extension)
    for file in files:
        pdf_path = file.with_suffix('.pdf')
        if extension == '.docx':
            convert_docx_to_pdf(file, pdf_path)
        elif extension == '.doc':
            convert_doc_to_pdf(file, pdf_path)

def create_directory_if_not_exists(directory):
    """Create directory if it does not exist."""
    if not os.path.exists(directory):
        try:
            os.makedirs(directory)
            logging.info(f"Created directory: {directory}")
        except Exception as e:
            logging.error(f"Error creating directory {directory}: {e}")

def backup_existing_files(directory):
    """Backup existing PDF files in the directory."""
    backup_dir = Path(directory) / 'backup'
    create_directory_if_not_exists(backup_dir)
    pdf_files = get_files(directory, '.pdf')
    for file in pdf_files:
        try:
            backup_path = backup_dir / file.name
            file.rename(backup_path)
            logging.info(f"Backed up {file} to {backup_path}.")
        except Exception as e:
            logging.error(f"Error backing up {file}: {e}")

def main():
    """Main function to handle user input and conversion process."""
    try:
        directory = input("Enter the directory containing the files: ").strip()
        if not os.path.isdir(directory):
            logging.error("Invalid directory provided.")
            print("Invalid directory")
            return

        print("Choose the type of files to convert:")
        print("1. .docx")
        print("2. .doc")
        
        choice = input("Enter your choice (1 or 2): ").strip()
        
        if choice == '1':
            extension = '.docx'
        elif choice == '2':
            extension = '.doc'
        else:
            logging.error("Invalid choice provided.")
            print("Invalid choice")
            return

        backup_existing_files(directory)
        convert_files(directory, extension)
        
        print("Conversion completed. Check the conversion.log file for details.")
    
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
