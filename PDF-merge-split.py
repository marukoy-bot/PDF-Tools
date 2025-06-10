import os
import sys
from pathlib import Path
from pypdf import PdfReader, PdfWriter

def collect_pdfs_for_merge():
    pdf_files = []
    while True:
        # Prompt the user for a directory or file path
        directory_input = input("Enter the path to a directory or drag and drop PDF files here (or press Enter to finish): ").strip()

        # If the user presses Enter without typing anything, stop the process
        if not directory_input:
            break

        # Handle paths with quotes (from drag-and-drop)
        directory_input = directory_input.strip('"')

        # Check if the input is a valid file or directory
        path_obj = Path(directory_input)

        if path_obj.is_file() and path_obj.suffix.lower() == '.pdf':
            # If it's a file, add it directly
            pdf_files.append(str(path_obj))
            print(f"✅ Added file: {path_obj}")

        elif path_obj.is_dir():
            # If it's a directory, collect PDFs in that folder
            pdf_list = list(sorted(path_obj.glob("*.pdf")))
            if pdf_list:
                print(f"\nFound the following PDF files in {directory_input}:")
                for pdf in pdf_list:
                    print(f"  - {pdf.name}")
                
                # Confirm if the user wants to add these files to the merge list
                proceed = input("\nDo you want to add these files to the merge list? (y/n): ").strip().lower()
                if proceed == 'y':
                    pdf_files.extend(str(pdf) for pdf in pdf_list)
                    print(f"✅ Added {len(pdf_list)} files.")
            else:
                print(f"\nNo PDF files found in {directory_input}.")
        else:
            print(f"Invalid file or directory: {directory_input}")

    return pdf_files

def merge_pdfs(pdf_list, output_filename="merged_output.pdf"):
    if not pdf_list:
        print("No PDF files to merge.")
        return

    merger = PdfWriter()

    for pdf in pdf_list:
        print(f"Adding: {pdf}")
        merger.append(pdf)

    merger.write(output_filename)
    merger.close()
    print(f"\n✅ Merged PDF saved as: {output_filename}")

def split_pdf(input_pdf_path):
    # Read the input PDF file
    input_pdf = Path(input_pdf_path)
    
    if not input_pdf.exists() or input_pdf.suffix.lower() != '.pdf':
        print("Please provide a valid PDF file.")
        return

    # Open the input PDF
    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)

    # Create a directory to store the split PDF pages
    output_dir = input_pdf.stem + "_split_pages"
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Loop through each page of the PDF and save it as a separate PDF
    for page_num in range(total_pages):
        writer = PdfWriter()
        writer.add_page(reader.pages[page_num])
        
        # Create a new file for each page
        output_pdf_path = Path(output_dir) / f"{input_pdf.stem}_page_{page_num + 1}.pdf"
        
        # Write the single-page PDF to a file
        with open(output_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)
        
        print(f"✅ Saved: {output_pdf_path}")
    
    print(f"\nAll pages have been split into individual PDFs in the directory: {output_dir}")

def main():
    print(f"Choose an action:\n\n 1. Merge PDFs \n 2. Separate PDFs (Split)")

    # Get user choice for action
    action_choice = input("Enter 1 or 2 to choose an action: ").strip()

    if action_choice == '1':
        # Merge PDFs
        print("You have chosen to merge PDFs.")
        pdfs_to_merge = collect_pdfs_for_merge()
        if pdfs_to_merge:
            print(f"\nReady to merge {len(pdfs_to_merge)} PDF files.")
            output_filename = input("Enter the output filename (default: merged_output.pdf): ").strip() or "merged_output.pdf"
            
            # Automatically add ".pdf" if not included
            if not output_filename.endswith('.pdf'):
                output_filename += '.pdf'
            
            merge_pdfs(pdfs_to_merge, output_filename)
        else:
            print("No files selected for merging.")
    
    elif action_choice == '2':
        # Split PDF
        input_pdf_path = input("Enter the path to the PDF file you want to split: ").strip()
        split_pdf(input_pdf_path)
    
    else:
        print("Invalid choice! Please enter 1 or 2 to choose a valid action.")

if __name__ == "__main__":
    main()
