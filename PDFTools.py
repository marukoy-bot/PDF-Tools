from pdf2docx import Converter
import os

def convert_pdf_to_word(pdf_file_path, word_file_path):
    if not os.path.exists(pdf_file_path):
        print(f"❌ Error: File '{pdf_file_path}' not found.")
        return
    
    try:
        print("🔄 Converting, please wait...")
        cv = Converter(pdf_file_path)
        cv.convert(word_file_path, start=0, end=None)
        cv.close()
        print(f"✅ Conversion successful! Word file saved at: {word_file_path}")
    except Exception as e:
        print(f"❌ Conversion failed: {str(e)}")

def main():
    while True:
        print("\n📎 Drag and drop your PDF file here or paste the full file path:")
        pdf_path = input("PDF File Path: ").strip('"').strip()

        if not pdf_path.lower().endswith(".pdf"):
            print("❌ Invalid file type. Please provide a PDF file.")
            continue

        if not os.path.isfile(pdf_path):
            print("❌ File not found. Please check the path and try again.")
            continue

        # Get default output path
        default_name = os.path.splitext(os.path.basename(pdf_path))[0] + ".docx"
        print(f"💾 Enter output Word filename (leave blank for '{default_name}'):")
        custom_name = input("Output filename (without .docx): ").strip()

        if custom_name:
            output_path = os.path.join(os.path.dirname(pdf_path), f"{custom_name}.docx")
        else:
            output_path = os.path.join(os.path.dirname(pdf_path), default_name)

        # Convert PDF to Word
        convert_pdf_to_word(pdf_path, output_path)

        # Ask to continue or exit
        print("\n🔁 Do you want to convert another file? (Y/n):")
        answer = input().strip().lower()
        if answer == 'n':
            print("👋 Exiting. Goodbye!")
            break

if __name__ == "__main__":
    main()
