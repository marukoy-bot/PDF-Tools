import os
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter

def get_unique_path(path):
    path = Path(path)
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent

    i = 1
    while True:
        new_name = f"{stem}({i}){suffix}"
        new_path = parent / new_name
        if not new_path.exists():
            return new_path
        i += 1

def convert_pdf_to_word(pdf_path, word_path):
    if not Path(pdf_path).exists():
        print(f"❌ Error: File '{pdf_path}' not found.")
        return

    try:
        print("🔄 Converting, please wait...")
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()
        print(f"✅ Conversion successful! Word file saved at: {word_path}")
    except Exception as e:
        print(f"❌ Conversion failed: {str(e)}")

def prompt_pdf_to_word():
    while True:
        pdf_path = input("📎 Drag and drop your PDF file here or paste the path: ").lstrip("& ").strip("'\"").strip()
        if pdf_path.lower().endswith(".pdf") and Path(pdf_path).is_file():
            break
        print("❌ Invalid PDF file. Please try again.")

    default_name = Path(pdf_path).stem + ".docx"
    custom_name = input(f"💾 Output filename (blank for '{default_name}'): ").strip()
    word_path = get_unique_path(Path(pdf_path).with_name(f"{custom_name or default_name}"))

    convert_pdf_to_word(pdf_path, str(word_path))

def collect_pdfs_for_merge():
    pdf_files = []
    while True:
        path_input = input("📂 Enter folder or drag/drop PDF file(s) (Enter to finish): ").lstrip("& ").strip("'\"").strip()
        if not path_input:
            break

        path = Path(path_input)

        if path.is_file() and path.suffix.lower() == '.pdf':
            pdf_files.append(str(path))
            print(f"✅ Added file: {path.name}")
        elif path.is_dir():
            while True:
                pdf_list = sorted(path.glob("*.pdf"))
                if not pdf_list:
                    print("❌ No PDFs found in directory. Try a different one.")
                    break

                print(f"📄 Found {len(pdf_list)} PDFs in {path}:")
                for pdf in pdf_list:
                    print(f" - {pdf.name}")
                if input("Add all to merge list? (y/n): ").strip().lower() == 'y':
                    pdf_files.extend(str(p) for p in pdf_list)
                    print("✅ Files added.")
                break
        else:
            print("❌ Invalid file or folder. Please try again.")

    return pdf_files

def merge_pdfs(pdf_list, output_filename="merged_output.pdf"):
    if not pdf_list:
        print("❌ No PDF files to merge.")
        return

    writer = PdfWriter()
    for pdf in pdf_list:
        print(f"➕ Adding: {pdf}")
        writer.append(pdf)

    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"✅ Merged PDF saved as: {output_filename}")

def prompt_merge_pdfs():
    pdfs = collect_pdfs_for_merge()
    if pdfs:
        output = input("📄 Output filename (default: merged_output.pdf): ").strip() or "merged_output.pdf"
        if not output.lower().endswith('.pdf'):
            output += '.pdf'
        output_path = get_unique_path(Path(pdfs[0]).parent / output)
        merge_pdfs(pdfs, str(output_path))


def split_pdf(input_path):
    input_pdf = Path(input_path)
    if not input_pdf.exists() or input_pdf.suffix.lower() != '.pdf':
        print("❌ Invalid PDF file.")
        return

    reader = PdfReader(input_pdf)
    output_dir = get_unique_path(input_pdf.parent / (input_pdf.stem + "_split_pages"))
    output_dir.mkdir()

    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        output_path = Path(output_dir) / f"{input_pdf.stem}_page_{i + 1}.pdf"
        with open(output_path, "wb") as f:
            writer.write(f)
        print(f"✅ Saved: {output_path}")

    print(f"\n📂 All pages saved to: {output_dir}")

def prompt_split_pdf():
    while True:
        input_path = input("📎 Enter the path of the PDF to split: ").lstrip("& ").strip("'\"").strip()
        if Path(input_path).is_file() and input_path.lower().endswith(".pdf"):
            break
        print("❌ Invalid PDF file. Please try again.")

    split_pdf(input_path)

def main():
    actions = {
        '1': ("Convert PDF to DOCX", prompt_pdf_to_word),
        '2': ("Merge PDFs", prompt_merge_pdfs),
        '3': ("Split PDF", prompt_split_pdf),
    }

    while True:
        print("📌 Choose an action:")
        for k, (desc, _) in actions.items():
            print(f"{k}. {desc}")

        choice = input("Action: ").strip()
        action = actions.get(choice)
        if action:
            action[1]()
        else:
            print("❌ Invalid choice.")

        again = input("\n🔁 Do you want to perform another task? (y/n): ").strip().lower()
        if again != 'y':
            print("Closing program...")
            break

if __name__ == "__main__":
    main()
