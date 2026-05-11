from pathlib import Path
from pdf2docx import Converter

# Safety setting: skip PDFs larger than this size.
MAX_FILE_SIZE_MB = 100

input_folder = Path("input-pdfs")
output_folder = Path("output-docx")

input_folder.mkdir(exist_ok=True)
output_folder.mkdir(exist_ok=True)


def has_suspicious_filename(file_path: Path) -> bool:
    """
    Check for filename characters that could make file handling confusing or unsafe.
    Since we only use file_path.name, this checks the filename itself, not the full path.
    """
    filename = file_path.name

    suspicious_characters = ["\n", "\r", "/", "\\"]

    return any(char in filename for char in suspicious_characters)


def get_file_size_mb(file_path: Path) -> float:
    """Return the file size in megabytes."""
    return file_path.stat().st_size / (1024 * 1024)


print()
print("================================")
print("   PDF to DOCX Folder Converter")
print("================================")
print()
print(f"Input folder:  {input_folder}")
print(f"Output folder: {output_folder}")
print(f"Max PDF size:  {MAX_FILE_SIZE_MB} MB")
print()

pdf_files = list(input_folder.glob("*.pdf"))

successful_count = 0
failed_count = 0
skipped_too_large_count = 0
skipped_exists_count = 0
skipped_suspicious_count = 0

if not pdf_files:
    print("No PDF files found in input-pdfs.")
    print("Put one or more PDFs in that folder, then run this script again.")
else:
    total_files = len(pdf_files)
    print(f"Found {total_files} PDF file(s).")
    print()

    for index, pdf_file in enumerate(pdf_files, start=1):
        output_file = output_folder / f"{pdf_file.stem}.docx"

        print(f"Converting {index} of {total_files}: {pdf_file.name}")

        # Safety check 1: reject suspicious filenames.
        if has_suspicious_filename(pdf_file):
            print(f"Skipped: suspicious filename: {pdf_file.name}")
            skipped_suspicious_count += 1
            print()
            continue

        # Safety check 2: skip very large PDFs.
        file_size_mb = get_file_size_mb(pdf_file)

        if file_size_mb > MAX_FILE_SIZE_MB:
            print(
                f"Skipped: file is too large "
                f"({file_size_mb:.2f} MB, limit is {MAX_FILE_SIZE_MB} MB)."
            )
            skipped_too_large_count += 1
            print()
            continue

        # Safety check 3: do not overwrite an existing DOCX.
        if output_file.exists():
            print(f"Skipped: output file already exists: {output_file}")
            skipped_exists_count += 1
            print()
            continue

        try:
            converter = Converter(str(pdf_file))
            converter.convert(str(output_file))
            converter.close()

            print(f"Saved: {output_file}")
            successful_count += 1

        except Exception as e:
            print(f"Failed to convert {pdf_file.name}: {e}")
            failed_count += 1

        print()

print("Finished.")
print(f"Successful conversions:        {successful_count}")
print(f"Failed conversions:            {failed_count}")
print(f"Skipped - too large:           {skipped_too_large_count}")
print(f"Skipped - output already exists: {skipped_exists_count}")
print(f"Skipped - suspicious filename: {skipped_suspicious_count}")