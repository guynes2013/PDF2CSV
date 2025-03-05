import os
import glob
import csv
import re
from datetime import datetime
try:
    from docx import Document
    import pdfplumber
except ImportError:
    print("Libraries missing. Install with: pip install python-docx pdfplumber")
    exit()

# Directory setup
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
base_dir = os.path.join(desktop, "IndexConverter")
input_dir = os.path.join(base_dir, "Input")
completed_dir = os.path.join(base_dir, "Completed")
csv_dir = os.path.join(base_dir, "CSV")
log_file = os.path.join(base_dir, "conversion_log.txt")
readme_file = os.path.join(base_dir, "README.txt")

# Create directories if they don't exist
for dir_path in [base_dir, input_dir, completed_dir, csv_dir]:
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
for file_type in ["docx", "pdf", "txt"]:
    type_dir = os.path.join(completed_dir, file_type)
    if not os.path.exists(type_dir):
        os.makedirs(type_dir)

# Create README if it doesn't exist
if not os.path.exists(readme_file):
    with open(readme_file, "w", encoding="utf-8") as f:
        f.write("""
Index Converter V1.1 Instructions:
1. Save this folder to your Desktop.
2. Place source files (.docx, .pdf, .txt) in the 'Input' folder.
3. Double-click 'index_converter.exe' to run.
4. Select file type (1-3), file number (1-N, 0 for all), or quit (4).
5. After conversion, optionally move files to 'Completed/[file_type]'.
6. Import .csv files from 'CSV' into Google Sheets and adjust formatting (e.g., 'Wrap Text') as needed.
7. For users without Python, download the pre-built bundle from [your source] and unzip to Desktop.
""")

print("Welcome to Index Converter V1.1!")

# File selection and conversion
def get_file_choice(file_type):
    files = glob.glob(os.path.join(input_dir, f"*.{file_type}"))
    if not files:
        print(f"No .{file_type} files found in {input_dir}")
        return None, None
    print("\nAvailable files:")
    for i, f in enumerate(files, 1):
        print(f"{i}. {os.path.basename(f)}")
    print("0. Convert All")
    print("4. Quit")
    choice = input("Select file number, '0' for all, or '4' to quit: ")
    if choice == "4":
        print("Exiting...")
        log_entry("User quit from file selection")
        exit()
    elif choice == "0":
        return files, True  # Convert all
    try:
        file_num = int(choice)
        return [files[file_num - 1]], False if 1 <= file_num <= len(files) else None
    except (ValueError, IndexError):
        print("Invalid selection. Try again.")
        log_entry("Invalid file selection")
        return None, None

# Text extraction
def extract_text(file_path):
    ext = os.path.splitext(file_path)[1].lower().lstrip(".")
    marker = "Index\nNote: The numbers indicate the book number, followed by the page number."
    text = ""
    if ext == "docx":
        doc = Document(file_path)
        full_text = "\n".join([para.text for para in doc.paragraphs])
        text = full_text.split(marker, 1)[1].strip() if marker in full_text else ""
    elif ext == "pdf":
        with pdfplumber.open(file_path) as pdf:
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        text = full_text.split(marker, 1)[1].strip() if marker in full_text else ""
    elif ext == "txt":
        with open(file_path, "r", encoding="utf-8") as f:
            full_text = f.read()
        text = full_text.split(marker, 1)[1].strip() if marker in full_text else ""
    return text.splitlines()

# Parse index with regex
def parse_index(lines):
    data = []
    current_subject = ""
    book_pages = {i: [] for i in range(1, 7)}
    ref_pattern = re.compile(r"(\d+)\s*:\s*(\S+)")  # Matches "Book#:Page" with flexible spacing
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if len(line) == 1 and line.isalpha():
            if current_subject:
                data.append([current_subject] + [format_pages(book_pages[i]) for i in range(1, 7)])
            data.append([line] + [""] * 6)
            current_subject = ""
            book_pages = {i: [] for i in range(1, 7)}
        else:
            parts = line.split("\t", 1) if "\t" in line else line.split(None, 1)
            if len(parts) > 1:
                if not current_subject and not line.endswith(","):
                    current_subject = parts[0].strip()
                refs = parts[1].rstrip(",").split(",")
                for ref in refs:
                    match = ref_pattern.match(ref.strip())
                    if match:
                        book, pages = match.groups()
                        book = int(book)
                        if 1 <= book <= 6:  # Ignore invalid book numbers
                            book_pages[book].append(pages)
            if not line.endswith(",") and current_subject:
                data.append([current_subject] + [format_pages(book_pages[i]) for i in range(1, 7)])
                current_subject = ""
                book_pages = {i: [] for i in range(1, 7)}
    if current_subject:
        data.append([current_subject] + [format_pages(book_pages[i]) for i in range(1, 7)])
    return data

def format_pages(pages):
    if not pages:
        return ""
    return f'"{", ".join(pages)}"'

# Write CSV and log, move file
def write_csv_and_move(data, input_file):
    output_name = os.path.splitext(os.path.basename(input_file))[0] + ".csv"
    output_path = os.path.join(csv_dir, output_name)
    print("Writing CSV...")
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Subject", "Book 1", "Book 2", "Book 3", "Book 4", "Book 5", "Book 6"])
        writer.writerows(data)
    log_entry(f"Processed {os.path.basename(input_file)} (V1.1)")
    print("Done! Note: Import .csv files from 'CSV' into Google Sheets and adjust formatting (e.g., 'Wrap Text') as needed.")

    # Offer to move file to Completed
    file_type = os.path.splitext(input_file)[1].lstrip(".").lower()
    completed_path = os.path.join(completed_dir, file_type, os.path.basename(input_file))
    move = input(f"Move {os.path.basename(input_file)} to Completed/{file_type}? (y/n/r for retry): ").lower()
    if move == 'y':
        try:
            os.rename(input_file, completed_path)
            log_entry(f"Moved {os.path.basename(input_file)} to Completed/{file_type}")
            print(f"Moved {os.path.basename(input_file)} to Completed/{file_type}")
        except Exception as e:
            log_entry(f"Failed to move {os.path.basename(input_file)}: {str(e)}")
            print(f"Failed to move file: {str(e)}")
    elif move == 'r':
        print("Retrying conversion...")
        try:
            lines = extract_text(input_file)
            if not lines:
                raise ValueError("No index data found after marker")
            parsed_data = parse_index(lines)
            write_csv_and_move(parsed_data, input_file)
        except Exception as e:
            log_entry(f"Retry failed for {os.path.basename(input_file)}: {str(e)}")
            print(f"Retry failed: {str(e)}")

def log_entry(message):
    with open(log_file, "a", encoding="utf-8") as f:
        timestamp = datetime.now().strftime("%b/%d %H:%M:%S")
        f.write(f"{timestamp} ----- {message}\n-----\n")

# Main execution
while True:
    print("\nSelect file type or action:")
    print("1. .docx\n2. .pdf\n3. .txt\n4. Quit")
    choice = input("Enter number (1-3) or '4' to quit: ")
    if choice == "4":
        print("Exiting...")
        log_entry("User quit from file type menu (V1.1)")
        exit()
    file_types = {"1": "docx", "2": "pdf", "3": "txt"}
    file_type = file_types.get(choice)
    if not file_type:
        print("Invalid selection. Try again.")
        log_entry("Invalid file type selection (V1.1)")
        continue

    files, convert_all = get_file_choice(file_type)
    if files is None:
        continue

    if convert_all:
        for file in files:
            attempt = 0
            max_attempts = 3
            while attempt < max_attempts:
                try:
                    lines = extract_text(file)
                    if not lines:
                        raise ValueError("No index data found after marker")
                    parsed_data = parse_index(lines)
                    write_csv_and_move(parsed_data, file)
                    break
                except Exception as e:
                    attempt += 1
                    log_entry(f"Attempt {attempt}/{max_attempts} failed for {os.path.basename(file)}: {str(e)}")
                    print(f"Error processing {os.path.basename(file)} (Attempt {attempt}/{max_attempts}): {str(e)}")
                    if attempt == max_attempts:
                        print(f"Max retries reached for {os.path.basename(file)}. Skipping.")
                        break
    else:
        try:
            lines = extract_text(files[0])
            if not lines:
                raise ValueError("No index data found after marker")
            parsed_data = parse_index(lines)
            write_csv_and_move(parsed_data, files[0])
        except Exception as e:
            log_entry(f"Error processing {os.path.basename(files[0])}: {str(e)}")
            print(f"Error processing {os.path.basename(files[0])}: {str(e)}")
            retry = input("Retry conversion? (y/n): ").lower()
            if retry == 'y':
                try:
                    lines = extract_text(files[0])
                    if not lines:
                        raise ValueError("No index data found after marker")
                    parsed_data = parse_index(lines)
                    write_csv_and_move(parsed_data, files[0])
                except Exception as e:
                    log_entry(f"Retry failed for {os.path.basename(files[0])}: {str(e)}")
                    print(f"Retry failed: {str(e)}")