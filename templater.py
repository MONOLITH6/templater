import os
import shutil
from datetime import datetime
from openpyxl import Workbook, load_workbook
import subprocess
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter

def main():
    # Base paths
    base_path = os.path.abspath(os.path.dirname(__file__))
    template_path = os.path.join(base_path, "template")
    jobs_path = os.path.join(base_path, "jobs")
    excel_file = os.path.join(base_path, "applications.xlsx")

    # Dynamically load template folders
    template_folders = sorted([
        name for name in os.listdir(template_path)
        if os.path.isdir(os.path.join(template_path, name))
    ])

    template_options = {
        str(i+1): (folder, os.path.join(template_path, folder))
        for i, folder in enumerate(template_folders)
    }

    # Prompt for company name (required)
    while True:
        company = input("Enter the company name: ").strip()
        if company:
            company = company.title()
            break
        else:
            print("[✗] Company name is required. Please enter a valid name.")

    # Excel setup
    columns = ["Company", "Role", "Status", "Link", "Date", "Date of Last Contact", "Other"]
    if not os.path.exists(excel_file):
        print("[+] Excel file does not exist. Creating...")
        wb = Workbook()
        ws = wb.active
        ws.append(columns)
        wb.save(excel_file)
        print("[✓] Excel file created.")

    # Load Excel and check for duplicates (case-insensitive)
    wb = load_workbook(excel_file)
    ws = wb.active

    duplicate_entries = [
        row for row in ws.iter_rows(min_row=2, values_only=True)
        if row[0] and row[0].strip().lower() == company.lower()
    ]

    if duplicate_entries:
        print(f"[!] Found {len(duplicate_entries)} existing entries for '{company}':")
        for i, entry in enumerate(duplicate_entries, 1):
            role = entry[1]
            date = entry[4]
            link = entry[3]
            print(f"  {i}. Role: {role} | Date: {date} | Link: {link}")
        confirm = input("\n[?] Do you still want to add another entry for this company? (y/n): ").lower()
        if confirm != 'y':
            print("[✗] Aborted. No changes made.")
            exit()

    # Continue prompts
    role = input("Enter the role you are applying for: ").strip()
    link_url = input("Enter the job application link: ").strip()
    other = input("Other notes (optional): ").strip()
    today = datetime.now().strftime("%Y-%m-%d")

    # Template selection
    print("\nSelect a template category:")
    for key, (name, _) in template_options.items():
        print(f"{key}. {name}")
    template_choice = input("Enter the number corresponding to your template: ").strip()

    if template_choice not in template_options:
        print("[✗] Invalid template selection.")
        exit()

    template_name, selected_template_path = template_options[template_choice]

    # Template file paths
    template_resume = os.path.join(selected_template_path, "Resume.docx")
    template_cover_letter = os.path.join(selected_template_path, "Cover_Letter.docx")

    # Create output folder
    company_folder = os.path.join(jobs_path, company)
    os.makedirs(company_folder, exist_ok=True)

    # Destination file paths
    resume_dest = os.path.join(company_folder, "Resume.docx")
    cover_letter_dest = os.path.join(company_folder, "Cover_Letter.docx")

    # Copy templates
    try:
        shutil.copy(template_resume, resume_dest)
        shutil.copy(template_cover_letter, cover_letter_dest)
        print(f"[✓] Templates copied from '{template_name}' to: {company_folder}")
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        exit()

    # Append entry to Excel
    ws.append([company, role, "Applied", link_url, today, "", other])
    wb.save(excel_file)
    print(f"[✓] Application entry for '{company}' added to Excel.")

    # Open RESUME for editing
    print("[+] Opening resume for redaction...")
    subprocess.Popen(["start", "", resume_dest], shell=True)
    input("[?] Press Enter when you're done redacting and have saved the resume...")

    # # Convert RESUME to PDF and trim
    # resume_pdf_path = resume_dest.replace(".docx", ".pdf")
    # try:
    #     convert(resume_dest, resume_pdf_path)
    #     print("[✓] Resume successfully converted to PDF.")

    #     # Trim to one page
    #     reader = PdfReader(resume_pdf_path)
    #     if len(reader.pages) > 1:
    #         writer = PdfWriter()
    #         writer.add_page(reader.pages[0])
    #         with open(resume_pdf_path, "wb") as f_out:
    #             writer.write(f_out)
    #         print("[✓] Trimmed resume PDF to only the first page.")
    # except Exception as e:
    #     print(f"[!] Resume PDF conversion or trimming failed: {e}")

    # Cover Letter
    generate_cl = input("\n[?] Do you want to generate a cover letter as well? (y/n): ").lower()
    if generate_cl == 'y':
        print("[+] Opening cover letter for redaction...")
        subprocess.Popen(["start", "", cover_letter_dest], shell=True)
        input("[?] Press Enter when you're done redacting and have saved the cover letter...")

        # Convert COVER LETTER to PDF
        cover_pdf_path = cover_letter_dest.replace(".docx", ".pdf")
        try:
            convert(cover_letter_dest, cover_pdf_path)
            print("[✓] Cover letter successfully converted to PDF.")
        except Exception as e:
            print(f"[!] Cover letter PDF conversion failed: {e}")
    else:
        print("[✓] Skipped cover letter generation.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[✗] Aborted by user (Ctrl+C).")
        exit()
