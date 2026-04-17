import os
import shutil
import random
import datetime
from datetime import timedelta
import pandas as pd
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import fitz  # PyMuPDF

# --- UTILITY FUNCTIONS ---

MONTHS_EN = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def format_date_time(dt):
    """Formats datetime objects into the specific string format required for the Audit Trail."""
    day = dt.day
    month = MONTHS_EN[dt.month]
    year = dt.year
    hour = dt.strftime("%I:%M %p").lstrip("0")
    return f"{day}-{month}-{year} {hour} EST"

def generate_audit_times(signature_date_str):
    """Generates realistic, randomized timestamps for the document audit trail."""
    try:
        base_dt = datetime.datetime.strptime(signature_date_str, "%m/%d/%Y")
    except ValueError:
        print("❌ Error: Invalid date format. Use mm/dd/yyyy.")
        return None
    
    # Randomize start time between 2 PM and 5 PM
    start_hour = random.randint(14, 17)
    start_minute = random.randint(0, 59)
    
    t1 = base_dt.replace(hour=start_hour, minute=start_minute)
    t2 = t1 + timedelta(minutes=random.randint(15, 25))  # Time spent viewing
    t3 = t2 + timedelta(minutes=random.randint(2, 5))    # Time spent signing
    
    return {
        "start": format_date_time(t1),
        "viewed": format_date_time(t2),
        "signed": format_date_time(t3),
        "completed": format_date_time(t3)
    }

def run_automation():
    print("=== MEDICAL RECORDS REQUEST AUTOMATION SYSTEM ===")
    
    # 1. Setup paths
    root_dir = os.getcwd()
    files_in_dir = os.listdir(root_dir)
    
    # Auto-detect necessary files
    signed_pdf = next((f for f in files_in_dir if f.lower().endswith('.pdf')), None)
    word_template = next((f for f in files_in_dir if f.lower().endswith('.docx')), None)
    excel_db = "Medical_request_facilities.xlsx"

    if not all([signed_pdf, word_template, os.path.exists(excel_db)]):
        print("❌ Error: Missing required files (.pdf HIPAA, .docx template, or Excel database).")
        return

    # 2. User Input for Case Specifics
    client_name = input("Enter Client Name: ").strip()
    p1 = int(input("First HIPAA page to extract (usually 1): "))
    p2 = int(input("Second HIPAA page to extract (usually 2): "))
    start_date = input("Service Start Date (mm/dd/yyyy): ")
    end_date = input("Service End Date (mm/dd/yyyy): ")
    sign_date = input("Signature Date (mm/dd/yyyy): ")
    client_email = input("Client Email: ")

    # Create Output Folder
    client_folder = os.path.join(root_dir, client_name)
    os.makedirs(client_folder, exist_ok=True)
    
    # 3. Process Master HIPAA PDF
    print(f"\nProcessing HIPAA Authorization from: {signed_pdf}...")
    audit_times = generate_audit_times(sign_date)
    
    # Extract specific pages
    reader = PdfReader(signed_pdf)
    writer = PdfWriter()
    writer.add_page(reader.get_page(p1 - 1))
    writer.add_page(reader.get_page(p2 - 1))
    
    temp_hipaa = "temp_extracted.pdf"
    with open(temp_hipaa, "wb") as f:
        writer.write(f)

    # Edit PDF Text (Audit Trail and Dates)
    doc = fitz.open(temp_hipaa)
    page_audit = doc[1] # Usually the second page contains the audit trail
    
    # coordinates based on standard HIPAA forms
    page_audit.insert_text((72, 442), audit_times["start"], fontsize=8, color=(0,0,0))
    page_audit.insert_text((72, 471), audit_times["viewed"], fontsize=8, color=(0,0,0))
    page_audit.insert_text((72, 501), audit_times["signed"], fontsize=8, color=(0,0,0))
    page_audit.insert_text((72, 530), audit_times["completed"], fontsize=8, color=(0,0,0))
    page_audit.insert_text((315, 501), client_email, fontsize=8, color=(0,0,0))
    
    master_hipaa_path = os.path.join(client_folder, "Master_HIPAA.pdf")
    doc.save(master_hipaa_path)
    doc.close()
    os.remove(temp_hipaa)

    # 4. Loop Through Facilities in Excel
    print("\nReading facility database and generating requests...")
    df = pd.read_excel(excel_db)
    
    for index, row in df.iterrows():
        facility = str(row['facility']).strip()
        print(f"🔄 Processing: {facility}")
        
        # Create facility subfolder
        facility_dir = os.path.join(client_folder, facility)
        os.makedirs(facility_dir, exist_ok=True)
        
        # Edit Word Template
        doc_word = Document(word_template)
        address_full = f"{row['address1']}, {row['address2']}" if pd.notna(row['address2']) else str(row['address1'])
        
        replacements = {
            "[FACILITY_NAME]": facility,
            "[ADDRESS]": address_full,
            "[OTHERS]": str(row['others']) if pd.notna(row['others']) else "",
            "[CLIENT_NAME]": client_name,
            "[START_DATE]": start_date,
            "[END_DATE]": end_date
        }

        for paragraph in doc_word.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)

        temp_word = os.path.join(facility_dir, "temp_request.docx")
        doc_word.save(temp_word)
        
        # Convert Word to PDF
        temp_pdf = os.path.join(facility_dir, "request_letter.pdf")
        convert(temp_word, temp_pdf)
        
        # Merge Letter + HIPAA
        final_filename = f"{datetime.datetime.now().strftime('%Y.%m.%d')} MR-MB Request {facility}.pdf"
        final_path = os.path.join(facility_dir, final_filename)
        
        merger = PdfMerger()
        merger.append(temp_pdf)
        merger.append(master_hipaa_path)
        merger.write(final_path)
        merger.close()
        
        # Cleanup temp facility files
        os.remove(temp_word)
        os.remove(temp_pdf)

    print(f"\n✨ DONE! All requests for {client_name} generated in the folder.")

if __name__ == "__main__":
    run_automation()
