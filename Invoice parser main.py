import re
import imaplib
import email
from email.header import decode_header
from io import BytesIO
import csv
import pytesseract
from pdf2image import convert_from_bytes
import fitz  # PyMuPDF
from invoice2data import extract_data
from invoice2data.extract.loader import read_templates
import camelot
import tempfile
import os

# Define patterns in French for regex matching
patterns = {
    "Montant Dû": r"(Montant\s*total\s*(HT|TTC|T\.T\.C|Net\s*à\s*payer|Total général|Somme\s*à\s*régler|Solde\s*dû)[\s:]*([\d\s,.€]+))",
    "IBAN": r"IBAN\s*:\s*([A-Z]{2}\d{2}(?:\s?\d{4})+)",
    "TVA": r"(TVA|Montant\s*de\s*la\s*TVA|Taux\s*de\s*TVA|Total\s*TVA|Taxe):?\s?([\d\s,.€]+)",
    "Quantité": r"(Quantité|Qté|Nombre\s*d'Unités|Total\s*Produits|Nombre\s*de\s*produits|Nombre\s*d'articles):?\s?(\d+)",
    "Échéance": r"(Échéance|Date\s*d'échéance|Délai\s*de\s*paiement|À\s*payer\s*avant\s*le)\s*:\s*(\d{2}/\d{2}/\d{4})",
}

def login_to_mail(username, password, imap_server="imap.gmail.com"):
    """
    Logs into the specified email account via IMAP.
    """
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(username, password)
        print(f"Successfully logged in as {username}")
        return mail
    except Exception as e:
        print(f"Failed to log in: {e}")
        return None

def extract_with_invoice2data(file_content):

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(file_content)
        temp_pdf_path = temp_pdf.name

    try:

        templates = read_templates('path_to_french_templates')
        result = extract_data(temp_pdf_path, templates=templates)
    finally:

        os.remove(temp_pdf_path)

    return result if result else {}


def extract_data_pymupdf(file_content):
    extracted_data = {}
    pdf_file = BytesIO(file_content)
    with fitz.open(stream=pdf_file, filetype="pdf") as pdf:
        for page_num in range(pdf.page_count):
            page = pdf.load_page(page_num)
            text = page.get_text("text")
            for field, pattern in patterns.items():
                match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                if match:
                    extracted_value = match.group(2) if len(match.groups()) > 1 else match.group()
                    extracted_data[field] = extracted_value
    return extracted_data


def extract_data_ocr(file_content):
    extracted_data = {}
    images = convert_from_bytes(file_content)
    full_text = ""
    for image in images:
        full_text += pytesseract.image_to_string(image, lang="fra")
    for field, pattern in patterns.items():
        match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
        if match:
            extracted_value = match.group(2) if len(match.groups()) > 1 else match.group()
            extracted_data[field] = extracted_value
    return extracted_data


def extract_table_data(file_content):
    pdf_file = BytesIO(file_content)
    tables = camelot.read_pdf(pdf_file, pages='all', flavor='stream')
    table_data = []
    for table in tables:
        table_data.append(table.df.to_dict(orient='records'))
    return table_data


def comprehensive_extraction(file_content):
    extracted_data = extract_with_invoice2data(file_content)
    if not extracted_data:
        extracted_data = extract_data_pymupdf(file_content)
    if not extracted_data:
        extracted_data = extract_data_ocr(file_content)
    return extracted_data

def save_to_csv(data, filename="extracted_data.csv", mode='a'):
    with open(filename, mode=mode, newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(["Field", "Value"])
        for field, value in data.items():
            writer.writerow([field, value])
    print(f"\nData appended to {filename}")

def read_and_extract_data(mail):
    mail.select("inbox")
    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()

    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                from_ = msg.get("From")

                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "application/pdf":
                            filename = part.get_filename()
                            if filename:
                                print(f"\n\nProcessing PDF attachment: {filename}")

                                pdf_content = part.get_payload(decode=True)

                                extracted_data = comprehensive_extraction(pdf_content)

                                if extracted_data:
                                    print(f"\nSubject: {subject}")
                                    print(f"From: {from_}")
                                    print("Detected Data:")
                                    for field, value in extracted_data.items():
                                        print(f"{field}: {value}")
                                    save_to_csv(extracted_data)
                                    print("\nExtraction complete for this PDF.")
                                else:
                                    print(f"No relevant data found in email: {subject}")

if __name__ == "__main__":

    email_user = ""
    email_pass = ""

    mail = login_to_mail(email_user, email_pass)

    if mail:
        read_and_extract_data(mail)
        mail.logout()
