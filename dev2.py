import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog, messagebox
from datetime import datetime

def extract_text_from_pdf():
    """Extract text from a selected PDF and save it to a uniquely named Excel file."""
    root = Tk()
    root.withdraw()  # Hide the main window

    # Select the PDF file
    pdf_file_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not pdf_file_path:
        messagebox.showinfo("Error", "No PDF selected!")
        return

    # Create a unique filename using timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel_path = pdf_file_path.replace(".pdf", f"_{timestamp}_Extracted_Text.xlsx")

    extracted_data = []
    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    extracted_data.append({"Page": page_number, "Extracted Text": text.strip()})

        # Save extracted text to Excel
        if extracted_data:
            df = pd.DataFrame(extracted_data)
            df.to_excel(output_excel_path, index=False, engine="openpyxl")
            messagebox.showinfo("Success", f"Text extraction complete!\nSaved at: {output_excel_path}")
        else:
            messagebox.showinfo("Error", "No text found in the PDF!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract text: {e}")

if __name__ == "__main__":
    extract_text_from_pdf()
