import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
from PyPDF2.generic import NameObject, TextStringObject
from openpyxl import Workbook


def extract_pdf_data(pdf_path):
    """Extract text from a PDF file."""
    pdf_data = {}
    try:
        with open(pdf_path, "rb") as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            for page_num, page in enumerate(reader.pages):
                pdf_data[f"Page {page_num + 1}"] = page.extract_text()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read PDF: {e}")
    return pdf_data


def write_to_spreadsheet(data, excel_path):
    """Write extracted data to an Excel spreadsheet."""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "PDF Data"

        ws.append(["Page", "Content"])
        for page, content in data.items():
            ws.append([page, content])

        wb.save(excel_path)
        messagebox.showinfo("Success", "Data successfully written to spreadsheet.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write to spreadsheet: {e}")


def extract_form_data(pdf_path):
    """Extract form data from a PDF form."""
    form_data = {}
    try:
        with open(pdf_path, "rb") as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            if "/AcroForm" in reader.trailer["/Root"]:
                fields = reader.trailer["/Root"]["/AcroForm"]["/Fields"]
                for field in fields:
                    field_data = field.get_object()
                    key = field_data.get(NameObject("/T"), "")
                    value = field_data.get(NameObject("/V"), "")
                    form_data[key] = value
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract form data: {e}")
    return form_data


def write_to_pdf_form(form_data, output_path, template_pdf_path, field_mapping):
    """Write data to a target PDF form using a mapping between source and target fields."""
    try:
        with open(template_pdf_path, "rb") as template_file:
            reader = PyPDF2.PdfReader(template_file)
            writer = PyPDF2.PdfWriter()

            # Copy all pages from the template
            for page in reader.pages:
                writer.add_page(page)

            # Apply field mapping
            for source_field, value in form_data.items():
                if source_field in field_mapping:  # Check if the source field has a target mapping
                    target_field = field_mapping[source_field]  # Get the corresponding target field
                    for page in writer.pages:
                        if "/Annots" in page:
                            for annotation in page["/Annots"]:
                                annot_object = annotation.get_object()
                                if annot_object.get(NameObject("/T")) == TextStringObject(target_field):
                                    annot_object.update({NameObject("/V"): TextStringObject(value)})

            # Save the new PDF with updated form fields
            with open(output_path, "wb") as output_file:
                writer.write(output_file)
            messagebox.showinfo("Success", "Data written to new PDF form with field mapping.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write to PDF form: {e}")


def open_pdf_and_process():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        data = extract_pdf_data(pdf_path)
        if data:
            excel_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
            )
            if excel_path:
                write_to_spreadsheet(data, excel_path)


def open_pdf_form_and_transfer():
    source_pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if source_pdf_path:
        form_data = extract_form_data(source_pdf_path)
        if form_data:
            template_pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if template_pdf_path:
                output_pdf_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
                )
                if output_pdf_path:
                    # Define the field mapping
                    field_mapping = {
                        "BrandName": "brand",
                        "StrainType": "strain",
                        "NetWeight": "weight",
                        "CaseSize": "case size",
                        "Dimensions": "case dimensions",
                        "TotalUnits": "units",
                        "PackageDate": "packaged on date",
                    }
                    write_to_pdf_form(form_data, output_pdf_path, template_pdf_path, field_mapping)


# GUI
root = tk.Tk()
root.title("PDF Data Processor")

frame = tk.Frame(root, padx=80, pady=10)
frame.pack()

extract_button = tk.Button(frame, text="Extract PDF to Spreadsheet", command=open_pdf_and_process)
extract_button.grid(row=0, column=0, padx=5, pady=5)

form_button = tk.Button(frame, text="Transfer PDF Form Data", command=open_pdf_form_and_transfer)
form_button.grid(row=1, column=0, padx=5, pady=5)

exit_button = tk.Button(frame, text="Exit", command=root.quit)
exit_button.grid(row=2, column=0, padx=5, pady=5)

root.mainloop()
