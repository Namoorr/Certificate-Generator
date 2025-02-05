from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert

# Load data from an XLSX file
try:
    ans = pd.read_excel('student_data1.xlsx', dtype={
        'Name': str, 'Enrollment': str, 'College': str,
        'Date_to': str, 'Date_from': str, 'Batch': str, 'sr': str
    })
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

def generate_certificate(student_data):
    try:
        tpl = DocxTemplate("certificate_template.docx")
        context = {
            "Name": student_data['Name'],
            "Enrollment": student_data['Enrollment'],
            "College": student_data['College'],
            "Date_to": student_data['Date_to'],
            "Date_from": student_data['Date_from']
        }
        output_docx_filename = f"{student_data['Batch']}-{student_data['sr']}.docx"
        tpl.render(context)
        tpl.save(output_docx_filename)
        convert_to_pdf(output_docx_filename)
    except Exception as e:
        print(f"Error generating certificate for {student_data['Name']}: {e}")

def convert_to_pdf(docx_filename):
    try:
        convert(docx_filename)
        print(f"Converted {docx_filename} to PDF successfully.")
    except Exception as e:
        print(f"Failed to convert {docx_filename} to PDF: {e}")

# Iterate through the rows in the DataFrame
for index, row in ans.iterrows():
    generate_certificate(row)

print("Certificate generation completed.")
