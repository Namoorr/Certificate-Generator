# **Certificate Generator**  

This script streamlines certificate generation by using a **Word template (`.docx`)** and **student data from an Excel file (`.xlsx`)**. It automatically fills in the required details, creates personalized certificates in `.docx` format, and then converts them into `.pdf` for easy distribution. Designed for efficiency, this tool simplifies bulk certificate processing while ensuring accuracy.  

## **Requirements**  
- **Python 3.x**  
- **Required Python packages:**  
  Install all dependencies with:  
  ```sh
  pip install docxtpl pandas docx2pdf
  ```  
  - `docxtpl` → For filling in the Word template with student details.  
  - `pandas` → For reading student data from an Excel file.  
  - `docx2pdf` → For converting `.docx` files into `.pdf` format.  

## **Files Required**  
- **`certificate_template.docx`** → Word template with placeholders.  
- **`student_data1.xlsx`** → Excel file containing student details with the following columns:  
  - **Name**  
  - **Enrollment**  
  - **College**  
  - **Date_to** (DD-MM-YYYY format)  
  - **Date_from** (DD-MM-YYYY format)  
  - **Batch**  
  - **sr**  

## **Usage**  
1. Ensure all required files are in the same directory as the script.  
2. Run the script:  
   ```sh
   python certificate_generator.py
   ```  
3. Generated certificates will be saved in **`.docx` format** and automatically converted to **`.pdf`**.  

## **Output**  
- Certificates are generated in both **DOCX** and **PDF** formats, named as:  
  ```
  Batch-sr.docx
  Batch-sr.pdf
  ```
  **Example:** `B1-001.docx` → `B1-001.pdf`  

## **Errors & Troubleshooting**  
- **File Not Found?**  
  - Ensure `certificate_template.docx` and `student_data1.xlsx` exist in the directory.  
- **Wrong Date Format?**  
  - Ensure `Date_to` and `Date_from` are in `DD-MM-YYYY` format.  
- **PDF Conversion Fails?**  
  - Try manually converting `.docx` files using **MS Word** or **LibreOffice**.  

## **Contribution**  
Feel free to **fork**, **improve**, and **contribute** to this project! 🚀  
