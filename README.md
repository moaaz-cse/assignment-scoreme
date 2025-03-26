PDF to Excel Extractor 
📌 Overview 
This Python script extracts structured text from a PDF file and saves it as a well-formatted Excel sheet. It ensures correct text alignment by grouping words based on their positions in the document.

🚀 Features ✅ Extracts text from PDF while maintaining structure ✅ Groups key-value pairs correctly ✅ Handles multi-line values properly ✅ Saves data in a clean, tabular format in Excel

📂 Installation Ensure you have the required dependencies installed: pip install pymupdf pandas openpyxl

📜 Usage 1️⃣ Place your PDF file in the same directory as the script. 2️⃣ Modify the paths in the script if needed. 3️⃣ Run the script: python pdf_to_excel.py

🔧 Example Code pdf_path = "test3.pdf" output_path = "output.xlsx" process_pdf(pdf_path, output_path)

📌 Notes Supports PDFs with structured text (not scanned images). If text extraction fails, ensure your PDF is text-based, not an image.
