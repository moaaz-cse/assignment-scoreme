import fitz  # PyMuPDF
import pandas as pd

def extract_text_by_lines(pdf_path):
    """
    Extracts text from the PDF while maintaining line structure.
    Groups text by Y-coordinates to maintain the correct reading order.
    """
    doc = fitz.open(pdf_path)
    extracted_lines = []

    for page in doc:
        words = page.get_text("words")  # Extract words with their positions
        words.sort(key=lambda w: (w[1], w[0]))  # Sort by Y (row), then X (column)

        line_dict = {}
        for w in words:
            x, y, text = w[0], w[1], w[4]  # Extract coordinates & text

            # Group words based on similar Y-coordinates (same line)
            if y not in line_dict:
                line_dict[y] = []
            line_dict[y].append(text)

        # Sort and store lines in extracted_lines
        for y in sorted(line_dict.keys()):
            extracted_lines.append(" ".join(line_dict[y]))  # Merge words in a row

    return extracted_lines

def structure_data(lines):
    """
    Cleans and structures extracted data into a proper table format.
    """
    structured_data = []
    for line in lines:
        # Splitting key-value pairs correctly
        if ":" in line:
            parts = line.split(":")
            if len(parts) > 2:  
                # Handling multiple colons by merging later parts
                structured_data.append([parts[0].strip(), ":".join(parts[1:]).strip()])
            else:
                structured_data.append([parts[0].strip(), parts[1].strip()])
        else:
            structured_data.append([line.strip()])  # Single-column data

    return structured_data

def save_to_excel(data, output_path):
    """
    Saves structured data into an Excel sheet with proper formatting.
    """
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)

def process_pdf(pdf_path, output_path):
    """
    Full pipeline to extract structured data from PDF and save it as an Excel file.
    """
    print(f"Processing PDF: {pdf_path}")

    # Step 1: Extract text while preserving line structure
    extracted_lines = extract_text_by_lines(pdf_path)

    # Step 2: Structure data into key-value pairs and tables
    structured_data = structure_data(extracted_lines)

    # Step 3: Save cleaned data into an Excel file
    save_to_excel(structured_data, output_path)

    print(f"✅ Data successfully saved to: {output_path}")

# Example usage:
pdf_path = "/test3 (1).pdf"
output_path = "/output.xlsx"
process_pdf(pdf_path, output_path)
