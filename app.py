import os
import re
import docx
from PyPDF2 import PdfReader
import glob

def clean_text(text):
    """
    Cleans text by removing NULL bytes, control characters, and normalizing whitespace.
    
    Args:
        text (str): Text to clean
        
    Returns:
        str: Cleaned text
    """
    # Remove NULL bytes and control characters
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]', ' ', text)
    # Normalize whitespace
    cleaned = ' '.join(cleaned.split())
    return cleaned.strip()

def extract_backlash_content(pdf_directory):
    """
    Scans a folder of PDF files and extracts content related to "backlash"
    
    Args:
        pdf_directory (str): Path to the directory containing PDF files
        
    Returns:
        dict: Dictionary of results with filenames as keys and lists of relevant text as values
    """
    results = {}
    
    # Get all PDF files in the specified directory
    pdf_files = glob.glob(os.path.join(pdf_directory, "*.pdf"))
    
    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        print(f"Processing {filename}...")
        
        try:
            reader = PdfReader(pdf_path)
            text_content = ""
            
            # Extract text from all pages
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:  # Only add if text was extracted
                    text_content += clean_text(page_text) + "\n"
                
            # Find all paragraphs containing the word "backlash"
            # Paragraphs are separated by blank lines
            paragraphs = re.split(r'\n\s*\n', text_content)
            
            backlash_paragraphs = []
            for paragraph in paragraphs:
                if re.search(r'\bbacklash\b', paragraph, re.IGNORECASE):
                    # Clean and add the paragraph
                    clean_paragraph = clean_text(paragraph)
                    if clean_paragraph:  # Only add if there's content
                        backlash_paragraphs.append(clean_paragraph)
            
            if backlash_paragraphs:
                results[filename] = backlash_paragraphs
                
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
    
    return results

def save_to_word(results, output_file):
    """
    Saves the results to a Word document
    
    Args:
        results (dict): Dictionary containing the extracted content
        output_file (str): Path to the output Word file
    """
    doc = docx.Document()
    
    # Add a title
    doc.add_heading('Backlash Content Extraction Results', 0)
    
    # If no results were found
    if not results:
        doc.add_paragraph('No content related to "backlash" was found in any of the PDF files.')
    
    # Add the results from each file
    for filename, paragraphs in results.items():
        # Add a heading for the file
        doc.add_heading(f'File: {filename}', level=1)
        
        # Add each paragraph containing "backlash"
        for i, paragraph in enumerate(paragraphs, 1):
            doc.add_heading(f'Match {i}:', level=2)
            
            # Create a new paragraph for the content
            p = doc.add_paragraph()
            
            # Split the paragraph to highlight "backlash" occurrences
            parts = re.split(r'(\bbacklash\b)', paragraph, flags=re.IGNORECASE)
            
            for j, part in enumerate(parts):
                if part:  # Only add if part is not empty
                    if j % 2 == 1:  # This is a "backlash" match
                        p.add_run(part).bold = True
                    else:
                        p.add_run(part)
            
            # Add a separator between matches
            if i < len(paragraphs):
                doc.add_paragraph('-' * 50)
    
    # Save the document
    try:
        doc.save(output_file)
        print(f"Results successfully saved to {output_file}")
    except Exception as e:
        print(f"Error saving Word document: {str(e)}")

def main():
    # Get input from user
    pdf_directory = r"D:\extract-content-from-file\inputs"
    output_file = r"D:\extract-content-from-file\outputs\output.docx"
    
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Validate directory
    if not os.path.isdir(pdf_directory):
        print(f"Error: Directory '{pdf_directory}' does not exist.")
        return
    
    # Extract content from PDFs
    print("Starting extraction process...")
    results = extract_backlash_content(pdf_directory)
    
    # Report results
    total_files = len(glob.glob(os.path.join(pdf_directory, "*.pdf")))
    files_with_matches = len(results)
    
    print(f"\nExtraction complete!")
    print(f"Examined {total_files} PDF files")
    print(f"Found matches in {files_with_matches} files")
    
    # Save results to Word document
    if results:
        save_to_word(results, output_file)
    else:
        print("No content related to 'backlash' was found in any of the PDF files.")

if __name__ == "__main__":
    main()