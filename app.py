# -*- coding: utf-8 -*-
import PyPDF2
from docx import Document
import json
import re
import os
import sys
from typing import List, Dict

# Fix console encoding for Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def clean_text(text: str) -> str:
    """Clean text from invalid XML characters and normalize spaces"""
    if not text:
        return ""
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    text = text.replace('\x0b', ' ').replace('\x0c', ' ')
    text = text.replace('\u0643', '\u06a9')  # ك to ک
    text = text.replace('\u0649', '\u06cc')  # ى to ی
    return text.strip()

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from PDF file while preserving paragraph structure"""
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += clean_text(page_text) + "\n\n"
    except Exception as e:
        print(f"Error reading PDF file: {str(e)}")
    return text

def split_into_paragraphs(text: str) -> List[str]:
    """Group lines into paragraphs separated by empty lines"""
    lines = text.splitlines()
    paragraphs = []
    current = []

    for line in lines:
        stripped = line.strip()
        if stripped == '':
            if current:
                paragraphs.append(' '.join(current))
                current = []
        else:
            current.append(stripped)

    if current:
        paragraphs.append(' '.join(current))

    return paragraphs

def find_backlash_paragraphs(text: str) -> List[Dict]:
    """Find all paragraphs containing the word 'backlash'"""
    paragraphs = split_into_paragraphs(text)

    results = []
    for i, para in enumerate(paragraphs, 1):
        if re.search(r'\bbacklash\b', para, re.IGNORECASE):
            sentences = re.split(r'(?<=[.!?]) +', para)
            relevant_sentences = [clean_text(s) for s in sentences 
                                  if re.search(r'\bbacklash\b', s, re.IGNORECASE)]
            results.append({
                "id": i,
                "paragraph": clean_text(para),
                "sentences": relevant_sentences,
                "full_paragraph": clean_text(para)
            })
    return results

def save_to_word(occurrences: List[Dict], output_path: str):
    doc = Document()
    doc.add_heading('گزارش پاراگراف‌های حاوی کلمه Backlash', 0)

    if not occurrences:
        doc.add_paragraph("هیچ پاراگرافی حاوی کلمه 'backlash' در سند پیدا نشد.")
        doc.save(output_path)
        return

    doc.add_paragraph(f"تعداد {len(occurrences)} پاراگراف حاوی کلمه 'backlash' پیدا شد:\n")

    for item in occurrences:
        try:
            doc.add_paragraph(f"پاراگراف شماره #{item['id']}:", style='Heading 2')
            doc.add_paragraph(item['paragraph'])
            doc.add_paragraph("جملات مرتبط:", style='Heading 3')
            for sent in item['sentences']:
                doc.add_paragraph(f"- {sent}")
            doc.add_paragraph()
        except Exception as e:
            print(f"Error saving paragraph {item['id']}: {str(e)}")
            continue

    try:
        doc.save(output_path)
    except Exception as e:
        print(f"Error saving Word file: {str(e)}")
        raise

def save_to_json(occurrences: List[Dict], output_path: str):
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({
                "search_term": "backlash",
                "total_paragraphs": len(occurrences),
                "paragraphs": [ {
                    "id": item["id"],
                    "full_paragraph": item["full_paragraph"],
                    "sentences": item["sentences"]
                } for item in occurrences ]
            }, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving JSON file: {str(e)}")
        raise

def main():
    pdf_path = "./inputs/Cylindrical_Gears_Calculation_Materials_Manufacturing_2016_Linke.pdf"

    if not os.path.exists(pdf_path):
        print("Error: PDF file not found.")
        return

    base_dir = os.path.dirname(pdf_path)
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    word_output = os.path.join(base_dir, f"{file_name}_backlash_paragraphs.docx")
    json_output = os.path.join(base_dir, f"{file_name}_backlash_paragraphs.json")

    try:
        print("Extracting text from PDF...")
        text = extract_text_from_pdf(pdf_path)

        print("Searching for 'backlash'...")
        occurrences = find_backlash_paragraphs(text)

        print("Saving results...")
        save_to_word(occurrences, word_output)
        save_to_json(occurrences, json_output)

        print("\nProcessing completed successfully:")
        print(f"- Found {len(occurrences)} paragraphs containing 'backlash'")
        print(f"- Word report saved to: {word_output}")
        print(f"- JSON report saved to: {json_output}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
