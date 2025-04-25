import json
from deep_translator import GoogleTranslator

# Load your input file
input_path = "Cylindrical_Gears_Calculation_Materials_Manufacturing_2016_Linke_backlash_paragraphs.json"
with open(input_path, "r", encoding="utf-8") as file:
    data = json.load(file)

# Translate each paragraph and add new field
for item in data['paragraphs']:
    try:
        translation = GoogleTranslator(source='auto', target='fa').translate(item['full_paragraph'])
        item['translated_full_paragraph'] = translation
    except Exception as e:
        print(f"Error translating paragraph {item.get('id', 'unknown')}: {e}")
        item['translated_full_paragraph'] = "خطا در ترجمه"

# Save updated data to a new file
output_path = "translated_backlash_paragraphs_persian.json"
with open(output_path, "w", encoding="utf-8") as file:
    json.dump(data, file, ensure_ascii=False, indent=2)

print(f"✅ Translated JSON saved to {output_path}")
