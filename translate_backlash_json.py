import json
from langchain.llms import Ollama

def translate_json_file(input_file, output_file):
    # ایجاد نمونه از Ollama با مدل mistral
    llm = Ollama(model="mistral")
    
    # خواندن فایل JSON ورودی
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # ترجمه فیلد paragraph-full برای هر آیتم
    for item in data:
        if 'full_paragraph' in item:
            # ساخت پرامپت برای ترجمه
            prompt = f"Translate this English text to Persian: {item['full_paragraph']}"
            
            # دریافت ترجمه از مدل
            translation = llm.predict(prompt)
            
            # اضافه کردن ترجمه به عنوان فیلد جدید
            item['paragraph-full-fa'] = translation.strip()
    
    # ذخیره نتیجه در فایل JSON جدید
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# استفاده از تابع
input_file = './salmple.json'
output_file = './output_mistral.py'
translate_json_file(input_file, output_file)
