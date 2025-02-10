import os
import csv
from datetime import datetime
from xml.sax.saxutils import escape

def get_current_datetime():
    return datetime.now().strftime("%Y%m%dT%H%M%S")

def generate_tmx_header():
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<tmx version="1.4">
  <header srclang="en_US" datatype="plaintext"/>
  <body>"""

def generate_tmx_footer():
    return """
  </body>
</tmx>"""

def process_csv_file(file_path, tmx_file_path, key_id, source_lang, target_lang):
    try:
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            rows = list(reader)
            total_rows = len(rows)
            
            if total_rows == 0:
                print(f"Skipping empty file: {file_path}")
                return
            
            current_datetime = get_current_datetime()
            
            with open(tmx_file_path, "w", encoding="utf-8") as tmxfile:
                tmxfile.write(generate_tmx_header(source_lang))
                
                for i, row in enumerate(rows):
                    key = escape(row.get(key_id, "N/A"))
                    source_text = escape(row.get(source_lang, "N/A"))
                    target_text = escape(row.get(target_lang, "N/A"))
                    prev_source = escape(rows[i - 1].get(source_lang, "N/A")) if i > 0 else ""
                    next_source = escape(rows[i + 1].get(source_lang, "N/A")) if i < total_rows - 1 else ""
                    
                    tmxfile.write(f"""
  <tu>
    <prop type="x-segment-id">{key}</prop>
    <prop type="x-previous-source-text">{prev_source}</prop>
    <prop type="x-next-source-text">{next_source}</prop>
    <tuv xml:lang="{source_lang}" creationdate="{current_datetime}" lastusagedate="{current_datetime}">
      <seg>{source_text}</seg>
    </tuv>
    <tuv xml:lang="{target_lang}" creationdate="{current_datetime}" lastusagedate="{current_datetime}">
      <seg>{target_text}</seg>
    </tuv>
  </tu>""")
                
                tmxfile.write(generate_tmx_footer())
                
            print(f"Finished writing: {tmx_file_path}")
    except Exception as e:
        print(f"Error processing {file_path}: {e}")

def main():
    key_id = "key"
    source_lang = "ja_JP"
    target_lang = "en_US"
    
    for root, _, files in os.walk("."):
        for file in files:
            if file.endswith(".csv"):
                file_path = os.path.join(root, file)
                tmx_file_path = os.path.join(root, os.path.splitext(file)[0] + ".tmx")
                print(f"Processing: {file_path}")
                process_csv_file(file_path, tmx_file_path, key_id, source_lang, target_lang)

if __name__ == "__main__":
    main()
