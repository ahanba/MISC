import os
import csv
from datetime import datetime

# Specify key, source and target languages
key_id = "key"
source_lang = "ja_JP"
target_lang = "en_US"

# Get current date-time in TMX format
current_datetime = datetime.now().strftime("%Y%m%dT%H%M%S")

# TMX template
TMX_HEADER = """<?xml version="1.0" encoding="UTF-8"?>
<tmx version="1.4">
  <header srclang="en_US" datatype="plaintext"/>
  <body>
"""

TMX_FOOTER = """
  </body>
</tmx>
"""

# Get all CSV files in the current directory and subdirectories
for root, _, files in os.walk("."):
    for file in files:
        if file.endswith(".csv"):
            file_path = os.path.join(root, file)
            tmx_file_path = os.path.join(root, os.path.splitext(file)[0] + ".tmx") 

            print(f"Reading: {file_path}")
            
            # Open the output text file
            with open(tmx_file_path, "w", encoding="utf-8") as tmxfile:
                tmxfile.write(TMX_HEADER)  # Write TMX header

                # Open and read the CSV file
                with open(file_path, newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)  # Read as dictionary

                    for row in reader:
                        key = row.get(key_id, "N/A")
                        source_text = row.get(source_lang, "N/A")
                        target_text = row.get(target_lang, "N/A")

                        # Write TMX translation unit with attributes and separate <seg> tags
                        tmxfile.write(f"""
    <tu tuid="{key}">
      <tuv xml:lang="{source_lang}" creationdate="{current_datetime}" lastusagedate="{current_datetime}">
        <seg>{source_text}</seg>
      </tuv>
      <tuv xml:lang="{target_lang}" creationdate="{current_datetime}" lastusagedate="{current_datetime}">
        <seg>{target_text}</seg>
      </tuv>
    </tu>
""")

                tmxfile.write(TMX_FOOTER)  # Write TMX footer

            print(f"Finished writing: {tmx_file_path}\n")