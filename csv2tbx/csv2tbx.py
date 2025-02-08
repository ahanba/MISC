import os
import csv
from datetime import datetime
import pytz
from pathlib import Path

# Constants
TBX_HEADER = """<?xml version='1.0'?>
<!DOCTYPE martif SYSTEM "TBXcoreStructV02.dtd">
<martif type="TBX" xml:lang="en">
<martifHeader>
<encodingDesc>
<p type="XCSURI">http://www.lisa.org/fileadmin/standards/tbx/TBXXCSV02.XCS</p>
</encodingDesc>
</martifHeader>
<text>
<body>
"""

TBX_FOOTER = """</body>
</text>
</martif>
"""

LANGUAGES = {"ja_JP", "en_US", "zh_CN"}  # List up languages to process

def get_current_iso_time():
    """Returns the current UTC time in ISO 8601 format with 'Z'."""
    return datetime.now(tz=pytz.utc).isoformat().replace("+00:00", "Z")

def convert_csv_to_tbx(csv_path, tbx_path, iso_time):
    """Converts a CSV file to TBX format and writes the output."""
    print(f"Processing: {csv_path}")
    
    with open(csv_path, newline='', encoding='utf-8') as csvfile, open(tbx_path, "w", encoding="utf-8") as tbxfile:
        reader = csv.DictReader(csvfile)
        tbxfile.write(TBX_HEADER)
        
        for row in reader:
            definition = row.get("Definition", "")
            pos = row.get("POS", "")
            
            tbxfile.write(f"""
<termEntry>
<descrip type="Creator">{Path(csv_path).name}</descrip>
<descrip type="xDate_CreateTime">{iso_time}</descrip>
<descrip type="definition">{definition}</descrip>
""")
            
            for lang, term in row.items():
                if lang in LANGUAGES and term:
                    tbxfile.write(f"""
<langSet xml:lang="{lang}">
<tig>
<term>{term}</term>
<termNote type="partOfSpeech">{pos}</termNote>
</tig>
</langSet>
""")
            
            tbxfile.write("</termEntry>\n")
        
        tbxfile.write(TBX_FOOTER)
    
    print(f"Finished writing: {tbx_path}\n")

def process_csv_files():
    """Finds all CSV files in the current directory and its subdirectories, then converts them to TBX."""
    iso_time = get_current_iso_time()
    
    for root, _, files in os.walk("."):
        for file in files:
            if file.endswith(".csv"):
                csv_path = os.path.join(root, file)
                tbx_path = os.path.splitext(csv_path)[0] + ".tbx"
                convert_csv_to_tbx(csv_path, tbx_path, iso_time)

if __name__ == "__main__":
    process_csv_files()