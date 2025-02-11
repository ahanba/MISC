# Misc  

## Table of Contents  
- [htmlTable2XLS](#htmltable2xls)  
- [csv2tmx](#csv2tmx)
- [json2csv](#json2csv)
- [csv2tbx](#csv2tbx)
- [mergexls](#mergexls)  

## htmlTable2XLS  

### Description  
This script converts HTML tables to an XLSX file, supporting merged rows and columns.  

### How to Use  
1. Install required packages.
   ```python
   pip install beautifulsoup4 pandas openpyxl
   ```
2. Save the HTML file to your PC (press **Ctrl+S** in your web browser).  
3. Place the Python file in the same folder.
4. Open the Python file in a text editor.  
5. Update the following value to match your target table header:  
   ```python
   HEADER_INDICATOR = 'HEADER OF YOUR TARGET TABLES'
   ```  
6. Run the script.

> [!TIP]
> Edit the following to change tables to select: ```all_tables = soup.select("div.table-wrap div.table-block table")```
   
## csv2tmx  

### Description  
This script converts CSV files to TMX format.  
It processes CSV files in the current directory and its subdirectories.  

### How to Use  
1. Open the Python file in a text editor.  
2. Update the following values to match your CSV headers:  
   ```python
   # Specify key, source, and target language headers
   key_id = "YOUR CSV'S KEY HEADER"
   source_lang = "YOUR CSV'S SOURCE HEADER"
   target_lang = "YOUR CSV'S TARGET HEADER"
   ```  
3. Run the script.  

> [!TIP]  
> For details on TMX files and their format, see the [TMX 1.4b specification](https://www.gala-global.org/tmx-14b), [Phrase](https://support.phrase.com/hc/ja/articles/6111346531484--TMX-Strings) and [Transifex](https://help.transifex.com/en/articles/6838724-tmx-files-and-format) documentation.  
  
## json2csv  

### Description  
This script merges data from multiple JSON files, each containing translations for different languages, into a single CSV file.  

### How to Use  
1. Place JSON files in a folder.
2. Run this script.
   ```python
   python json2csv.py input_folder output.csv
   ```

## csv2tbx  

### Description  
This script converts CSV files to TBX format.  
It processes CSV files in the current directory and its subdirectories.  

### How to Use  
1. Open the Python file in a text editor.
2. To edit target languages, edit the following line:
   ```
   LANGUAGES = {"ja_JP", "en_US", "zh_CN"}  # List up languages to process
   ```
3. To edit the Part Of Speech and Definition columns, edit the following lines to match with the headers of your source CSV:
   ```
   definition = row.get("Definition", "")
   pos = row.get("POS", "")
   ```
4. Run the script.  

> [!TIP]  
> For details on TBX files and their format, see the [TBX specification](https://www.gala-global.org/sites/default/files/migrated-pages/docs/tbx_oscar_0.pdf) and [memoQ](https://docs.memoq.com/9-9/api-docs/wsapi/memoqservices/tbservice.importexport.tbx.html) documentation.  

## mergexls  

### Description  
This script unzip a ZIP file containing XLS files, and then merge the XLS files into one XLS file.  
Column A and B are used as keys to match multiple XLS files.  

### How to Use  
1. Install required packages.
   ```python
   pip install pandas openpyxl
   ```
2. Place a ZIP file and the Python file in the same folder.
3. Run the script.
