import os
import pandas as pd
from bs4 import BeautifulSoup
import csv
import html
from typing import List, Optional, Dict
import logging
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class HTMLTableConverter:
    HEADER_INDICATOR = 'タイトル'

    def __init__(self, input_dir: str = "."):
        self.input_dir = Path(input_dir)
        
    def get_html_files(self) -> List[Path]:
        """Return all HTML files in the input directory and its subdirectories."""
        return list(self.input_dir.rglob("*.html"))

    def _has_required_header(self, table) -> bool:
        """Check if table contains the required header indicator."""
        headers = table.find_all(['th', 'td'])
        return any(self.HEADER_INDICATOR in html.unescape(header.get_text(strip=True)) 
                  for header in headers)

    def _get_cell_text(self, cell) -> str:
        """Extract text from a cell, handling paragraphs and other elements."""
        # Get only the direct text of this cell, not its children
        texts = []
        for content in cell.contents:
            if isinstance(content, str):
                text = content.strip()
                if text:
                    texts.append(text)
            elif content.name in ['p', 'li', 'pre', 'a']:
                text = content.get_text(strip=True)
                if text:
                    texts.append(text)
        
        return html.unescape(' '.join(texts))

    def _extract_table_data(self, table) -> List[List[str]]:
        """
        Extract data from a table, handling colspan and rowspan.
        Returns a list of rows, where each row is a list of cell values.
        """
        rows = table.find_all('tr')
        
        # First pass: determine the maximum number of columns
        max_cols = 0
        for row in rows:
            col_count = 0
            for cell in row.find_all(['td', 'th']):
                colspan = int(cell.get('colspan', 1))
                col_count += colspan
            max_cols = max(max_cols, col_count)

        # Initialize the grid with None
        grid = [[None] * max_cols for _ in range(len(rows))]
        
        # Process the table
        for row_idx, row in enumerate(rows):
            col_idx = 0
            
            for cell in row.find_all(['td', 'th']):
                # Skip cells that are already filled
                while col_idx < max_cols and grid[row_idx][col_idx] is not None:
                    col_idx += 1
                
                if col_idx >= max_cols:
                    break

                # Get cell properties
                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))
                cell_value = self._get_cell_text(cell)

                # Fill the spanned cells
                for i in range(rowspan):
                    for j in range(colspan):
                        if row_idx + i < len(grid) and col_idx + j < len(grid[0]):
                            grid[row_idx + i][col_idx + j] = cell_value

                col_idx += colspan

        # Convert None to empty string in the final output
        return [[cell if cell is not None else '' for cell in row] for row in grid]

    def read_html(self, html_file: Path) -> Optional[Path]:
        """
        Read HTML file and convert tables to CSV.
        Only processes tables containing the required header indicator.
        Returns the path to the created CSV file or None if no valid tables found.
        """
        csv_filename = html_file.with_suffix('.csv')
        
        try:
            with open(html_file, "r", encoding="utf-8") as file:
                soup = BeautifulSoup(file, "html.parser")

            all_tables = soup.select("div.table-wrap div.table-block table")
            tables = [table for table in all_tables if self._has_required_header(table)]
            
            if not tables:
                logging.warning(f"No tables with header '{self.HEADER_INDICATOR}' found in {html_file}")
                return None

            with open(csv_filename, "w", encoding="utf-8", newline="") as csvfile:
                writer = csv.writer(csvfile)

                for table in tables:
                    table_data = self._extract_table_data(table)
                    writer.writerows(table_data)

            logging.info(f"Successfully created CSV: {csv_filename}")
            return csv_filename

        except Exception as e:
            logging.error(f"Error processing {html_file}: {str(e)}")
            return None

    def write_excel(self, html_file: Path, csv_file: Path) -> Optional[Path]:
        """Convert CSV to Excel file."""
        try:
            df = pd.read_csv(csv_file, encoding="utf-8", header=None)
            xlsx_path = html_file.with_suffix('.xlsx')
            df.to_excel(xlsx_path, index=False, header=False, engine="openpyxl")
            logging.info(f"Successfully created Excel file: {xlsx_path}")
            return xlsx_path
        except Exception as e:
            logging.error(f"Error creating Excel file from {csv_file}: {str(e)}")
            return None

    def process_files(self):
        """Process all HTML files in the input directory."""
        html_files = self.get_html_files()
        
        if not html_files:
            logging.warning(f"No HTML files found in {self.input_dir}")
            return

        for html_file in html_files:
            logging.info(f"Processing: {html_file}")
            csv_file = self.read_html(html_file)
            
            if csv_file:
                self.write_excel(html_file, csv_file)

def main():
    converter = HTMLTableConverter()
    converter.process_files()

if __name__ == "__main__":
    main()