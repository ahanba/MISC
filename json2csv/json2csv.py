import sys
import json
import csv
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Any
from dataclasses import dataclass

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class TranslationConfig:
    """Configuration settings for translation processing."""
    json_folder: Path
    output_csv: Path
    separator: str = '.'

class TranslationError(Exception):
    """Custom exception for translation-related errors."""
    pass

class JsonTranslationHandler:
    """Handles the processing of JSON translation files to CSV format."""
    
    def __init__(self, config: TranslationConfig):
        """Initialize the handler with configuration settings."""
        self.config = config
        self.data: Dict[str, Dict[str, Any]] = {}
        self.languages: List[str] = []
        
    def flatten_json(self, nested_json: Dict, parent_key: str = '') -> Dict[str, Any]:
        """
        Flatten a nested JSON object into a single-level dictionary.
        
        Args:
            nested_json: The nested JSON dictionary to flatten
            parent_key: The parent key for nested structures
            
        Returns:
            A flattened dictionary with dot-notation keys
        """
        items: Dict[str, Any] = {}
        for key, value in nested_json.items():
            new_key = f"{parent_key}{self.config.separator}{key}" if parent_key else key
            
            if isinstance(value, dict):
                items.update(self.flatten_json(value, new_key))
            else:
                items[new_key] = value
        return items

    def load_json_files(self) -> None:
        """
        Load and process all JSON files from the specified folder.
        
        Raises:
            TranslationError: If there are issues with file processing
        """
        try:
            json_files = list(self.config.json_folder.glob("*.json"))
            if not json_files:
                raise TranslationError(f"No JSON files found in {self.config.json_folder}")

            for json_file in json_files:
                self._process_json_file(json_file)
            
            self.languages.sort()
            logger.info(f"Successfully loaded {len(json_files)} JSON files")
            
        except Exception as e:
            raise TranslationError(f"Error loading JSON files: {str(e)}") from e

    def _process_json_file(self, json_file: Path) -> None:
        """
        Process a single JSON file and update the translation data.
        
        Args:
            json_file: Path to the JSON file to process
        """
        lang = json_file.stem # Use filename as language identifier
        self.languages.append(lang)
        
        try:
            with json_file.open('r', encoding='utf-8') as f:
                json_data = json.load(f)
                flat_data = self.flatten_json(json_data)
                
                for key, value in flat_data.items():
                    if key not in self.data:
                        self.data[key] = {}
                    self.data[key][lang] = value
                    
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in file {json_file}: {str(e)}")
            raise TranslationError(f"Invalid JSON format in {json_file}") from e
        except Exception as e:
            logger.error(f"Error processing file {json_file}: {str(e)}")
            raise TranslationError(f"Error processing {json_file}") from e

    def write_csv_output(self) -> None:
        """
        Write the processed translation data to a CSV file.
        
        Raises:
            TranslationError: If there are issues writing the CSV file
        """
        try:
            with self.config.output_csv.open('w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                header = ["Key"] + self.languages
                writer.writerow(header)
                
                for key, translations in self.data.items():
                    row = [key] + [translations.get(lang, '') for lang in self.languages]
                    writer.writerow(row)
                    
            logger.info(f"Successfully created CSV file: {self.config.output_csv}")
            
        except Exception as e:
            raise TranslationError(f"Error writing CSV file: {str(e)}") from e

def main() -> None:
    """Main entry point for the translation processor."""
    try:
        if len(sys.argv) != 3:
            raise ValueError("Incorrect number of arguments")
        
        config = TranslationConfig(
            json_folder=Path(sys.argv[1]),
            output_csv=Path(sys.argv[2])
        )
        
        if not config.json_folder.exists():
            raise TranslationError(f"Input folder does not exist: {config.json_folder}")
        
        handler = JsonTranslationHandler(config)
        handler.load_json_files()
        handler.write_csv_output()
        
    except (ValueError, TranslationError) as e:
        logger.error(str(e))
        print(f"Error: {str(e)}")
        print("Usage: python merge_json_to_csv.py <json_folder> <output_csv>")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        print(f"An unexpected error occurred. Check the logs for details.")
        sys.exit(1)

if __name__ == "__main__":
    main()
