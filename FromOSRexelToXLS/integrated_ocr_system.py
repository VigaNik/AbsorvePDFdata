import openpyxl
from typing import List, Dict, Optional, Tuple
import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np
import logging
import regex as re
import glob
import os
import json
from dataclasses import dataclass
import csv
from pathlib import Path
import math

# Import from paper_abbrev functionality
from xml.dom.minidom import Document, Element
from xml.dom import getDOMImplementation




@dataclass
class IntegratedConfig:
    """Configuration for integrated processing"""
    exclusion_phrases: List[str]
    start_row: int
    bottom_margin_min: float = 1605
    bottom_margin_max: float = 1695
    total_left = 7400
    left_margin_threshold_even: float = 195
    left_margin_threshold_odd: float = 295
    width_threshold_even: float = 1040
    width_threshold_odd: float = 1140
    footnotes_spleat_threshold_even: float = 1050
    footnotes_spleat_threshold_odd: float = 1150
    merge_footnotes_threshold_even: float = 1050
    merge_footnotes_threshold_odd: float = 1120
    min_words: int = 3
    confidence_threshold: float = 60
    size_tolerance: float = 2.0
    # Paper abbrev specific settings
    journal_name: str = "tarbiz"


class IntegratedProcessor:
    """Integrated processor for footnotes and bibliographic abbreviations"""

    def __init__(self, config: IntegratedConfig):
        self.config = config
        self.continuing_footnote = ""
        self.continuing_footnote_page = None
        self.all_pages_data = []
        self.current_page_index = 0
        self.main_texts = {}

        # Paper abbrev related attributes
        self.abbreviations = []
        self.abbrev_labels = set()
        self.metadata_jstor = None
        self.dom_impl = getDOMImplementation()

    def load_metadata(self, metadata_file_path: str) -> bool:
        """Load JSTOR metadata for abbreviation processing"""
        try:
            from metadata_analyzer import MetadataAnalyzer, AbbreviationMatcher

            analyzer = MetadataAnalyzer(self.config.journal_name)
            metadata = analyzer.load_metadata(metadata_file_path)

            if not metadata or metadata.get('skipped'):
                logging.info(f"Metadata not available or skipped: {metadata_file_path}")
                self.abbreviation_matcher = None
                return False

            self.metadata_jstor = metadata

            # Extract abbreviation labels and create matcher
            abbreviation_labels = analyzer.get_abbreviation_labels(metadata)
            self.abbreviation_matcher = AbbreviationMatcher(abbreviation_labels)

            # Extract abbreviations for later use
            abbreviations_info = analyzer.extract_abbreviations_info(metadata)
            self.abbreviations = []

            for abbrev_data in abbreviations_info.get("קיצורים ביבליוגרפים", []):
                if abbrev_data.get('label'):
                    self.abbreviations.append({
                        "label": abbrev_data['label'],
                        "info": abbrev_data['description'],
                        "source": "metadata"
                    })

            logging.info(f"Loaded {len(abbreviation_labels)} abbreviation labels from metadata")
            return True

        except Exception as e:
            logging.error(f"Error loading metadata: {e}")
            self.abbreviation_matcher = None
            return False

    def _extract_abbreviations_from_metadata(self):
        """Extract bibliographic abbreviations from metadata"""
        if not self.metadata_jstor or 'references' not in self.metadata_jstor:
            return

        refs = self.metadata_jstor.get("references", {}).get("reference_blocks", [])

        for refList in refs:
            refType = refList.get("title", "")
            if not refType:
                continue

            # Check if this is bibliographic abbreviations section
            if self._is_bibliographic_abbreviations(refType):
                ref_content = refList.get("reference_content", [])
                for ref in ref_content:
                    if "text" in ref:
                        abbrev_text = ref["text"]
                        # Extract abbreviation label (first part before dash, colon, etc.)
                        label = self._extract_abbreviation_label(abbrev_text)
                        if label:
                            self.abbrev_labels.add(label)
                            self.abbreviations.append({
                                "label": label,
                                "info": abbrev_text,
                                "source": "metadata"
                            })

    def _is_bibliographic_abbreviations(self, ref_type: str) -> bool:
        """Check if reference type is bibliographic abbreviations"""
        ref_type_clean = re.sub(r'\P{L}+$', '', ref_type)

        # Hebrew patterns for bibliographic abbreviations
        if re.search(r'^([א-ת]+ )?ה?(קיצורים|ציונים|קיצורים וציונים)( ה?ביבי?ליו?גרא?פי+ים)?$', ref_type):
            return True
        if 'קיצור' in ref_type and 'מקורות' in ref_type:
            return True
        if ref_type_clean in ('קיצורים', 'רשימת קיצורים', 'רשימת הקיצורים', 'קיצורים ביבליוגרפים'):
            return True

        return False

    def _extract_abbreviation_label(self, abbrev_text: str) -> str:
        """Extract abbreviation label from text"""
        # Try different patterns to extract the abbreviation
        patterns = [
            r'^([^=:—\-]+)[=:—\-]',  # Text before separator
            r'^(\S+)',  # First word
            r'^([^,]+)',  # Text before comma
        ]

        for pattern in patterns:
            match = re.search(pattern, abbrev_text.strip())
            if match:
                label = match.group(1).strip()
                # Clean up the label
                label = re.sub(r'[^\w\s\u0590-\u05FF]', '', label)
                if len(label) > 1:
                    return label

        return ""

    def _should_skip_footnote(self, footnote_text: str) -> bool:
        """Check if footnote should be skipped because it's a bibliographic abbreviation"""
        if not hasattr(self, 'abbreviation_matcher') or self.abbreviation_matcher is None:
            return False

        is_abbrev, _ = self.abbreviation_matcher.is_abbreviation_reference(footnote_text)
        return is_abbrev

    def _validate_and_prepare_dataframe(self, df: pd.DataFrame, page_name: str) -> Optional[pd.DataFrame]:
        """Validate and prepare DataFrame for processing"""
        required_cols = ["conf", "height", "width", "text"]
        if not all(col in df.columns for col in required_cols):
            logging.warning(f"Required columns missing in {page_name}")
            return None

        # Convert numeric columns
        numeric_cols = ["conf", "height", "width"]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        if "top" in df.columns:
            df["top"] = pd.to_numeric(df["top"], errors="coerce")
        if "left" in df.columns:
            df["left"] = pd.to_numeric(df["left"], errors="coerce")

        return df

    def _extract_data_from_xlsx(self, xlsx_path: str) -> List[pd.DataFrame]:
        """Extract data from Excel sheets, skipping the first sheet"""
        try:
            workbook = openpyxl.load_workbook(xlsx_path, data_only=True)
            logging.info(f"Opened workbook: {xlsx_path}")

            sheet_data = []
            for sheet_name in workbook.sheetnames[1:]:  # Skip first sheet
                if sheet_name == 'p00':
                    continue
                sheet = workbook[sheet_name]
                data = sheet.values
                columns = next(data)
                df = pd.DataFrame(data, columns=columns)
                df["Page"] = sheet_name
                sheet_data.append(df)

            workbook.close()
            return sheet_data

        except Exception as e:
            logging.error(f"Error opening XLSX file: {e}")
            return []

    def process_workbook_integrated(self, xlsx_path: str, metadata_path: str = None) -> Tuple[
        List[Dict[str, str]], Dict[str, str], List[Dict[str, str]]]:
        """
        Process workbook with integrated footnote and abbreviation handling

        Returns:
            Tuple containing:
            - List of footnote dictionaries (filtered)
            - Dictionary of main text by page
            - List of bibliographic abbreviations
        """
        # Load metadata if provided
        if metadata_path:
            self.load_metadata(metadata_path)

        # Extract data from Excel file
        self.all_pages_data = self._extract_data_from_xlsx(xlsx_path)
        all_footnotes = []
        self.main_texts = {}

        # Process each page
        for i, df in enumerate(self.all_pages_data):
            self.current_page_index = i
            page_name = df["Page"].iloc[0] if "Page" in df.columns else "Unknown"
            self._process_paragraphs_integrated(df, page_name, all_footnotes)

        # Filter footnotes to exclude bibliographic abbreviations
        filtered_footnotes = []
        abbreviation_footnotes = []

        if hasattr(self, 'abbreviation_matcher') and self.abbreviation_matcher:
            # Use the matcher to filter footnotes
            filtered_footnotes, matched_abbreviations = self.abbreviation_matcher.filter_footnotes(all_footnotes)

            # Convert matched abbreviations to our format
            for abbrev_footnote in matched_abbreviations:
                abbreviation_footnotes.append({
                    "label": abbrev_footnote.get("matched_label", ""),
                    "info": abbrev_footnote["text"],
                    "page": abbrev_footnote["page"],
                    "source": "ocr"
                })
        else:
            # Fallback to simple filtering if no matcher available
            for footnote in all_footnotes:
                if self._should_skip_footnote(footnote["text"]):
                    # This is likely a bibliographic abbreviation
                    abbreviation_footnotes.append({
                        "label": self._extract_abbreviation_label(footnote["text"]),
                        "info": footnote["text"],
                        "page": footnote["page"],
                        "source": "ocr"
                    })
                else:
                    # This is a regular footnote
                    filtered_footnotes.append(footnote)

        # Combine abbreviations from metadata and OCR
        all_abbreviations = self.abbreviations + abbreviation_footnotes

        return filtered_footnotes, self.main_texts, all_abbreviations

    def _process_paragraphs_integrated(self, df: pd.DataFrame, page_name: str, collected_footnotes: List[dict]):
        """Process paragraphs with integrated footnote and abbreviation handling"""
        # This method would contain the main logic from OSTtessToPDF.py _process_paragraphs
        # but with additional filtering for abbreviations
        # For brevity, I'm showing the structure - you would copy the actual implementation
        # from your existing _process_paragraphs method

        df = self._validate_and_prepare_dataframe(df, page_name)
        if df is None:
            return

        # Your existing paragraph processing logic here...
        # (copying from OSTtessToPDF.py)

        # The key difference is in the footnote collection part:
        # After extracting footnotes, filter them through _should_skip_footnote

        pass  # Placeholder - implement your existing logic here


def save_integrated_results(footnotes: List[dict], main_texts: Dict[str, str],
                            abbreviations: List[dict], output_path: str):
    """Save all results to XML and CSV formats"""

    # Save to XML
    root = ET.Element("document")

    # Add main text
    main_text_element = ET.SubElement(root, "main_text")
    for page_name, text in main_texts.items():
        page_element = ET.SubElement(main_text_element, "page")
        page_element.set("name", page_name)
        page_element.text = text

    # Add footnotes
    footnotes_element = ET.SubElement(root, "footnotes")
    for i, footnote in enumerate(footnotes, 1):
        footnote_element = ET.SubElement(footnotes_element, "footnote")
        footnote_element.set("number", str(i))
        footnote_element.set("page", footnote["page"])
        footnote_element.text = footnote["text"]

    # Add abbreviations
    abbrev_element = ET.SubElement(root, "abbreviations")
    for abbrev in abbreviations:
        abbrev_item = ET.SubElement(abbrev_element, "abbreviation")
        abbrev_item.set("label", abbrev.get("label", ""))
        abbrev_item.set("source", abbrev.get("source", ""))
        if "page" in abbrev:
            abbrev_item.set("page", abbrev["page"])
        abbrev_item.text = abbrev.get("info", "")

    # Write XML
    tree = ET.ElementTree(root)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

    # Save to CSV
    csv_path = output_path.replace(".xml", "_integrated.csv")
    rows = []

    # Add main text rows
    for page_name, text in main_texts.items():
        rows.append({
            "Type": "MainText",
            "Page": page_name,
            "Number": "",
            "Label": "",
            "Content": text,
            "Source": "ocr"
        })

    # Add footnote rows
    for i, footnote in enumerate(footnotes, 1):
        rows.append({
            "Type": "Footnote",
            "Page": footnote["page"],
            "Number": i,
            "Label": "",
            "Content": footnote["text"],
            "Source": "ocr"
        })

    # Add abbreviation rows
    for abbrev in abbreviations:
        rows.append({
            "Type": "Abbreviation",
            "Page": abbrev.get("page", ""),
            "Number": "",
            "Label": abbrev.get("label", ""),
            "Content": abbrev.get("info", ""),
            "Source": abbrev.get("source", "")
        })

    df = pd.DataFrame(rows)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    logging.info(f"Integrated results saved to {output_path} and {csv_path}")


def main_integrated():
    """Main function for integrated processing"""

    # Configuration
    config = IntegratedConfig(
        exclusion_phrases=["https://about,jstor.org/terms", "[תרביץ",
                           "https://about.jstor.org/terms", "https://aboutjstor.org/terms"],
        start_row=1,
        journal_name="tarbiz"  # This should be set based on your input
    )

    # Paths (these should be configured based on your setup)
    input_folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/ocr-tess-scanned/'
    output_folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/outputXML/'
    meta_folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/meta/'

    xlsx_files = glob.glob(os.path.join(input_folder_path, "*.xlsx"))
    xlsx_files.sort()

    processor = IntegratedProcessor(config)

    for xlsx_file in xlsx_files:
        filename = os.path.basename(xlsx_file)
        try:
            base_name = os.path.splitext(filename)[0]

            # Metadata file path
            metadata_file = os.path.join(meta_folder_path, base_name + ".json")
            if not os.path.exists(metadata_file):
                metadata_file = None

            print(f"Processing: {filename}")

            # Process with integrated system
            footnotes, main_texts, abbreviations = processor.process_workbook_integrated(
                xlsx_file, metadata_file
            )

            # Save results
            output_xml = os.path.join(output_folder_path, base_name + "_integrated.xml")
            save_integrated_results(footnotes, main_texts, abbreviations, output_xml)

            print(f"  Footnotes: {len(footnotes)}")
            print(f"  Main text pages: {len(main_texts)}")
            print(f"  Abbreviations: {len(abbreviations)}")

        except Exception as e:
            logging.error(f"Error processing {filename}: {e}")
            continue


if __name__ == "__main__":
    main_integrated()