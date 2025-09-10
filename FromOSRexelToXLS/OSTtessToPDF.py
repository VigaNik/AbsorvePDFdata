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
#from difflib import SequenceMatcher

# Unicode direction marks
uni_ltr = '\u200e'  # left to right mark
uni_rtl = '\u200f'  # right to left mark

# Additional bidi symbols that may appear
BIDI_CHARS = [
    '\u200e',  # LEFT-TO-RIGHT MARK
    '\u200f',  # RIGHT-TO-LEFT MARK
    '\u202a',  # LEFT-TO-RIGHT EMBEDDING
    '\u202b',  # RIGHT-TO-LEFT EMBEDDING
    '\u202c',  # POP DIRECTIONAL FORMATTING
    '\u202d',  # LEFT-TO-RIGHT OVERRIDE
    '\u202e',  # RIGHT-TO-LEFT OVERRIDE
    '\u2066',  # LEFT-TO-RIGHT ISOLATE
    '\u2067',  # RIGHT-TO-LEFT ISOLATE
    '\u2068',  # FIRST STRONG ISOLATE
    '\u2069',  # POP DIRECTIONAL ISOLATE
]



@dataclass
class footnoteConfig:
    """Configuration settings for footnote extraction"""
    exclusion_phrases: List[str]
    start_row: int
    bottom_margin_min: float = 1605
    # for tarbitz 1680(1605better) for lecohotenu 1655 meghillot 1667 shenmishivri 1663 sibra 1703 zion 1689
    bottom_margin_max: float = 1695
    # for tarbitz 1695 for lecohotenu 1665 meghillot 1675 shenmishivri 1671 sibra 1712 zion 1698
    total_left: float = 7400
    left_margin_threshold_even: float = 195
    left_margin_threshold_odd: float = 295
    # for tarbitz: 195 for lecohotenu: 220 meghillot 220 shenmishivri 210 sibra 195 zion 225
    width_threshold_even: float = 1040
    width_threshold_odd: float = 1140
    # width_threshold_odd for cheking if we have to combain referenses from diferent pages
    # for tarbitz 1095 for lecohotenu 1077 meghillot 1075 shenmishivri 1090 sibra ? zion (1080/1005?)
    #for printed docs we have divide every page to odd and even, in the scanned doce  merge_footnotes_threshold_even = merge_footnotes_threshold_odd
    footnotes_spleat_threshold_even: float = 1050
    footnotes_spleat_threshold_odd:  float = 1150
    merge_footnotes_threshold_even: float = 1050
    merge_footnotes_threshold_odd: float = 1120
    # merge_footnotes_threshold_odd threshold fot checking if we have merge footnotes from seme page
    # for tarbitz 1070 for lecohotenu ? meghillot ? shenmishivri ? sibra ? zion ?
    min_words: int = 3
    confidence_threshold: float = 60
    size_tolerance: float = 2.0


class footnoteProcessor:
    def __init__(self, config: footnoteConfig):
        self.config = config
        self.continuing_footnote = ""
        self.continuing_footnote_page = None
        self.all_pages_data = []
        self.current_page_index = 0
        self.check_current_page_index = 0
        self.main_texts = {}  # Dictionary to store main text for each page

    def _validate_and_prepare_dataframe(self, df: pd.DataFrame, page_name: str) -> Optional[pd.DataFrame]:
        # Проверка на None или пустой DataFrame
        if df is None:
            logging.warning(f"DataFrame is None for {page_name}")
            return None

        if df.empty:
            logging.warning(f"DataFrame is empty for {page_name}")
            return None

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

        # Проверяем что после конвертации остались валидные данные
        valid_data = df[df["conf"] >= 0]
        if valid_data.empty:
            logging.warning(f"No valid data after processing for {page_name}")
            return None

        return df

    def typeset_words(self, words: list):
        """
        Reorder the word list according to bidi marks
        (switch to RTL when the words begins with a Hebrew letter)
        """
        reordered_words = []
        reverse_span = []
        word_num = len(words)

        if word_num <= 1:
            return words

        # DEBUGGING: Let's add logging for verification
        # print(f"DEBUG: typeset_words called with {word_num} words")
        # print(f"DEBUG: First word: '{words[0].text}', left: {words[0].left}")

        # sort the words from right to left
        words_sorted_rl = sorted(words, key=lambda w: -w.left)
        words_sorted_lr = sorted(words, key=lambda w: w.left)

        # FIX 3: Improved sort order selection logic
        if len(words) >= 2:
            # Check if there is Hebrew text in the first words
            has_hebrew_in_first_words = any(re.search(r'[א-ת]', w.text) for w in words[:3])

            if has_hebrew_in_first_words:
                # For Hebrew text we use right to left sorting
                words = words_sorted_rl
            else:
                # For English/Latin - left to right
                words = words_sorted_lr
        else:
            words = words_sorted_rl

        # We define the main direction
        if len(words) >= 2 and words[1].left < words[0].left:
            dir = 'rtl'
        else:
            dir = 'ltr'

        main_dir = dir

        for (i_w, w) in enumerate(words):
            if w.text == '':
                continue

            prev_dir = dir

            if dir == main_dir:
                reordered_words.append(w)
            else:
                reverse_span.append(w)

            # Improved direction switching logic
            if w.text.endswith(uni_ltr):
                dir = 'ltr'
            elif w.text.endswith(uni_rtl):
                dir = 'rtl'
            elif i_w + 1 < word_num:
                next_word = words[i_w + 1]
                # Checking Hebrew letters to determine direction
                if re.search(r'[א-ת]', next_word.text):
                    dir = 'rtl'
                elif re.search(r'[a-zA-Z]', next_word.text):
                    dir = 'ltr'

            if dir != prev_dir and prev_dir != main_dir:
                if main_dir == 'rtl':
                    reordered_words = reordered_words + reverse_span[::-1]
                else:
                    reordered_words = reverse_span[::-1] + reordered_words
                reverse_span = []

        if len(reverse_span) > 0:
            if main_dir == 'rtl':
                reordered_words = reordered_words + reverse_span[::-1]
            else:
                reordered_words = reverse_span[::-1] + reordered_words

        return reordered_words

    def calc_font_size(self, word_span):
        """
        Calculates appropriate font size based on the vertical characteristics of text

        Args:
            word_span: An object with height and text properties

        Returns:
            Integer representing the calculated font size
        """
        height = word_span.height
        text = word_span.text


        # Symbols with UPPER ascenders
        upper_chars_hebrew = 'ל'  # lamed
        upper_chars_latin = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'  # Все заглавные латинские
        upper_chars_latin_ascenders = 'bdfhklt'  # Строчные с верхними выносными

        # Symbols with DECLINERS
        lower_chars_hebrew = 'ךןףץק'  # final forms + qof
        lower_chars_latin = 'gjpqy'  # Строчные с нижними выносными

        # Check for the presence of upper extensions
        has_upper = (re.search(f'[{upper_chars_hebrew}{upper_chars_latin}{upper_chars_latin_ascenders}]',
                               text) is not None)

        # Check for the presence of lower extensions
        has_lower = (re.search(f'[{lower_chars_hebrew}{lower_chars_latin}]', text) is not None)

        # Font size calculation logic
        if not has_upper and not has_lower:
            # Base height characters only (eg: אבגדהוזחטיכמנסעפצרשת, aemnorsuvwxz, numbers)
            if ',' in text:
                font_size = height - 3  # Запятая может быть ниже базовой линии
            else:
                font_size = height  # Прямое соответствие высоты и размера шрифта
        else:
            # There are remote elements
            if has_upper and has_lower:
                # Both upper and lower extensions
                font_size = height / 2
            else:
                # Only ascenders OR only descenders
                font_size = height * 2 / 3

            font_size = int(font_size)

        return font_size

    def only_full_line(self, word):
        """
        Checks if a word contains only characters that are on the standard line height
        (no ascending or descending characters)

        Args:
            word: Text string to check

        Returns:
            Boolean indicating if the word only has standard-height characters
        """

        # Symbols with UPPER ascenders
        upper_chars_hebrew = 'ל'  # lamed
        upper_chars_latin = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'  # Все заглавные
        upper_chars_latin_ascenders = 'bdfhklt'  # Строчные с верхними выносными

        # Symbols with DECLINERS
        lower_chars_hebrew = 'ךןףץק'  # final forms + qof
        lower_chars_latin = 'gjpqy'  # Строчные с нижними выносными

        has_upper = (re.search(f'[{upper_chars_hebrew}{upper_chars_latin}{upper_chars_latin_ascenders}]',
                               word) is not None)
        has_lower = (re.search(f'[{lower_chars_hebrew}{lower_chars_latin}]', word) is not None)

        # Check that the word does not consist only of 'י' (yod)
        has_not_yod = (re.search(r'[^י]', word) is not None)

        # A word has only a base height if:
        # - No ascenders AND
        # - No descenders AND
        # - There are characters other than yod

        return (not has_upper) and (not has_lower) and has_not_yod

    def _extract_data_from_xlsx(self, xlsx_path: str) -> List[pd.DataFrame]:
        """Extract data from Excel sheets, skipping the first sheet"""
        try:
            workbook = openpyxl.load_workbook(xlsx_path, data_only=True)
            logging.info(f"Opened workbook: {xlsx_path}")

            sheet_data = []
            for sheet_name in workbook.sheetnames[0:]:
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

    def _split_into_paragraphs(self, df: pd.DataFrame, page_name: str) -> List[dict]:
        """Split DataFrame into paragraphs based on confidence values and font sizes"""
        current_paragraph = []
        paragraph_number = 1
        paragraph_data = []
        temp = 0
        previous_row = None
        previous_previous_row = None
        previous_previous_previous_row = None

        # Create a class to store word information for font size calculation
        class WordSpan:
            def __init__(self, height, text):
                self.height = height
                self.text = text

        def is_valid_paragraph(para_df: pd.DataFrame, paragraph_index: int = None,
                               total_paragraphs: int = None) -> bool:
            word_count = sum(len(str(text).split()) for text in para_df["text"].dropna())
            if word_count < self.config.min_words:
                return False

            # Calculate adjusted heights for information
            adjusted_heights = []
            for _, row in para_df.iterrows():
                if pd.notna(row["height"]) and pd.notna(row["text"]):
                    word_span = WordSpan(row["height"], str(row["text"]))
                    adjusted_font_size = self.calc_font_size(word_span)
                    adjusted_heights.append(adjusted_font_size)

            # Check additional criteria
            for text in para_df["text"].dropna().astype(str):
                if page_name == "p01" and ("לשוננו" or "מגילות") in text:
                    return False

                # Checking publication titles ONLY for the penultimate paragraph
                journal_names = ["מגילות", "magilot"]
                if (page_name == "p01" and
                        paragraph_index is not None and
                        total_paragraphs is not None and
                        paragraph_index == total_paragraphs - 2):  # Предпоследний параграф (индекс = total-2)
                    if any(journal_name in text.lower() for journal_name in journal_names):
                        return False

                # Checking publication titles ONLY for the penultimate paragraph
                journal_names = ["לשוננו", "lecohotenu"]
                if (page_name == "p01" and
                        paragraph_index is not None and
                        total_paragraphs is not None and
                        paragraph_index == total_paragraphs - 2):  # Предпоследний параграф (индекс = total-2)
                    if any(journal_name in text.lower() for journal_name in journal_names):
                        return False

                if any(phrase in text.lower() for phrase in self.config.exclusion_phrases):
                    return False
                if (para_df["top"].max() - para_df["top"].min() >= 60) and (word_count < 4):
                    return False

            return True

        # Collect all paragraphs first without validation
        all_paragraphs = []

        for _, row in df.iterrows():
            if (row["conf"] == -1 and
                previous_row is not None and previous_row["conf"] == -1 and
                previous_previous_row is not None and previous_previous_row["conf"] == -1) and (
                    previous_previous_previous_row is not None and row["top"] - previous_previous_previous_row[
                "top"] >= 35):
                temp = -1

            if temp == -1 and current_paragraph:
                paragraph_df = pd.DataFrame(current_paragraph)
                all_paragraphs.append(paragraph_df)
                paragraph_number += 1
                current_paragraph = []
                temp = 0

            current_paragraph.append(row)
            previous_previous_previous_row = previous_previous_row
            previous_previous_row = previous_row
            previous_row = row

        # Add the last paragraph
        if current_paragraph:
            paragraph_df = pd.DataFrame(current_paragraph)
            all_paragraphs.append(paragraph_df)

        # Now we validate paragraphs with information about their position
        total_paragraphs = len(all_paragraphs)

        for paragraph_index, paragraph_df in enumerate(all_paragraphs):
            if is_valid_paragraph(paragraph_df, paragraph_index, total_paragraphs):
                # Use our new functions for improved height analysis
                std_height = paragraph_df.loc[paragraph_df["conf"] != -1, "height"].std()
                median_height = float(np.mean(paragraph_df.loc[paragraph_df["conf"] != -1, "height"]))
                avg_height = paragraph_df["height"].mean()
                avg_width = paragraph_df["width"].mean()

                # Calculate adjusted height metrics
                adjusted_heights = []
                adjusted_median = median_height
                adjusted_std = std_height

                for _, p_row in paragraph_df.iterrows():
                    if pd.notna(p_row["height"]) and pd.notna(p_row["text"]):
                        word_span = WordSpan(p_row["height"], str(p_row["text"]))
                        adjusted_heights.append(self.calc_font_size(word_span))

                # We calculate adjusted metrics only if we have data
                if adjusted_heights:
                    adjusted_median = float(np.mean(adjusted_heights))
                    adjusted_std = np.std(adjusted_heights)
                """
                logging.info(
                    f"Page: {page_name}, Paragraph: {paragraph_index + 1}/{total_paragraphs}, "
                    f"Avg_Height: {avg_height:.2f}, Avg_Width: {avg_width:.2f}, "
                    f"Adj_Median: {adjusted_median:.2f}, Adj_Std: {adjusted_std:.2f}"
                )
                """
                paragraph_data.append({
                    "median_height": median_height,
                    "std_height": std_height,
                    "adjusted_median": adjusted_median,
                    "adjusted_std": adjusted_std,
                    "number": paragraph_index + 1,
                    "avg_height": avg_height,
                    "avg_width": avg_width,
                    "data": paragraph_df
                })

        return paragraph_data

    def _extract_main_text(self, paragraphs: List[dict]) -> str:
        """Extract main text from paragraphs with improved logic"""
        if not paragraphs:
            return ""

        # By default, include ALL paragraphs in the main text
        # Exclude only those that are EXACTLY defined as footnotes
        main_text_paragraphs = paragraphs.copy()

        # If there are several paragraphs, check the last one for footnotes
        if len(paragraphs) > 1:
            last_paragraph = paragraphs[-1]
            paragraph_df = last_paragraph["data"]

            # Check if the last paragraph is in the footnote zone
            is_in_footnote_zone = False
            if "top" in paragraph_df.columns:
                max_top = paragraph_df["top"].max()
                if max_top > self.config.bottom_margin_min:
                    is_in_footnote_zone = True

            # TEMPORARILY exclude the last paragraph from the main text
            # It will be added back if it does not pass the footnote check
            if is_in_footnote_zone:
                main_text_paragraphs = paragraphs[:-1]

        # Collect text from main text paragraphs
        main_text = ""
        for paragraph in main_text_paragraphs:
            paragraph_text = self._get_paragraph_text(paragraph["data"])
            if paragraph_text:
                main_text += paragraph_text + "\n\n"

        return main_text.strip()

    def _get_paragraph_text(self, paragraph_df: pd.DataFrame) -> str:
        """Extract text from a paragraph DataFrame with bidi text support and line separation."""
        # Group by 'top' coordinate to identify separate lines
        if "top" in paragraph_df.columns:
            # Sort by top coordinate to maintain line order
            sorted_df = paragraph_df.sort_values('top')

            # Group consecutive rows with similar 'top' values (same line)
            lines = []
            current_line_rows = []
            previous_top = None
            line_tolerance = 10  # pixels tolerance for considering rows as same line

            for _, row in sorted_df.iterrows():
                if pd.notna(row["text"]) and str(row["text"]).strip():
                    current_top = row["top"]

                    # If this is a new line (different top coordinate)
                    if (previous_top is not None and
                            abs(current_top - previous_top) > line_tolerance):
                        # Process the previous line
                        if current_line_rows:
                            line_text = self._process_line_text(current_line_rows)
                            if line_text:
                                lines.append(line_text)
                            current_line_rows = []

                    current_line_rows.append(row)
                    previous_top = current_top

            # Process the last line
            if current_line_rows:
                line_text = self._process_line_text(current_line_rows)
                if line_text:
                    lines.append(line_text)

            # Join lines with "|" separator
            return " | ".join(lines)

        else:
            # Fallback to original method if no 'top' coordinate available
            return self._process_line_text(paragraph_df.to_dict('records'))

    def _process_line_text(self, line_rows) -> str:
        """Process a single line of text with bidi support."""
        if not line_rows:
            return ""

        # Convert to DataFrame if it's a list of dictionaries
        if isinstance(line_rows, list):
            line_df = pd.DataFrame(line_rows)
        else:
            line_df = line_rows

        text_parts = []

        #Improved checking for Hebrew/RTL text
        has_rtl = False
        for _, row in line_df.iterrows():
            if pd.notna(row["text"]):
                text = str(row["text"])
                # Check for Hebrew letters (more reliable than \p{bc=R})
                if re.search(r'[א-ת]', text):
                    has_rtl = True
                    break
                # Additional check for RTL symbols
                if re.search(r'\p{bc=R}', text):
                    has_rtl = True
                    break

        # FIX 2: ALWAYS use typeset_words if left coordinates are present
        # (in older versions this could always work)
        if "left" in line_df.columns and len(line_df) > 1:
            # Convert DataFrame rows to word objects
            class WordObj:
                def __init__(self, text, left):
                    self.text = text
                    self.left = left

            word_objs = []
            for _, row in line_df.iterrows():
                if pd.notna(row["text"]):
                    word_objs.append(WordObj(str(row["text"]), row["left"]))

            # ALWAYS apply bidi reordering if there are coordinates
            reordered_words = self.typeset_words(word_objs)
            text_parts = [w.text for w in reordered_words]
        else:
            # Fallback: Regular processing без координат
            for _, row in line_df.iterrows():
                if pd.notna(row["text"]):
                    text_parts.append(str(row["text"]))

        return " ".join(text_parts).strip()

    def _should_merge_footnotes(self, segments: List[pd.DataFrame], page_name: str = None) -> List[pd.DataFrame]:
        """
        Merges link segments if the next segment does not contain cells with left > threshold.
        Different thresholds are used for odd and even pages.

        Args:
            segments (List[pd.DataFrame]): List of DataFrames representing link segments.
            page_name (str, optional): Name of the page. If None, tries to extract from segments.

        Returns:
            List[pd.DataFrame]: The updated list of segments, merged if necessary.
        """
        if len(segments) <= 1:
            return segments

        # If page_name is not passed, we try to extract from segments
        if page_name is None and "Page" in segments[0].columns:
            page_name = segments[0]["Page"].iloc[0]

        # Determine the parity of a page if the name is known
        even = False
        if page_name is not None:
            try:
                # Extract the number from the page name
                page_num = int(''.join([c for c in page_name if c.isdigit()]))
                even = (page_num % 2 == 0)
            except (ValueError, TypeError):
                # If we couldn't extract the number, we use the threshold for odd pages
                print(f"Warning: Could not determine if page {page_name} is even or odd. Using odd page threshold.")

        # Select the threshold value depending on the page parity
        threshold = (self.config.merge_footnotes_threshold_even if even
                     else self.config.merge_footnotes_threshold_odd)

        # If there are no special thresholds, use the usual one
        if not hasattr(self.config, 'merge_footnotes_threshold_even'):
            threshold = self.config.merge_footnotes_threshold_even

        merged_segments = [segments[0]]
        for i in range(1, len(segments)):
            segment = segments[i]
            prev_segment = merged_segments[-1]

            has_large_left = not segment[(segment["conf"] != -1) & (segment["left"] > threshold)].empty

            if has_large_left:
                # New link, save separately
                merged_segments.append(segment)
            else:
                # This is a continuation of the previous one
                merged_segments[-1] = pd.concat([prev_segment, segment])

        return merged_segments

    def _extract_footnotes(self, paragraph_df: pd.DataFrame, page_name: str = None) -> List[str]:
        """Extract individual footnotes from a paragraph, with line separation using |"""
        # Stage 1: basic segmentation by double -1 confidence
        footnote_segments = []
        current_rows = []
        consecutive_minus_one = 0
        for idx, row in paragraph_df.iterrows():
            if row['conf'] == -1:
                consecutive_minus_one += 1
                if consecutive_minus_one == 2:
                    if current_rows:
                        footnote_segments.append(paragraph_df.loc[current_rows])
                        current_rows = []
                    consecutive_minus_one = 0
            else:
                consecutive_minus_one = 0
                if str(row['text']).strip():
                    current_rows.append(idx)
        if current_rows:
            footnote_segments.append(paragraph_df.loc[current_rows])
        if not footnote_segments:
            return []

        # Stage 2: merge adjacent if needed
        merged_segments = self._should_merge_footnotes(footnote_segments, page_name)

        # Stage 3: first splitting by left threshold
        split_once = []
        for segment in merged_segments:
            parts = self._split_by_left_threshold(segment, page_name)
            split_once.extend(parts)

        # Stage 4: second splitting pass on all parts
        split_twice = []
        for segment in split_once:
            parts = self._split_by_left_threshold(segment, page_name)
            split_twice.extend(parts)

        # Reconstruct texts from final segments with line separation
        footnotes = []
        for i, seg in enumerate(split_twice, 1):
            # Use the new line-aware text extraction
            text = self._get_paragraph_text(seg)
            footnotes.append(text)
        return footnotes

    #this function created to spliat combine footnotes
    def _split_by_left_threshold(self, segment: pd.DataFrame, page_name: str = None) -> List[pd.DataFrame]:
        """
        Split a footnote segment into subsegments based on left-position threshold.

        Args:
            segment: DataFrame of a single footnote segment.
            page_name: Optional page name to determine odd/even threshold.
        Returns:
            List of DataFrame subsegments after splitting.
        """
        # Determine even/odd page
        try:
            num = int(''.join(c for c in (page_name or '') if c.isdigit()))
            even = (num % 2 == 0)
        except Exception:
            even = False

        # Select threshold
        threshold = (self.config.footnotes_spleat_threshold_even if even
                     else self.config.footnotes_spleat_threshold_odd)

        indices = segment.index.tolist()
        large_positions = [pos for pos, (_, row) in enumerate(segment.iterrows())
                           if row['conf'] != -1 and row.get('left', 0) > threshold]

        # If insufficient split points, return original segment
        if len(large_positions) < 2:
            return [segment]

        result = []
        # First part until just before second large
        end_pos = large_positions[1] - 1
        if end_pos >= 0:
            result.append(segment.loc[indices[0]:indices[end_pos]])

        # Subsequent parts
        for idx in range(1, len(large_positions)):
            start = large_positions[idx]
            start_idx = indices[start]
            if idx + 1 < len(large_positions):
                end = large_positions[idx + 1] - 1
                end_idx = indices[end] if end >= 0 else start_idx
            else:
                end_idx = indices[-1]
            result.append(segment.loc[start_idx:end_idx])

        return result

    def _get_footnote_lines(self, paragraph_df: pd.DataFrame, footnote_text: str) -> pd.DataFrame:
        """Get the specific lines that make up a footnote using word overlap matching and optional merging logic"""
        footnotes = []
        current_footnote = []
        consecutive_minus_one = 0
        footnote_segments = []
        current_rows = []

        # Step 1: Segment into individual footnotes
        for i, row in paragraph_df.iterrows():
            if row["conf"] == -1:
                consecutive_minus_one += 1
                if consecutive_minus_one == 2:
                    if current_footnote:
                        ref_text = " ".join(current_footnote)
                        footnotes.append(ref_text)
                        footnote_segments.append(paragraph_df.loc[current_rows])
                        current_footnote = []
                        current_rows = []
                    consecutive_minus_one = 0
            else:
                consecutive_minus_one = 0
                if pd.notna(row["text"]) and str(row["text"]).strip():
                    current_footnote.append(str(row["text"]))
                    current_rows.append(i)

        if current_footnote:
            ref_text = " ".join(current_footnote)
            footnotes.append(ref_text)
            footnote_segments.append(paragraph_df.loc[current_rows])

        if not footnotes:
            return pd.DataFrame()


        # Step 2: Merge segments where necessary
        footnote_segments = self._should_merge_footnotes(footnote_segments)


        # Recreate the list of link texts after merging
        footnotes = []
        for i, seg in enumerate(footnote_segments):
            combined_text = " ".join(str(t) for t in seg["text"].tolist() if pd.notna(t))
            footnotes.append(combined_text)

        # Step 3: Try to find best matching segment
        footnote_words = set(w for w in footnote_text.split() if w.strip())

        best_match = None
        best_score = 0

        for idx, ref_df in enumerate(footnote_segments):
            current_text = " ".join(str(t) for t in ref_df["text"].tolist() if pd.notna(t))
            current_words = set(w for w in current_text.split() if w.strip())
            common_words = footnote_words.intersection(current_words)
            overlap_percent = len(common_words) / max(len(footnote_words), 1) * 100


            if overlap_percent > 80 and overlap_percent > best_score:
                best_match = ref_df
                best_score = overlap_percent

        if best_match is not None:
            return best_match

        return footnote_segments[0] if footnote_segments else pd.DataFrame()

    def _get_next_page_first_footnote(self) -> Optional[pd.DataFrame]:
        """Get the first footnote from the next page if it exists"""
        if self.current_page_index + 1 >= len(self.all_pages_data):
            return None

        next_df = self.all_pages_data[self.current_page_index + 1]
        next_page_name = next_df["Page"].iloc[0] if "Page" in next_df.columns else "Unknown"

        next_df = self._validate_and_prepare_dataframe(next_df, next_page_name)
        if next_df is None:
            return None

        paragraph_data = self._split_into_paragraphs(next_df, next_page_name)
        if not paragraph_data:
            return None

        all_footnotes = self.process_paragraphs_with_numerical_data(next_df, next_page_name)

        if not all_footnotes:
            return None

        for ref in all_footnotes:
            footnotes = self._extract_footnotes(ref["data"])
            if footnotes:
                ref_lines = self._get_footnote_lines(ref["data"], footnotes[0])
                if not ref_lines.empty:
                    return ref_lines

        return None

    def _check_width_threshold(self, df: pd.DataFrame) -> bool:
        """Check if the first footnote's total width is within the threshold. Due to combine ore slit footnotes from diferent pages"""

        footnotes = self._extract_footnotes(df)
        page_name = None
        if 'Page' in df.columns and not df['Page'].empty:
            page_name = df['Page'].iloc[0]
        try:
            num = int(''.join(c for c in (page_name or '') if c.isdigit()))
            even = (num % 2 == 0)
        except Exception:
            even = False
        if even:
            width_threshold = self.config.width_threshold_even
        else:
            width_threshold = self.config.width_threshold_odd
        if footnotes:
            ref_lines = self._get_footnote_lines(df, footnotes[0])

            if not ref_lines.empty:
                max_left_index = ref_lines["left"].idxmax()
                max_line = ref_lines.loc[max_left_index]
                text = max_line["text"]
                left_val = float(max_line["left"]) if not pd.isnull(max_line["left"]) else 0.0
                width_val = float(max_line["width"]) if not pd.isnull(max_line["width"]) else 0.0
                total_width = left_val + width_val
                print(f"total_wight: {total_width}, wight_treshold: {width_threshold}, left_val: {left_val}")
                return total_width <= width_threshold
        return False

    def split_combined_first_footnote(self, current_ref_lines: pd.DataFrame, page_name: str):
        current_ref_lines = self._validate_and_prepare_dataframe(current_ref_lines, page_name)
        if current_ref_lines is None:
            return None

        paragraph_data = self._split_into_paragraphs(current_ref_lines, page_name)
        if not paragraph_data:
            return None

        total_first_line_left = 0
        first_footnote = []

        for idx, f_line in current_ref_lines[1:].iterrows():
            if f_line["conf"] == -1:
                break

            total_first_line_left += f_line["left"]
            first_footnote.append(f_line["text"])

        if self.config.total_left >= total_first_line_left:
            #print(f"Position check passed - top: {current_ref_lines['top'].iloc[0]}, left: {current_ref_lines['left'].iloc[0]}, page: {page_name}, total left: {total_first_line_left}")
            return first_footnote

        return None  # Do not need to split

    def _check_footnote_continuation(self, current_ref_lines: pd.DataFrame,
                                      footnotes: List[str], initial_page: Optional[str],
                                      page_name: str) -> bool:
        # this finction chek if referense continue in next page

        if not current_ref_lines.empty and "top" in current_ref_lines.columns and "left" in current_ref_lines.columns:
            last_line = current_ref_lines["top"].max()
            min_left_value = current_ref_lines["left"].min()
            matches = re.findall(r"\d+", page_name)
            page_num = int(matches[0]) if matches else None
            if page_num%2 == 0:
                left_margin_threshold = self.config.left_margin_threshold_even
                #print(f"if left_margin_threshold: {left_margin_threshold}")
            else:
                left_margin_threshold = self.config.left_margin_threshold_odd
                #print(f"else left_margin_threshold: {left_margin_threshold}")

            if (self.config.bottom_margin_min <= last_line <= self.config.bottom_margin_max and
                    min_left_value < left_margin_threshold):

                print(f"Position left check passed for combaine footnotes from diferent pages - top: {last_line}, left: {min_left_value}, page: {page_name}")
                next_page_ref = self._get_next_page_first_footnote()

                if next_page_ref is None:
                    print("No next page footnote found - saving current footnote separately")
                    return False

                next_page_check = self._check_width_threshold(next_page_ref)

                if next_page_check:
                    print(f"Width threshold checks passed - will combine footnotes, page: {page_name}")
                    self.continuing_footnote = footnotes.pop(-1)

                    self.continuing_footnote_page = initial_page if initial_page and not footnotes else page_name
                    return True
                else:
                    #(f"Width threshold check failed - saving footnotes separately, page {page_name}")
                    return False

        return False

    def process_paragraphs_with_numerical_data(self, df: pd.DataFrame, page_name: str) -> List[dict]:
        """Process paragraphs in a sheet to extract footnotes with numerical data and proper validation"""
        all_paragraphs_data = []
        df = self._validate_and_prepare_dataframe(df, page_name)
        if df is None:
            return []
        # print(f"{page_name} process_paragraphs_with_numerical_data")
        paragraph_data = self._split_into_paragraphs(df, page_name)

        # ADD: Calculate statistics for the main text (excluding the last paragraph)
        # Collect words from the main text for statistics
        all_main_text_words = []
        all_main_text_adjusted_heights = []

        # Exclude the last paragraph from the average calculation
        # main_text_paragraphs = paragraph_data[:-1] if len(paragraph_data) > 1 else []

        # print(f"process_paragraphs_with_numerical_data: Processing {len(main_text_paragraphs)} main text paragraphs in page {page_name}")

        # Collect words from the main text (excluding the last paragraph)
        for paragraph_idx, paragraph in enumerate(paragraph_data):
            # Skip the last paragraph (potential footnotes)
            if len(paragraph_data) > 1 and paragraph_idx == len(paragraph_data) - 1:
                continue

            for _, row in paragraph["data"].iterrows():
                if pd.notna(row["height"]) and pd.notna(row["text"]) and str(row["text"]).strip():

                    class WordSpan:
                        def __init__(self, height, text):
                            self.height = height
                            self.text = text

                    word_span = WordSpan(row["height"], str(row["text"]))

                    all_main_text_words.append({
                        'text': str(row["text"]),
                        'height': row["height"]
                    })

                    if self.only_full_line(str(row["text"])):
                        adjusted_font_size = row["height"]
                    else:
                        adjusted_font_size = self.calc_font_size(word_span)

                    all_main_text_adjusted_heights.append(adjusted_font_size)

        # Calculate statistics of the main text
        if all_main_text_adjusted_heights:
            # adjusted_mean = np.mean(all_main_text_adjusted_heights)
            # print(f"adjusted_mean PN: {adjusted_mean:.2f}")
            adjusted_median_paragraph = float(np.mean(all_main_text_adjusted_heights))
            # print(f"Main text adjusted median ND: {adjusted_median_paragraph:.2f}")
        else:
            adjusted_median_paragraph = 0
            # print("No main text found for statistics")

        # ADD: Footnotes checking logic like in _process_paragraphs
        if len(paragraph_data) > 1:
            # Find largest paragraph among main text paragraphs (excluding last)
            largest_paragraph = None
            max_avg_size = 0

            for paragraph_idx, paragraph in enumerate(paragraph_data[:-1]):  # Exclude last paragraph
                avg_size = max(paragraph["avg_height"], paragraph["avg_width"])
                if avg_size > max_avg_size:
                    max_avg_size = avg_size
                    largest_paragraph = paragraph

            last_paragraph = paragraph_data[-1]

            # ADD: Calculate statistics for the last paragraph
            last_paragraph_adjusted_heights = []
            for _, row in last_paragraph["data"].iterrows():
                if pd.notna(row["height"]) and pd.notna(row["text"]) and str(row["text"]).strip():

                    class WordSpan:
                        def __init__(self, height, text):
                            self.height = height
                            self.text = text

                    word_span = WordSpan(row["height"], str(row["text"]))

                    if self.only_full_line(str(row["text"])):
                        adjusted_font_size = row["height"]
                    else:
                        adjusted_font_size = self.calc_font_size(word_span)

                    last_paragraph_adjusted_heights.append(adjusted_font_size)

            if last_paragraph_adjusted_heights:
                last_adjusted_median_ND = float(np.mean(last_paragraph_adjusted_heights))
                # last_adjusted_mean = np.mean(last_paragraph_adjusted_heights)

                # print(f"LAST PARAGRAPH STATISTICS:")
                # print(f"LAST last_adjusted_median_ND: {last_adjusted_median_ND:.2f}")
                # print(f"LAST last_adjusted_mean NP: {last_adjusted_median:.2f}")
            else:
                last_adjusted_median_ND = 0

            # ADD: Font size check (like in _process_paragraphs)
            font_size_check_passed = True
            if (adjusted_median_paragraph > 0 and
                    last_adjusted_median_ND > adjusted_median_paragraph - 0.35):
                # print( f"Font size check FAILED - last paragraph adjusted median {last_adjusted_median_ND:.2f} exceeds main text adjusted median {adjusted_median_paragraph:.2f}")
                font_size_check_passed = False

            # ADD: Check paragraph size (like in _process_paragraphs)
            size_check_passed = False
            if largest_paragraph is not None:
                if (not np.isclose(last_paragraph["avg_height"], largest_paragraph["avg_height"],
                                   atol=self.config.size_tolerance) or
                        not np.isclose(last_paragraph["avg_width"], largest_paragraph["avg_width"],
                                       atol=self.config.size_tolerance)):
                    size_check_passed = True
                    # print("Size check PASSED - last paragraph differs significantly from main text")

            # UPDATE: Process as footnotes only if BOTH tests pass
            if font_size_check_passed and size_check_passed:
                # print("process_paragraphs_with_numerical_data: Processing last paragraph as FOOTNOTES")

                footnotes = self._extract_footnotes(last_paragraph["data"])

                if self.continuing_footnote:
                    if footnotes:
                        footnotes[0] = self.continuing_footnote + footnotes[0]
                        initial_page = self.continuing_footnote_page
                    else:
                        footnotes = [self.continuing_footnote]
                        initial_page = self.continuing_footnote_page
                    self.continuing_footnote = ""
                    self.continuing_footnote_page = None
                else:
                    initial_page = None

                if footnotes:
                    for i, ref in enumerate(footnotes):
                        if i == 0 and initial_page:
                            all_paragraphs_data.append({
                                "page": initial_page,
                                "text": ref,
                                "number": last_paragraph["number"],
                                "avg_height": last_paragraph["avg_height"],
                                "avg_width": last_paragraph["avg_width"],
                                "data": last_paragraph["data"]
                            })
                        else:
                            all_paragraphs_data.append({
                                "page": page_name,
                                "text": ref,
                                "number": last_paragraph["number"],
                                "avg_height": last_paragraph["avg_height"],
                                "avg_width": last_paragraph["avg_width"],
                                "data": last_paragraph["data"]
                            })

                    # print(f"process_paragraphs_with_numerical_data: Successfully processed {len(footnotes)} footnotes")

        return all_paragraphs_data

    def _process_paragraphs(self, df: pd.DataFrame, page_name: str, collected_footnotes: List[dict]):
        """Process paragraphs in a sheet to extract footnotes and main text using word-level font size calculation"""
        df = self._validate_and_prepare_dataframe(df, page_name)
        if df is None:
            return

        paragraph_data = self._split_into_paragraphs(df, page_name)

        # Extract main text before processing footnotes (предварительно)
        main_text = self._extract_main_text(paragraph_data)

        # Flag to track if the last paragraph was processed as a footnote
        last_paragraph_processed_as_footnote = False

        # Collect ALL words from the main text (excluding potential footnotes paragraph)
        all_main_text_words = []
        all_main_text_adjusted_heights = []

        # Determine which paragraphs to exclude (only potential footnotes)
        paragraphs_to_exclude = set()
        if len(paragraph_data) > 1:
            paragraphs_to_exclude.add(len(paragraph_data) - 1)  # Last paragraph INDEX (not count)

        #print(f"Processing {len(paragraph_data)} paragraphs in page {page_name}")
        # print(f"Temporarily excluding paragraphs: {paragraphs_to_exclude}")

        # Collect words from the main text (excluding the last paragraph)
        for paragraph_idx, paragraph in enumerate(paragraph_data):
            if paragraph_idx in paragraphs_to_exclude:
                # print(f"Temporarily skipping paragraph {paragraph_idx} from main text calculation")
                continue

            # Process each word in the paragraph
            for _, row in paragraph["data"].iterrows():
                if pd.notna(row["height"]) and pd.notna(row["text"]) and str(row["text"]).strip():

                    # Создаем объект word_span для calc_font_size
                    class WordSpan:
                        def __init__(self, height, text):
                            self.height = height
                            self.text = text

                    word_span = WordSpan(row["height"], str(row["text"]))

                    # Add the original height
                    all_main_text_words.append({
                        'text': str(row["text"]),
                        'height': row["height"],
                        'paragraph': paragraph_idx
                    })

                    # Calculate adjusted height using calc_font_size and only_full_line
                    if self.only_full_line(str(row["text"])):
                        adjusted_font_size = row["height"]
                    else:
                        adjusted_font_size = self.calc_font_size(word_span)

                    all_main_text_adjusted_heights.append(adjusted_font_size)

        # Calculate general statistics for the main text
        if all_main_text_words:
            all_heights = [word['height'] for word in all_main_text_words]
            # median_height = np.median(all_heights)
            # mean_height = np.mean(all_heights)

            adjusted_median = float(np.mean(all_main_text_adjusted_heights))
            # adjusted_mean = np.mean(all_main_text_adjusted_heights)

            print("__________________")
            print(page_name)
            print(f"ADJ_MEDIAN height PP: {adjusted_median:.2f}")
            # print(f"ADJ_MEAN height PP: {adjusted_mean:.3f}")

        else:
            adjusted_median = 0
            # print(f"No main text words found in {page_name}")

        # Now we process the last paragraph (potential footnotes)
        if len(paragraph_data) > 1:
            last_paragraph = paragraph_data[-1]

            # Calculate statistics for the last paragraph
            last_paragraph_adjusted_heights = []

            for _, row in last_paragraph["data"].iterrows():
                if pd.notna(row["height"]) and pd.notna(row["text"]) and str(row["text"]).strip():

                    class WordSpan:
                        def __init__(self, height, text):
                            self.height = height
                            self.text = text

                    word_span = WordSpan(row["height"], str(row["text"]))

                    if self.only_full_line(str(row["text"])):
                        adjusted_font_size = row["height"]
                    else:
                        adjusted_font_size = self.calc_font_size(word_span)

                    last_paragraph_adjusted_heights.append(adjusted_font_size)

            if last_paragraph_adjusted_heights:
                last_adjusted_median = float(np.mean(last_paragraph_adjusted_heights))
                # last_adjusted_mean = np.mean(last_paragraph_adjusted_heights)

                # print(f"LAST PARAGRAPH STATISTICS:")
                # print(f"LAST adj_median PP: {last_adjusted_median:.2f}")
                print(f"LAST last_adjusted_mean PP: {last_adjusted_median:.2f}")

                # Check 1: font size
                font_size_check_passed = True
                if (adjusted_median > 0 and
                        last_adjusted_median > adjusted_median - 0.35):
                    # print(f"Font size check FAILED - last paragraph adjusted median {last_adjusted_median:.2f} exceeds main text adjusted median {adjusted_median:.2f}")
                    font_size_check_passed = False

                # Check 2: paragraph size
                size_check_passed = False
                if len(paragraph_data) > 1:
                    # Find the paragraph with the largest average size among the main text
                    largest_paragraph = None
                    max_avg_size = 0

                    for paragraph_idx, paragraph in enumerate(paragraph_data[:-1]):  # Исключаем последний
                        if paragraph_idx in paragraphs_to_exclude:
                            continue
                        avg_size = max(paragraph["avg_height"], paragraph["avg_width"])
                        if avg_size > max_avg_size:
                            max_avg_size = avg_size
                            largest_paragraph = paragraph

                    if largest_paragraph is not None:
                        # Check if the last paragraph is a footnote by size comparison
                        if (not np.isclose(last_paragraph["avg_height"], largest_paragraph["avg_height"],
                                           atol=self.config.size_tolerance) or
                                not np.isclose(last_paragraph["avg_width"], largest_paragraph["avg_width"],
                                               atol=self.config.size_tolerance)):
                            size_check_passed = True
                            # print("Size check PASSED - last paragraph differs significantly from main text")

                # SOLUTION: Process as footnotes only if BOTH tests pass
                if font_size_check_passed and size_check_passed:
                    # print("Processing last paragraph as FOOTNOTES")

                    footnotes = self._extract_footnotes(last_paragraph["data"])

                    if page_name == "p01" and footnotes and '*' in footnotes[0]:
                        footnotes.pop(0)

                    # Check for first footnote split using only_full_line
                    first_split_footnote = self.split_combined_first_footnote(last_paragraph["data"], page_name)
                    if first_split_footnote:
                        collected_footnotes.append({
                            "page": page_name,
                            "text": " ".join(first_split_footnote)
                        })
                        if footnotes:
                            footnotes[0] = " ".join(footnotes[0].split()[len(first_split_footnote):])
                            if not footnotes[0].strip():
                                footnotes.pop(0)

                    if self.continuing_footnote:
                        if footnotes:
                            footnotes[0] = self.continuing_footnote + footnotes[0]
                            initial_page = self.continuing_footnote_page
                        else:
                            footnotes = [self.continuing_footnote]
                            initial_page = self.continuing_footnote_page
                        self.continuing_footnote = ""
                        self.continuing_footnote_page = None
                    else:
                        initial_page = None

                    if footnotes:
                        last_ref = footnotes[-1]
                        ref_lines = self._get_footnote_lines(last_paragraph["data"], last_ref)
                        self._check_footnote_continuation(ref_lines, footnotes, initial_page, page_name)

                        # print(f" page_name: {page_name},")
                        # f"\t footnotes: {footnotes}")

                        for i, ref in enumerate(footnotes):
                            if i == 0 and initial_page:
                                # print(f" i: {i},\t initial_page: {initial_page}")
                                collected_footnotes.append(
                                    {"page": initial_page, "text": ref}
                                )
                            else:
                                collected_footnotes.append(
                                    {"page": page_name, "text": ref}
                                )

                        last_paragraph_processed_as_footnote = True
                        # print(f"Successfully processed {len(footnotes)} footnotes from last paragraph")

        # FINAL SAVE MAIN TEXT
        if not last_paragraph_processed_as_footnote and len(paragraph_data) > 1:
            # The last paragraph was NOT processed as a footnote - add it to the main text
            # print("Adding last paragraph to main text (not processed as footnote)")
            last_paragraph_text = self._get_paragraph_text(paragraph_data[-1]["data"])
            if last_paragraph_text:
                if main_text:
                    main_text += "\n\n" + last_paragraph_text
                else:
                    main_text = last_paragraph_text

        # ALWAYS save main text
        self.main_texts[page_name] = main_text if main_text else ""
        # print(f"Final main text length for {page_name}: {len(self.main_texts[page_name])} characters")

    def process_workbook(self, xlsx_path: str) -> Tuple[List[Dict[str, str]], Dict[str, str]]:
        """Process the entire workbook for footnotes and main text

        Returns:
            Tuple containing:
            - List of footnote dictionaries
            - Dictionary of main text by page
        """
        self.all_pages_data = self._extract_data_from_xlsx(xlsx_path)
        all_footnotes = []
        self.main_texts = {}  # Reset main texts

        for i, df in enumerate(self.all_pages_data):
            self.current_page_index = i
            page_name = df["Page"].iloc[0] if "Page" in df.columns else "Unknown"
            self._process_paragraphs(df, page_name, all_footnotes)

        return all_footnotes, self.main_texts

    def extract_meta_info(self, meta_file_path: str) -> dict:
        """
        Extract number of references and last label number from meta JSON file

        Args:
            meta_file_path: Path to the meta JSON file

        Returns:
            Dictionary with meta information
        """
        meta_info = {
            "number_of_references": 0,
            "biggest_label_number": 0,
            "has_meta_file": False
        }

        try:
            if os.path.exists(meta_file_path):
                with open(meta_file_path, "r", encoding="utf-8") as f:
                    meta_data = json.load(f)

                meta_info["has_meta_file"] = True

                # Extract number of references from the specific JSON structure
                # Look for content.references.number_of_references
                if "content" in meta_data and "references" in meta_data["content"]:
                    references_data = meta_data["content"]["references"]
                    if "number_of_references" in references_data:
                        meta_info["number_of_references"] = references_data["number_of_references"]

                # Extract last label number - find the highest reference number
                # Look through reference_blocks and reference_content for numbered labels
                if "content" in meta_data and "references" in meta_data["content"]:
                    references_data = meta_data["content"]["references"]
                    max_label = 0

                    if "reference_blocks" in references_data:
                        for block in references_data["reference_blocks"]:
                            if "reference_content" in block:
                                for ref_content in block["reference_content"]:
                                    if "label" in ref_content and ref_content["label"]:
                                        try:
                                            # Try to convert label to integer
                                            label_num = int(ref_content["label"])
                                            max_label = max(max_label, label_num)
                                        except (ValueError, TypeError):
                                            # If label is not a number, skip it
                                            continue

                    meta_info["biggest_label_number"] = max_label

        except Exception as e:
            print(f"Error reading meta file {meta_file_path}: {e}")

        return meta_info


def clean_bidi_marks_regex(text: str) -> str:
    """Removes bidi characters using regex (more efficient)"""
    if not text:
        return text

    import regex as re
    # Удаляем все Unicode bidi control characters
    cleaned_text = re.sub(r'[\u200e\u200f\u202a-\u202e\u2066-\u2069]', '', text)
    return cleaned_text

def save_footnotes_to_csv(footnotes: List[dict], main_texts: Dict[str, str], output_path: str):
    """Save footnotes and main texts to a CSV file."""
    rows = []

    # First, let's add the main text for each page
    for page_name, text in main_texts.items():
        rows.append({
            "Type": "MainText",
            "Page": page_name,
            "Number": "",
            "Content": text
        })

    # Now let's add all the links
    ref_number = 1
    for ref in footnotes:
        rows.append({
            "Type": "footnote",
            "Page": ref["page"],
            "Number": ref_number,
            "Content": ref["text"]
        })
        ref_number += 1

    df = pd.DataFrame(rows)
    csv_path = output_path.replace(".xml", ".csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    logging.info(f"footnotes and main text saved to {csv_path}")


def save_footnotes_to_xml(footnotes: List[dict], main_texts: Dict[str, str], output_path: str):
    """Save footnotes and main text to XML file with global sequential numbering"""
    root = ET.Element("footnotes")

    # Add main text for each page
    pages_with_content = set(page for ref in footnotes for page in [ref["page"]])
    pages_with_content.update(main_texts.keys())

    # Sort pages for consistent output
    all_pages = sorted(pages_with_content)

    for page_name in all_pages:
        page_element = ET.SubElement(root, "Page")
        page_element.set("name", page_name)

        # Add main text if available - ОЧИЩАЕМ ОТ BIDI МАРКЕРОВ
        if page_name in main_texts:
            main_text_element = ET.SubElement(page_element, "MainText")
            cleaned_main_text = clean_bidi_marks_regex(main_texts[page_name])
            main_text_element.text = cleaned_main_text

    # Add footnotes with global sequential numbering
    ref_number = 1
    pages = {}
    for ref in footnotes:
        if ref["page"] not in pages:
            pages[ref["page"]] = []
        pages[ref["page"]].append(ref)

    for page_name in sorted(pages.keys()):
        page_element = None

        # Find existing page element or create new one
        for existing in root.findall("Page"):
            if existing.get("name") == page_name:
                page_element = existing
                break

        if page_element is None:
            page_element = ET.SubElement(root, "Page")
            page_element.set("name", page_name)

        for ref in pages[page_name]:
            ref_element = ET.SubElement(page_element, "footnote")
            ref_element.set("number", str(ref_number))
            ref_element.set("page", ref["page"])
            # CLEANING FOOTNOTE TEXT FROM BIDI MARKERS
            cleaned_footnote_text = clean_bidi_marks_regex(ref["text"])
            ref_element.text = cleaned_footnote_text
            ref_number += 1

    tree = ET.ElementTree(root)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    logging.info(f"footnotes and main text saved to {output_path}")



def extract_issue_number_from_filename(filename: str) -> str:
    """
    Extract issue number from filename

    Args:
        filename: The filename to extract issue number from

    Returns:
        Issue number as string
    """
    # Remove file extension and common suffixes
    base_name = os.path.splitext(filename)[0]
    base_name = base_name.replace("_footnotes", "").replace("_references", "")

    # Extract numbers from filename
    numbers = re.findall(r'\d+', base_name)
    if numbers:
        return numbers[-1]  # Return the last number found

    return base_name


def extract_journal_name_from_path(file_path: str) -> str:
    """
    Extract journal name from file path

    Args:
        file_path: Full path to the file

    Returns:
        Journal name
    """
    path_lower = file_path.lower()

    journal_mapping = {
        "tarbiz": "Tarbiz",
        "meghillot": "Meghillot",
        "shenmishivri": "Shenmishivri",
        "sidra": "Sidra",
        "leshonenu": "Leshonenu",
        "zion": "Zion"
    }

    for pattern, name in journal_mapping.items():
        if pattern in path_lower:
            return name

    return "Unknown"


def create_csv_report(report_data: list, output_folder: str, journal_name: str):
    """
    Create CSV report with processing results

    Args:
        report_data: List of dictionaries with processing results
        output_folder: Folder to save the report
        journal_name: Name of the journal
    """
    if not report_data:
        #print("No data to create report")
        return

    csv_filename = f"{journal_name}_processing_report.csv"
    csv_path = os.path.join(output_folder, csv_filename)

    # Define CSV headers
    headers = [
        "Issue_Number",
        "Filename",
        "Meta_References_Count",
        "Meta_biggest_label_number",
        "Collected_Footnotes_Count",
        "Has_Meta_File",
        "Processing_Status"
    ]

    try:
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers)
            writer.writeheader()

            for row in report_data:
                writer.writerow(row)

        print(f"CSV report saved to: {csv_path}")

    except Exception as e:
        print(f"Error creating CSV report: {e}")


def main():
    global bottom_margin_max, footnotes_spleat_threshold_odd, merge_footnotes_threshold_odd, \
        footnotes_spleat_threshold_even, merge_footnotes_threshold_even, width_threshold_even, width_threshold_odd, \
        bottom_margin_min, left_margin_threshold_even, left_margin_threshold_odd, total_left

    # Ask user for processing mode
    import tkinter as tk
    from tkinter import messagebox, filedialog

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Ask user to choose processing mode
    choice = messagebox.askyesnocancel(
        "Processing Mode",
        "Choose processing mode:\n\nYes = Process entire folder\nNo = Process single file\nCancel = Exit"
    )

    if choice is None:  # User clicked Cancel
        root.destroy()
        return

    if choice:  # Process entire folder
        input_folder_path = filedialog.askdirectory(title="Select Input Folder")
        if not input_folder_path:
            root.destroy()
            return

        output_folder_path = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder_path:
            root.destroy()
            return

        meta_folder_path = filedialog.askdirectory(title="Select Metadata Folder")
        if not meta_folder_path:
            root.destroy()
            return

        xlsx_files = glob.glob(os.path.join(input_folder_path, "*.xlsx"))

    else:  # Process single file
        single_file = filedialog.askopenfilename(
            title="Select XLSX File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not single_file:
            root.destroy()
            return

        output_folder_path = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder_path:
            root.destroy()
            return

        # Set input folder path to the directory of the selected file
        input_folder_path = os.path.dirname(single_file)
        meta_folder_path = None  # No metadata for single file processing
        xlsx_files = [single_file]

    root.destroy()  # Clean up the tkinter root window

    xlsx_files.sort()

    total_processed_files = 0
    total_footnotes_found = 0
    total_meta_references = 0

    # Extract journal name for the report
    journal_name = extract_journal_name_from_path(input_folder_path)

    # List to store report data
    report_data = []

    for xlsx_file in xlsx_files:
        filename = os.path.basename(xlsx_file)
        base_name = os.path.splitext(filename)[0]
        issue_number = extract_issue_number_from_filename(filename)

        mode = 'printed' if 'ocr-tess-printed' in input_folder_path.lower() else 'scanned'
        path = input_folder_path.lower()

        # Set parameters based on journal and mode (existing logic)
        if 'tarbiz' in path:
            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1605, 1670
                left_margin_threshold_even, left_margin_threshold_odd = 195, 295
                width_threshold_even, width_threshold_odd = 1070, 1160
                merge_footnotes_threshold_even, merge_footnotes_threshold_odd = 1050, 1080
                footnotes_spleat_threshold_even, footnotes_spleat_threshold_odd = 1070, 1170
                total_left = 7200 #?
            else:  # scanned
                bottom_margin_min, bottom_margin_max = 1605, 1695
                left_margin_threshold_even = left_margin_threshold_odd = 195
                width_threshold_even = width_threshold_odd = 1095
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1070
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1085
                total_left = 7200


        elif 'meghillot' in path:
            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1660, 1670
                left_margin_threshold_even = left_margin_threshold_odd = 220
                width_threshold_even = width_threshold_odd = 1025
                merge_footnotes_threshold_even, merge_footnotes_threshold_odd = 1050, 1140
                footnotes_spleat_threshold_even, footnotes_spleat_threshold_odd = 1080, 1150
                total_left = 6700  # ?
            else:
                bottom_margin_min, bottom_margin_max = 1670, 1680
                left_margin_threshold_even = left_margin_threshold_odd = 220
                width_threshold_even = width_threshold_odd = 1030
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1027
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1050
                total_left = 6700


        elif 'shenmishivri' in path:

            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1645, 1720  # Slightly adjusted for scanned variance
                left_margin_threshold_even = left_margin_threshold_odd = 208  # Slightly reduced for scanned
                width_threshold_even = width_threshold_odd = 1075  # Slightly reduced for scanned
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1045  # More conservative merging
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1055  # More conservative splitting
                total_left = 7000  # Reduced for scanned documents

            else:  # scanned - TUNED PARAMETERS
                bottom_margin_min, bottom_margin_max = 1645, 1720  # Slightly adjusted for scanned variance
                left_margin_threshold_even = left_margin_threshold_odd = 208  # Slightly reduced for scanned
                width_threshold_even = width_threshold_odd = 1075  # Slightly reduced for scanned
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1045  # More conservative merging
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1055  # More conservative splitting
                total_left = 7000  # Reduced for scanned documents

        elif 'sibra' in path:
            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1680, 1712
                left_margin_threshold_even, left_margin_threshold_odd = 195, 295
                width_threshold_even = width_threshold_odd = 1090
                merge_footnotes_threshold_even, merge_footnotes_threshold_odd = 1050, 1140
                footnotes_spleat_threshold_even, footnotes_spleat_threshold_odd = 1080, 1150
                total_left = 7200  # ?
            else:
                bottom_margin_min, bottom_margin_max = 1680, 1712
                left_margin_threshold_even = left_margin_threshold_odd = 195
                width_threshold_even = width_threshold_odd = 1090
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1070
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1080
                total_left = 7200  # ?

        elif 'lecohotenu' in path:
            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1655, 1675
                left_margin_threshold_even = left_margin_threshold_odd = 220
                width_threshold_even = width_threshold_odd = 1080
                merge_footnotes_threshold_even, merge_footnotes_threshold_odd = 1050, 1140
                footnotes_spleat_threshold_even, footnotes_spleat_threshold_odd = 1070, 1140
                total_left = 6700  # ?
            else:
                bottom_margin_min, bottom_margin_max = 1655, 1675
                left_margin_threshold_even = left_margin_threshold_odd = 220
                width_threshold_even = width_threshold_odd = 1060
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 950
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1050
                total_left = 7200  # ?

        elif 'zion' in path:
            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1680, 1698
                left_margin_threshold_even = left_margin_threshold_odd = 225
                width_threshold_even = width_threshold_odd = 1080
                merge_footnotes_threshold_even, merge_footnotes_threshold_odd = 1050, 1140
                footnotes_spleat_threshold_even, footnotes_spleat_threshold_odd = 1080, 1150
                total_left = 7200  # ?
            else:
                bottom_margin_min, bottom_margin_max = 1680, 1698
                left_margin_threshold_even = left_margin_threshold_odd = 225
                width_threshold_even = width_threshold_odd = 1080
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1070
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1080
                total_left = 7200  # ?

        else:
            # Default settings
            if mode == 'printed':
                bottom_margin_min, bottom_margin_max = 1671, 1680
                left_margin_threshold_even = left_margin_threshold_odd = 220
                width_threshold_even = width_threshold_odd = 1077
                merge_footnotes_threshold_even, merge_footnotes_threshold_odd = 1050, 1140
                footnotes_spleat_threshold_even, footnotes_spleat_threshold_odd = 1080, 1150
                total_left = 7200
            else:
                bottom_margin_min, bottom_margin_max = 1671, 1680
                left_margin_threshold_even = left_margin_threshold_odd = 220
                width_threshold_even = width_threshold_odd = 1077
                merge_footnotes_threshold_even = merge_footnotes_threshold_odd = 1070
                footnotes_spleat_threshold_even = footnotes_spleat_threshold_odd = 1080
                total_left = 7200

        # Check for JSON with metadata (only for folder processing)
        meta_info = {"number_of_references": 0, "biggest_label_number": 0, "has_meta_file": False}

        if meta_folder_path:  # Only check metadata for folder processing
            json_file = os.path.join(meta_folder_path, base_name + ".json")
            if os.path.exists(json_file):
                with open(json_file, "r", encoding="utf-8") as jf:
                    meta_data = json.load(jf)
                if meta_data.get("skipped", False) is True:
                    logging.info(f"File {filename} skipped according to meta file {base_name}.json")
                    continue

                # Extract meta information
                processor = footnoteProcessor(footnoteConfig(exclusion_phrases=[]))
                meta_info = processor.extract_meta_info(json_file)

        processor = footnoteProcessor(footnoteConfig(
            exclusion_phrases=["https://about,jstor.org/terms", "[תרביץ", "(תרביץ", "https://about.jstor.org/terms",
                               "https://aboutjstor.org/terms"],
            start_row=1,
            bottom_margin_min=bottom_margin_min,
            bottom_margin_max=bottom_margin_max,
            left_margin_threshold_even=left_margin_threshold_even,
            left_margin_threshold_odd=left_margin_threshold_odd,
            width_threshold_even=width_threshold_even,
            width_threshold_odd=width_threshold_odd,
            merge_footnotes_threshold_even=merge_footnotes_threshold_even,
            merge_footnotes_threshold_odd=merge_footnotes_threshold_odd,
            footnotes_spleat_threshold_even=footnotes_spleat_threshold_even,
            footnotes_spleat_threshold_odd=footnotes_spleat_threshold_odd,
            total_left=total_left
        ))

        # Initialize row data for CSV report
        row_data = {
            "Issue_Number": issue_number,
            "Filename": filename,
            "Meta_References_Count": meta_info["number_of_references"],
            "Meta_biggest_label_number": meta_info["biggest_label_number"],
            "Collected_Footnotes_Count": 0,
            "Has_Meta_File": meta_info["has_meta_file"],
            "Processing_Status": "Processed"
        }

        try:
            # Process the workbook
            all_footnotes, main_texts = processor.process_workbook(xlsx_file)

            # Save in xml and csv
            output_xml = os.path.join(output_folder_path, base_name + "_footnotes.xml")
            save_footnotes_to_xml(all_footnotes, main_texts, output_xml)
            save_footnotes_to_csv(all_footnotes, main_texts, output_xml)

            ref_count = len(all_footnotes)
            total_footnotes_found += ref_count
            total_processed_files += 1
            total_meta_references += meta_info["number_of_references"]

            # Update row data
            row_data["Collected_Footnotes_Count"] = ref_count

            print("====================================")
            print(f"File: {filename}")
            print(f"Issue Number: {issue_number}")
            print(f"Number of references (meta): {meta_info['number_of_references']}")
            print(f"Last label number (meta): {meta_info['biggest_label_number']}")
            print(f"Collected footnotes: {ref_count}")
            print(f"Main text pages: {len(main_texts)}")
            print("====================================")

        except Exception as e:
            print(f"Error processing {filename}: {e}")
            row_data["Processing_Status"] = f"Error: {str(e)}"

        report_data.append(row_data)

    create_csv_report(report_data, output_folder_path, journal_name)

    # general statistics
    print("\n====================================")
    print("TOTAL SUMMARY:")
    print(f"Journal: {journal_name}")
    print(f"Processed files: {total_processed_files}")
    print(f"Total footnotes found: {total_footnotes_found}")
    print(f"Total meta references: {total_meta_references}")
    if total_processed_files > 0:
        print(f"Average footnotes per file: {total_footnotes_found / total_processed_files:.2f}")
        print(f"Average meta references per file: {total_meta_references / total_processed_files:.2f}")
    print("====================================")


if __name__ == "__main__":
    main()


