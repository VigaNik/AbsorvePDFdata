"""
Processing Raw PDF files from JSTOR
"""

# PDF processing package
# See the documenation of the page object in 
# https://pymupdf.readthedocs.io/en/latest/textpage.html

from typing import List
import pandas
import fitz
from collections import Counter, namedtuple

from re import Match
import sys
import json
from xml.dom.minidom import Document, Element
from unicodedata import bidirectional
import regex as re
import traceback

import math

from pathlib import Path

from openpyxl import Workbook, worksheet
from openpyxl.styles import Font, Alignment, NamedStyle
from openpyxl.styles.borders import Border, Side

from xml.dom import getDOMImplementation

from traitlets import Bool

from . import LineType
from . import load_metadata_url, print_to_string, add_span_to_blocks
from . import split_words_by_col, split_words_by_regex,is_centered
from . import uni_ltr, uni_rtl, ref_subtypes_regex, quoteTranslate
from . import is_asterik_comment, get_text_letters, typeset_words
from . import abbrev_indentation, abbrev_single_line_indentation, column_tolerance
from . import dom_impl

class paper_abbrev:

    def __init__(self, paper_ocr, journal_name, pdf_file:Path, jstor_metadata_file:Path, trace_file, pdf_dir=None):

        if pdf_dir is None:
            pdf_dir = pdf_file.parent()
        
        self.journal_name = journal_name
        self.pdf_file = pdf_file
        self.pdf_file_rel_pdf = pdf_file.relative_to(pdf_dir)
        
        self.paper_ocr = paper_ocr
        self.top = None
        self.bottom = None

        # If there is no OCR data, set the skip flag
        if paper_ocr is None:
            self.skip = f"Missing OCR data for {pdf_file.name}"
            return

        self.metadata = None
        self.metadata_jstor = None
        self.trace_file = trace_file
        self.trace = ''

        self.abbrev_list_found = False
        self.references = dict()
        self.reference_labels = dict()
        self.abbrev_lines = list()
        self.abbrev = None

        self.skip = None
        self.next_fn = 0
        self.line_type = None
        self.abbrev_list_found = False

        self.paper_pdf_document = fitz.open(pdf_file)
        self.paper_page_num = len(self.paper_pdf_document)
        cover_page = self.paper_pdf_document.load_page(0)
        self.get_paper_metadata(jstor_metadata_file, cover_page)

    def skip_paper(self):
        return self.skip
    
    """
    print the trace to the trace file (if it was specified in the constructor)
    """
    def print_trace(self):
        if self.trace_file is not None:
            if self.trace != '':
                print(self.trace, file=self.trace_file)
            self.trace_file.flush()

    def canonize_reference_labels(self):

        refs = self.metadata_jstor["references"]["reference_blocks"]

        self.references = dict()
        self.reference_labels = dict()
        
        self.reference_labels['קיצורים אחרים'] = 'קיצורים אחרים'
        self.reference_labels['קיצורים נוספים'] = 'קיצורים נוספים'

        for refList in refs:
            refType = refList["title"]
            if refType == '':
                refType = 'הערות שוליים'
            refType_orig = refType

            # convert variations of reference types with similar meaning to a single reference type label

            if re.search(r'^([א-ת]+ )?ה?(קיצורים|ציונים|קיצורים וציונים)( ה?ביבי?ליו?גרא?פי+ים)?$', refType):
                refType = "קיצורים ביבליוגרפים"
            if 'קיצור' in refType and 'מקורות' in refType:
                refType = "קיצורים ביבליוגרפים"
            elif refType in {"ספרות", "מראי מקום", "ביבליוגרפיה", "ביבליוגראפיה", "רשימת מקורות", "המקורות", "מקורות", "רשימת המקורות", "רשימה ביבליוגרפית", "רשימה ביבליוגראפית"}:
                refType = "ביבליוגרפיה"
            elif 'footnotes' in refType.lower():
                refType = 'הערות שוליים'

            refType = re.sub(r'(^\[|\]$)', '', refType)

            if refType != 'הערות שוליים':
                self.references[refType] = refList["reference_content"]
                self.reference_labels[refType_orig] = refType

    def get_paper_metadata(self, jstor_metadata_file:Path, cover_page:fitz.Page):

        if jstor_metadata_file.exists():
            # load the metadata from a file
            metadata_str = jstor_metadata_file.read_text(encoding='utf-8', errors='replace')
        else:
            # load the metadata from a URL and save them to a file

            cover_page_text = cover_page.get_text("text")
            stable_urls = [ u for u in re.findall(r'http[^\s]*', cover_page_text) if 'stable' in u]
            if len(stable_urls) == 0:
                return
            
            # make sure the stable URL points to the new domain with https
            stable_url = stable_urls[0].replace('http://', 'https://').replace('jstor.com', 'jstor.org')
            
            # replace with the URL for the metadata
            metadata_url = stable_url.replace('stable', 'content-service/content-data')
            md_response = load_metadata_url(metadata_url)

            # md_response contains either 'status' or 'text'
            if 'status' in md_response:
                self.skip = f'Could not get URL {metadata_url}, status: {md_response["status"]}' 
                if md_response['status'] == 403:
                    self.trace += print_to_string(self.skip)
                    print(self.skip) 
                    sys.exit(1)
                return

            # Save the metadata to a file
            metadata_str = md_response['text']
            jstor_metadata_file.write_text(metadata_str, errors='replace')

        # If a PDF file was skipped, the metadata file starts with 'skipped:'

        if metadata_str.startswith('skipped: '):
            self.skip = metadata_str.strip()
            return

        metadata_jstor = json.loads(metadata_str)['content']

        metadata = dict()
        abstract_list = metadata_jstor['abstract']
        if len(abstract_list) == 0:
            abstract = ''
        else:
            abstract = abstract_list[0].strip()
        abstract = re.sub(r'\[illegible text\]', '`', abstract)
        metadata['abstract'] = re.sub(r' +', ' ', abstract)

        authors = metadata_jstor['authors']
        la_half = int(len(authors)/2)
        metadata['authors_heb'] = ', '.join(authors[0:la_half])
        metadata['authors_eng'] = ', '.join(authors[la_half:])

        title = metadata_jstor['displayTitle'].strip()
        title = re.sub(r'\[illegible text\]', '`', title)
        title = re.sub(r'\s+', ' ', title)

        sep = ''
        if re.search(r'[\p{Script=Latin}]', title) == None:
            title_heb = title
            title_eng = ''
        elif re.search(r'[א-ת]', title) == None:
            title_heb = ''
            title_eng = title
        else:
            if ' / ' in title:
                sep = ' / '
                title_parts = title.split(sep)
            elif ' - ' in title:
                sep = ' - '
                title_parts = title.split(sep)
            elif re.search(r'\p{Script=Latin}{2}/ [א-ת]{2}', title) is not None:
                (title_eng, sep_str, title_heb) = re.split(r'(\p{Script=Latin}{2}/ [א-ת]{2})', title)[:3]
                title_eng += sep_str[:2]
                title_heb = sep_str[-2:]+title_heb
                title_parts = [title_eng, title_heb]
            else:
                title_parts = []
                self.trace += print_to_string(f'Setting title"{title}" as hebrew in file {self.pdf_file.name}')
                title_heb = title
                title_eng = ''
        
            num_title_parts = len(title_parts)
            if num_title_parts != 2:
                self.trace += print_to_string(f'title parts:{num_title_parts} sep: "{sep}" title: "{title}"')
        
            if num_title_parts == 2:
                (title_eng, title_heb) = title_parts
            elif num_title_parts == 1:
                if re.search(r'[\'\p{Script=Latin}]{3}', title) == None:
                    title_heb = title
                    title_eng = ''
                else:
                    title_heb = ''
                    title_eng = title
            elif num_title_parts > 2:
                title_eng = title_parts[0]
                for i in range(1, num_title_parts):
                    if re.search(r'[\'\p{Script=Latin}]{3}', title_parts[i]):
                        title_eng += sep+title_parts[i]
                    else:
                        break
                title_heb = sep.join(title_parts[i:])

        if len(title_heb) > 0 and len(title_eng) > 0:
            h_in_e = len(re.findall(r'[א-ת]', title_eng))
            h_in_h = len(re.findall(r'[א-ת]', title_heb))
            l_in_e = len(re.findall(r'\p{L}', title_eng))
            l_in_h = len(re.findall(r'\p{L}', title_heb))

            r_h = h_in_h/l_in_h
            r_e = h_in_e/l_in_e

            # switch title_heb and title_eng according to the letter frequency

            if r_e > r_h:
                temp = title_heb
                title_heb = title_eng
                title_eng = temp
                self.trace += print_to_string(f'switched heb/eng heb:"{title_heb}" eng:"{title_eng}" in file {self.pdf_file.name}')
                
            heb_cont_in_eng_words:Match = re.search(r'\p{Script=Latin}{2}: [א-ת\s]{5}', title_eng)
            if heb_cont_in_eng_words is not None:
                additional_heb = title_eng[heb_cont_in_eng_words.start()+4:]
                if not title_heb.endswith(additional_heb):
                    title_heb += ': '+additional_heb
                    title_eng = title_eng[:heb_cont_in_eng_words.start()+2]
                    self.trace += print_to_string(f'moved heb part from eng word:"{title_heb}" eng:"{title_eng}" in file {self.pdf_file.name}')
                
            heb_cont_in_eng_paren:Match = re.search(r'\): [א-ת\s]{9}', title_eng)
            if heb_cont_in_eng_paren is not None:
                additional_heb = title_eng[heb_cont_in_eng_paren.start()+3:]
                if not title_heb.endswith(additional_heb):
                    title_heb += ': '+additional_heb
                    title_eng = title_eng[:heb_cont_in_eng_paren.start()+1]
                    self.trace += print_to_string(f'moved heb part from eng paren:"{title_heb}" eng:"{title_eng}" in file {self.pdf_file.name}')

        # We are interested only in full papers in Hebrew.
        # Filter out other types of texts

        if re.search('Abstracts', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.search('Summary', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.search('Errata', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.search('Preface', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.search('Front Matter', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.search('Back Matter', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.search('Books Received', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.match('Review', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if re.match('(English )?Summaries', title_eng, flags=re.IGNORECASE):
            self.skip = title

        if title.startswith('['):
            self.skip = title

        if title_heb == 'שער קדמי':
            self.skip = title

        if title_heb == 'שער אחורי':
            self.skip = title

        if title_heb == 'חומר קדמי':
            self.skip = title

        if title_heb == 'חומר אחורי':
            self.skip = title

        if title_heb == 'הקדמה':
            self.skip = title

        if title_heb == 'פתח דבר':
            self.skip = title

        if title_heb == 'אחרית דבר':
            self.skip = title

        if title_heb == 'אחרית דבר':
            self.skip = title

        if re.search(r'^תקצירים', title_heb):
            self.skip = title

        if re.search("ספרים ש.תקבלו ב?מערכת", title):
            self.skip = title

        if re.search("הערת העורך|הערת המערכת|מאת המערכת", title):
            self.skip = title

        if self.skip is not None:
            return

        page_range = metadata_jstor['pageRange']
        if page_range is None:
            self.skip = 'no pages'
            return

        metadata['title_heb'] = title_heb
        metadata['title_eng'] = title_eng

        page_range = page_range.strip()
        single_page_match = re.search(r'^p\. (.+)$', page_range, flags=re.IGNORECASE)
        if single_page_match is not None:
            metadata['page_from'] = single_page_match.group(1)
            metadata['page_to'] = metadata['page_from']
        else: # page range
            page_range_match = re.search(r'^pp. (\P{Pd}+)\p{Pd}(\P{Pd}+)', page_range, flags=re.IGNORECASE)
            if page_range_match is not None:
                (metadata['page_from'], metadata['page_to']) = page_range_match.groups()
            else:
                metadata['page_from'] = page_range
                metadata['page_to'] = metadata['page_from']
                self.trace += print_to_string(f'unexpected page range: "{page_range}" in file {self.pdf_file.name}')

        metadata['volume'] = ''
        metadata['num'] = ''

        if 'volume' in metadata_jstor:
            volume = metadata_jstor['volume']
            if volume is None:
                volume = ''
            vol_match = re.search(r'(vol\.|כרך)\s+([^\s]+)', volume, flags=re.IGNORECASE)
            if vol_match is not None:
                metadata['volume'] = vol_match.group(2)
            else:
                metadata['volume'] = volume

        if 'issue' in metadata_jstor:
            issue_num = metadata_jstor['issue']
            if issue_num is None:
                issue_num = ''
            num_match = re.search(r'(num\.|no\.|חוברת)\s+([^\s]+)', issue_num, flags=re.IGNORECASE)
            if num_match is not None:
                metadata['num'] = num_match.group(2)
            else:
                metadata['num'] = issue_num

        """
        heb_year_match = re.search(r'[א-ת]+(?:"|&quot;)[א-ת])$')
        metadata['heb_month'] = ''
        metadata['heb_year'] = ''
        if heb_year_match is not None:
            metadata['heb_year'] = heb_year_match.group()
            heb_pub_date = re.sub(' '+self.paper_info['heb_year']+'$', '', heb_pub_date)
        metadata['heb_month'] = heb_pub_date
        """

        metadata['pubdate'] = metadata_jstor['publishedDate']
        metadata['year'] = metadata_jstor['year']
        metadata['url'] = 'https://www.jstor.org'+metadata_jstor['stable']
        metadata['scanned'] = 'לא' if metadata_jstor["hasRendition"] else 'כן'

        self.metadata = metadata
        self.metadata_jstor = metadata_jstor

        self.canonize_reference_labels()

        return metadata_jstor

    # check whether the distance from the first word to the second
    # is larger than the discance between words in the running text

    def line_columns(self, line):

        # the line words are already sorted
        # top to bottom and right to left

        if len(line["words"]) == 1:
            return [1]
        
        space_widths = Counter()
        for (i, sw) in enumerate(line["words"]):
            if i > 0:
                space_width = line["words"][i-1].left-line["words"][i].left-line["words"][i].width
                space_widths.update([space_width])
        
        # print(space_widths)
        self.trace += print_to_string(' spaces between words: '+str(space_widths))
        pass

    def divide_page(self, lines:List):
    
        lines_y_top = [ l["bbox"][1] for l in lines]
        lines_y_bottom = [ l["bbox"][3] for l in lines]

        global_page_height = self.top-self.bottom
        page_height = lines_y_bottom[-1]-lines_y_top[0]
        if page_height <= global_page_height*0.7:
            return -1
        
        vert_spaces = [lines_y_top[i]-lines_y_bottom[i-1] for i in range(1, len(lines_y_bottom))]
        if len(vert_spaces) == 0:
            pass
        
        av = sum(vert_spaces)/len(vert_spaces)
        av2 = sum([ s*s for s in vert_spaces])/len(vert_spaces)
        vert_sd = math.sqrt(av2-av*av)/len(vert_spaces)

        sep_line = -1
        last_space = lines_y_top[-1]-lines_y_bottom[-2]

        space_prev = 0
        for i in range(len(lines)-1, 2, -1):
            space_i = lines_y_top[i]-lines_y_bottom[i-1]
            # if space_i-space_prev >= 10:
            if space_i-space_prev-av > 10*vert_sd:
                text_test_ref_label = re.sub(r'\P{L}+$', '', lines[i]["text"])
                if text_test_ref_label in self.reference_labels:
                    return -1
                if re.search(ref_subtypes_regex, text_test_ref_label):
                    return -1
                text_test_ref_label_m1 = re.sub(r'\P{L}+$', '', lines[i-1]["text"])
                if text_test_ref_label_m1 in self.reference_labels:
                    return -1
                if re.search(ref_subtypes_regex, text_test_ref_label_m1):
                    return -1
                if i+1 < len(lines):
                    text_test_ref_label_1 = re.sub(r'\P{L}+$', '', lines[i+1]["text"])
                    if text_test_ref_label_1 in self.reference_labels:
                        return -1
                    if re.search(ref_subtypes_regex, text_test_ref_label_1):
                        return -1
                if i+2 < len(lines):
                    text_test_ref_label_2 = re.sub(r'\P{L}+$', '', lines[i+2]["text"])
                    if text_test_ref_label_2 in self.reference_labels:
                        return -1
                    if re.search(ref_subtypes_regex, text_test_ref_label_2):
                        return -1
                if is_centered(lines[i], self.page_width):
                    return -1
                
                sep_line = i
                break
            space_prev = space_i

        return sep_line

    # Analyze bibliographic abbreviation data

    def get_abbrev(self) -> list:

        if self.abbrev is not None:
            return self.abbrev

        if not self.paper_has_abbrev():
            return list()
        
        self.abbrev = list()

        right_margins = Counter()

        abbrev_text_dir = 'ltr'
        for abbrev_line in self.abbrev_lines:
            spans = abbrev_line["spans"]
            right_margins.update([spw.left+spw.width for sp in spans for spw in sp["words"]])
            if re.search(r'[א-ת]', abbrev_line["text"]):
                abbrev_text_dir = 'rtl'
        
        print(f'{len(self.abbrev_lines)} abbrev. lines', right_margins.most_common(4))
        
        bins = dict()
        col_toler = column_tolerance[self.journal_name]

        if abbrev_text_dir == 'rtl':
            bbox_ent = 2
            page_margin = max([line["bbox"][bbox_ent] for line in self.abbrev_lines]) 
        else:
            bbox_ent = 0
            page_margin = min([line["bbox"][bbox_ent] for line in self.abbrev_lines]) 

        # self.trace += print_to_string(f"abbrev_text_dir={abbrev_text_dir}, page_margin={page_margin}")

        first_ref_lines = [l['text'] for l in self.abbrev_lines if abs(l["bbox"][bbox_ent]-page_margin) <= col_toler]
        refs_with_dash = [t for t in first_ref_lines if re.search(r'=|:|\p{Pd}{2,3}\s', t) is not None]
        refs_with_single_dash = [t for t in first_ref_lines if re.search(r'^\p{L}+,( \p{L}+){,3} \p{Pd}', t) is not None]
        
        single_dash = False

        if len(first_ref_lines)-len(refs_with_dash) <= 3 or len(refs_with_dash)/len(first_ref_lines) > 0.45:
            num_columns = 1
        elif len(first_ref_lines)-len(refs_with_single_dash) <= 3 or len(refs_with_single_dash)/len(first_ref_lines) > 0.45:
            num_columns = 1
            single_dash = True
        else:
            for (margin, count) in right_margins.items():
                found = False
                for b in bins.keys():
                    if abs(b-margin) <= col_toler:
                        bins[b] += count
                        found = True
                        break
                if not found:
                    bins[margin] = count

            bins_by_count = sorted([bi for bi in bins.items()], key=lambda bi:bi[1], reverse=True)
            print(bins_by_count[:5])

            col_info = -1

            if abs(bins_by_count[0][0]-page_margin) > col_toler:
                num_columns = 2
                col_info = bins_by_count[0][0]
                self.trace += print_to_string(f"Abbreviation abbrev_text_dir={abbrev_text_dir}, page_margin={page_margin} column bins 2 cols case 1:", str(bins_by_count[:5]))
            elif abs(bins_by_count[0][1]-bins_by_count[1][1]) <= 2:
                col_info = bins_by_count[1][0]
                num_columns = 2
                self.trace += print_to_string(f"Abbreviation abbrev_text_dir={abbrev_text_dir}, page_margin={page_margin} column bins 2 cols case 2:", str(bins_by_count[:5]))
            else:
                num_columns = 1
                self.trace += print_to_string(f"Abbreviation abbrev_text_dir={abbrev_text_dir}, page_margin={page_margin} column bins 1 col case 3:", str(bins_by_count[:5]))

        abbrev = dict()
        had_rtl = False

        for abbrev_line in self.abbrev_lines:
            if 'Bauer' in abbrev_line["text"]:
                pass

            spans = abbrev_line["spans"]
            words = abbrev_line["words"]
            if len(words) == 0:
                continue
            
            # for printing only
            abbrev_line['text'] = abbrev_line['text'].translate(quoteTranslate)

            if abbrev_text_dir == 'rtl':
                line_margin = max([w.left+w.width for w in words])
                first_span_x_offset = max(abbrev_line["page_right_margin"]-line_margin, 0)
            else:
                line_margin = min([w.left for w in words]) 
                first_span_x_offset = max(line_margin-abbrev_line["page_left_margin"], 0)

            self.trace += print_to_string(f"Line '{abbrev_line['text']}' first_span_x_offset={first_span_x_offset}")

            if num_columns == 2:
                words = typeset_words(words, had_rtl)
                if first_span_x_offset >= abbrev_indentation[self.journal_name]:
                    # continuation of abbrev information
                    abbrev["info"] += ' '+' '.join([w.text for w in words])
                else:
                    # check for indented label
                    if (first_span_x_offset > 20) and (first_span_x_offset < 40):
                        (label_text, info_text) = split_words_by_col(words, col_info+col_toler, had_rtl)
                        abbrev["label"] += ' '+label_text
                        abbrev["info"] += ' '+info_text
                    else:
                        # a new abbreviation
                        had_rtl = False
                        (label_text, info_text) = split_words_by_col(words, col_info+col_toler, had_rtl)
                        if "label" in abbrev:
                            self.trace += print_to_string(f'Abbrev: "{abbrev["label"]}" Info: {abbrev["info"]}')
                            self.abbrev.append(abbrev)
                        abbrev = dict()
                        abbrev["label"] = label_text
                        abbrev["info"] = info_text
            else: # num_columns == 1
                if first_span_x_offset >= abbrev_single_line_indentation[self.journal_name]:
                    # continuation of abbrev information
                    words = typeset_words(words, had_rtl)
                    if 'info' not in abbrev:
                        continue
                    abbrev["info"] += ' '+' '.join([w.text for w in words])
                elif re.search(r'24352704', self.pdf_file.name) and not re.search(r' = ', abbrev_line['text']):
                    # file specific: in this file the continutations are not indented
                    # continuation of abbrev information
                    words = typeset_words(words, had_rtl)
                    if 'info' not in abbrev:
                        continue
                    abbrev["info"] += ' '+' '.join([w.text for w in words])
                else:
                    # a new abbreviation
                    had_rtl = False
                    if "label" in abbrev:
                        self.trace += print_to_string(f'Abbrev: "{abbrev["label"]}" Info: {abbrev["info"]}')
                        self.abbrev.append(abbrev)
                    if re.search(r' = ', abbrev_line['text']):
                        if abbrev_text_dir == 'rtl':
                            (label_text, info_text, self.trace) = split_words_by_regex(words, had_rtl, r'=', self.trace)
                        else:
                            (info_text, label_text, self.trace) = split_words_by_regex(words, had_rtl, r'=', self.trace)
                    elif single_dash:
                        if abbrev_text_dir == 'rtl':
                            (label_text, info_text, self.trace) = split_words_by_regex(words, had_rtl, r'\p{Pd}{1,3}', self.trace)
                        else:
                            (info_text, label_text, self.trace) = split_words_by_regex(words, had_rtl, r'\p{Pd}{1,3}', self.trace)
                        
                        # The OCR can easily miss the single dash
                        if info_text == '':
                            label_text = 'unknown'
                            info_text = abbrev_line['text']
                    elif re.search(r'\p{L}: ', abbrev_line['text']):
                        if abbrev_text_dir == 'rtl':
                            (label_text, info_text, self.trace) = split_words_by_regex(words, had_rtl, r'.*\p{L}:', self.trace, include_matching=True)
                        else:
                            (info_text, label_text, self.trace) = split_words_by_regex(words, had_rtl, r'.*\p{L}:', self.trace, include_matching=True)
                    elif re.search(r'\s(\p{Pd}{2,3}|—)\s', abbrev_line['text']):
                        if abbrev_text_dir == 'rtl':
                            (label_text, info_text, self.trace) = split_words_by_regex(words, had_rtl, r'\p{Pd}{2,3}|—', self.trace)
                        else:
                            (info_text, label_text, self.trace) = split_words_by_regex(words, had_rtl, r'\p{Pd}{2,3}|—', self.trace)
                    else:
                        if re.search(ref_subtypes_regex, abbrev_line["text"]):
                            self.trace += print_to_string(f'Ignoring line "{abbrev_line["text"]}"')
                            label_text = None
                        else:
                            self.trace += print_to_string(f'!! Cannot analyze abbrev line "{abbrev_line["text"]}"')
                            label_text = 'unknown'
                            info_text = abbrev_line['text']
                    if label_text is not None:
                        abbrev = dict()
                        abbrev["label"] = label_text
                        abbrev["info"] = info_text

            if "info" in abbrev and re.search(r'[א-ת]', abbrev["info"]):
                had_rtl = True

        if "label" in abbrev:
            self.trace += print_to_string(f'Abbrev: "{abbrev["label"]}": {abbrev["info"]}')
            self.abbrev.append(abbrev)

        # post editing

        for abbrev in self.abbrev:
            abbrev["info"] = abbrev["info"].replace(' 150 על ספר', ' ספרי על ספר')
            abbrev["info"] = (' '+abbrev["info"]).replace(' 150 ', ' ספר ').strip()
            abbrev["info"] = abbrev["info"].replace(' 50 דבי רב', ' ספרי דבי רב')

            abbrev["label"] = (' '+abbrev["label"]).replace(' 150 ', ' ספר ').strip()
            abbrev["label"] = abbrev["label"].replace('50 במדבר', 'ספרי במדבר')

            # Remove bidi characters
            abbrev["label"] = re.sub(fr'[{uni_rtl}{uni_ltr}]', '', abbrev["label"])
            abbrev["info"] = re.sub(fr'[{uni_rtl}{uni_ltr}]', '', abbrev["info"])

            # replace all forms of double quotes with '"'
            abbrev["label"] = abbrev["label"].translate(quoteTranslate)
            abbrev["info"] = abbrev["info"].translate(quoteTranslate)


        del self.abbrev_lines

        return self.abbrev

    # create a list of bibliographic abbreviations in XML format

    def create_abbrev_list(self, abbrev_dir, pdf_dir, pdf_file):

        if len(self.abbrev) > 0:
            paper_abbrev_doc:Document = dom_impl.createDocument(None, "paper", None)
            paper_abbrev_element:Element = paper_abbrev_doc.documentElement
            paper_abbrev_element.setAttribute("file", self.pdf_file.name)
            paper_abbrev_element.setAttribute("url", self.metadata["url"])

            abbrev_list_element:Element = paper_abbrev_doc.createElement("abbreviations")
            paper_abbrev_element.appendChild(abbrev_list_element)
            for abbrev in self.abbrev:
                abbrev["label"] = re.sub(fr'[{uni_rtl}{uni_ltr}]', '', abbrev["label"])
                abbrev["info"] = re.sub(fr'[{uni_rtl}{uni_ltr}]', '', abbrev["info"])

                abbrev["label"] = abbrev["label"].translate(quoteTranslate)
                abbrev["info"] = abbrev["info"].translate(quoteTranslate)

                abbrev_element:Element = paper_abbrev_doc.createElement("abbrev")
                abbrev_list_element.appendChild(abbrev_element)
                abbrev_element.setAttribute("label", abbrev["label"])
                abbrev_element.setAttribute("file", self.pdf_file.name)
                abbrev_element.setAttribute("journal", abbrev_dir.resolve().parent.name)
                abbrev_element.appendChild(paper_abbrev_doc.createTextNode(abbrev["info"]))
            out_abbrev_file = Path(abbrev_dir, self.pdf_file.relative_to(pdf_dir)).with_suffix('.xml').open('w', encoding='utf8')
            paper_abbrev_doc.writexml(out_abbrev_file, encoding='utf8', newl='\n', addindent=' '*4)
        else:
            print(f'No abbreviations in file {pdf_file} (when writing abbrevs) labels: {[(s["title"], s["reference_content"][:3]) for s in self.metadata_jstor["references"]["reference_blocks"] if "הערות" not in s["title"] and s["title"] != ""]}')

    def paper_has_abbrev(self) -> bool:
        if "קיצורים ביבליוגרפים" not in self.references and "ביבליוגרפיה" not in self.references:
            # print(f'No abbreviations/bibliography in file {pdf_file} labels: {[(s["title"], s["reference_content"][:3]) for s in metadata_jstor["references"]["reference_blocks"] if "הערות" not in s["title"] and s["title"] != ""]}')
            return False
        return True

    def analyze_page_abbrev(self, page_num):

        self.trace += print_to_string(f'--- page {page_num}')
        # Check for None
        if self.paper_ocr is None:
            self.trace += print_to_string(f'No OCR data for this paper')
            return False

        page_ocr_tess_sheet_name = 'p%02d' % page_num
        r"""
        # this is the agent that was used
        # ocr_agent = lp.TesseractAgent(languages=['heb', 'eng'])

        page_image = np.asarray(paper_image[page_num])
        page_ocr = ocr_agent.detect(page_image, return_response=True, agg_output_level=lp.TesseractFeatureType.WORD)
        page_ocr_data = list(page_ocr['data'].itertuples())
        page_ocr['data'].to_excel(paper_tess_excel_writer, sheet_name=page_ocr_tess_sheet_name)
        """
        page_ocr_sheet = self.paper_ocr[page_ocr_tess_sheet_name]
        # fix common OCR errors            

        for (i_t, t) in enumerate(page_ocr_sheet.text):
            # When a text in an unexpected language appears, it is garbage or 'nan'
            # convert to an empty word
            if type(t) is float:
                if math.isnan(t):
                    t = ''
                else:
                    t = str(t)
                page_ocr_sheet.loc[i_t, 'text'] = t

            if page_ocr_sheet.loc[i_t, 'width'] in range(20, 31) and t=='' and self.paper_page_num-page_num <= 4:
                # print(f"Page {page_num+1}/{len(self.paper_pdf_document)-1}: changed empty string w={page_ocr_sheet.loc[i_t, 'width']} to =")
                page_ocr_sheet.loc[i_t, 'text'] = '='

            # 'pp.' is sometime detected as Hebrew
            if t == '.קק'+uni_ltr:
                page_ocr_sheet.loc[i_t, 'text'] = 'pp.'

            if t == 'קק'+uni_ltr:
                page_ocr_sheet.loc[i_t, 'text'] = 'pp.'

            if t == '/ס'+uni_ltr:
                page_ocr_sheet.loc[i_t, 'text'] = 'of'

            if t == 'מס'+uni_ltr:
                page_ocr_sheet.loc[i_t, 'text'] = 'on'

            if t == 'מסץ'+uni_ltr:
                page_ocr_sheet.loc[i_t, 'text'] = 'von'

            if t == 'by'+uni_rtl:
                page_ocr_sheet.loc[i_t, 'text'] = 'על'

            if t == '‘ay'+uni_rtl:
                page_ocr_sheet.loc[i_t, 'text'] = "עמ'"

            if re.match(rf'[nmBhdD]y{uni_rtl}$', t):
                page_ocr_sheet.loc[i_t, 'text'] = "עמ'"

            if re.match(rf'[‘⸂⸄‛⸌][nmBhdD]y{uni_rtl}?$', t):
                page_ocr_sheet.loc[i_t, 'text'] = "עמ'"

        page_ocr_data = list(page_ocr_sheet.itertuples())
        if page_num == 21:
            pass

        (page_width, page_height) = (page_ocr_data[0].width, page_ocr_data[0].height)
        self.page_width = page_width

        blocks = get_scanned_page(page_ocr_data)

        end_page = False

        lines = []
        page_footnote_refs = []
        page_footnote_lines = []
        max_line_width = 0
        font_size_widest_line = 0

        # - store all the lines in one array
        # - if a footnote reference is found, its bounding box is outside the line because it is a bit above
        #   so change the top-y of bounding box to match the line

        for block in blocks:

            if end_page:
                break

            # process only text blocks (type 0)

            if block['type'] != 0:
                continue

            for line in block["lines"]:
                if end_page:
                    break
                
                if "size" in line:
                    font_size = line["size"]
                else:
                    font_size = line["spans"][0]["size"]

                if font_size > 30:
                    self.trace += print_to_string('Ignoring high line '+str(line))
                    continue

                # bbox is x1, y1, x2, y2
                # bbox[2] is x2 (the right side) and here we sort in reverse order
                # because the text is right to left
                
                line["spans"].sort(key=lambda span: -span["bbox"][2])
                line_spans = line["spans"]

                y1 = line_spans[0]["bbox"][1]
                text_size = line_spans[0]["size"]

                line_bbox = line["bbox"]
                line_width = line_bbox[2]-line_bbox[0]

                # assume that the main text font is the font of the longest line
                # but o not count the top line (page header)
                if line_width > max_line_width and len(lines) > 0:
                    max_line_width = line_width
                    font_size_widest_line = text_size

                for span in line_spans:
                    # pages of JStor papers end with this statement, no need to process further
                
                    if "text" not in span:
                        continue

                    if "ent downloaded from" in span["text"]:
                        line_spans = []
                        end_page = True
                        break

                    # footnote reference, change the y1 to match that of the line
                    # and insert a reference text for further processing

                    if span["text"].isdigit():
                        pass

                    if span["size"] < 0.7*text_size and span["text"].isdigit():
                        if math.fabs(span["bbox"][1]-y1) < 1:
                            span["bbox"] = (span["bbox"][0], y1, span["bbox"][2], span["bbox"][3])
                            page_footnote_refs.append(span["text"])
                            span["text"] = f' <fn-{span["text"]}> '

                line["text"] = ' '.join([span["text"] for span in line_spans if "text" in span])
                line["text"] = re.sub(r'\s+', ' ', line["text"]).strip()
                # lines[-1] += revert_text(text, debug=False)
                
                # Nikkud that has to be reversed

                line["text"] = re.sub(r'(\p{Mn})(\p{L})\p{Zs}(\p{Mn})', r'\2\1\3 ', line["text"])
                line["text"] = re.sub(r'(\p{Mn})(\p{L})', r'\2\1', line["text"])

                if line["text"].startswith("פסוקים אלה"):
                    pass
                
                # used for finding the number of columns

                line["words"] = []
                for span in line_spans:
                    if 'words' not in span:
                        continue
                    line["words"] += span["words"]

                # u200e - left to right mark
                # u200f - right to left mark
                # line["words"].sort(key=lambda sw: -sw.left)

                if not end_page:
                    lines.append(line)

        # the font of the main text is the font size of the first longest line
        # encountered in the page

        # remove empty lines (could be with figures)
        lines = [ l for l in lines if len(l["text"]) > 0]
        if len(lines) < 2:
            # only the line containing the page number
            return False

        # sort the lines of the page according to the top Y
        lines.sort(key=lambda line: line["bbox"][1])
        # find the rightmost text margin
        # ignoring the top line as the page number may be outside the text margins
        page_right_margin = max([line["bbox"][2] for line in lines[1:]]) 
        for line in lines:
            line['page_right_margin'] = page_right_margin
        
        page_left_margin = min([line["bbox"][0] for line in lines[1:] if line["bbox"][0]>0]) 
        for line in lines:
            line['page_left_margin'] = page_left_margin

        last_line = lines[-1]
        if '343' in last_line["text"]:
            pass

        if page_num == 8:
            pass

        # check if the last line is a page number
        # digits, optionally between brackets, centered


        if re.search(r'^\p{P}*\d+\p{P}*$', last_line["text"]):
            if is_centered(last_line, page_width):
                lines.pop()

        if len(lines) < 2:
            return False


        line_spaces = Counter()
        
        for i_line in range(1, len(lines)):
            line_spaces.update([lines[i_line]['bbox'][1]-lines[i_line-1]['bbox'][3]])
        self.trace += print_to_string(f'{len(lines)} lines, Line spaces: {str(line_spaces)}')

        # remove title/author lines at the front page

        if page_num == 1:
            n1 = len(lines)
            for i_line in range(n1-1, 1, -1):
                line_right_offset = page_right_margin-lines[i_line]['bbox'][2]
                if line_right_offset < 130:
                    break
                self.trace += print_to_string(f' Ignoring line {i_line} page {page_num}: "{lines[i_line]["text"]}"')
                lines.pop()

        if self.bottom is None and len(lines) > 0 and page_num <= 5:
            # calculate for the first page and the actual value from the 2nd page
            self.top = max([ l["bbox"][1] for l in lines])
            self.bottom = max([ l["bbox"][3] for l in lines])

        self.trace += print_to_string(f'page {page_num}: bottom {max([ l["bbox"][3] for l in lines])}/{self.bottom} ')

        if page_num == 11:
            pass

        fn_sep_line = self.divide_page(lines)

        # debug
        # [(i, t) for (i, t) in enumerate([(l['size'], l['bbox'][2]-l['bbox'][0], l['text']) for l in lines])]
        
        main_font_size = font_size_widest_line
        num_text_lines = 0

        self.trace += print_to_string(f'main font size: {main_font_size}\n')

        # Abbreviation lists continue to the end of the paper

        line_type = LineType.HEADER
        if self.abbrev_list_found:
            line_type = LineType.ABBREV
        
        num_text_lines = 0

        for (i_line, line) in enumerate(lines):

            if i_line == 24:
                pass
            
            if i_line == fn_sep_line:
                self.trace += '--- footnote sep ---\n'
                line_type = LineType.FOOTNOTE

            if line_type == LineType.FOOTNOTE:
                page_footnote_lines.append(line)
                continue

            line_bbox = []
            # self.trace += print_to_string(str(line["bbox"])+'\n')

            # reduce the precision of the bounding box coordinates to two decimal digits
            for coord in line["bbox"]:
                coord = int(coord*100)/100
                line_bbox.append(coord)
            line_width = line_bbox[2]-line_bbox[0]

            line_right_margin = line_bbox[2]
            line_right_offset = page_right_margin-line_right_margin

            # if the font size is less than 0.9 X main font:
            # - signal that we are in a footnote from now
            # - print a separator line in the trace

            # find the first text line (after the page header, may span several lines).
            # the first line is the page header (hopefully)

            if line_type == LineType.HEADER and (line["size"] == main_font_size and i_line > 0):
                line_type = LineType.TEXT
                self.trace += '--- Text '+ '-'*50+'\n'

            if 'בספרות הבלשנית' in line["text"]:
                pass
            
            line_centered = is_centered(line, page_width)

            # consider as a title if there are 4 or less words
            if len(line["words"]) > 4:
                line_centered = False

            line_right_justified = (line_right_offset <= 2*column_tolerance[self.journal_name])
            line_new_block = True
            if i_line > 0:
                line_prev = lines[i_line-1]
                if (line['words'][0].level, line['words'][0].block_num) == (line_prev['words'][0].level, line_prev['words'][0].block_num):
                    line_new_block = False

            found_abbrev_label = self.test_abbrev_label(line)
            if found_abbrev_label and line_right_justified:
                if not line_new_block:
                    found_abbrev_label = False

            if found_abbrev_label:
                line_type = LineType.ABBREV
            elif line_centered and line_type == LineType.ABBREV:
                line_type = LineType.TEXT
                self.abbrev_list_found = False
                self.trace += '--- Text '+ '-'*50+'\n'

            if line_type == LineType.ABBREV:
                self.abbrev_list_found = True

            if line_type == LineType.ABBREV and not line_centered:
                if i_line > 0:
                    self.abbrev_lines.append(line)
                continue

            """
            elifls  line_type != LineType.HEADER and (line["size"] < 0.9*main_font_size):
                line_type = LineType.FOOTNOTE
                self.trace += '--- Footnotes '+ '-'*50+'\n'
            """

            if line_width == max_line_width:
                max_width_sign = '*'
            else:
                max_width_sign = ' '

            self.trace += print_to_string("--- line --- font size %5.2f width %d %s %s " % (line["size"], line_width, max_width_sign, line["text"])+str(line_bbox))
            if i_line > 0:
                self.trace += print_to_string(' space from previous line: ', lines[i_line]['bbox'][1]-lines[i_line-1]['bbox'][3])
            self.line_columns(line)

            if line_type == LineType.TEXT:
                num_text_lines += 1

        if len(lines) == 0:
            # no text, can also be an empty page
            return False

        if self.trace_file is not None:
            print(self.trace, file=self.trace_file)
            self.trace_file.flush()
            self.trace = ''

        return (len(self.abbrev_lines) > 0)
        
    def test_abbrev_label(self, line):
        line_text_nonletter_end_removed = re.sub(r'\P{L}+$', '', line["text"])
        line_text_blank_nonletter_end_removed = re.sub(r'\s+', '', line_text_nonletter_end_removed)

        found_abbrev_label = False

        if line_text_nonletter_end_removed in self.reference_labels:
            refType = self.reference_labels[line_text_nonletter_end_removed]
            if refType == "קיצורים ביבליוגרפים":
                found_abbrev_label = True
            if self.journal_name == 'מגילות' and refType == "ביבליוגרפיה":
                found_abbrev_label = True
        if line_text_blank_nonletter_end_removed in self.reference_labels:
            refType = self.reference_labels[line_text_blank_nonletter_end_removed]
            if refType == "קיצורים ביבליוגרפים":
                found_abbrev_label = True
            if self.journal_name == 'מגילות' and refType == "ביבליוגרפיה":
                found_abbrev_label = True
        if self.journal_name == 'מגילות' or True:
            if line_text_nonletter_end_removed in ('קיצורים', 'רשימת קיצורים', 'רשימת הקיצורים', 'מקורות מודפסים ומחקרים'):
                found_abbrev_label = True
            # allow a few characters as in file 23438303.pdf
            if re.search(r'^רשימת ה?קיצורים.{,5}', line["text"]):
                found_abbrev_label = True
        if self.journal_name == 'מגילות' or True:
            if line_text_blank_nonletter_end_removed in ('קיצורים', 'רשימת קיצורים', 'רשימת הקיצורים', 'מקורות מודפסים ומחקרים'):
                found_abbrev_label = True
            # allow a few characters as in file 23438303.pdf
            if re.search(r'^רשימת ה?קיצורים.{,5}', line["text"]):
                found_abbrev_label = True
        if self.journal_name == 'לשוננו' or True:
            if line_text_nonletter_end_removed in ('מחקרים', 'קיצורים'):
                found_abbrev_label = True
            if len(line_text_nonletter_end_removed) < 25 and re.search(r'^(ה?מקורות|ביבליוגרא?פיה)', line_text_nonletter_end_removed):
                found_abbrev_label = True
            if re.search(r'^רשימת.{,10}(ה?מקורות|ביבליוגרא?פי|ה?קיצור|מאמרים|ספרים|מחקרים)', line_text_nonletter_end_removed):
                found_abbrev_label = True
        if self.journal_name == 'לשוננו' or True:
            if line_text_blank_nonletter_end_removed in ('מחקרים', 'קיצורים'):
                found_abbrev_label = True
            if len(line_text_blank_nonletter_end_removed) < 25 and re.search(r'^(ה?מקורות|ביבליוגרא?פיה)', line_text_blank_nonletter_end_removed):
                found_abbrev_label = True
            if re.search(r'^רשימת.{,10}(ה?מקורות|ביבליוגרא?פי|ה?קיצור)', line_text_blank_nonletter_end_removed):
                found_abbrev_label = True

        return found_abbrev_label    

def get_scanned_page(page_ocr_data:List)->List:

    blocks = list()
    block = None
    line = None
    span = None

    prev_paragraph = -1
    prev_block = -1
    prev_line = -1

    # Add dummy data so that the last block will be processed
     
    last_data = namedtuple('ld', ['block_num', 'line_num', 'par_num', 'conf'])
    setattr(last_data, 'block_num', -1)
    setattr(last_data, 'par_num', -1)
    setattr(last_data, 'line_num', -1)
    setattr(last_data, 'conf', 0)
    page_ocr_data.append(last_data)

    for ocr_word_data in page_ocr_data:
        if ocr_word_data.conf < 0:
            continue

        if (ocr_word_data.block_num, ocr_word_data.par_num) != (prev_block, prev_paragraph):
            if span is not None:
                add_span_to_blocks(blocks, span)
            span = dict(words=[], block_num=ocr_word_data.block_num, par_num=ocr_word_data.par_num)
            prev_line = -1
        elif (ocr_word_data.line_num != prev_line):
            if span is not None:
                add_span_to_blocks(blocks, span)
            span = dict(words=[], block_num=ocr_word_data.block_num, par_num=ocr_word_data.par_num)
        
        span["words"].append(ocr_word_data)

        prev_block = ocr_word_data.block_num
        prev_paragraph = ocr_word_data.par_num
        prev_line = ocr_word_data.line_num

    blocks.sort(key=lambda block: block["top"])
    return blocks

