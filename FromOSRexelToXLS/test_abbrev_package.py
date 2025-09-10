"""
Processing Raw PDF files from JSTOR
"""

# PDF processing package
# See the documenation of the page object in 
# https://pymupdf.readthedocs.io/en/latest/textpage.html

import argparse
import sys
import regex as re
import pandas

from pathlib import Path
from xml.dom import getDOMImplementation

from abbreviations import paper_abbrev

def get_paper_ocr(ocr_dir, pdf_file):
    paper_ocr_tess_xls_path = Path(ocr_dir, pdf_file.name).with_suffix('.xlsx')

    if not paper_ocr_tess_xls_path.exists():
        sys.stderr.write("No Tesseract output in "+str(paper_ocr_tess_xls_path.resolve())+'\n')
        return None  # Возвращаем None, но не прерываем выполнение
    
    try:
        paper_ocr_tess = pandas.read_excel(paper_ocr_tess_xls_path, sheet_name=None)
        return paper_ocr_tess
    except Exception as e:
        sys.stderr.write(f"Error reading OCR file {paper_ocr_tess_xls_path}: {str(e)}\n")
        return None  

# main

parser = argparse.ArgumentParser(prog='process-raw-pdf', description='Convert JSTOR PDF to text and separate footnotes')

parser.add_argument('-j', help='Journal name in Hebrew', dest='journal', type=str, required=True)
parser.add_argument('-i', help='Input PDF directory', dest='pdf_dir', type=Path, required=True)

# Directory containing metadata JSON files downloaded from JStor and metadata.xlsx for all the papers
parser.add_argument('-m', help='Input Metadata directory', dest='meta_dir', type=Path, required=True)

parser.add_argument('-a', help='Output Abbreviation directory, default: %(default)s', dest='abbrev_dir', type=Path, required=False, default=Path(Path.cwd(), 'abbrev').resolve())

# Trace information about the processing of each PDF file
parser.add_argument('-T', help='Output Trace directory', dest='trace_dir', type=Path, required=True)

# Trace information about the processing of each PDF file
parser.add_argument('-O', help='Tesseract OCR directory', dest='ocr_tess_dir', type=Path, required=False)

# Whether or not to proces a specific file (specified in the code)
parser.add_argument('-s', help='Specific file?', dest='specific_file', action='store_true', default=False, required=False)

if len(sys.argv) == 1:
    parser.print_help()
    sys.exit(0)

args = parser.parse_args()

journal_name = args.journal

# Assigning values from the arguments

pdf_dir:Path = args.pdf_dir
if (not pdf_dir.exists()):
    sys.stderr.write(f"The input directory {str(pdf_dir.resolve())} does not exist!\n")
    sys.exit(1)

if args.ocr_tess_dir is not None:
    ocr_tess_dir = args.ocr_tess_dir
else:
    ocr_tess_dir:Path = Path(pdf_dir.parent, 'ocr-tess')
ocr_tess_raw_dir = ocr_tess_dir.with_name(ocr_tess_dir.name+'-raw')

trace_dir:Path = args.trace_dir
trace_dir.mkdir(exist_ok=True, parents=True)

meta_dir:Path = args.meta_dir
meta_dir.mkdir(exist_ok=True)

abbrev_dir:Path = args.abbrev_dir
abbrev_dir.mkdir(exist_ok=True, parents=True)

# specific_file = False | True
specific_file = args.specific_file
dom_impl = getDOMImplementation()

ocr_dir:Path = Path(pdf_dir.parent, 'ocr-tess')
# ocr_tess_raw_dir = ocr_tess_dir.with_name(ocr_tess_dir.name+'-raw')

n_files = -1
i_file = 0
pdf_files = [f for f in pdf_dir.iterdir() if f.suffix == '.pdf']

if n_files == -1:
    n_files = len(pdf_files)

for pdf_file in pdf_files:
    
    """
    if pdf_file.name < '26264395.pdf':
        continue
    """

    # 27101916 24371684 26694434 24704335 24350320 23438242 23438234 24174727 24164394 
    # 24173476 - two columns of bibliographic abbreviations, from p.40
    # if specific_file and not re.search(r'24173476|2637768[78]|zx0834|z0985', pdf_file.name):
    # if specific_file and not re.search(r'24164389|z0202', pdf_file.name):
    #    continue
    # if specific_file and not re.search(r'24328321|24331279|24360523', pdf_file.name):
    # if specific_file and not re.search(r'27068633', pdf_file.name):

    #                                    sss
    if specific_file and not re.search(r'24327775', pdf_file.name):
       continue

    override = True

    out_abbrev_path = Path(abbrev_dir, pdf_file.relative_to(pdf_dir)).with_suffix('.xml')
    if out_abbrev_path.exists() and not override:
        continue

    # if re.search(r'27101916|27101921|27101940', file.name):
    #    continue
       
    if i_file == n_files:
        break
    i_file += 1

    # Data for the paper as a whole

    trace_path = Path(trace_dir, pdf_file.relative_to(pdf_dir)).with_suffix('.txt')
    trace_file = trace_path.open('w', encoding='utf8')

    # print("=-"*6)

    print("reading file ", pdf_file, end='')
    if i_file % 10 == 0:
        print(f' {i_file}/{n_files}', end='')
    print('')

    paper_ocr = get_paper_ocr(ocr_tess_dir, pdf_file)
    # Check if there is OCR data
    if paper_ocr is None:
        print(f"Skipping file {pdf_file.name} due to missing OCR data")
        continue # Skip this file and move on to the next one
    paper_meta_path = Path(meta_dir, pdf_file.relative_to(pdf_dir)).with_suffix('.json')

    paper_abbreviations = paper_abbrev.paper_abbrev(paper_ocr, journal_name, pdf_file, paper_meta_path, trace_file, pdf_dir)

    if paper_abbreviations.skip is not None:
        # pdf_file.replace(Path(pdf_skipped_dir, pdf_file.name))
        print(paper_abbreviations, file=trace_file)
        trace_file.close()
        continue

    if not paper_abbreviations.paper_has_abbrev():
        print('No abbreviations')

    for page_num in range(1, paper_abbreviations.paper_page_num):
        paper_abbreviations.analyze_page_abbrev(page_num)

    abbrevs = paper_abbreviations.get_abbrev()

    for abbrev in abbrevs:
        print(abbrev)

