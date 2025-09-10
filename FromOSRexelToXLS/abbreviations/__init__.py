import io
import enum
import unicodedata
import pandas
import numpy as np
import layoutparser as lp
from editdistance import distance

from fuzzysearch import find_near_matches
from bidi import algorithm
from enum import Enum

from xml.dom.minidom import Document, Element
from xml.dom import getDOMImplementation

from re import Match
import regex as re

import math
import time
import random
import httpx

# package functions

# static

heb_full_height = 'אבגדהוזחטכםמנסעפצרשת'
uni_ltr = '\u200e'
uni_rtl = '\u200f'

uniQuotes ="«»“”„‟❝❞❠〞〟＂🙶🙷🙸״"
quoteTranslate = dict( [ (ord(x), ord(y)) for x,y in zip( uniQuotes,  '"'*len(uniQuotes)) ] )

abbrev_indentation = {'תרביץ': 245, 'לשוננו': 35, 'מגילות': 35, 'שנתון': 35, 'סידרא': 245}
abbrev_single_line_indentation = {'תרביץ': 30, 'לשוננו': 10, 'מגילות': 35, 'שנתון': 10, 'סידרא': 30}
column_tolerance = {'תרביץ': 10, 'לשוננו': 5, 'מגילות': 3, 'שנתון': 3, 'סידרא': 10}

LineType = Enum('LineType', ['HEADER', 'TEXT', 'FOOTNOTE', 'ABBREV'])

Fields = Enum('Fields', ['TITLE', 'AUTHOR', 'SOURCE', 'PUBLISHER', 'URL'])
reverse_paren = dict( [ (ord(x), ord(y)) for x,y in {'[': ']', ']':'[', '(': ')', ')':'('}.items()])

heb_months = [ 'תשרי', 'חשוו?ן', 'כסלי?ו', 'טבת', 'שבט', r"אדר(?:\s[אב]'?)?", 'ניסן', 'אייר', 'סיוו?ן', 'תמוז', 'אב', 'אלול' ]
heb_month_regex = '(?P<month>(?:'+'|'.join(heb_months)+r')(?:\s*\p{Pd}\s*'+'(?:'+'|'.join(heb_months)+'))?)'
ref_subtypes_regex = r'^([א-ת]\p{P}?)?\s*(מקורות|מחקרים|טקסטים|עדי נוסח|כתבי יד|מאמרים|ספרות|ספרי |ספרים|עי?תונות|לשון הדיבור|מדרשים?|משנה|תלמוד|מילונים|מקרא.{,15}עדי נוסח|מקרא.{,15}כתבי יד)\s*.{,10}$'

par_indent = 31
quote_indent = par_indent*2

asterik_start = {
    'ה?(הרצאה|מחקר|מאמר|חיבור) ',
    'הער[הת] ה?[המ]ערכת',
    'מתוך (הרצאה|מחקר|מאמר|חיבור) ',
    '(ראשיתו|עיקרו) של (מחקר|מאמר|חיבור)',
    'מוקדש',
    'להלן רשימת',
    '(תודה|תודתי)',
    'אנו מודים ',
    '(הריני|הרינו|הנני|הננו|אני|אנו|אנחנו|המחבר|המחברים|המחברות) מבקש(ת|ים|ות) להודות ',
    '(הריני|הרינו|הנני|הננו|אני|אנו|אנחנו|המחבר|המחברים|המחברות) (מודה|מודים|מודות) ',
    'ברצוני ',
    'דברים ש',
    'ה?דברים ',
    'חלק(ים)? (מה?|מן ה)(דברים|מאמר|עבודה)',
    'מוקדש ל',
    'מתוך ',
    'עיבוד של ',
    'עיקרי ה?דברים ',
    'פרק(ים)? מתוך'
}

dom_impl = getDOMImplementation()

asterik_start_regex = '|'.join(asterik_start)

def clearAllChildren(element:Element):
    children = element.childNodes
    while len(children) > 0:
        element.removeChild(children[0])

def is_asterik_comment(fn_text:str)->bool:
    fn_text = re.sub(r'^\P{L}*', '', fn_text)
    return (re.search(asterik_start_regex, fn_text) is not None)

def get_text_letters(text:str)->str:
    letters_all = ''.join(re.findall(r'\p{L}', text))
    letters_heb = ''.join(re.findall(r'[א-ת]', letters_all))
    letters_lat = ''.join(re.findall(r'[^א-ת]', letters_all))
    return {'heb':letters_heb, 'lat':letters_lat, 'combined-20':letters_heb[:10]+letters_lat[:10]}

"""
Compare two outputs of get_text_letters
"""
def compare_letters(l1, l2, pref_len=10):

    max_l1 = max([len(v) for v in l1.values()])
    max_l2 = max([len(v) for v in l2.values()])

    max_k1 = [k for k in l1 if len(l1[k]) == max_l1]
    max_k2 = [k for k in l2 if len(l2[k]) == max_l2]

    for k in max_k1+max_k2:
        if distance(l1[k][:pref_len], l2[k][:pref_len]) <= 2:
            return True
    
    return False

def adjust_punct(reverse_span):
    if len(reverse_span) <= 1:
        return
    last_rtl = reverse_span[-1]
    punct_match_last = re.search(r'[\p{P}\p{S}]*\)'+f'[{uni_rtl}{uni_ltr}]?$', last_rtl.text)

    first_rtl = reverse_span[0]
    punct_match_first = re.search(r'^\([\p{P}\p{S}]*', first_rtl.text)

    if punct_match_last is not None:
        punct = punct_match_last.group(0).translate(reverse_paren)[::-1]
        punct = re.sub(f'[{uni_rtl}{uni_ltr}]', '', punct)
        last_word = re.sub(r'[\p{P}\p{S}]*\)'+f'[{uni_rtl}{uni_ltr}]?$', '', last_rtl.text)
        rep_last = replace_ocr_word_text(last_rtl, last_word)
        reverse_span[-1] = rep_last
        reverse_span[0] = replace_ocr_word_text(reverse_span[0], punct+reverse_span[0].text)

    if punct_match_first is not None:
        punct = punct_match_first.group(0).translate(reverse_paren)[::-1]
        first_word = re.sub('^\([\p{P}\p{S}]*', '', first_rtl.text)
        rep_first = replace_ocr_word_text(first_rtl, first_word)
        reverse_span[0] = rep_first
        reverse_span[-1] = replace_ocr_word_text(reverse_span[-1], reverse_span[-1].text+punct)

def replace_ocr_word_text(w, new_text):
    w_new = pandas.Series(data={'left':w.left, 'top':w.top, 'height':w.height, 'width':w.width, 'text':new_text, 'block_num':w.block_num, 'line_num':w.line_num, 'word_num':w.word_num}) 
    return w_new

def revert_no_blanks(text:str):
    if len(text) <= 1:
        return text
    m_blank_start = re.search(r'^\s+', text)

    if m_blank_start is not None:
        l = m_blank_start.end()
        b_start = text[:l]
        text = text[l:]
    else:
        b_start = ''

    m_blank_end = re.search(r'\s+$', text)
    if m_blank_end is not None:
        l = m_blank_end.end()-m_blank_end.start()
        b_end = text[-l:]
        text = text[:-l]
    else:
        b_end = ''

    return b_start+text[::-1]+b_end

def revert_digits(text:str):

    text_rev_dig = ''
    cur_pos = 0

    digit_matches = list(re.finditer(r'[0-9]+(\p{P}+[0-9]+)?', text))
    if len(digit_matches) == 0:
        return text

    for digit_match in digit_matches:
        pos = digit_match.start()
        if pos > cur_pos:
            text_rev_dig += text[cur_pos:pos]
        text_rev_dig += digit_match.group(0)[::-1]
        cur_pos = digit_match.end()

    if cur_pos < len(text):
        text_rev_dig += text[cur_pos:]

    return text_rev_dig

def revert_text(text:str, debug=False):

    text = ' '+text+' '
    # rtl_matches = list(re.finditer(r'[\p{Mn}\p{P}\p{Bidi_Class=R}](?:[\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]+)[\p{Bidi_Class=R}\p{P}]', text))
    # rtl_matches = list(re.finditer(r'[\p{Mn}\p{P}\p{Bidi_Class=R}][\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]*\p{Bidi_Class=R}[\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]*[\p{Bidi_Class=R}\p{P}]', text))
    # rtl_matches = list(re.finditer(r'[\p{Mn}\p{P}\p{Bidi_Class=R}][\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]*[\p{Bidi_Class=R}\p{Mn}\p{P}]+', text))
    # rtl_matches = list(re.finditer(r'[\p{Mn}\p{P}]*\p{Bidi_Class=R}[\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]*[\p{Bidi_Class=R}\p{Mn}\p{P}]+', text))
    # rtl_matches = list(re.finditer(r'(?:[0-9]+[\p{P}\p{Zs}]*)*[\p{Mn}\p{P}]*\p{Bidi_Class=R}[\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]*[\p{Bidi_Class=R}\p{Mn}\p{P}]+', text))
    rtl_matches = list(re.finditer(r'\p{Mn}?\p{Bidi_Class=R}[\p{Zs}\p{P}\p{Mn}\p{Bidi_Class=R}]*[\p{Bidi_Class=R}\p{Mn}\p{P}]+', text))

    text_rev = ''
    cur_pos = 0

    for rtl_match in rtl_matches:
        pos = rtl_match.start()
        if pos > cur_pos:
            # text_rev += text[cur_pos:pos]
            text_rev = text[cur_pos:pos]+text_rev
            # text_rev += revert_no_blanks(text[cur_pos:pos])
            if debug:
                print(f'appending "{text[cur_pos:pos]}"')
        # text_rev += revert_digits(rtl_match.group(0))
        # text_rev += rtl_match.group(0)
        rtl_text = rtl_match.group(0)
        text_rev = rtl_text+text_rev
        if debug:
            print(f'appending match "{rtl_text}"')
        cur_pos = rtl_match.end()
    if cur_pos < len(text):
        # text_rev += text[cur_pos:]
        text_rev = text[cur_pos:]+text_rev
        # text_rev += revert_no_blanks(text[cur_pos:])
        if debug:
            print(f'appending "{text[cur_pos:]}"')

    # text_rev = text_rev.strip()

    return text_rev

def print_to_string(*args, **kwargs):
    output = io.StringIO()
    print(*args, file=output, **kwargs)
    contents = output.getvalue()
    output.close()
    return contents

def only_full_line(word:str):
    has_upper = (re.search('ל', word) is not None)
    has_lower = (re.search(r'[ךןףץקygqQ]', word) is not None)
    has_not_yod = (re.search(r'[^י]', word) is not None)

    return (not has_upper) and (not has_lower) and has_not_yod

def calc_font_size(word_span):

    height = word_span.height
    text = word_span.text

    if re.search(r'[ךןףץקלygqQ]', text) is None:
        if ',' in text:
            font_size = height-3
        else:
            font_size = height
    else:
        has_upper = (re.search('ל', text) is not None)
        has_lower = (re.search(r'[ךןףץקygqQ]', text) is not None)

        # one of them is True! because of the previous 'if' condition

        if has_upper and has_lower:
            font_size = height/2
        else:
            font_size = height*2/3
        font_size = int(font_size)

    """
    if font_size > 15:
        print(f'"{text}" height:{height} fs:{font_size}')
    """

    return font_size

def update_line_bbox(line, span_bbox):
    line_bbox = line["bbox"]
    x1 = min(line_bbox[0], span_bbox[0])
    y1 = min(line_bbox[1], span_bbox[1])
    x2 = max(line_bbox[2], span_bbox[2])
    y2 = max(line_bbox[3], span_bbox[3])
    line["bbox"] = (x1, y1, x2, y2)

def load_metadata_url(metadata_url):

    headers = dict()
    
    headers['accept'] = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7'
    headers['accept-language'] = 'en-US,en;q=0.9,he;q=0.8'
    headers['cache-control'] = 'no-cache'
    headers['cookie'] = 'UUID=6fd52e49-fe28-4907-bca3-c004e90e4bf4; pxcts=f2835bd1-ed12-11ee-ae09-ad26d74b5577; _pxvid=f1dbe368-ed12-11ee-bd2c-0b21b9ae95a5; csrftoken=UTeIxgRSKx5Imw1aF0nghw15HKmt5o8rAq45lEgODXTlUXHIqpkt2BUbcZPQUkiE; AccessToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzZXNzaW9uSWQiOiI4ZmRkNzAyNDU1MGQ0ZGIwYmE4Mjg0MDUxYmM5MTMxMSIsInV1aWQiOiI2ZmQ1MmU0OS1mZTI4LTQ5MDctYmNhMy1jMDA0ZTkwZTRiZjQiLCJ1c2VyIjp7ImlkIjoiIiwibG9nZ2VkSW4iOmZhbHNlLCJhZG1pbiI6ZmFsc2V9LCJpYXQiOjE3MTI1ODAyMjcsImV4cCI6MTcxMjU4MDUyN30.umXMvhqmx16DMu8lF9pnaGHTqRgHWtcJ9UtZ2R-DvHc; ReferringRequestId=fastly-default:f0a3283605ce7227567bc37fe90f9f37; OptanonAlertBoxClosed=2024-04-08T12:45:00.570Z; OptanonConsent=isGpcEnabled=0&datestamp=Mon+Apr+08+2024+15%3A45%3A00+GMT%2B0300+(Israel+Daylight+Time)&version=202303.1.0&browserGpcFlag=0&isIABGlobal=false&hosts=&consentId=b464be44-8ca8-47d3-8471-db892aa6155d&interactionCount=2&landingPath=NotLandingPage&groups=C0001%3A1%2CC0002%3A1%2CC0005%3A1%2CC0004%3A1%2CC0003%3A1&AwaitingReconsent=false; _ga=GA1.1.736159652.1712580301; _ga_JPYYW8RQW6=GS1.1.1712580300.1.1.1712580300.0.0.0; _pxhd=lWx5vS6B-RKkOxFisPx4YWjEneWQIbBMev/1mXaG9YByr9jDOkUcfDWCtjk3XrfbR6RLH-7eyJTKvHPc-h0Nxg==:6wE5z0/am5FFKRheqmoJkFY2qd8Euk1OqR-e9mJuIopIl55OH4yWVyxcJmszUYmdy1/obvtuQEVjPQ2V4psWVxFB2V12RwcijS-gsexKziI=; _px2=eyJ1IjoiYTQyNDM3YjAtZjVhNS0xMWVlLWFkMmUtODVjNTMzNzNkYWM4IiwidiI6ImYxZGJlMzY4LWVkMTItMTFlZS1iZDJjLTBiMjFiOWFlOTVhNSIsInQiOjE1NjE1MDcyMDAwMDAsImgiOiJkY2RkZWE0OTI5NDQyMTU2Yzk1NjczMzhjOTZiMzgwNzNjMTBmN2M4YzNiMDcwNzQwOTI2MTI5NDg2ZmU4NWRjIn0=; AccessSessionTimedSignature=b8442f81882f3e10471350a4efff7014e6b1b642f7a0e2d034e6803ce15119fd; AccessSession=H4sIAAAAAAAA_42S3W4TMRCF38XXcTT-t_cuRAIqQEJtuKCoqrzecbqwTaKNtwiivDt2NoQUIZrVXlgzxzPH38yODEPbkIro2CiO0tGI3FLpwNA6eEEDgEQHKOsoyYS0m6zlTEyZMFPDp9qWYClga64xcFQBQQawnkXDFEMrpLLIfNb1o9BDjNICVZFpKqXj1FoE2hjhNfPKqUZncedTaQVcUsi_XTBeKVFJMQXjbotg-FsgKi4rZqZKOCcV5wfZ9oUyKZAq-m6LE_Lku4PDc7U-1czvMGKs6VPqt6Takfk8y2c3OTRf5NMGk3-4T-23J3__oQQ_5-DV-3y6nv8-3RTh1eLt7N2M7HOpIT3MQki53JcdST82WNKrbWrTkNr1qmBbd1jSdxPy2K7ax_Ynvu78klSpH_BInwOYMorxAUyB1FxnCDmEY6iOwnClqWqkpNKgpXUdBBUmCJ1H5pqA5f7RQfvMQeixGQ1evCv0woaZ-YCX79_R3qt-_X2L_Uff9u1qSUaMI5D93T6D8iekB8fCSCO5PdFQPphYC0adC4zKJgpqg88rKZzVoRYcxFm3a1yOHP6Fvzs2EJZLOCPYHy59yi7zmI82GACIkwmHFkyDgWZQNZXBKOqD5RQCs86rUIOHPybe9Oth818P_Kz7sqjH5pnG10MeXvieY_wFaTjQmxoEAAA; AccessSessionSignature=a0ead39a6b37538e8c82dc51f81000a875c2cd2dbaccae9402d1468da1031d7d'
    headers['dnt'] = '1'
    headers['pragma'] = 'no-cache'
    headers['sec-ch-ua'] = '"Microsoft Edge";v="123", "Not:A-Brand";v="8", "Chromium";v="123"'
    headers['sec-ch-ua-mobile'] = '?0'
    headers['sec-ch-ua-platform'] = '"Windows"'
    headers['sec-fetch-dest'] = 'document'
    headers['sec-fetch-mode'] = 'navigate'
    headers['sec-fetch-site'] = 'none'
    headers['sec-fetch-user'] = '?1'
    headers['upgrade-insecure-requests'] = '1'
    headers['user-agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0'

    wait_sec = random.uniform(0.05, 0.2)
    time.sleep(wait_sec)
    
    jstor_response = httpx.get(metadata_url, headers=headers, follow_redirects=True)

    if jstor_response.status_code != 200:
        return (dict(status=jstor_response.status_code))

    unicode_escapes = set(re.findall(r'\\u[0-9a-f]{4}', jstor_response.text))
    resp = jstor_response.text
    for u_e in unicode_escapes:
        u_ch = bytes(u_e, 'ascii').decode('unicode-escape')
        resp = resp.replace(u_e, u_ch)
    
    return (dict(text=resp))

def add_span_to_blocks(blocks, span):

    # dbg_word = 'ביתרביץי'
    # dbg_word = 'יצורה'
    dbg_words = ['הנמן', 'הנמן,']

    for dbg_word in dbg_words:
        if dbg_word in [w.text for w in span["words"]] or dbg_word[::-1] in [w.text for w in span["words"]]:
            pass

    line_toler = 6

    words_letter_digit = [w for w in span["words"] if re.search(r'[\p{L}\p{Nd}]', w.text)]
    if len(words_letter_digit) == 0:
        span["size"] = max([w.height for w in span["words"]])
        right_margin = max([w.left+w.width for w in span["words"]])
        left_margin = min([w.left for w in span["words"]])
        letter_top = max([w.top for w in span["words"]])
        letter_bottom = min([w.top+w.height for w in span["words"]])
    else:
        words_letter_heb = [w for w in span["words"] if re.search(r'[א-ת]', w.text)]
        if len(words_letter_heb) > 0:
            words_letters = words_letter_heb
        else:
            words_letters = words_letter_digit
        letter_top = max([w.top for w in words_letters])
        letter_bottom = min([w.top+w.height for w in words_letters])

        spans_for_bbox = span["words"]

        # if there are spans tha contain letters that use the full line height 
        # then use them to get the font height
        # search for text that does not contain letters that go above or below the line.

        span_containing_full_height = [ sp for sp in spans_for_bbox if re.search(f'[{heb_full_height}0-9a-zA-Z]', sp.text)]
        if len(span_containing_full_height) > 0:
            span_no_above_letters = [ sp for sp in span_containing_full_height if re.search('ל', sp.text) is None]
            if len(span_no_above_letters) > 0:
                letter_top = min([sp.top for sp in span_no_above_letters])
            span_no_below_letters = [ sp for sp in span_containing_full_height if re.search(r'[ךןףץק]', sp.text) is None]
            if len(span_no_below_letters) > 0:
                letter_bottom = min([w.top+w.height for w in span_no_below_letters])

        font_height_letters = letter_bottom-letter_top
        
        span_max_font_size = min([calc_font_size(w) for w in words_letters])

        right_margin = max([w.left+w.width for w in span["words"]])
        left_margin = min([w.left for w in span["words"]])

        span["size"] = min(font_height_letters, span_max_font_size)

        """
        print('-----', span["size"])
        print(font_height_letters, ' '.join([w.text for w in span_containing_full_height]))
        print('\n')
        """
        
    # bbox = (left_margin, letter_top, right_margin, letter_top+span["size"])
    bbox = (left_margin, letter_top, right_margin, letter_bottom)

    span["bbox"] = bbox
    span["text"] = ' '.join([w.text for w in span["words"]])
    (span_tops_bottoms) = [ (w.top, w.top+w.height) for w in span["words"]]

    # del span["words"]

    span_line = None
    span_block = None

    # check if the line already appeared in another block.

    for block in blocks:
        for cur_line in block["lines"]:
            for cur_span in cur_line["spans"]:
                for cur_word in cur_span["words"]:
                    for span_word_top_bottom in span_tops_bottoms:
                        if math.fabs(cur_word.top-span_word_top_bottom[0]) <= 3 or math.fabs(cur_word.top+cur_word.height-span_word_top_bottom[1]) <= 3:
                            span_line = cur_line
                            break
                if span_line is not None:
                    break
            if span_line is not None:
                break
        if span_line is not None:
            span_block = block
            break

    # if it is a new line, check if it is in an existing block

    if span_block is None:
        for block in blocks:
            if block["block_num"] == span["block_num"] and block["par_num"] == span["par_num"]:
                span_block = block
                break

    if span_block is None:
        span_block = dict(lines=[], type=0, block_num=span["block_num"], par_num=span["par_num"], top=letter_top)
        blocks.append(span_block)

    if span_line is None:
        span_line = dict(spans=[], bbox=(99999, 99999, -1, -1))
        span_block["lines"].append(span_line)

    span_line["spans"].append(span)
    span_line["size"] = max([sp["size"] for sp in span_line["spans"]])

    # span_line["spans"].sort(key=lambda span: (span["bbox"][1], span["bbox"][2]))
    update_line_bbox(span_line, span["bbox"])

    # set y1 to the same value for consistency and future sorting
    for sp in span_line["spans"]:
        sp["bbox"] = (sp["bbox"][0], span_line["bbox"][1], sp["bbox"][2], sp["bbox"][3])

# reorder the word list according to bidi marks
# (switch to RTL when the words begins with a Hebrew letter)

def typeset_words(words:list, had_rtl=False):

    # u200e - left to right mark
    # u200f - right to left mark

    # sort the words from right to left
    # sorted_words = sorted(words, key=lambda w: -w.left)
    # sort the words from left to right
    # sorted_words = sorted(words, key=lambda w: w.left)

    reordered_words = []

    all_text = ' '.join([w.text for w in words])
    if 'Sahot' in all_text:
        pass
    if 'Cordoba' in all_text:
        pass

    reverse_span = []
    word_num = len(words)

    if word_num <= 1:
        return words

    # sort the words from right to left
    words = sorted(words, key=lambda w: -w.left)
    if re.search('(?=\p{Bidi_Class=Right_to_Left})\p{General_Category=Letter}', words[0].text):
        dir = 'rtl'
    else:
        dir = 'ltr'

    for (i_w, w) in enumerate(words):
        if w.text == '':
            continue
        prev_dir = dir
        if re.search('(?=\p{Bidi_Class=Right_to_Left})\p{General_Category=Letter}', w.text):
            dir = 'rtl'
        elif re.search(r'^\P{L}+$', w.text):
            pass
        else:
            dir = 'ltr'

        if dir != prev_dir:
            if had_rtl:
                adjust_punct(reverse_span)
            reordered_words = reordered_words+reverse_span
            reverse_span = []

        if dir == 'rtl':
            had_rtl = True
            reordered_words.append(w)
        else:
            reverse_span.insert(0, w)

    if had_rtl:
        adjust_punct(reverse_span)
    reordered_words = reordered_words+reverse_span

    """
    for (i_w, w) in enumerate(sorted_words):
        if dir == 'rtl':
            reordered_words.append(w)
        else:
            if ltr_len == 0:
                reordered_words.append(w)
            else:
                reordered_words.insert(-ltr_len, w)
            ltr_len += 1
        new_dir = dir
        if w.text[-1] == uni_ltr:
            new_dir = 'rtl'
        elif w.text[-1] == uni_rtl:
            # we scan from right to left so here we see an LTR word
            new_dir = 'ltr'
        elif i_w < word_num-1:
            next_w = sorted_words[i_w+1].text
            if re.search('\p{Bidi_Class=Right_to_Left}{2}', next_w):
                new_dir = 'rtl'
            elif re.search('\p{Bidi_Class=Left_to_Right}', next_w):
                new_dir = 'ltr'
        if new_dir != dir:
            dir = new_dir
            if dir == 'ltr':
                ltr_len = 0
    """
    return reordered_words

def typeset_words_prev(words:list):

    # u200e - left to right mark
    # u200f - right to left mark

    # sort the words from right to left
    # sorted_words = sorted(words, key=lambda w: -w.left)
    # sort the words from left to right
    # sorted_words = sorted(words, key=lambda w: w.left)

    reordered_words = []

    all_text = ' '.join([w.text for w in words])
    if 'בנעט' in all_text:
        pass
    if 'Granada 1986' in all_text:
        pass

    reverse_span = []
    word_num = len(words)

    if word_num <= 1:
        return words

    # sort the words from right to left
    words_sorted_rl = sorted(words, key=lambda w: -w.left)
    words_sorted_lr = sorted(words, key=lambda w: w.left)

    if words_sorted_rl[0].left != words[0].left and words_sorted_rl[0].left != words[-1].left:
        words = words_sorted_rl
    elif words_sorted_rl[0].left == words[0].left and words_sorted_rl[-1].left == words[-1].left:
        words = words_sorted_rl
    else:
        words = words_sorted_lr

    if words[1].left < words[0].left:
        dir = 'rtl'
    else:
        dir = 'ltr'

    main_dir = dir

    if main_dir == 'rtl' and words[0].text[-1] == uni_rtl:
        dir = 'ltr'
    if main_dir == 'ltr' and words[0].text[-1] == uni_ltr:
        dir = 'rtl'

    """
    if re.search('\p{Bidi_Class=Right_to_Left}{2}', sorted_words[0].text):
        dir = 'rtl'
    """

    for (i_w, w) in enumerate(words):
        if w.text == '':
            continue
        prev_dir = dir
        if dir == main_dir:
            reordered_words.append(w)
        else:
            reverse_span.append(w)
        if w.text[-1] == uni_ltr and i_w > 0:
            dir = 'ltr'
        elif w.text[-1] == uni_rtl and i_w > 0:
            dir = 'rtl'
        elif i_w+1 < word_num and main_dir == 'rtl':
            if re.search('(?=\p{Bidi_Class=Right_to_Left})\p{General_Category=Letter}', words[i_w+1].text):
                dir = 'rtl'
        # elif i_w+1 < word_num and main_dir == 'ltr':
        #     if re.search('(?=\p{Bidi_Class=Left_to_Right})\p{General_Category=Letter}', words[i_w+1].text):
        #         dir = 'ltr'
        if dir != prev_dir and prev_dir != main_dir:
            if main_dir == 'rtl':
                reordered_words = reordered_words+reverse_span[::-1]
            else:
                reordered_words = reverse_span[::-1]+reordered_words
            reverse_span = []

    if len(reverse_span) > 0:
        if main_dir == 'rtl':
            reordered_words = reordered_words+reverse_span[::-1]
        else:
            reordered_words = reverse_span[::-1]+reordered_words

    """
    for (i_w, w) in enumerate(sorted_words):
        if dir == 'rtl':
            reordered_words.append(w)
        else:
            if ltr_len == 0:
                reordered_words.append(w)
            else:
                reordered_words.insert(-ltr_len, w)
            ltr_len += 1
        new_dir = dir
        if w.text[-1] == uni_ltr:
            new_dir = 'rtl'
        elif w.text[-1] == uni_rtl:
            # we scan from right to left so here we see an LTR word
            new_dir = 'ltr'
        elif i_w < word_num-1:
            next_w = sorted_words[i_w+1].text
            if re.search('\p{Bidi_Class=Right_to_Left}{2}', next_w):
                new_dir = 'rtl'
            elif re.search('\p{Bidi_Class=Left_to_Right}', next_w):
                new_dir = 'ltr'
        if new_dir != dir:
            dir = new_dir
            if dir == 'ltr':
                ltr_len = 0
    """
    return reordered_words

def split_words_by_regex(words, had_rtl, regex, trace:str, include_matching:bool=False):
    # this function is called only when we are sure that a match exists
    matching_words = [w for w in words if re.fullmatch(regex, w.text)]
    if len(matching_words) == 0:
        trace += print_to_string("No Matching word in split_words_by_regex")
        return('', '', trace)
    matching_word = matching_words[0]
    if include_matching:
        # include the matching word in the label
        (label_text, info_text) = split_words_by_col(words, matching_word.left-1, had_rtl, remove_word=None)
    else:
        (label_text, info_text) = split_words_by_col(words, matching_word.left, had_rtl, remove_word=matching_word)
    return (label_text, info_text, trace)

def split_words_by_col(words, col, had_rtl, remove_word=None):
    w_label =[ w for w in words if w.left+w.width > col]
    w_info = [ w for w in words if w.left+w.width <= col]
    if remove_word is not None:
        if any(w.text == remove_word.text for w in w_label):
            w_label = [w for w in w_label if w.text != remove_word.text]
        elif any(w.text == remove_word.text for w in w_info):
            w_info = [w for w in w_info if w.text != remove_word.text]
    w_label = typeset_words(w_label, had_rtl)
    w_info = typeset_words(w_info, had_rtl)

    label_text = ' '.join([w.text for w in w_label])
    info_text = ' '.join([w.text for w in w_info])

    return (label_text, info_text)

def is_centered(line, page_width):
    right_offset = page_width-line["bbox"][2]
    left_offset = line["bbox"][0]
    # text_centered = (abs(left_offset-right_offset) <= column_tolerance[self.journal_name]*10)
    # text_centered = (min(left_offset, right_offset) > page_width/4)
    text_centered = (min(left_offset, right_offset)/max(left_offset, right_offset) >= 0.9 and min(left_offset, right_offset) > page_width/4)
    # print(f'right_offset={right_offset} left_offset={left_offset} r={"%.2f" % (min(left_offset, right_offset)/max(left_offset, right_offset))} text_centered={text_centered}')
    return text_centered

    """
    if text_centered:
        if left_offset > 400:
            return True
    return False
    """
