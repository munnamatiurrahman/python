"""
JSON to MCQ PowerPoint Generator (SutonnyMJ / Bijoy 52 Font)
============================================================
This script scans the current folder for all .json files and converts them
into formatted PowerPoint presentations (.pptx) using MCQ Format.pptx layout.
Output files are saved in an 'output' folder with the same base name as the JSON files.
English text will automatically use Times New Roman font, while Bengali uses SutonnyMJ.

Usage:
    python b.py
"""

import os
import sys
import copy
import json
import re
import glob
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

# Fix Windows console UTF-8 output if available
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# --- Bijoy 52 / SutonnyMJ Converter ---
DIGITS = {'০':'0', '১':'1', '২':'2', '৩':'3', '৪':'4', '৫':'5', '৬':'6', '৭':'7', '৮':'8', '৯':'9'}
VOWELS = {'অ':'A', 'আ':'Av', 'ই':'B', 'ঈ':'C', 'উ':'D', 'ঊ':'E', 'ঋ':'F', 'এ':'G', 'ঐ':'H', 'ও':'I', 'ঔ':'J'}
CONSONANTS = {
    'ক':'K', 'খ':'L', 'গ':'M', 'ঘ':'N', 'ঙ':'O',
    'চ':'P', 'ছ':'Q', 'জ':'R', 'ঝ':'S', 'ঞ':'T',
    'ট':'U', 'ঠ':'V', 'ড':'W', 'ঢ':'X', 'ণ':'Y',
    'ত':'Z', 'থ':'_', 'দ':'`', 'ধ':'a', 'ন':'b',
    'প':'c', 'ফ':'d', 'ব':'e', 'ভ':'f', 'ম':'g',
    'য':'h', 'র':'i', 'ল':'j', 'শ':'k', 'ষ':'l',
    'স':'m', 'হ':'n', 'ড়':'o', 'ঢ়':'p', 'য়':'q',
    'ৎ':'r', 'ং':'s', 'ঃ':'t', 'ঁ':'u'
}

JUNKT = [
    ('গু', '¸'), ('শু', 'ï'), ('হু', 'û'),
    ('রু', 'iæ'), ('রূ', 'iƒ'),
    ('ক্রু', 'µæ'), ('ত্রু', 'Îæ'), ('দ্রু', '`Öæ'), ('ধ্রু', 'aÖæ'),
    ('প্রু', 'cÖæ'), ('ব্রু', 'eÖæ'), ('ভ্রু', 'fÖæ'), ('গ্রু', 'MÖæ'), ('ম্রু', 'gÖæ'), ('শ্রু', 'kÖæ'),
    ('ন্ত্র', 'šÍÖ'), ('স্ত্র', '¯ÍÖ'), ('ন্ড্র', '›WÖ'), ('ন্ট্র', '›UÖ'), ('স্ট্র', '¯UÖ'),
    ('ক্ষ্ম', '²¥'), ('ক্ষ', '²'), ('জ্ঞ', 'Á'), ('ঞ্চ', 'Â'), ('ঞ্ছ', 'Ã'), ('ঞ্জ', 'Ä'),
    ('চ্চ', '”P'), ('চ্ছ', '”Q'), ('জ্জ', '¾'), ('জ্ঝ', '¿'),
    ('ক্ট', '±'), ('ক্ত', '³'), ('ক্ব', '´'), ('ক্ল', 'K¬'), ('ক্স', '·'),
    ('ঙ্ক', '¼'), ('ঙ্গ', '½'), ('ক্ক', '°'), ('ট্ট', 'Æ'), ('ট্ব', 'Ç'),
    ('ট্ম', 'È'), ('ড্ড', 'É'), ('ণ্ট', '›U'), ('ণ্ঠ', 'Ú'), ('ণ্ড', 'Ð'), ('ণ্ণ', 'Í'),
    ('ত্ত', 'Ë'), ('ত্ম', 'Z¥'), ('ত্ব', 'Z¡'), ('ত্র', 'Î'), ('থ্ব', '_¡'),
    ('দ্গ', '`Î'), ('দ্দ', 'Ï'), ('দ্ধ', '×'), ('দ্ব', 'Ø'), ('দ্ভ', '™'), ('দ্ম', '®'),
    ('ধ্বনি', 'a¡wb'), ('ধ্ব', 'a¡'),
    ('ন্ট', '›U'), ('ন্ঠ', '›V'), ('ন্ড', '›W'), ('ন্ত', 'šÍ'), ('ন্থ', 'š’'),
    ('ন্দ', '›`'), ('ন্ধ', '›a'), ('ন্ন', 'bœ'), ('ন্ব', 'b¡'), ('ন্ম', 'b¥'),
    ('প্ট', 'Þ'), ('প্ত', 'ß'), ('প্ন', 'cœ'), ('প্প', 'à'), ('প্ল', 'cø'), ('প্স', 'â'),
    ('ফ্ল', 'd¬'), ('ব্জ', 'ã'), ('ব্দ', 'ä'), ('ব্ধ', 'å'),
    ('ব্ব', 'eŸ'),
    ('ব্ল', 'eø'),
    ('ভ্ল', 'f¬'), ('ম্ন', 'gœ'), ('ম্প', '¤c'), ('ম্ফ', '¤d'), ('ম্ব', '¤e'), ('ম্ভ', '¤f'),
    ('ম্ম', '¤g'), ('ম্ল', 'gø'),
    ('ল্ক', 'é'), ('ল্গ', 'ê'), ('ল্ট', 'ë'), ('ল্ড', 'ì'), ('ল্প', 'í'), ('ল্ফ', 'î'),
    ('ল্ব', 'j¡'), ('ল্ম', 'j¥'), ('ল্ল', 'jø'), ('শ্চ', 'ð'), ('শ্ব', 'k¡'), ('শ্ম', 'k¥'),
    ('শ্ল', 'k¬'), ('ষ্ক', '®K'), ('ষ্ট', 'ó'), ('ষ্ঠ', 'ô'), ('ষ্ণ', 'ò'),
    ('স্প', '¯ú'), ('স্ফ', 'ù'), ('স্ত', '¯Í'), ('স্থ', '¯’'), ('স্ন', 'mœ'),
    ('স্ম', '¯§'), ('স্ব', 'm¡'), ('স্ল', 'm¬'), ('স্ক', '¯‹'), ('স্ট', '÷'), ('স্খ', '¯L'),
    ('হ্ণ', 'ý'), ('হ্ন', 'ý'), ('হ্ম', 'þ'), ('হ্ল', 'n¬'), ('হ্ব', 'n¡'),
    ('গ্র', 'MÖ'), ('প্র', 'cÖ'), ('ব্র', 'eÖ'), ('ভ্র', 'fÖ'), ('ম্র', 'gÖ'), ('শ্র', 'kÖ'),
    ('ক্র', 'µ'), ('ফ্র', 'dÖ'), ('ধ্র', 'aÖ'), ('দ্র', '`Ö'), ('জ্ব', 'R¡')
]

def to_bijoy(text):
    if not text:
        return ""
    text = text.replace('য়', 'য়').replace('ড়', 'ড়').replace('ঢ়', 'ঢ়').replace('অা', 'আ')
    text = text.replace('–', '-').replace('—', '-')
    text = re.sub(r'্য', '¨', text)
    for u, b in DIGITS.items():
        text = text.replace(u, b)
    reph_targets = list(CONSONANTS.keys()) + [j[0] for j in JUNKT]
    reph_targets.sort(key=len, reverse=True)
    reph_pattern = r'র্(' + '|'.join(reph_targets) + r')'
    text = re.sub(reph_pattern, r'\1©', text)
    for u, b in JUNKT:
        text = text.replace(u, b)
    ansi_conjuncts = [re.escape(j[1]) for j in JUNKT]
    uni_consonants = list(CONSONANTS.keys())
    combined_targets = uni_consonants + ansi_conjuncts
    combined_targets.sort(key=len, reverse=True)
    consonant_pattern = r'(' + '|'.join(combined_targets) + r')(?:©)?(?:¨|Ö)*'
    text = re.sub(consonant_pattern + r'ো', r'‡\1v', text)
    text = re.sub(consonant_pattern + r'ৌ', r'‡\1Š', text)
    text = re.sub(consonant_pattern + r'ি', r'w\1', text)
    text = re.sub(consonant_pattern + r'ে', r'‡\1', text)
    text = re.sub(consonant_pattern + r'ৈ', r'ˆ\1', text)
    text = text.replace('া', 'v').replace('ী', 'x').replace('ু', 'z').replace('ূ', '~').replace('ৃ', '„').replace('্য', '¨').replace('্র', 'Ö')
    for u, b in VOWELS.items():
        text = text.replace(u, b)
    for u, b in CONSONANTS.items():
        text = text.replace(u, b)
    return text

def parse_multi_statement(q_text):
    lines = [l.strip() for l in q_text.split('\n') if l.strip()]
    stem = ""
    i_text = ""
    ii_text = ""
    iii_text = ""
    bottom_text = "নিচের কোনটি সঠিক?"
    for line in lines:
        if line.startswith('i.') or line.startswith('i '):
            i_text = re.sub(r'^i[\.\s]+', '', line).strip()
        elif line.startswith('ii.') or line.startswith('ii '):
            ii_text = re.sub(r'^ii[\.\s]+', '', line).strip()
        elif line.startswith('iii.') or line.startswith('iii '):
            iii_text = re.sub(r'^iii[\.\s]+', '', line).strip()
        elif 'নিচের কোনটি সঠিক' in line:
            bottom_text = line
        else:
            if not i_text:
                if stem:
                    stem += "\n" + line
                else:
                    stem = line
    return stem, i_text, ii_text, iii_text, bottom_text


def update_shape_runs(shape, runs_spec, force_size=None):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    
    # Extract original formatting (rPr) to preserve White Color, Bold, Size, etc.
    orig_rPr = None
    if tf.paragraphs:
        p0 = tf.paragraphs[0]
        if p0.runs:
            r0 = p0.runs[0]
            orig_rPr = copy.deepcopy(r0._r.find(qn('a:rPr')))
        
        # Fallback to paragraph default run properties if no run exists
        if orig_rPr is None:
            pPr = p0._p.get_or_add_pPr()
            defRPr = pPr.find(qn('a:defRPr'))
            if defRPr is not None:
                orig_rPr = copy.deepcopy(defRPr)
        
        # Safely clear the text frame without destroying paragraph properties
        for r in list(p0.runs):
            r._r.getparent().remove(r._r)
            
        for br in p0._p.findall(qn('a:br')):
            br.getparent().remove(br)
            
        for p in list(tf.paragraphs)[1:]:
            p._p.getparent().remove(p._p)
    else:
        p0 = tf.paragraphs[0]

    for text_str, font_name in runs_spec:
        if not text_str:
            continue
        
        # Handle newlines properly by inserting <a:br/> elements
        lines = text_str.split('\n')
        for i, line in enumerate(lines):
            if i > 0:
                br = p0._p.makeelement(qn('a:br'), {})
                p0._p.append(br)
            
            if line:
                r = p0.add_run()
                r.text = line
                
                # Apply the copied original formatting (preserves white color)
                if orig_rPr is not None:
                    new_rPr = copy.deepcopy(orig_rPr)
                    existing_rPr = r._r.find(qn('a:rPr'))
                    if existing_rPr is not None:
                        r._r.remove(existing_rPr)
                    r._r.insert(0, new_rPr)
                
                # Set font name across all possible font slots (Latin, Complex Script, East Asian)
                rPr = r._r.get_or_add_rPr()
                for tag in ('a:latin', 'a:cs', 'a:ea'):
                    el = rPr.find(qn(tag))
                    if el is None:
                        el = rPr.makeelement(qn(tag), {})
                        rPr.append(el)
                    el.set('typeface', font_name)
                
                # Override font size if explicitly forced
                if force_size:
                    r.font.size = force_size


def create_mixed_runs(text, bengali_font="SutonnyMJ", english_font="Times New Roman"):
    """
    Splits text into Bengali and English/Other parts.
    Bengali parts are converted to Bijoy and use SutonnyMJ.
    English parts remain unchanged and use Times New Roman.
    """
    if not text:
        return []
    
    runs = []
    # Split by Bengali Unicode characters
    parts = re.split(r'([\u0980-\u09FF]+)', text)
    
    for part in parts:
        if not part:
            continue
        if re.search(r'[\u0980-\u09FF]', part):
            runs.append((to_bijoy(part), bengali_font))
        else:
            runs.append((part, english_font))
            
    return runs


def generate_mcq_pptx(json_path, template_path, output_path):
    if not os.path.exists(json_path):
        print(f"Error: JSON file '{json_path}' not found.")
        return
    if not os.path.exists(template_path):
        print(f"Error: Template PowerPoint file '{template_path}' not found.")
        sys.exit(1)

    print(f"  Loading JSON data...")
    with open(json_path, 'r', encoding='utf-8') as f:
        mcqs = json.load(f)

    print(f"  Loading PowerPoint template...")
    prs = Presentation(template_path)

    if len(prs.slides) < 3:
        print("Error: Template PowerPoint must contain at least 3 slides.")
        return

    std_template = prs.slides[1]
    multi_template = prs.slides[2]

    answer_letters = ["K", "L", "M", "N"]

    print(f"  Processing {len(mcqs)} MCQs...")

    for idx, q in enumerate(mcqs, start=1):
        q_text = q.get('question', '')
        stimulus = q.get('stimulus', '').strip()
        board_year = q.get('board_and_year', '').strip()
        options = q.get('options', [])
        ans_idx = q.get('correctAnswer', 0)
        ans_letter = answer_letters[ans_idx] if 0 <= ans_idx < len(answer_letters) else "K"

        is_multi = ('i.' in q_text and 'ii.' in q_text and 'iii.' in q_text and 'নিচের কোনটি সঠিক' in q_text)
        source_slide = multi_template if is_multi else std_template

        new_slide = prs.slides.add_slide(source_slide.slide_layout)
        for shape in list(new_slide.shapes):
            sp = shape._element
            sp.getparent().remove(sp)

        for shape in source_slide.shapes:
            new_sp = copy.deepcopy(shape._element)
            new_slide.shapes._spTree.append(new_sp)

        for shape in new_slide.shapes:
            sname = shape.name

            if is_multi:
                if sname == "Rectangle 10":
                    by_str = f"[ {board_year} ]" if board_year else ""
                    update_shape_runs(shape, create_mixed_runs(by_str))
                elif sname == "Rectangle 3":
                    stem, _, _, _, _ = parse_multi_statement(q_text)
                    full_stem = f"{stimulus}\n\n{stem}" if stimulus else stem
                    update_shape_runs(shape, create_mixed_runs(f" {full_stem} "))
                elif sname == "Rounded Rectangle 8":
                    _, i_t, _, _, _ = parse_multi_statement(q_text)
                    update_shape_runs(shape, [(" i. ", "Times New Roman")] + create_mixed_runs(i_t))
                elif sname == "Rounded Rectangle 5":
                    _, _, ii_t, _, _ = parse_multi_statement(q_text)
                    update_shape_runs(shape, [(" ii. ", "Times New Roman")] + create_mixed_runs(ii_t))
                elif sname == "Rounded Rectangle 2":
                    _, _, _, iii_t, _ = parse_multi_statement(q_text)
                    update_shape_runs(shape, [(" iii. ", "Times New Roman")] + create_mixed_runs(iii_t))
                elif sname == "Rectangle 4":
                    update_shape_runs(shape, create_mixed_runs(" নিচের কোনটি সঠিক? "))
                elif sname == "Rounded Rectangle 6":
                    opt0 = options[0] if len(options) > 0 else ""
                    update_shape_runs(shape, [("  K) ", "SutonnyMJ")] + create_mixed_runs(opt0), force_size=Pt(66))
                elif sname == "Rounded Rectangle 11":
                    opt1 = options[1] if len(options) > 1 else ""
                    update_shape_runs(shape, [("  L) ", "SutonnyMJ")] + create_mixed_runs(opt1), force_size=Pt(66))
                elif sname == "Rounded Rectangle 12":
                    opt2 = options[2] if len(options) > 2 else ""
                    update_shape_runs(shape, [("  M) ", "SutonnyMJ")] + create_mixed_runs(opt2), force_size=Pt(66))
                elif sname == "Rounded Rectangle 13":
                    opt3 = options[3] if len(options) > 3 else ""
                    update_shape_runs(shape, [("  N) ", "SutonnyMJ")] + create_mixed_runs(opt3), force_size=Pt(66))
                elif sname == "Rectangle 9":
                    update_shape_runs(shape, [(ans_letter, "SutonnyMJ")])
                elif sname == "Rectangle 14":
                    update_shape_runs(shape, [("  eûwbev©Pwb mgvavb", "SutonnyMJ")])
            else:
                if sname == "Rectangle 1":
                    by_str = f"[  {board_year}  ]" if board_year else ""
                    update_shape_runs(shape, create_mixed_runs(by_str))
                elif sname == "Rectangle 5":
                    full_q = f"{stimulus}\n\n{q_text}" if stimulus else q_text
                    update_shape_runs(shape, create_mixed_runs(f" {full_q} "))
                elif sname == "Rectangle: Rounded Corners 6":
                    opt0 = options[0] if len(options) > 0 else ""
                    update_shape_runs(shape, [(" K) ", "SutonnyMJ")] + create_mixed_runs(opt0), force_size=Pt(66))
                elif sname == "Rectangle: Rounded Corners 7":
                    opt1 = options[1] if len(options) > 1 else ""
                    update_shape_runs(shape, [(" L) ", "SutonnyMJ")] + create_mixed_runs(opt1), force_size=Pt(66))
                elif sname == "Rectangle: Rounded Corners 8":
                    opt2 = options[2] if len(options) > 2 else ""
                    update_shape_runs(shape, [(" M) ", "SutonnyMJ")] + create_mixed_runs(opt2), force_size=Pt(66))
                elif sname == "Rectangle: Rounded Corners 9":
                    opt3 = options[3] if len(options) > 3 else ""
                    update_shape_runs(shape, [(" N) ", "SutonnyMJ")] + create_mixed_runs(opt3), force_size=Pt(66))
                elif sname == "Rectangle 2":
                    update_shape_runs(shape, [(ans_letter, "SutonnyMJ")])
                elif sname == "Rectangle 10":
                    update_shape_runs(shape, [("eûwbev©Pwb mgvavb", "SutonnyMJ")])

    rId_list = [prs.slides._sldIdLst[i] for i in range(3)]
    for rId in rId_list:
        prs.slides._sldIdLst.remove(rId)

    prs.save(output_path)
    print(f"  -> Saved successfully!")


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "MCQ Format.pptx")
    
    if not os.path.exists(template_path):
        print(f"Error: Template file 'MCQ Format.pptx' not found in {script_dir}")
        sys.exit(1)

    json_files = glob.glob(os.path.join(script_dir, "*.json"))
    
    if not json_files:
        print("No .json files found in the current folder.")
        return

    output_dir = os.path.join(script_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"Found {len(json_files)} JSON file(s). Starting batch processing...\n")
    
    for json_path in json_files:
        base_name = os.path.splitext(os.path.basename(json_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}.pptx")
        
        print(f"Processing: {os.path.basename(json_path)}")
        try:
            generate_mcq_pptx(json_path, template_path, output_path)
        except Exception as e:
            print(f"  Error processing {os.path.basename(json_path)}: {e}")
        
        print("-" * 50)
        
    print(f"\nBatch Complete! All generated files are saved in the 'output' folder.")


if __name__ == "__main__":
    main()