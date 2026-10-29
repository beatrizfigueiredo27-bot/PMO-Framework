"""Build the complete PMI Predictive Project Management Toolkit.
Usage:  python build_toolkit.py [output_dir]
Edit config.py first to stamp your own project metadata on all artefacts."""
import os, sys, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import REGISTER
import framework_book as FB, docs_word_a as A, docs_word_b as Bw, books_a as XA, books_b as XB, deck

OUT = sys.argv[1] if len(sys.argv) > 1 else "PMI Predictive PM Toolkit - PMBOK 8"
os.makedirs(OUT, exist_ok=True)
builders = {
    "01": A.build_01, "02": A.build_02, "03": A.build_03, "04": XA.build_04, "05": XA.build_05,
    "06": A.build_06, "07": XA.build_07, "08": A.build_08, "09": XA.build_09, "10": XB.build_10,
    "11": A.build_11, "12": XB.build_12, "13": Bw.build_13, "14": XB.build_14, "15": Bw.build_15,
    "16": XB.build_16, "17": Bw.build_17, "18": Bw.build_18, "19": XB.build_19, "20": Bw.build_20,
    "21": XB.build_21, "22": deck.build_22_pptx, "23": XB.build_23, "24": XB.build_24, "25": Bw.build_25,
}
files = []
for no, title, code, ext, *_ in list(REGISTER):
    if no == "00":
        continue
    p = os.path.join(OUT, f"{no} — {title}.{ext}")
    builders[no](p); files.append(p)
p = os.path.join(OUT, "22 — Project Status Report.docx"); Bw.build_22_docx(p); files.append(p)
p = os.path.join(OUT, "00 — Master PMO Documentation Framework.xlsx"); FB.build_framework(p); files.append(p)
p = os.path.join(OUT, "00 — Toolkit Guide & Document Standards.docx"); A.build_00(p); files.append(p)
print(f"Built {len(files)} files in {OUT}")
