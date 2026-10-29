"""Assemble the PMO Documentation Framework portal.

Modes
  python portal.py <toolkit_dir> <out.html>
      Self-contained page with the templates embedded (Claude artifact).
  python portal.py <toolkit_dir> <site_dir> --static [--zip <toolkit.zip>]
      Static website for GitHub Pages / any host: site_dir/index.html + site_dir/files/.
  python portal.py <toolkit_dir> <out.html> --base <https://host/path/files/>
      Single page whose Download buttons link to files already hosted at --base.
"""
import argparse, base64, json, os, re, shutil, sys, unicodedata
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import PROJECT, DOC_DEFAULTS, DOMAINS, FOCUS_AREAS, GLOSSARY, VERSIONING, PRINCIPLES
from framework import artefacts, GATES, CATS, AUTHORITY, ESCALATION, RAG, DOC_STANDARD

ap = argparse.ArgumentParser()
ap.add_argument("toolkit_dir"); ap.add_argument("out")
ap.add_argument("--static", action="store_true"); ap.add_argument("--base")
ap.add_argument("--zip", help="toolkit zip to publish alongside the files")
args = ap.parse_args()


def slug(name):
    stem, ext = os.path.splitext(name)
    s = unicodedata.normalize("NFKD", stem.replace("×", "x").replace("&", "and").replace("+", "and"))
    s = re.sub(r"[^A-Za-z0-9]+", "-", s.encode("ascii", "ignore").decode()).strip("-").lower()
    return s + ext.lower()


arts = artefacts()
data = dict(
    project=PROJECT, version=DOC_DEFAULTS["version"], status=DOC_DEFAULTS["status"],
    artefacts=arts, domains=DOMAINS, focus=FOCUS_AREAS, principles=PRINCIPLES,
    gates=[dict(id=g[0], name=g[1], fa=g[2], appr=g[3], crit=g[4]) for g in GATES],
    cats=CATS, authority=AUTHORITY, escalation=ESCALATION, rag=RAG, docStandard=DOC_STANDARD,
    glossary=GLOSSARY, versioning=VERSIONING,
    meta={"Organization": PROJECT["org"], "Project name": PROJECT["name"], "Project ID": PROJECT["id"],
          "Sponsor": PROJECT["sponsor"], "Project Manager": PROJECT["pm"], "PMO / Portfolio": PROJECT["pmo"],
          "Planned start / finish": f"{PROJECT['start']} / {PROJECT['finish']}", "Currency": PROJECT["currency"],
          "Delivery approach": PROJECT["approach"], "Classification": PROJECT["classification"]},
)

if args.static:
    site = args.out
    files_dir = os.path.join(site, "files")
    shutil.rmtree(site, ignore_errors=True); os.makedirs(files_dir)
    out_html = os.path.join(site, "index.html")
    base, external = "files/", False
elif args.base:
    out_html, base, external = args.out, args.base, True
else:
    out_html, base, external = args.out, "", False

files = {"list": [], "data": {}, "base": base, "external": external}
for name in sorted(os.listdir(args.toolkit_dir)):
    p = os.path.join(args.toolkit_dir, name)
    if not os.path.isfile(p) or name.split(".")[-1] not in ("docx", "xlsx", "pptx"):
        continue
    m = re.match(r"(\d\d) — (.+)\.(\w+)$", name)
    no, label, ext = m.group(1), m.group(2), m.group(3)
    raw = open(p, "rb").read()
    entry = dict(name=name, path=slug(name), label=label + (" (Word)" if no == "22" and ext == "docx" else ""),
                 ext=ext, size=len(raw), no=None if no == "00" else no)
    if no == "00":
        entry["id"] = PROJECT["id"] + ("-MDF" if ext == "xlsx" else "-GDE")
    files["list"].append(entry)
    if not base:
        files["data"][name] = base64.b64encode(raw).decode()
    elif args.static:
        shutil.copyfile(p, os.path.join(files_dir, entry["path"]))

zip_name = "PMI_Predictive_PM_Toolkit_PMBOK8.zip"
if base:
    files["zip"] = dict(name=zip_name, path=zip_name)
    if args.static and args.zip:
        shutil.copyfile(args.zip, os.path.join(files_dir, zip_name))

j = lambda o: json.dumps(o, ensure_ascii=False).replace("</", "<\\/")
tpl = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "portal_template.html"), encoding="utf-8").read()
html = tpl.replace("__DATA__", j(data)).replace("__FILES__", j(files))

if args.static:
    # a static host serves the file as-is, so give it a full document skeleton
    title = re.search(r"<title>.*?</title>", html, re.S).group(0)
    head_bits = re.findall(r"<meta name=\"claude[^>]*>|<link [^>]*>", html)
    body = html.replace(title, "", 1)
    for b in head_bits:
        body = body.replace(b, "", 1)
    style = re.search(r"<style>.*?</style>", body, re.S).group(0)
    body = body.replace(style, "", 1)
    html = ("<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
            f"{title}\n<meta name=\"description\" content=\"Master PMO Documentation Register and free predictive "
            "project management templates aligned to the PMBOK Guide - Eighth Edition.\">\n"
            + "\n".join(b for b in head_bits if b.startswith("<link")) + "\n" + style +
            "\n<style>:root{padding-top:env(safe-area-inset-top);padding-bottom:env(safe-area-inset-bottom)}"
            "body{margin:0}img{max-width:100%}</style>\n</head>\n<body>\n" + body + "\n</body>\n</html>\n")
    open(os.path.join(site, ".nojekyll"), "w").close()

open(out_html, "w", encoding="utf-8").write(html)
print(f"{out_html}: {len(html)/1e6:.2f} MB, {len(files['list'])} files, base={base or 'embedded'}")
