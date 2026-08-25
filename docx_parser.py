"""
docx_parser.py
Converts a .docx document into a flat list of "actions" that the automation
layer replays into the VC.ru (Osnova) editor.

Design goal: DO NOT paste. Instead produce a stream of instructions that mimic
a human typing markdown, which the Osnova editor converts live:
    "## "  -> H2
    "### " -> H3
    "- "   -> bullet list
    "1. "  -> numbered list
    "> "   -> quote
    "**x**" (inline) -> bold
    "*x*"  (inline) -> italic

Each action is a dict. Types:
  {"type": "heading", "level": 2, "runs": [...]}
  {"type": "paragraph", "runs": [...]}
  {"type": "bullet", "runs": [...]}
  {"type": "number", "runs": [...]}
  {"type": "quote", "runs": [...]}
  {"type": "image", "path": "<abs path to extracted png/jpg>"}

A "run" is {"text": str, "bold": bool, "italic": bool}.
Consecutive runs with identical style are merged so we type as few markers as
possible (cleaner and faster in the editor).
"""

import os
import zipfile
import tempfile
from docx import Document
from docx.oxml.ns import qn


def _iter_block_items(document):
    """Yield paragraphs in document order (body-level)."""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield ('p', child)
        elif child.tag == qn('w:tbl'):
            yield ('tbl', child)


def _merge_runs(runs):
    """Merge adjacent runs that share bold/italic styling."""
    merged = []
    for r in runs:
        if not r["text"]:
            continue
        if merged and merged[-1]["bold"] == r["bold"] and merged[-1]["italic"] == r["italic"]:
            merged[-1]["text"] += r["text"]
        else:
            merged.append(dict(r))
    return merged


def _extract_runs(paragraph):
    """Pull styled runs out of a python-docx paragraph."""
    runs = []
    for run in paragraph.runs:
        text = run.text
        if text == "":
            continue
        bold = bool(run.bold)
        italic = bool(run.italic)
        # Some docx use style-level bold/italic; check font too
        runs.append({"text": text, "bold": bold, "italic": italic})
    return _merge_runs(runs)


def _para_image_rels(paragraph, document):
    """Return list of embedded image relationship ids inside a paragraph."""
    rels = []
    for blip in paragraph._p.iter(qn('a:blip')):
        rid = blip.get(qn('r:embed'))
        if rid:
            rels.append(rid)
    return rels


def _classify(paragraph):
    """Return (kind, level) for a paragraph based on its style name."""
    name = (paragraph.style.name or "").lower()
    # Headings
    if "heading 1" in name or name == "title":
        return ("heading", 2)   # VC top heading is H2
    if "heading 2" in name:
        return ("heading", 2)
    if "heading 3" in name or "heading 4" in name:
        return ("heading", 3)
    # Quotes
    if "quote" in name or "intense quote" in name:
        return ("quote", 0)
    # Lists
    fmt = paragraph._p.find(qn('w:pPr'))
    if fmt is not None:
        numpr = fmt.find(qn('w:numPr'))
        if numpr is not None:
            # Try to tell bullet vs numbered — default to bullet
            return ("list", 0)
    if "list bullet" in name:
        return ("list", 0)
    if "list number" in name:
        return ("number", 0)
    return ("paragraph", 0)


def parse_docx(path, image_out_dir=None):
    """
    Parse a .docx file into a list of actions.
    Images are extracted to image_out_dir (a temp dir if None) as real files.
    Returns (actions, image_dir).
    """
    if image_out_dir is None:
        image_out_dir = tempfile.mkdtemp(prefix="vc_imgs_")
    os.makedirs(image_out_dir, exist_ok=True)

    document = Document(path)

    # Map relationship id -> extracted file path
    rid_to_file = {}
    idx = 0
    for rid, rel in document.part.rels.items():
        if "image" in rel.reltype:
            try:
                blob = rel.target_part.blob
                ext = os.path.splitext(rel.target_part.partname)[1] or ".png"
                fname = os.path.join(image_out_dir, f"img_{idx}{ext}")
                with open(fname, "wb") as f:
                    f.write(blob)
                rid_to_file[rid] = fname
                idx += 1
            except Exception:
                pass

    actions = []
    for kind_tag, element in _iter_block_items(document):
        if kind_tag != 'p':
            continue
        from docx.text.paragraph import Paragraph
        para = Paragraph(element, document)

        # Images first (if the paragraph holds any)
        img_rels = _para_image_rels(para, document)
        for rid in img_rels:
            if rid in rid_to_file:
                actions.append({"type": "image", "path": rid_to_file[rid]})

        runs = _extract_runs(para)
        if not runs:
            continue

        kind, level = _classify(para)
        if kind == "heading":
            actions.append({"type": "heading", "level": level, "runs": runs})
        elif kind == "list":
            actions.append({"type": "bullet", "runs": runs})
        elif kind == "number":
            actions.append({"type": "number", "runs": runs})
        elif kind == "quote":
            actions.append({"type": "quote", "runs": runs})
        else:
            actions.append({"type": "paragraph", "runs": runs})

    return actions, image_out_dir


def runs_to_markdown(runs):
    """Convert styled runs to inline markdown string."""
    out = []
    for r in runs:
        t = r["text"]
        # Trim markers to inside the visible (non-space) text so the editor
        # recognizes them. Move leading/trailing spaces outside the markers.
        lead = len(t) - len(t.lstrip(" "))
        trail = len(t) - len(t.rstrip(" "))
        core = t[lead: len(t) - trail] if trail else t[lead:]
        prefix = t[:lead]
        suffix = t[len(t) - trail:] if trail else ""
        if core and (r["bold"] or r["italic"]):
            marker = ""
            if r["bold"]:
                marker += "**"
            if r["italic"]:
                marker += "*"
            core = f"{marker}{core}{marker[::-1]}"
        out.append(f"{prefix}{core}{suffix}")
    return "".join(out)


if __name__ == "__main__":
    import sys
    acts, d = parse_docx(sys.argv[1])
    for a in acts:
        if a["type"] == "image":
            print("IMAGE:", a["path"])
        else:
            print(f'{a["type"]}({a.get("level","")}):', runs_to_markdown(a["runs"]))
