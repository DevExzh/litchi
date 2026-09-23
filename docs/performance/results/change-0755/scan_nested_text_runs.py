#!/usr/bin/env python3
"""Count slide-like parts of the repository's PPTX fixtures that contain a
DrawingML text element (`a:t`, start or empty) inside an open `a:t`.

Run from the repository root. Slide-like parts are the owners the scene reader
and the opened text edits read: slides, layouts, masters, notes and handout
masters. SmartArt data parts are excluded on purpose: their `dgm:t` element
legitimately contains `a:t` and is not read by the scene reader.
"""

import glob
import re
import zipfile

OWNERS = re.compile(
    r"^ppt/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|handoutMasters)/[^/]+\.xml$"
)
TEXT = re.compile(rb"<(/?)a:t(?=[\s/>])[^>]*?(/?)>")

fixtures = sorted(glob.glob("test-data/**/*.pptx", recursive=True))
parts = 0
nested = []
for fixture in fixtures:
    archive = zipfile.ZipFile(fixture)
    for name in archive.namelist():
        if not OWNERS.match(name):
            continue
        parts += 1
        open_text = False
        for match in TEXT.finditer(archive.read(name)):
            closing, empty = match.groups()
            if closing:
                open_text = False
            elif open_text:
                nested.append(f"{fixture} {name}")
                break
            elif not empty:
                open_text = True
print(f"fixtures={len(fixtures)} slide_like_parts={parts} with_nested_text={len(nested)}")
for entry in nested:
    print(entry)
