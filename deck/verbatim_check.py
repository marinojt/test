"""Check that every locked-draft paragraph for the built slides appears verbatim in the deck text."""
import re, sys, json
from pptx import Presentation
deck, src = sys.argv[1], sys.argv[2]
norm = lambda t: re.sub(r'\s+', ' ', t).strip().lower()
text = norm(' '.join(sh.text_frame.text for sl in Presentation(deck).slides for sh in sl.shapes if sh.has_text_frame))
expected = json.load(open(src))
missing = [e for e in expected if norm(e) not in text]
print(f'{len(expected) - len(missing)}/{len(expected)} locked paragraphs found verbatim')
for m in missing: print('MISSING:', m)
