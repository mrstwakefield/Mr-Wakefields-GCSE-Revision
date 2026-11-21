"""Keeps index.html pointing at the current content.js.

The question cards and exam-skills material live in content.js. index.html loads it as
content.js?v=<hash>, so after any change to content.js run:

    python3 tools/stamp-content.py

which updates the hash. A new hash means students' browsers fetch the new copy
straight away instead of using an old cached one.
"""
import hashlib, pathlib, re

root = pathlib.Path(__file__).resolve().parent.parent
content = (root / "content.js").read_bytes()
h = hashlib.sha256(content).hexdigest()[:10]
index = root / "index.html"
html = index.read_text(encoding="utf8")
new, n = re.subn(r'<script src="content\.js\?v=[0-9a-f]*"></script>', f'<script src="content.js?v={h}"></script>', html)
assert n == 1, "content.js script tag not found in index.html"
index.write_text(new, encoding="utf8")
print("content.js version", h)
