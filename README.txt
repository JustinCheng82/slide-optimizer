Lucid Slides

Files
- index.html: the website and upload interface
- simplify-engine.js: the browser-side PPTX simplification engine
- jszip.min.js: JSZip 3.10.1, used to read and package PPTX files

Run locally
1. Open Terminal in this folder.
2. Run: python3 -m http.server 8000
3. Open: http://localhost:8000

Use
Upload a .pptx file. Processing stays in the browser. When it finishes,
download the file named "(simplified).pptx" and review the listed manual flags.

Important
The automatic edits are intentionally limited. The tool trims long text lines,
bolds their opening words, enlarges standalone statistics, and removes very small
images. It flags charts, dense slides, title clarity, visual hierarchy, animation,
and slide splitting for human review rather than editing them automatically.
