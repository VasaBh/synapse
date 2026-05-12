"""
Extracts content from docx files and injects HTML into topic pages.
Uses ElementTree for proper XML parsing and code-block detection.
"""
import zipfile, re, os, sys
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def wtag(name):
    return f'{{{W}}}{name}'

def parse_doc(path):
    """Return list of (style, text, is_code) tuples for each paragraph."""
    with zipfile.ZipFile(path) as z:
        raw = z.read('word/document.xml')
    root = ET.fromstring(raw)
    body = root.find(f'{{{W}}}body')
    results = []
    for para in body.iter(f'{{{W}}}p'):
        pPr = para.find(f'{{{W}}}pPr')
        style = 'Normal'
        is_code = False
        if pPr is not None:
            ps = pPr.find(f'{{{W}}}pStyle')
            if ps is not None:
                style = ps.get(f'{{{W}}}val', 'Normal')
            shd = pPr.find(f'{{{W}}}shd')
            if shd is not None:
                fill = shd.get(f'{{{W}}}fill', '')
                if fill.upper() in ('F1F5F9', 'E2E8F0', 'EFF6FF', 'F8FAFC', 'F0F9FF', '1E293B'):
                    is_code = True
        # Also detect code by Consolas font in any run
        for run in para.findall(f'{{{W}}}r'):
            rPr = run.find(f'{{{W}}}rPr')
            if rPr is not None:
                rFonts = rPr.find(f'{{{W}}}rFonts')
                if rFonts is not None:
                    a = rFonts.get(f'{{{W}}}ascii', '') or rFonts.get(f'{{{W}}}hAnsi', '')
                    if 'Consolas' in a or 'Courier' in a or 'Mono' in a:
                        is_code = True
        # Collect text from all w:t elements
        texts = []
        for t in para.findall(f'.//{{{W}}}t'):
            texts.append(t.text or '')
        text = ''.join(texts).strip()
        results.append((style, text, is_code))
    return results

# ── HTML helpers ────────────────────────────────────────────────────────────

def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def is_subsection(text):
    return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+\S', text)) and len(text) < 120

def is_main_section(text):
    return bool(re.match(r'^\d+\.\s+[A-Z]', text)) and len(text) < 120

def paragraphs_to_html(paras):
    """Convert list of (style, text, is_code) to HTML string."""
    html_parts = []
    i = 0
    while i < len(paras):
        style, text, is_code = paras[i]
        if not text:
            i += 1
            continue

        # ── Code block ──────────────────────────────────────────────────────
        if is_code:
            code_lines = []
            while i < len(paras) and (paras[i][2] or (not paras[i][1] and code_lines)):
                if paras[i][1]:
                    code_lines.append(esc(paras[i][1]))
                i += 1
            if code_lines:
                html_parts.append(f'<pre><code>{chr(10).join(code_lines)}</code></pre>')
            continue

        # ── Heading styles ──────────────────────────────────────────────────
        lvl = 0
        if style in ('Heading1', 'heading1'): lvl = 1
        elif style in ('Heading2', 'heading2'): lvl = 2
        elif style in ('Heading3', 'heading3'): lvl = 3

        if lvl == 1:
            html_parts.append(f'<h2>{esc(text)}</h2>')
            i += 1; continue
        if lvl == 2:
            html_parts.append(f'<h3>{esc(text)}</h3>')
            i += 1; continue
        if lvl == 3:
            html_parts.append(f'<h4>{esc(text)}</h4>')
            i += 1; continue

        # ── Numbered subsection headers in Normal text ──────────────────────
        if is_main_section(text):
            html_parts.append(f'<h2>{esc(text)}</h2>')
            i += 1; continue
        if is_subsection(text):
            depth = text.count('.')
            tag = 'h3' if depth == 1 else 'h4'
            html_parts.append(f'<{tag}>{esc(text)}</{tag}>')
            i += 1; continue

        # ── Bullet list ─────────────────────────────────────────────────────
        if 'Bullet' in style or 'ListParagraph' in style or text.startswith('- '):
            items = []
            while i < len(paras):
                s2, t2, _ = paras[i]
                if 'Bullet' in s2 or 'ListParagraph' in s2:
                    items.append(t2)
                    i += 1
                elif t2.startswith('- '):
                    items.append(t2[2:])
                    i += 1
                else:
                    break
            if items:
                lis = '\n'.join(f'  <li>{esc(x)}</li>' for x in items if x)
                html_parts.append(f'<ul>\n{lis}\n</ul>')
            continue

        # ── Interview Tip / callout ─────────────────────────────────────────
        if text.startswith('💡') or text.startswith('⚠️') or text.startswith('✅'):
            html_parts.append(f'<div class="callout">{esc(text)}</div>')
            i += 1; continue

        # ── Table detection: 3+ consecutive short lines (<= 5 words) ────────
        # Look ahead to see if next several lines are also short (table cells)
        def looks_like_cell(t):
            return t and len(t.split()) <= 6 and not t.startswith('-') and not is_subsection(t)

        if looks_like_cell(text):
            j = i
            row_lines = []
            while j < len(paras) and looks_like_cell(paras[j][1]) and not paras[j][2]:
                row_lines.append(paras[j][1])
                j += 1
            if len(row_lines) >= 3:
                # Render as a table (assume groups of N cols)
                # Find column count from first header-like row count
                # Simple: detect columns by grouping 3 cells per row
                # Look at the first run for column names
                # Heuristic: use 3 columns if divisible else just output as list
                col_count = 2
                for c in (3, 4, 2):
                    if len(row_lines) % c == 0:
                        col_count = c
                        break
                rows = [row_lines[k:k+col_count] for k in range(0, len(row_lines), col_count)]
                tbl = ['<table class="data-table">']
                for ri, row in enumerate(rows):
                    cells = ''.join(
                        f'<{"th" if ri == 0 else "td"}>{esc(c)}</{("th" if ri==0 else "td")}>'
                        for c in row
                    )
                    tbl.append(f'  <tr>{cells}</tr>')
                tbl.append('</table>')
                html_parts.append('\n'.join(tbl))
                i = j
                continue

        # ── Regular paragraph ───────────────────────────────────────────────
        html_parts.append(f'<p>{esc(text)}</p>')
        i += 1

    return '\n'.join(html_parts)


def split_sections(paras):
    """Split paragraphs into {section_title: [paras]} dict keyed by Heading1."""
    sections = {}
    current = '_intro'
    buf = []
    for style, text, is_code in paras:
        is_h1 = style in ('Heading1', 'heading1') or (style == 'Normal' and is_main_section(text))
        if is_h1 and text:
            if buf:
                sections[current] = buf
            current = text
            buf = [(style, text, is_code)]
        else:
            buf.append((style, text, is_code))
    if buf:
        sections[current] = buf
    return sections


def inject_into_html(html_path, section_map):
    """
    section_map: {section_id: html_content}
    Replace each <div id="X" class="section-placeholder"> block.
    """
    with open(html_path, encoding='utf-8') as f:
        src = f.read()

    for sec_id, content in section_map.items():
        # Match the entire div block for this id
        pattern = (
            rf'(<div id="{re.escape(sec_id)}"[^>]*>)\s*<h2>[^<]*</h2>\s*'
            rf'Content coming soon…\s*(</div>)'
        )
        replacement = rf'\1\n{content}\n\2'
        src, n = re.subn(pattern, replacement, src, flags=re.DOTALL)
        if n == 0:
            print(f'  WARNING: section id "{sec_id}" not found in {os.path.basename(html_path)}')

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(src)


# ── Document → section-id mapping ──────────────────────────────────────────

DOCS_DIR = r'D:\InterviePreparation\docs'
TOPICS_DIR = r'D:\InterviePreparation\topics'

# (docx, html_file, [(section_title_regex, section_id), ...])
DOC_CONFIG = [
    (
        'Senior Architect Interview Prep.docx',
        'senior-architect.html',
        [
            (r'CORE ARCHITECTURE', 'core-architecture'),
            (r'MICROSERVICES', 'microservices'),
            (r'ROUTING.*LOAD', 'openshift-routing'),
            (r'HORIZONTAL.*VERTICAL|SCALING', 'openshift-scaling'),
            (r'MCP|MODEL CONTEXT', 'mcp'),
            (r'AGENTIC', 'agentic-ai'),
            (r'INTERVIEW.*Q|CHEAT', 'interview-qa'),
        ]
    ),
    (
        'react-interview-guide.docx',
        'react.html',
        [
            (r'React Fundamentals', 'fundamentals'),
            (r'React Hooks', 'hooks'),
            (r'How React Works', 'internals'),
            (r'Node\.js|Vite|Ecosystem', 'ecosystem'),
            (r'Advanced Patterns', 'patterns'),
            (r'TypeScript', 'typescript'),
            (r'Interview Questions|Rapid Fire', 'interview-questions'),
        ]
    ),
    (
        'Python Advanced Features Reference.docx',
        'python.html',
        [
            (r'Advanced Language', 'advanced-language'),
            (r'Decorators', 'decorators'),
            (r'Generators', 'generators'),
            (r'Async.*Await|Await', 'async-await'),
            (r'Context Managers', 'context-managers'),
            (r'Async Context', 'async-context-managers'),
            (r'Async Generators', 'async-generators'),
            (r'map.*reduce|Functional', 'functional'),
            (r'Dunder|Magic.*Method', 'dunder'),
            (r'Duck Typing|Protocol', 'duck-typing'),
            (r'Data Science|Ecosystem', 'data-science'),
            (r'Metaprogramming', 'metaprogramming'),
            (r'Testing', 'testing'),
            (r'Async I.O|Async IO', 'async-io'),
            (r'GIL|Global Interpreter', 'gil'),
            (r'Quick Reference|Cheatsheet', 'cheatsheet'),
        ]
    ),
    (
        'load-balancing-openshift-guide.docx',
        'load-balancing.html',
        [
            (r'Fundamentals', 'fundamentals'),
            (r'Inside OpenShift', 'inside-openshift'),
            (r'Outside OpenShift|External\b', 'external'),
            (r'External Load Balancers.*WITH|How to Use', 'external-with-openshift'),
            (r'Canary', 'canary'),
            (r'Production.*COB|COB.*DR|Prod.*COB', 'prod-cob'),
            (r'End.to.End|Complete.*Stack', 'end-to-end'),
            (r'Decision Guide', 'decision-guide'),
        ]
    ),
    (
        'autoscaling-openshift-guide.docx',
        'autoscaling.html',
        [
            (r'Scaling Fundamentals', 'fundamentals'),
            (r'Horizontal Pod Autoscaler|HPA', 'hpa'),
            (r'Vertical Pod Autoscaler|VPA', 'vpa'),
            (r'Custom Metrics|KEDA', 'keda'),
            (r'Cluster.*Machine|Node Scaling', 'cluster-autoscaler'),
            (r'Knative|Serverless', 'knative'),
            (r'FastAPI', 'fastapi-metrics'),
            (r'NOT Possible|Limitations', 'limitations'),
            (r'Decision Guide', 'decision-guide'),
            (r'End.to.End', 'end-to-end'),
        ]
    ),
]


for docx_name, html_name, mappings in DOC_CONFIG:
    docx_path = os.path.join(DOCS_DIR, docx_name)
    html_path = os.path.join(TOPICS_DIR, html_name)
    print(f'\n--- {docx_name} → {html_name}')

    paras = parse_doc(docx_path)
    sections = split_sections(paras)

    # For each mapping, find the matching section and build HTML
    section_html = {}
    for pattern, sec_id in mappings:
        matched_key = None
        for key in sections:
            if re.search(pattern, key, re.IGNORECASE):
                matched_key = key
                break
        if matched_key:
            content = paragraphs_to_html(sections[matched_key])
            section_html[sec_id] = content
            print(f'  ✓ [{sec_id}] ← "{matched_key[:60]}"')
        else:
            print(f'  ✗ [{sec_id}] — no match for /{pattern}/')

    inject_into_html(html_path, section_html)
    print(f'  Written: {html_path}')

print('\nDone.')
