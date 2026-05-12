"""
build_pages.py  — re-extracts each docx and generates individual HTML
section pages in topics/{topic}/ folders. Also regenerates the topic index
pages with links to the individual section files.
"""
import zipfile, re, os, sys
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ── docx parsing ─────────────────────────────────────────────────────────────

def _parse_para(para):
    """Extract (style, text, is_code) from a single <w:p> element."""
    pPr = para.find(f'{{{W}}}pPr')
    style, is_code = 'Normal', False
    if pPr is not None:
        ps = pPr.find(f'{{{W}}}pStyle')
        if ps is not None:
            style = ps.get(f'{{{W}}}val', 'Normal')
        shd = pPr.find(f'{{{W}}}shd')
        if shd is not None and shd.get(f'{{{W}}}fill','').upper() in (
                'F1F5F9','E2E8F0','EFF6FF','F8FAFC','F0F9FF','FFFBF5'):
            is_code = True
    for run in para.findall(f'{{{W}}}r'):
        rPr = run.find(f'{{{W}}}rPr')
        if rPr is not None:
            rf = rPr.find(f'{{{W}}}rFonts')
            if rf is not None:
                a = rf.get(f'{{{W}}}ascii','') or rf.get(f'{{{W}}}hAnsi','')
                if 'Consolas' in a or 'Courier' in a:
                    is_code = True
    text = ''.join(t.text or '' for t in para.findall(f'.//{{{W}}}t')).strip()
    return style, text, is_code

def parse_doc(path):
    """Return mixed list:
      ('p', style, text, is_code)  — normal paragraph
      ('table', [[cell, ...], ...]) — Word table (rows × cols of text)
    Processes body children directly so tables are not flattened.
    """
    with zipfile.ZipFile(path) as z:
        raw = z.read('word/document.xml')
    root = ET.fromstring(raw)
    body = root.find(f'{{{W}}}body')
    out = []

    def process_node(node):
        tag = node.tag
        if tag == f'{{{W}}}p':
            style, text, is_code = _parse_para(node)
            out.append(('p', style, text, is_code))
        elif tag == f'{{{W}}}tbl':
            rows = []
            for tr in node.findall(f'{{{W}}}tr'):
                row = []
                for tc in tr.findall(f'{{{W}}}tc'):
                    parts = []
                    for p in tc.findall(f'.//{{{W}}}p'):
                        t = ''.join(x.text or '' for x in p.findall(f'.//{{{W}}}t')).strip()
                        if t:
                            parts.append(t)
                    row.append(' '.join(parts))
                if any(row):
                    rows.append(row)
            if rows:
                out.append(('table', rows))
        elif tag == f'{{{W}}}sdt':
            # Structured document tag — recurse into content
            content = node.find(f'{{{W}}}sdtContent')
            if content is not None:
                for child in content:
                    process_node(child)
        # w:body children can also include w:sectPr (section props) — ignore those

    for child in body:
        process_node(child)
    return out

# ── HTML helpers ──────────────────────────────────────────────────────────────

def esc(s):
    return (s.replace('&','&amp;').replace('<','&lt;')
             .replace('>','&gt;').replace('"','&quot;'))

def is_h1(style, text):
    return style in ('Heading1','heading1') or (
        style == 'Normal' and bool(re.match(r'^\d+\.\s+[A-Z–]', text))
        and len(text) < 120)

def is_subsec(text):
    return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+\S', text)) and len(text) < 120

def is_ascii_diag(text):
    BOX = set('┌┐└┘│─├┤┬┴┼╔╗╚╝║═╠╣╦╩╬+|')
    return bool(text) and sum(c in BOX for c in text) > len(text) * 0.15

def callout_class(text):
    if text.startswith('💡'):  return 'tip'
    if text.startswith('⚠'):  return 'warn'
    if text.startswith('✅'):  return 'info'
    if text.startswith('❌'):  return 'danger'
    return ''

def paras_to_html(paras):
    """Convert list of items (from parse_doc) to HTML string.
    Items are either ('p', style, text, is_code) or ('table', rows).
    """
    parts = []
    i = 0
    while i < len(paras):
        item = paras[i]

        # ── Word table (properly extracted) ──────────────────────────────────
        if item[0] == 'table':
            rows = item[1]
            tbl = ['<div class="table-wrap"><table>']
            for ri, row in enumerate(rows):
                tag = 'th' if ri == 0 else 'td'
                cells_html = ''.join(f'<{tag}>{esc(c)}</{tag}>' for c in row)
                tbl.append(f'  <tr>{cells_html}</tr>')
            tbl.append('</table></div>')
            parts.append('\n'.join(tbl))
            i += 1; continue

        # Unpack paragraph tuple
        _, style, text, is_code = item
        if not text:
            i += 1; continue

        # ── Code block (detected by doc style) ───────────────────────────────
        if is_code:
            buf = []
            while i < len(paras):
                row = paras[i]
                if row[0] == 'table': break
                _, _, t2, c2 = row
                if c2 or (buf and not t2):
                    if t2: buf.append(esc(t2))
                    i += 1
                else:
                    break
            if buf:
                parts.append(f'<pre><code>{chr(10).join(buf)}</code></pre>')
            continue

        # ── Styled headings ───────────────────────────────────────────────────
        lvl = {'Heading1':1,'heading1':1,'Heading2':2,'heading2':2,
               'Heading3':3,'heading3':3,'Heading4':4,'heading4':4}.get(style, 0)
        if lvl:
            tag = f'h{lvl + 1}'        # h1→h2, h2→h3 (h1 reserved for page title)
            parts.append(f'<{tag}>{esc(text)}</{tag}>')
            i += 1; continue

        # ── Numbered subsection headers in Normal style ───────────────────────
        if is_h1(style, text):
            parts.append(f'<h2>{esc(text)}</h2>')
            i += 1; continue
        if is_subsec(text):
            depth = text.count('.')
            tag = 'h3' if depth == 1 else 'h4'
            parts.append(f'<{tag}>{esc(text)}</{tag}>')
            i += 1; continue

        # ── Bullet list ───────────────────────────────────────────────────────
        if 'Bullet' in style or 'ListParagraph' in style or text.startswith('- '):
            items = []
            while i < len(paras):
                row = paras[i]
                if row[0] == 'table': break
                _, s2, t2, _ = row
                if 'Bullet' in s2 or 'ListParagraph' in s2:
                    items.append(t2); i += 1
                elif t2.startswith('- '):
                    items.append(t2[2:]); i += 1
                else:
                    break
            if items:
                lis = '\n'.join(f'  <li>{esc(x)}</li>' for x in items if x)
                parts.append(f'<ul>\n{lis}\n</ul>')
            continue

        # ── Callout ───────────────────────────────────────────────────────────
        cc = callout_class(text)
        if cc:
            parts.append(f'<div class="callout {cc}">{esc(text)}</div>')
            i += 1; continue

        # ── ASCII diagram block ───────────────────────────────────────────────
        if is_ascii_diag(text):
            buf = []
            while i < len(paras):
                row = paras[i]
                if row[0] == 'table': break
                t2 = row[2]
                if is_ascii_diag(t2) or (buf and t2 and t2[0] in '│|+┌'):
                    buf.append(esc(t2)); i += 1
                else:
                    break
            if buf:
                parts.append(f'<pre class="diagram">{chr(10).join(buf)}</pre>')
            continue

        # ── Regular paragraph ─────────────────────────────────────────────────
        parts.append(f'<p>{esc(text)}</p>')
        i += 1

    return '\n'.join(parts)

# ── Section splitting ──────────────────────────────────────────────────────��──

def split_sections(paras):
    """Return list of (title, [items]) in document order."""
    sections, current, buf = [], '_intro', []
    for item in paras:
        if item[0] == 'table':
            buf.append(item)
        else:
            _, style, text, _ = item
            if is_h1(style, text) and text:
                if buf: sections.append((current, buf))
                current, buf = text, [item]
            else:
                buf.append(item)
    if buf: sections.append((current, buf))
    return sections

# ── Slug helper ───────────────────────────────────────────────────────────────

def slugify(text):
    text = re.sub(r'[^\w\s-]', '', text.lower())
    text = re.sub(r'\s+', '-', text.strip())
    return text[:60].strip('-')

# ── HTML page template ────────────────────────────────────────────────────────

PAGE_TPL = """\
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{page_title}</title>
  <link rel="stylesheet" href="../../assets/style.css" />
</head>
<body>
  <header class="top-bar">
    <a href="../../index.html" class="brand">&#8592; Interview Prep</a>
    <nav class="breadcrumb">
      <a href="../{topic_file}">{topic_name}</a>
      <span class="sep">&#8250;</span>
      <span class="current">{title_short}</span>
    </nav>
  </header>

  <div class="page-layout">
    <aside class="section-sidebar">
      <div class="sidebar-header">{topic_name}</div>
      <ul>
{sidebar_items}
      </ul>
    </aside>

    <main class="content-area">
      <article class="content-body">
        <h1>{section_title}</h1>
{body}
      </article>

      <nav class="page-nav">
        {prev_link}
        <span class="spacer"></span>
        {next_link}
      </nav>
    </main>
  </div>

  <footer><p>{topic_name}</p></footer>
</body>
</html>
"""

INDEX_TPL = """\
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{topic_name} &#8211; Interview Prep</title>
  <link rel="stylesheet" href="../assets/style.css" />
</head>
<body>
  <header class="guide-header">
    <a href="../index.html" class="back-home">&#8592; All Topics</a>
    <h1>{topic_name}</h1>
  </header>

  <div class="topic-page">
    <aside class="sidebar">
      <p class="sidebar-title">{topic_name}</p>
      <nav><ul>
{sidebar_nav}
      </ul></nav>
    </aside>

    <main class="content">
      <h1>{topic_name}</h1>
      <p class="page-description">{description}</p>
{section_cards}
    </main>
  </div>

  <footer><p>{topic_name}</p></footer>
</body>
</html>
"""

# ── Config ────────────────────────────────────────────────────────────────────

DOCS_DIR   = r'D:\InterviePreparation\docs'
TOPICS_DIR = r'D:\InterviePreparation\topics'

TOPICS = [
    {
        'docx':        'Senior Architect Interview Prep.docx',
        'folder':      'senior-architect',
        'topic_file':  'senior-architect.html',
        'topic_name':  'Senior Architect',
        'description': 'Architecture patterns, microservices, OpenShift routing & scaling, MCP, and agentic AI — for senior-level technical interviews.',
    },
    {
        'docx':        'react-interview-guide.docx',
        'folder':      'react',
        'topic_file':  'react.html',
        'topic_name':  'React',
        'description': 'React fundamentals, hooks, internals, advanced patterns, and TypeScript — for front-end technical interviews.',
    },
    {
        'docx':        'Python Advanced Features Reference.docx',
        'folder':      'python',
        'topic_file':  'python.html',
        'topic_name':  'Python',
        'description': 'Advanced Python features for technical interviews — decorators, generators, async/await, metaprogramming, GIL, and more.',
    },
    {
        'docx':        'load-balancing-openshift-guide.docx',
        'folder':      'load-balancing',
        'topic_file':  'load-balancing.html',
        'topic_name':  'Load Balancing',
        'description': 'API & traffic load balancing in and around OpenShift — algorithms, routes, service mesh, canary, and DR.',
    },
    {
        'docx':        'autoscaling-openshift-guide.docx',
        'folder':      'autoscaling',
        'topic_file':  'autoscaling.html',
        'topic_name':  'Autoscaling',
        'description': 'HPA, VPA, KEDA, Knative, cluster autoscaler, and FastAPI metrics — complete OpenShift autoscaling guide.',
    },
]

# ── Build ─────────────────────────────────────────────────────────────────────

for topic in TOPICS:
    docx_path  = os.path.join(DOCS_DIR, topic['docx'])
    out_dir    = os.path.join(TOPICS_DIR, topic['folder'])
    os.makedirs(out_dir, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  {topic['topic_name']}  ←  {topic['docx']}")
    print(f"{'='*60}")

    paras    = parse_doc(docx_path)
    sections = split_sections(paras)

    # Build list of (filename, short_title) for navigation
    file_list = []
    for idx, (title, _) in enumerate(sections, start=1):
        slug  = slugify(title)
        fname = f'{idx:02d}-{slug}.html'
        # Short title: strip leading "N. " or "N.N " prefix for sidebar
        short = re.sub(r'^\d+[\.\s]+', '', title).strip()
        file_list.append((fname, title, short))

    # ── Generate each section HTML page ──────────────────────────────────────
    for idx, (title, paras_sec) in enumerate(sections):
        fname, full_title, short_title = file_list[idx]
        out_path = os.path.join(out_dir, fname)

        # Skip the first paragraph when it IS the section heading (already in <h1>)
        body_paras = paras_sec[1:] if (title != '_intro' and paras_sec) else paras_sec
        body_html = paras_to_html(body_paras)

        # Sidebar items
        sidebar_items = []
        for j, (fn, ft, st) in enumerate(file_list):
            active_li  = ' class="active"' if j == idx else ''
            active_a   = ' class="active"' if j == idx else ''
            label      = esc(st or ft)
            sidebar_items.append(
                f'        <li{active_li}><a href="{fn}"{active_a}>{label}</a></li>'
            )

        # Prev / Next
        prev_link = ''
        next_link = ''
        if idx > 0:
            pf, _, ps = file_list[idx - 1]
            prev_link = f'<a href="{pf}">&#8592; {esc(ps)}</a>'
        if idx < len(file_list) - 1:
            nf, _, ns = file_list[idx + 1]
            next_link = f'<a href="{nf}">{esc(ns)} &#8594;</a>'

        page_html = PAGE_TPL.format(
            page_title   = f'{esc(short_title)} – {esc(topic["topic_name"])}',
            topic_file   = topic['topic_file'],
            topic_name   = esc(topic['topic_name']),
            title_short  = esc(short_title),
            section_title= esc(title),
            sidebar_items= '\n'.join(sidebar_items),
            body         = body_html,
            prev_link    = prev_link,
            next_link    = next_link,
        )

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(page_html)
        print(f'  [{idx+1:02d}] {fname}')

    # ── Regenerate topic index page ───────────────────────────────────────────
    index_path = os.path.join(TOPICS_DIR, topic['topic_file'])

    sidebar_nav_items = []
    section_cards     = []
    for fn, ft, st in file_list:
        sidebar_nav_items.append(f'        <li><a href="{topic["folder"]}/{fn}">{esc(st)}</a></li>')
        section_cards.append(
            f'      <div class="section-placeholder">\n'
            f'        <h2><a href="{topic["folder"]}/{fn}" style="text-decoration:none;color:inherit">'
            f'{esc(ft)}</a></h2>\n'
            f'        <ul class="section-link-list"><li><a href="{topic["folder"]}/{fn}">'
            f'Open section &#8594;</a></li></ul>\n'
            f'      </div>'
        )

    index_html = INDEX_TPL.format(
        topic_name   = esc(topic['topic_name']),
        description  = esc(topic['description']),
        sidebar_nav  = '\n'.join(sidebar_nav_items),
        section_cards= '\n'.join(section_cards),
    )
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(index_html)
    print(f'  Index → {index_path}')

print('\n✓ Done.')
