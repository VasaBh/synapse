import zipfile, re, sys, os, html

sys.stdout.reconfigure(encoding='utf-8')

def get_paragraphs(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    paras = re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL)
    results = []
    for p in paras:
        style_m = re.search(r'<w:pStyle w:val="([^"]+)"', p)
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', p)
        line = ''.join(texts).strip()
        style = style_m.group(1) if style_m else 'Normal'
        results.append((style, html.unescape(line)))
    return results

def style_to_level(style):
    if style == 'Heading1' or style == 'heading1': return 1
    if style == 'Heading2' or style == 'heading2': return 2
    if style == 'Heading3' or style == 'heading3': return 3
    return 0

def is_section_header(style, text):
    """Detect numbered section headers in Normal style (e.g. '1. TITLE', '1.1 Title')"""
    if style != 'Normal':
        return False
    if re.match(r'^\d+\.\s+[A-Z]', text) and len(text) < 100:
        return True
    return False

def get_section_level_from_text(text):
    if re.match(r'^\d+\.\s+', text) and not re.match(r'^\d+\.\d+', text):
        return 1
    if re.match(r'^\d+\.\d+\s+', text) and not re.match(r'^\d+\.\d+\.\d+', text):
        return 2
    if re.match(r'^\d+\.\d+\.\d+\s+', text):
        return 3
    return 0

def para_to_md(style, text):
    if not text:
        return ''
    level = style_to_level(style)
    if level == 1:
        return f'# {text}'
    if level == 2:
        return f'## {text}'
    if level == 3:
        return f'### {text}'
    if is_section_header(style, text):
        lvl = get_section_level_from_text(text)
        if lvl == 1: return f'# {text}'
        if lvl == 2: return f'## {text}'
        if lvl == 3: return f'### {text}'
    if 'Bullet' in style or 'ListParagraph' in style or style == 'ListBullet':
        return f'- {text}'
    # Code-like (indented blocks with no style)
    return text

def slugify(text):
    text = re.sub(r'[^\w\s-]', '', text.lower())
    text = re.sub(r'[\s]+', '-', text.strip())
    return text[:60]

def split_into_sections(paragraphs):
    """Split paragraphs into top-level sections (Heading1 or level-1 numbered)."""
    sections = []
    current_title = 'Introduction'
    current_paras = []

    for style, text in paragraphs:
        level = style_to_level(style)
        is_h1 = level == 1
        is_norm_h1 = is_section_header(style, text) and get_section_level_from_text(text) == 1

        if is_h1 or is_norm_h1:
            if current_paras:
                sections.append((current_title, current_paras))
            current_title = text
            current_paras = [(style, text)]
        else:
            current_paras.append((style, text))

    if current_paras:
        sections.append((current_title, current_paras))

    return sections

def write_section_file(out_dir, index, title, paragraphs):
    slug = slugify(title)
    filename = f'{index:02d}-{slug}.md'
    filepath = os.path.join(out_dir, filename)

    lines = []
    prev_was_empty = False
    for style, text in paragraphs:
        if not text:
            if not prev_was_empty:
                lines.append('')
            prev_was_empty = True
            continue
        prev_was_empty = False
        md = para_to_md(style, text)
        lines.append(md)

    content = '\n'.join(lines).strip()
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    return filename

# Document config: (docx filename, output folder name)
docs_config = [
    ('Senior Architect Interview Prep.docx', 'senior-architect'),
    ('react-interview-guide.docx', 'react'),
    ('Python Advanced Features Reference.docx', 'python'),
    ('load-balancing-openshift-guide.docx', 'load-balancing'),
    ('autoscaling-openshift-guide.docx', 'autoscaling'),
]

docs_dir = r'D:\InterviePreparation\docs'
topics_dir = r'D:\InterviePreparation\topics'

os.makedirs(topics_dir, exist_ok=True)

summary_lines = []

for docx_name, folder_name in docs_config:
    path = os.path.join(docs_dir, docx_name)
    out_dir = os.path.join(topics_dir, folder_name)
    os.makedirs(out_dir, exist_ok=True)

    print(f'\nProcessing: {docx_name} -> {folder_name}/')
    paragraphs = get_paragraphs(path)
    sections = split_into_sections(paragraphs)

    summary_lines.append(f'\n## {folder_name} ({len(sections)} sections from {docx_name})')

    for i, (title, paras) in enumerate(sections, start=1):
        fname = write_section_file(out_dir, i, title, paras)
        print(f'  [{i:02d}] {fname}  ({len(paras)} paragraphs)')
        summary_lines.append(f'  - [{fname}](topics/{folder_name}/{fname})')

# Write index
index_path = os.path.join(topics_dir, 'INDEX.md')
with open(index_path, 'w', encoding='utf-8') as f:
    f.write('# Interview Preparation — Topic Index\n')
    f.write('\n'.join(summary_lines))
print(f'\nIndex written: {index_path}')
