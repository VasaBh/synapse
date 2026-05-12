import zipfile, re, sys, os

sys.stdout.reconfigure(encoding='utf-8')

def get_headings(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    paras = re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL)
    results = []
    for p in paras:
        style_m = re.search(r'<w:pStyle w:val="([^"]+)"', p)
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', p)
        line = ''.join(texts).strip()
        if not line:
            continue
        style = style_m.group(1) if style_m else 'Normal'
        # Only include headings or numbered section headers
        is_heading = 'Heading' in style or 'heading' in style
        is_section = bool(re.match(r'^\d+\.\s', line)) or bool(re.match(r'^\d+\.\d+\s', line))
        if is_heading or is_section:
            results.append((style, line))
    return results

docs = [
    'Senior Architect Interview Prep.docx',
    'react-interview-guide.docx',
    'Python Advanced Features Reference.docx',
    'load-balancing-openshift-guide.docx',
    'autoscaling-openshift-guide.docx',
]

docs_dir = r'D:\InterviePreparation\docs'

for doc in docs:
    path = os.path.join(docs_dir, doc)
    headings = get_headings(path)
    print(f'\n{"="*70}')
    print(f'DOC: {doc}')
    print('='*70)
    for s, l in headings:
        indent = '  ' if 'Heading2' in s or 'heading2' in s or re.match(r'^\d+\.\d+', l) else ''
        indent2 = '    ' if 'Heading3' in s or 'heading3' in s else ''
        print(f'{indent2 or indent}[{s}] {l[:100]}')
