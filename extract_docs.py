import zipfile, re, sys, os

sys.stdout.reconfigure(encoding='utf-8')

def get_text_with_styles(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    paras = re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL)
    results = []
    for p in paras:
        style_m = re.search(r'<w:pStyle w:val="([^"]+)"', p)
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', p)
        line = ''.join(texts).strip()
        if line:
            style = style_m.group(1) if style_m else 'Normal'
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
    lines = get_text_with_styles(path)
    print(f'\n{"="*80}')
    print(f'DOC: {doc}  (total lines: {len(lines)})')
    print('='*80)
    # Print first 150 lines
    for i, (s, l) in enumerate(lines[:150]):
        print(f'[{s}] {l[:110]}')
