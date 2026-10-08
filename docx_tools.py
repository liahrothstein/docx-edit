# docx_tools.py
import sys, os
from docx import Document

def outline(path):
    """Структура: заголовки и их индексы."""
    doc = Document(path)
    for i, p in enumerate(doc.paragraphs):
        n = p.style.name.lower()
        if "heading" in n or "заголовок" in n:
            print(f"[{i}] {p.style.name}: {p.text}")

def find(path, query):
    doc = Document(path)
    for i, p in enumerate(doc.paragraphs):
        if query.lower() in p.text.lower():
            print(f"[{i}] {p.text[:150]}")

def _replace_in_paragraph(p, old, new):
    # 1) фраза целиком в одном run — форматирование сохраняется
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            return True
    # 2) фраза разбита по runs — пересобираем текст в первом run
    if old in p.text and p.runs:
        full = p.text.replace(old, new)
        for r in p.runs:
            r.text = ""
        p.runs[0].text = full
        return True
    return False

def replace(path, old, new, out=None):
    doc = Document(path)
    targets = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                targets += cell.paragraphs
    for s in doc.sections:                       # колонтитулы
        targets += s.header.paragraphs + s.footer.paragraphs
    n = sum(1 for p in targets if _replace_in_paragraph(p, old, new))
    out = out or path
    doc.save(out)
    print(f"Заменено абзацев: {n} -> {out}")

if __name__ == "__main__":
    cmd, path = sys.argv[1], sys.argv[2]
    if cmd == "outline": outline(path)
    elif cmd == "find":  find(path, sys.argv[3])
    elif cmd == "replace":
        replace(path, sys.argv[3], sys.argv[4],
                sys.argv[5] if len(sys.argv) > 5 else None)