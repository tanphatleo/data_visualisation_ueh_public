"""
convert_word_to_md.py
---------------------
Converts a Word (.docx) file to Markdown, preserving all images and tables.

Output structure:
    <filename>/
    ├── <filename>.md
    └── images/
        ├── image_001.png
        └── ...

Usage:
    python convert_word_to_md.py report.docx
    python convert_word_to_md.py report.docx --output my_folder
"""

import sys
import re
import argparse
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.text.run import Run as DocxRun
    from docx.table import Table as DocxTable
except ImportError:
    sys.exit("python-docx not found.  Run: pip install python-docx")


# ── Document traversal ────────────────────────────────────────────────────────

def iter_block_items(doc):
    """Yield paragraphs and tables from the document body in document order."""
    body = doc.element.body
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            yield DocxParagraph(child, doc)
        elif tag == 'tbl':
            yield DocxTable(child, doc)


# ── Image extraction ──────────────────────────────────────────────────────────

def extract_run_image(run, images_dir: Path, counter: list, inline: bool = False):
    """
    If the run contains an embedded image, save it and return a markdown tag.
    `inline=True` suppresses the surrounding blank lines (for use inside tables).
    Returns None if the run has no image.
    """
    drawing = run._r.find(qn('w:drawing'))
    if drawing is None:
        drawing = run._r.find(qn('w:pict'))
    if drawing is None:
        return None

    blip = drawing.find('.//' + qn('a:blip'))
    if blip is None:
        return None

    r_embed = blip.get(qn('r:embed'))
    if r_embed is None:
        return None

    try:
        image_part = run.part.related_parts[r_embed]
    except KeyError:
        return None

    ext = image_part.content_type.split('/')[-1].lower()
    ext = {'jpeg': 'jpg', 'tiff': 'tif', 'svg+xml': 'svg'}.get(ext, ext)
    counter[0] += 1
    filename = f"image_{counter[0]:03d}.{ext}"
    (images_dir / filename).write_bytes(image_part.blob)

    tag = f"![Figure {counter[0]}](images/{filename})"
    return tag if inline else f"\n\n{tag}\n\n"


# ── Inline run → markdown ─────────────────────────────────────────────────────

def run_to_md(run, images_dir: Path, counter: list, in_table: bool = False):
    """Convert a single run to a markdown string."""
    img = extract_run_image(run, images_dir, counter, inline=in_table)
    if img is not None:
        return img

    text = run.text
    if not text:
        return ''

    # Always escape pipe so table cells aren't broken
    text = text.replace('|', '\\|')

    bold   = run.bold
    italic = run.italic

    if bold and italic:
        return f'***{text}***'
    if bold:
        return f'**{text}**'
    if italic:
        return f'*{text}*'
    return text


# ── Paragraph inline content → markdown ──────────────────────────────────────

def para_inline_to_md(para, images_dir: Path, counter: list, in_table: bool = False):
    """
    Walk all inline children of a paragraph (runs + hyperlinks) and
    return the combined markdown string.
    """
    parts = []
    for child in para._p:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'r':
            run = DocxRun(child, para)
            parts.append(run_to_md(run, images_dir, counter, in_table=in_table))

        elif tag == 'hyperlink':
            r_id = child.get(qn('r:id'))
            href = ''
            if r_id:
                try:
                    href = para.part.target_ref(r_id)
                except Exception:
                    href = ''
            link_text = ''.join(
                t.text or ''
                for r_el in child.findall('.//' + qn('w:r'))
                for t in r_el.findall(qn('w:t'))
            )
            parts.append(f'[{link_text}]({href})' if href else link_text)

    return ''.join(parts)


# ── Paragraph → block markdown ────────────────────────────────────────────────

_HEADING_PREFIXES = {
    'Heading 1': '#',  'heading 1': '#',
    'Heading 2': '##', 'heading 2': '##',
    'Heading 3': '###','heading 3': '###',
    'Heading 4': '####',
    'Heading 5': '#####',
    'Heading 6': '######',
    'Title':     '#',
    'Subtitle':  '##',
}


def para_to_md(para, images_dir: Path, counter: list):
    """Convert a paragraph to a markdown block string."""
    style = (para.style.name or 'Normal') if para.style else 'Normal'
    inline = para_inline_to_md(para, images_dir, counter, in_table=False)
    text   = inline.strip()

    if not text:
        return ''

    # Headings
    for prefix, hashes in _HEADING_PREFIXES.items():
        if style.startswith(prefix):
            return f'{hashes} {text}'

    # Lists — prefer numPr detection over style name
    num_pr = para._p.find('.//' + qn('w:numPr'))
    if num_pr is not None:
        ilvl_el = num_pr.find(qn('w:ilvl'))
        level   = int(ilvl_el.get(qn('w:val'), '0')) if ilvl_el is not None else 0
        indent  = '  ' * level
        bullet  = '1.' if ('Number' in style or 'Ordered' in style) else '-'
        return f'{indent}{bullet} {text}'

    if 'List Bullet' in style:
        return f'- {text}'
    if 'List Number' in style:
        return f'1. {text}'

    # Block quote
    if 'Quote' in style or 'Quotation' in style:
        return f'> {text}'

    # Preformatted / code
    if any(k in style for k in ('Code', 'Preformatted', 'verbatim', 'Monograph')):
        return f'```\n{text}\n```'

    # Caption (italic)
    if 'Caption' in style:
        return f'*{text}*'

    return inline  # preserve original spacing for Normal paragraphs


# ── Table → markdown ──────────────────────────────────────────────────────────

def _unique_row_cells(row):
    """Return cells in the row, skipping horizontally merged duplicates."""
    seen, result = set(), []
    for cell in row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid)
            result.append(cell)
    return result


def _cell_to_text(cell, images_dir: Path, counter: list):
    """
    Flatten all paragraphs inside a cell into a single-line markdown string.
    Multiple paragraphs are joined with ` / ` so the table row stays on one line.
    """
    parts = []
    for para in cell.paragraphs:
        chunk = para_inline_to_md(para, images_dir, counter, in_table=True).strip()
        # Collapse any embedded newlines (e.g. from image tags)
        chunk = re.sub(r'\s*\n\s*', ' ', chunk)
        if chunk:
            parts.append(chunk)
    return ' / '.join(parts) if parts else ''


def table_to_md(table, images_dir: Path, counter: list):
    """Convert a docx Table to a GitHub-Flavoured Markdown table string."""
    if not table.rows:
        return ''

    # Build a 2D list of cell texts, handling horizontal merges
    grid = []
    for row in table.rows:
        cells = _unique_row_cells(row)
        grid.append([_cell_to_text(c, images_dir, counter) for c in cells])

    # Normalise column count across all rows
    col_count = max(len(r) for r in grid)
    for row in grid:
        while len(row) < col_count:
            row.append('')

    # Compute per-column max width for padding
    widths = [0] * col_count
    for row in grid:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(cell))
    widths = [max(w, 3) for w in widths]  # at least 3 chars for '---'

    def fmt_row(cells):
        padded = [c.ljust(widths[i]) for i, c in enumerate(cells)]
        return '| ' + ' | '.join(padded) + ' |'

    lines = []
    for row_idx, row in enumerate(grid):
        lines.append(fmt_row(row))
        if row_idx == 0:                        # separator after header row
            sep = ['-' * widths[i] for i in range(col_count)]
            lines.append('| ' + ' | '.join(sep) + ' |')

    return '\n'.join(lines)


# ── Post-processing ───────────────────────────────────────────────────────────

def clean_markdown(text: str) -> str:
    """Remove trailing whitespace and collapse excess blank lines."""
    # Strip trailing spaces/tabs on every line
    text = re.sub(r'[ \t]+$', '', text, flags=re.MULTILINE)
    # Collapse 3+ consecutive blank lines to 2
    text = re.sub(r'\n{3,}', '\n\n', text)
    # Ensure blank line before headings when preceding content runs right up to them
    text = re.sub(r'([^\n])\n(#{1,6} )', r'\1\n\n\2', text)
    # Ensure blank line after headings when content immediately follows
    text = re.sub(r'(#{1,6} [^\n]+)\n([^\n#\-\|> ])', r'\1\n\n\2', text)
    # NOTE: table spacing is NOT handled here — '\n\n'.join(blocks) already
    # guarantees one blank line between every block (including before/after tables).
    # Adding table-specific regexes would incorrectly fire inside table rows.
    return text.strip() + '\n'


# ── Main conversion ───────────────────────────────────────────────────────────

def convert(docx_path: Path, output_dir: Path) -> None:
    images_dir = output_dir / 'images'
    output_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(exist_ok=True)

    doc     = Document(docx_path)
    counter = [0]
    blocks  = []

    for item in iter_block_items(doc):
        if isinstance(item, DocxParagraph):
            md = para_to_md(item, images_dir, counter)
            if md:
                blocks.append(md)
        elif isinstance(item, DocxTable):
            md = table_to_md(item, images_dir, counter)
            if md:
                blocks.append(md)

    text    = '\n\n'.join(blocks)
    text    = clean_markdown(text)
    md_path = output_dir / f'{docx_path.stem}.md'
    md_path.write_text(text, encoding='utf-8')

    saved = list(images_dir.iterdir())
    print(f'\nConversion complete')
    print(f'  Source   : {docx_path}')
    print(f'  Markdown : {md_path}  ({md_path.stat().st_size:,} bytes)')
    print(f'  Images   : {len(saved)} saved -> {images_dir}')
    print(f'  Blocks   : {len(blocks)} (paragraphs + tables)')


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Convert a Word .docx to Markdown with images and tables.'
    )
    parser.add_argument('docx', help='Path to the input .docx file')
    parser.add_argument(
        '--output', '-o', default=None,
        help='Output folder  (default: <docx stem>/ beside the source file)',
    )
    args = parser.parse_args()

    docx_path  = Path(args.docx).resolve()
    if not docx_path.exists():
        sys.exit(f'File not found: {docx_path}')
    if docx_path.suffix.lower() != '.docx':
        sys.exit(f'Expected a .docx file, got: {docx_path.suffix}')

    output_dir = Path(args.output).resolve() if args.output else docx_path.parent / docx_path.stem
    convert(docx_path, output_dir)


if __name__ == '__main__':
    main()
