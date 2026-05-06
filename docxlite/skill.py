'''Read, render, and edit .docx files with tracked changes.

Opens Word documents via `DocxDocument(path)`, renders them as markdown and supports editing with automatic revision tracking (insertions and deletions attributed to "AI Editor").

## Document structure

A .docx body contains two types of content elements: paragraphs (`w:p`) and tables (`w:tbl`).
They are siblings at the top level. The full hierarchy:

    Document → Paragraph / Table
    Table → Row → Cell → Paragraph → Run

Table cells are mini document bodies — they contain paragraphs which contain runs,
the same structure all the way down. A Run is the leaf node where text and formatting live.

## Formatting and styles

Formatting inherits down a chain: Run → Paragraph Style → Document defaults.
A property returns `None` when inherited (not explicitly set at that level).

Run-level properties: `run.font.bold`, `.italic`, `.underline`, `.name`, `.size`, `.strike`.

Font sizes are in EMUs (English Metric Units): 12700 EMU = 1pt. Use `Pt(12)` from `docx.shared`
instead of doing the math.

`p.style.name` gives the paragraph style name (`'Normal'`, `'Heading 1'`, etc.).

When inserting new content, match the formatting of nearby paragraphs. `insert_before` and
`insert_after` do this automatically by copying style and numbering from the reference paragraph.

## Lists and numbering

Lists in docx are paragraphs with numbering properties (`w:numPr`). There is no list container.
Each bullet or number is its own paragraph with `w:ilvl` (indent level, 0-indexed) and
`w:numId` (pointer to a numbering definition).

Style-based lists use `'List Bullet'` or `'List Number'` paragraph styles.
Legal documents often use raw numbering properties without list styles.

To add an item to an existing list, use `insert_before` or `insert_after` on a neighboring
list paragraph — they copy the numbering properties automatically.

## Reading

    doc = DocxDocument('contract.docx')
    doc  # renders full document as markdown with tracked changes visible

    p = doc.paragraphs[3]
    p  # renders single paragraph
    p.runs  # all runs including those inside tracked changes
    r = p.runs[0]
    r.revision  # Revision(typ='ins', author='Virgil', date='...') or None

## Editing

All edits are automatically wrapped in tracked changes (w:ins / w:del).

`add_md` parses markdown and creates runs with the right formatting.
It handles `**bold**`, `*italic*`, `***both***`, `<u>underline</u>`,
and nesting like `<u>**bold underline**</u>`.

`insert_before` and `insert_after` copy style and numbering from self
(or an optional `ref` paragraph) and use `add_md` for the text content.

    p.add_run('appended text')
    p.add_md('**bold** and *italic* and <u>underlined</u>')
    p.insert_after('New paragraph text')
    p.insert_before('Another paragraph', ref=other_p)
    p.runs[0].delete()  # marks as tracked deletion
    p.delete()          # marks all runs as tracked deletions

## Tables

    tbl = doc.tables[0]
    tbl  # renders as markdown table
    tbl.add_row('col1', 'col2', 'col3')
    tbl.delete()

`insert_table_before` and `insert_table_after` create a new table positioned relative
to an existing paragraph or table. The table width is computed automatically from the
document's page layout (page width minus margins).

    tbl = p.insert_table_after(rows=3, cols=4)   # insert after a paragraph
    tbl = p.insert_table_before(rows=2, cols=3)  # insert before a paragraph
    tbl = existing_tbl.insert_after(rows=5, cols=2)  # insert after a table
    tbl = existing_tbl.insert_before(rows=0, cols=6) # insert before a table that is empty

    # Populate after insertion
    for i, h in enumerate(['Name', 'Value', 'Notes']):
        tbl.cell(0, i).paragraphs[0].add_md(f'**{h}**')
    tbl.add_row('Alpha', '42', 'First entry')

## Saving

    doc.save('output.docx')
'''

from docxlite.core import *
from docxlite.core import DocxDocument, set_tracking, _copy_para_fmt
from docx.document import Document
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from docx.table import Table
from pyskills.core import allow, PosAllowPolicy

__all__ = ['DocxDocument', 'set_tracking', 'Document', 'Run', 'Paragraph', 'Table']

set_tracking('AI Editor')

allow(DocxDocument)
allow({Run: ...})
allow({Paragraph: ...})
allow({Table: ...})
allow({Document: ...})
