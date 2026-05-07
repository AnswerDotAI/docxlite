'''Read, render, search, and edit .docx files with tracked changes.

This skill is for LLM-friendly Word document editing inside notebooks and solveit dialogs.
It builds on `python-docx`, adding markdown previews, simple markdown-to-runs insertion,
document-order block iteration, raw-text search, and tracked changes.

Importing this skill enables tracked changes by default:

    set_tracking('AI Editor')

Use `set_tracking(None)` to disable revision tracking.

## Opening and rendering

    doc = DocxDocument('contract.docx')
    doc  # renders the full document as markdown

    p = doc.paragraphs[0]
    p  # renders one paragraph

    tbl = doc.tables[0]
    tbl  # renders one table as a markdown table

Markdown rendering is a readable preview, not a full docx serializer. It shows common
formatting and tracked insertions/deletions, but complex Word features may not appear.

## Document structure

A .docx body contains two top-level block types:

    Document -> Paragraph / Table
    Table -> Row -> Cell -> Paragraph -> Run

Paragraphs and tables are siblings in the document body. `doc.paragraphs` and
`doc.tables` lose that interleaving, so use document iteration when order matters.

## Document-order blocks

Documents are iterable. Iteration yields top-level paragraphs and tables in the order
they appear in the body.

    for block in doc:
        print(block)

    doc.blocks  # materialized list of top-level Paragraph/Table objects

The yielded objects are real `python-docx` objects patched by docxlite, so they can be
edited directly:

    block = doc.blocks[3]
    block.insert_after('New text')

## Search

`doc.search()` searches document-order blocks and returns the matching Paragraph/Table
objects themselves. It searches raw text (`._text`), not markdown preview text, so bold,
italic, underline, and revision spans do not affect matching.

Exact search is case-insensitive by default:

    hits = doc.search('termination')
    first = hits[0]
    first.insert_after('New language here.')

Regex search is also supported:

    hits = doc.search(r'terminat\\w+', regex=True)

Use `case=True` for case-sensitive matching:

    hits = doc.search('Effective Date', case=True)

Current search includes text from normal runs, inserted runs, and deleted runs. That is
usually what you want when reviewing tracked-change documents.

## Runs and revisions

Runs are the leaf text objects. docxlite patches `Paragraph.runs` so it includes runs
inside tracked insertions and deletions.

    r = p.runs[0]
    r.text
    r.revision  # Revision(typ='ins'/'del', author='...', date='...') or None

Tracked changes are represented with WordprocessingML elements:

    w:ins      inserted content
    w:del      deleted content
    w:delText  deleted run text

## Formatting and markdown insertion

`add_md` parses a small markdown subset and creates formatted runs.

Supported inline formatting:

    **bold**
    *italic*
    ***bold italic***
    <u>underline</u>
    <u>**bold underline**</u>

Examples:

    p.add_md('This is **bold** and *italic*')
    p.add_md('<u>underlined</u> text')

Unsupported markdown is not intended to round-trip perfectly. Keep inserted text simple.

## Inserting paragraphs

`insert_before` and `insert_after` create new paragraphs near an existing paragraph or
table. They copy nearby style/formatting and use `add_md` for content.

    p.insert_after('New paragraph text')
    p.insert_before('Another paragraph')
    p.insert_after('Uses another paragraph as style reference', ref=other_p)

When tracking is enabled, inserted runs are wrapped in tracked insertions.

## Deleting

Deletion methods are tracking-aware.

With tracking disabled, `.delete()` physically removes content.

With tracking enabled:

    run.delete()   # marks a normal run as deleted
    p.delete()     # marks all runs and the paragraph mark as deleted
    row.delete()   # marks the row and cell contents as deleted
    tbl.delete()   # marks all rows and contents as deleted

Deleting content that was itself a tracked insertion removes the insertion instead of
creating a deletion of an insertion.

## Tables

Tables render as markdown tables for preview.

    tbl = doc.tables[0]
    tbl.add_row('Name', 'Value', 'Notes')
    tbl.delete()

To insert a table before or after a paragraph/table, pass `(rows, cols)` to
`insert_before` or `insert_after`:

    tbl = p.insert_after((0, 3))          # empty 3-column table after paragraph
    tbl = p.insert_before((1, 2))         # 1-row, 2-column table before paragraph
    tbl = existing_tbl.insert_after((0, 4))

Populate cells through normal python-docx access:

    tbl.add_row('Alpha', '42', 'First entry')
    tbl.cell(0, 0).paragraphs[0].add_md('**Name**')

Table cells contain paragraphs, and those paragraphs support the same docxlite methods.

## Formatting notes

Formatting inherits through Word's normal chain:

    Run -> Paragraph style -> Document defaults

Run-level properties include:

    run.font.bold
    run.font.italic
    run.font.underline
    run.font.name
    run.font.size
    run.font.strike

Use `Pt(12)` from `docx.shared` for font sizes.

Lists are paragraphs with numbering properties, not list containers. To add an item to
an existing list, insert before or after a nearby list paragraph so numbering properties
are copied.

## Saving

    doc.save('output.docx')
'''

from docx import Document as DocxDocument
from docx.document import Document
from docx.shared import Emu, Inches, Pt
from docx.table import Table, _Row, _Cell
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from docxlite.core import *
from pyskills.core import allow

__all__ = [
    'DocxDocument', 'set_tracking',
    'Document', 'Paragraph', 'Run', 'Table', '_Row', '_Cell',
    'Revision', 'Underline',
    'Pt', 'Inches', 'Emu',
]

set_tracking('AI Editor')

allow(DocxDocument)
allow(Pt)
allow(Inches)
allow(Emu)

allow({Run: ...})
allow({Paragraph: ...})
allow({Table: ...})
allow({_Row: ...})
allow({_Cell: ...})
allow({Document: ...})