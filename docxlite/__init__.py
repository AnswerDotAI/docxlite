__version__ = "0.0.2"
from docx import Document as DocxDocument
from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.shared import Emu, Inches, Pt
from docx.table import Table, _Row, _Cell
from docx.text.font import Font
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run
from docxlite.core import *
__all__ = [
    'DocxDocument', 'set_tracking',
    'Document', 'Paragraph', 'Run', 'Table', '_Row', '_Cell',
    'Revision', 'Underline',
    'Pt', 'Inches', 'Emu',
    'WD_ALIGN_PARAGRAPH', 'WD_CELL_VERTICAL_ALIGNMENT', 'WD_TABLE_ALIGNMENT',
]
