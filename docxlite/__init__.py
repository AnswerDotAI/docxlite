__version__ = "0.0.2"
from docx import Document as DocxDocument
from docx.document import Document
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from docx.table import Table
from docxlite.core import *
__all__ = ['DocxDocument', 'Paragraph', 'Run', 'Table']