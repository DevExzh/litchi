"""
Litchi - High-performance Office file format parser

Type stubs for the litchi_py Python extension module.

Errors are raised as ``LitchiError`` subclasses, which carry the error's
``code``, ``message`` and ``suggestion``. Every family class also derives from
the built-in exception that fits it (``OSError``, ``ValueError`` or
``LookupError``). Handles are immutable and can be used from any thread; opening
and the document-wide reads release the GIL while they run.
"""

from pathlib import Path
from typing import ClassVar, Optional, List

# Exceptions

class LitchiError(Exception):
    """Base class of the errors litchi_py raises for a file it cannot open or read.

    ``str()`` gives the message, followed by the suggestion when there is one.
    Each family subclass also derives from the built-in exception that fits it,
    so ``except OSError`` and ``except ValueError`` keep working.
    """

    def __init__(
        self, message: str, code: Optional[str] = None, suggestion: Optional[str] = None
    ) -> None:
        """Create an error. ``code`` defaults to the class's own code."""
        ...

    @property
    def code(self) -> str:
        """The stable error code, for example "unsupported_format"."""
        ...

    @property
    def message(self) -> str:
        """The human-readable message."""
        ...

    @property
    def suggestion(self) -> Optional[str]:
        """How to recover, or None when there is no known way."""
        ...

class PathError(LitchiError, LookupError):
    """A path or selector is malformed or resolves to nothing (codes such as
    "path_not_found", "invalid_path" and "out_of_range")."""

class PropertyError(LitchiError, ValueError):
    """A property, argument or payload value is invalid, missing or not
    supported (codes such as "invalid_value" and "unsupported_property")."""

class UnsupportedError(LitchiError, ValueError):
    """The input's format, or the requested operation, is not supported (codes
    "unsupported_format", "unsupported_capability", "unsupported_element",
    "unsupported_type" and "unsupported_command")."""

class IoError(LitchiError, OSError):
    """Reading or writing the file failed, or the file changed after it was
    opened (codes "file_not_found", "file_locked", "file_exists",
    "file_required" and "io_error")."""

class CorruptFileError(LitchiError, ValueError):
    """The file is malformed (codes "corrupt_file" and "invalid_input")."""

class ProtectedError(LitchiError, ValueError):
    """The document is encrypted, signed or edit-protected (code
    "document_protected")."""

# Formats and values

class FileFormat:
    """File format enumeration

    One value per format detection can report. The values are the class
    attributes below; ``str()`` gives the name and ``int()`` a distinct number.
    """
    Doc: ClassVar[FileFormat]  # Microsoft Word 97-2003 (.doc)
    Docx: ClassVar[FileFormat]  # Microsoft Word 2007+ (.docx)
    Ppt: ClassVar[FileFormat]  # Microsoft PowerPoint 97-2003 (.ppt)
    Pptx: ClassVar[FileFormat]  # Microsoft PowerPoint 2007+ (.pptx)
    Xls: ClassVar[FileFormat]  # Microsoft Excel 97-2003 (.xls)
    Xlsx: ClassVar[FileFormat]  # Microsoft Excel 2007+ (.xlsx)
    Xlsb: ClassVar[FileFormat]  # Microsoft Excel Binary 2007+ (.xlsb)
    Odt: ClassVar[FileFormat]  # OpenDocument Text (.odt)
    Ods: ClassVar[FileFormat]  # OpenDocument Spreadsheet (.ods)
    Odp: ClassVar[FileFormat]  # OpenDocument Presentation (.odp)
    Odg: ClassVar[FileFormat]  # OpenDocument Drawing (.odg, .otg)
    Odc: ClassVar[FileFormat]  # OpenDocument Chart (.odc, .otc)
    Odf: ClassVar[FileFormat]  # OpenDocument Formula (.odf, .otf)
    Odi: ClassVar[FileFormat]  # OpenDocument Image (.odi, .oti)
    Odm: ClassVar[FileFormat]  # OpenDocument Master Document (.odm)
    Oth: ClassVar[FileFormat]  # OpenDocument Web Document (.oth)
    Odb: ClassVar[FileFormat]  # OpenDocument Database Front End (.odb)
    Pages: ClassVar[FileFormat]  # Apple Pages (.pages)
    Keynote: ClassVar[FileFormat]  # Apple Keynote (.key)
    Numbers: ClassVar[FileFormat]  # Apple Numbers (.numbers)
    Rtf: ClassVar[FileFormat]  # Rich Text Format (.rtf)

class RGBColor:
    """RGB color representation
    
    Represents a color in RGB format with values from 0-255.
    """
    
    def __init__(self, r: int, g: int, b: int) -> None:
        """Create a new RGB color
        
        Args:
            r: Red component (0-255)
            g: Green component (0-255)
            b: Blue component (0-255)
        """
        ...
    
    @property
    def r(self) -> int:
        """Red component (0-255)"""
        ...
    
    @property
    def g(self) -> int:
        """Green component (0-255)"""
        ...
    
    @property
    def b(self) -> int:
        """Blue component (0-255)"""
        ...

class Length:
    """Length with units
    
    Represents a measurement with associated units (EMUs, points, inches, etc.).
    """
    
    @staticmethod
    def from_emus(emus: int) -> Length:
        """Create a length from EMUs (English Metric Units)
        
        Args:
            emus: Length in EMUs (914400 EMUs = 1 inch)
        """
        ...
    
    @staticmethod
    def from_points(points: float) -> Length:
        """Create a length from points
        
        Args:
            points: Length in points (72 points = 1 inch)
        """
        ...
    
    @staticmethod
    def from_inches(inches: float) -> Length:
        """Create a length from inches
        
        Args:
            inches: Length in inches
        """
        ...
    
    def to_emus(self) -> int:
        """Convert to EMUs"""
        ...
    
    def to_points(self) -> float:
        """Convert to points"""
        ...
    
    def to_inches(self) -> float:
        """Convert to inches"""
        ...

def detect_file_format(path: Path | str) -> Optional[FileFormat]:
    """Detect file format from file path

    Detection reads the file's content; the file name is not consulted. The
    GIL is released while the file is read.

    Args:
        path: Path to the file

    Returns:
        The detected FileFormat, or None if format cannot be determined
    """
    ...

def detect_file_format_from_bytes(data: bytes) -> Optional[FileFormat]:
    """Detect file format from bytes

    The GIL is released while the bytes are inspected.

    Args:
        data: File content as bytes

    Returns:
        The detected FileFormat, or None if format cannot be determined
    """
    ...

# Document API

class Run:
    """A run of text with consistent formatting
    
    Represents a contiguous section of text that shares the same formatting properties.
    """
    
    def text(self) -> str:
        """Extract text from the run"""
        ...
    
    def bold(self) -> Optional[bool]:
        """Check if the run is bold
        
        Returns:
            True if bold, False if not bold, None if unspecified
        """
        ...
    
    def italic(self) -> Optional[bool]:
        """Check if the run is italic
        
        Returns:
            True if italic, False if not italic, None if unspecified
        """
        ...
    
    def underline(self) -> Optional[bool]:
        """Check if the run is underlined

        .docx reports the run's own underline (False for an explicit "none");
        .doc reports True for any underline and None otherwise; .rtf always
        answers True or False. .odt and .pages runs report None.

        Returns:
            True if underlined, False if not underlined, None if unspecified
        """
        ...

class Paragraph:
    """A paragraph in a document
    
    Represents a single paragraph with text and formatting.
    """
    
    def text(self) -> str:
        """Extract text from the paragraph"""
        ...
    
    def runs(self) -> List[Run]:
        """Get all runs in the paragraph

        A run is a contiguous section of text with the same formatting. A
        Pages paragraph is one run without formatting.

        Returns:
            List of Run objects
        """
        ...

class TableCell:
    """A cell in a table
    
    Represents a single cell containing text and possibly other content.
    """
    
    def text(self) -> str:
        """Extract text from the cell"""
        ...

class TableRow:
    """A row in a table
    
    Represents a single row containing cells.
    """
    
    def cells(self) -> List[TableCell]:
        """Get all cells in the row"""
        ...

class Table:
    """A table in a document
    
    Represents a table with rows and cells.
    """
    
    def row_count(self) -> int:
        """Get the number of rows in the table"""
        ...
    
    def rows(self) -> List[TableRow]:
        """Get all rows in the table"""
        ...

class Document:
    """Unified Word document interface

    Provides a format-agnostic interface for word-processing documents:
    .doc, .docx, .rtf, .odt and .pages. The format is detected from the
    file's content when it is opened.

    A Document is immutable and can be shared between threads. Opening and
    the document-wide reads release the GIL while they run.

    Example:
        >>> from litchi_py import Document
        >>> doc = Document.open("document.docx")
        >>> text = doc.text()
        >>> for para in doc.paragraphs():
        ...     print(para.text())
    """
    
    @staticmethod
    def open(path: Path | str) -> Document:
        """Open a Word document from a file path

        The format (.doc, .docx, .rtf, .odt or .pages) is detected from the
        file's content, not from its name.

        Args:
            path: Path to the document file

        Returns:
            Document instance

        Raises:
            IoError: (an OSError) If the file cannot be read
            UnsupportedError: (a ValueError) If the file is not a
                word-processing document, is not a recognized format, or
                uses a variant that cannot be read (flat OpenDocument, Word 6)
            CorruptFileError: (a ValueError) If the file is malformed
            ProtectedError: (a ValueError) If the document is encrypted
        """
        ...

    def text(self) -> str:
        """Extract all text from the document

        .docx, .rtf, .odt and .pages separate paragraphs with a line break;
        .doc keeps Word's paragraph marks (carriage returns).

        Returns:
            All text content as a single string

        Raises:
            IoError: (an OSError) If the file changed after it was opened
        """
        ...
    
    def paragraphs(self) -> List[Paragraph]:
        """Get all paragraphs in the document
        
        Returns:
            List of Paragraph objects
        """
        ...
    
    def tables(self) -> List[Table]:
        """Get all tables in the document

        Pages tables are not exposed; the list is empty for .pages files.

        Returns:
            List of Table objects
        """
        ...

# Presentation API

class Slide:
    """A slide in a presentation
    
    Represents a single slide with text and shapes.
    """
    
    def text(self) -> str:
        """Extract all text from the slide"""
        ...

class Presentation:
    """Unified PowerPoint presentation interface

    Provides a format-agnostic interface for presentations: .ppt, .pptx,
    .odp and .key. The format is detected from the file's content when it is
    opened.

    A Presentation is immutable and can be shared between threads. Opening
    and the presentation-wide reads release the GIL while they run.

    Example:
        >>> from litchi_py import Presentation
        >>> pres = Presentation.open("presentation.pptx")
        >>> print(f"Slides: {pres.slide_count()}")
        >>> for slide in pres.slides():
        ...     print(slide.text())
    """
    
    @staticmethod
    def open(path: Path | str) -> Presentation:
        """Open a PowerPoint presentation from a file path

        The format (.ppt, .pptx, .odp or .key) is detected from the file's
        content, not from its name.

        Args:
            path: Path to the presentation file

        Returns:
            Presentation instance

        Raises:
            IoError: (an OSError) If the file cannot be read
            UnsupportedError: (a ValueError) If the file is not a
                presentation, is not a recognized format, or uses a variant
                that cannot be read (flat OpenDocument)
            CorruptFileError: (a ValueError) If the file is malformed
            ProtectedError: (a ValueError) If the presentation is encrypted
        """
        ...

    def text(self) -> str:
        """Extract all text from the presentation

        Returns:
            All text content from all slides as a single string

        Raises:
            IoError: (an OSError) If the file changed after it was opened
        """
        ...
    
    def slide_count(self) -> int:
        """Get the number of slides in the presentation"""
        ...
    
    def slides(self) -> List[Slide]:
        """Get all slides in the presentation
        
        Returns:
            List of Slide objects
        """
        ...

# Sheet API

class Worksheet:
    """A worksheet in a workbook

    Note: This is a placeholder for future worksheet-level API. No method
    returns one yet; use Workbook.worksheet_names() and Workbook.text() for
    data access.
    """
    pass

class Workbook:
    """Excel workbook interface

    Provides support for workbooks in various formats
    (.xls, .xlsx, .xlsb, .ods, .numbers).
    The format is detected from the file's content when it is opened.

    A Workbook is immutable and can be shared between threads. Opening and
    the reads release the GIL while they run.

    Example:
        >>> from litchi_py import Workbook
        >>> wb = Workbook.open("workbook.xlsx")
        >>> print(f"Worksheets: {wb.worksheet_count()}")
        >>> for name in wb.worksheet_names():
        ...     print(f"Sheet: {name}")
        >>> text = wb.text()
        >>> print(text)
    """
    
    @staticmethod
    def open(path: Path | str) -> Workbook:
        """Open an Excel workbook from a file path

        The format (.xls, .xlsx, .xlsb, .ods, .numbers) is detected from the
        file's content, not from its name.

        Args:
            path: Path to the workbook file

        Returns:
            Workbook instance

        Raises:
            IoError: (an OSError) If the file cannot be read
            UnsupportedError: (a ValueError) If the file is not a workbook,
                is not a recognized format, or uses a variant that cannot be
                read (flat OpenDocument)
            CorruptFileError: (a ValueError) If the file is malformed
            ProtectedError: (a ValueError) If the workbook is encrypted
        """
        ...
    
    def worksheet_count(self) -> int:
        """Get the number of worksheets in the workbook"""
        ...
    
    def worksheet_names(self) -> List[str]:
        """Get all worksheet names
        
        Returns:
            List of worksheet names
        """
        ...
    
    def text(self) -> str:
        """Extract all text from all worksheets

        Returns:
            All text content as a single string

        Raises:
            IoError: (an OSError) If the file changed after it was opened
        """
        ...

__all__ = [
    "LitchiError",
    "PathError",
    "PropertyError",
    "UnsupportedError",
    "IoError",
    "CorruptFileError",
    "ProtectedError",
    "FileFormat",
    "RGBColor",
    "Length",
    "detect_file_format",
    "detect_file_format_from_bytes",
    "Document",
    "Paragraph",
    "Run",
    "Table",
    "TableRow",
    "TableCell",
    "Presentation",
    "Slide",
    "Workbook",
    "Worksheet",
]

