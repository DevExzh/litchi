"""
Litchi - High-performance Office file format parser

Type stubs for the litchi_py Python extension module.
"""

from pathlib import Path
from typing import Optional, List
from enum import Enum

class FileFormat(Enum):
    """File format enumeration
    
    Represents the different Office file formats supported by Litchi.
    """
    Doc: int  # Microsoft Word 97-2003 (.doc)
    Docx: int  # Microsoft Word 2007+ (.docx)
    Ppt: int  # Microsoft PowerPoint 97-2003 (.ppt)
    Pptx: int  # Microsoft PowerPoint 2007+ (.pptx)
    Xls: int  # Microsoft Excel 97-2003 (.xls)
    Xlsx: int  # Microsoft Excel 2007+ (.xlsx)
    Xlsb: int  # Microsoft Excel Binary 2007+ (.xlsb)
    Odt: int  # OpenDocument Text (.odt)
    Ods: int  # OpenDocument Spreadsheet (.ods)
    Odp: int  # OpenDocument Presentation (.odp)
    Pages: int  # Apple Pages (.pages)
    Keynote: int  # Apple Keynote (.key)
    Numbers: int  # Apple Numbers (.numbers)
    Rtf: int  # Rich Text Format (.rtf)

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
    
    Args:
        path: Path to the file
    
    Returns:
        The detected FileFormat, or None if format cannot be determined
    """
    ...

def detect_file_format_from_bytes(data: bytes) -> Optional[FileFormat]:
    """Detect file format from bytes
    
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
        
        A run is a contiguous section of text with the same formatting.
        
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
    
    Provides format-agnostic interface for both .doc and .docx files.
    The format is automatically detected when opening a file.
    
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
        
        The file format (.doc or .docx) is automatically detected.
        
        Args:
            path: Path to the document file
        
        Returns:
            Document instance
        
        Raises:
            IOError: If the file cannot be read
            ValueError: If the file format is invalid or unsupported
        """
        ...
    
    def text(self) -> str:
        """Extract all text from the document
        
        Returns:
            All text content as a single string
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
    
    Provides format-agnostic interface for both .ppt and .pptx files.
    The format is automatically detected when opening a file.
    
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
        
        The file format (.ppt or .pptx) is automatically detected.
        
        Args:
            path: Path to the presentation file
        
        Returns:
            Presentation instance
        
        Raises:
            IOError: If the file cannot be read
            ValueError: If the file format is invalid or unsupported
        """
        ...
    
    def text(self) -> str:
        """Extract all text from the presentation
        
        Returns:
            All text content from all slides as a single string
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
    
    Note: This is a placeholder for future worksheet-level API.
    Currently, use Workbook.worksheet_names() and Workbook.text() for data access.
    """
    pass

class Workbook:
    """Excel workbook interface
    
    Provides support for Excel workbooks in various formats 
    (.xls, .xlsx, .xlsb, .ods, .numbers).
    The format is automatically detected when opening a file.
    
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
        
        The file format (.xls, .xlsx, .xlsb, .ods, .numbers) is automatically detected.
        
        Args:
            path: Path to the workbook file
        
        Returns:
            Workbook instance
        
        Raises:
            IOError: If the file cannot be read
            ValueError: If the file format is invalid or unsupported
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
        """
        ...

__all__ = [
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

