"""
Litchi - High-performance Office file format parser

Type stubs for the litchi_py Python extension module.

``open``, ``open_bytes`` and ``create`` return a ``File``: the unified API,
with OfficeCLI's paths, selectors, node views and semantic views, and its
write verbs (``set``, ``add``, ``remove``, ``move``, ``swap``, ``edit``,
``batch``, ``apply``, ``save``). A File is immutable: every write returns a
new File. ``Document``, ``Presentation`` and ``Workbook`` are the per-family
read API.

Errors are raised as ``LitchiError`` subclasses, which carry the error's
``code``, ``message`` and ``suggestion``. Every family class also derives from
the built-in exception that fits it (``OSError``, ``ValueError`` or
``LookupError``). Handles are immutable and can be used from any thread; opening
and the document-wide reads release the GIL while they run.

Names that start with an underscore (``_NodeView``, ``_Capabilities``,
``_Support``, ``_Metadata``, ``_ItemResult``, ``_PartInfo``, ``_Validation``,
``_ValidationIssue``, ``_FormatValue``, ``_Props``, ``_PropValue``) are typing helpers that exist only in this stub: they
describe the dicts the API takes and returns.
"""

from pathlib import Path
from types import TracebackType
from typing import (
    Any,
    Callable,
    ClassVar,
    Dict,
    List,
    Literal,
    Mapping,
    Optional,
    Sequence,
    Tuple,
    Type,
    Union,
    overload,
)

from typing_extensions import NotRequired, TypedDict

# Exceptions

class LitchiError(Exception):
    """Base class of the errors litchi_py raises for a file it cannot open, read,
    edit or save.

    ``str()`` gives the message, followed by the suggestion when there is one.
    Each family subclass also derives from the built-in exception that fits it,
    so ``except OSError`` and ``except ValueError`` keep working.
    """

    def __init__(
        self,
        message: str,
        code: Optional[str] = None,
        suggestion: Optional[str] = None,
        item: Optional[int] = None,
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

    @property
    def item(self) -> Optional[int]:
        """The zero-based index of the batch item or edit operation that
        failed, or None when the error is not about one."""
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
    """The document is encrypted, signed or edit-protected, or a save would
    change its protection (code "document_protected")."""

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

# Unified read API

#: A value of a node view's ``format`` dict: the JSON types the node view
#: writes. Lengths and colours are text, as OfficeCLI writes them ("2.54cm",
#: "11pt", "#1F4E79"). The metadata ``keywords`` are a list of str.
_FormatValue = Union[bool, int, float, str, List[str], None]

class _NodeView(TypedDict):
    """A node view: OfficeCLI's ``DocumentNode`` shape (stub-only type).

    The keys come in this order. ``text``, ``preview``, ``style`` and
    ``childCount`` are left out when the node has none, as the JSON form
    leaves them out, so a node view equals ``json.loads`` of its JSON.
    """

    path: str
    """The canonical path; ``File.get(path)`` returns the same node."""
    type: str
    """OfficeCLI's element type: "paragraph", "run", "table", "cell", "row",
    "sheet", "workbook", ...; Word parts and marks too: "style", "level",
    "header", "footer", "footnote", "endnote", "tab", "break", "bookmark",
    "field", "toc", "watermark", "textbox"."""
    text: NotRequired[str]
    """The element's text."""
    preview: NotRequired[str]
    """A short label: the first 50 characters of a paragraph, a cell's A1
    reference, a sheet's name."""
    style: NotRequired[str]
    """Word paragraphs: the style id."""
    childCount: NotRequired[int]
    """How many children the node has; left out when counting would read a
    part the call did not need."""
    format: Dict[str, _FormatValue]
    """The element's own (not inherited) properties, in OfficeCLI's names."""
    children: List[_NodeView]
    """The child views, ``depth`` levels down; empty at depth 0."""

class _Support(TypedDict):
    """Whether a verb or view mode is available (stub-only type)."""

    supported: bool
    code: NotRequired[str]
    """The error code a refusal raises, such as "unsupported_capability"."""
    followUp: NotRequired[str]
    """What makes it available: a milestone ("M2"), a leaf request, or an
    action."""
    reason: NotRequired[str]

class _Capabilities(TypedDict):
    """``File.capabilities()``: OfficeCLI's help schema for the file's format,
    with the file's state (stub-only type)."""

    format: Optional[str]
    """The format, such as "docx"; None for an encrypted OOXML package whose
    format is hidden."""
    state: Literal["open", "signed", "locked"]
    retainsSource: NotRequired[bool]
    """Whether the handle kept what an editor is built from."""
    verbs: Dict[str, _Support]
    """Each verb ("get", "find", "query", "view", "set", ...)."""
    views: Dict[str, _Support]
    """Each view mode ("text", "annotated", "outline", "stats", ...)."""
    elements: List[Dict[str, Any]]
    """Each element the format's reader knows, with its paths, children and
    properties, in OfficeCLI's help schema vocabulary."""
    gaps: List[Dict[str, Any]]
    """Each element, path or property Litchi cannot read yet, with its code
    and follow-up."""

class _Schema(TypedDict):
    """``schema(format)``: the format's property registry (stub-only type).

    Each element is in OfficeCLI's help-schema shape; each of its
    ``properties`` has ``type``, ``add``, ``set`` and ``get``, and, where
    they apply, ``aliases``, ``values``, ``valueAliases``, ``unit``,
    ``colorContext``, ``constraint``, ``clearable`` and ``writeGap``.
    """

    format: str
    """The format, such as "docx"."""
    pending: NotRequired[_Support]
    """Present when this build has no reader for the format: why, and the
    follow-up."""
    elements: List[Dict[str, Any]]
    """Each element kind with its properties, ``propertyGaps`` and ``gap``."""
    gaps: List[Dict[str, Any]]
    """The format's axis, operation, view and undeclared element gaps."""
    setKeywords: NotRequired[Dict[str, Any]]
    """Present when the format has a reader: the find/replace keywords of
    ``set`` (``find``, ``replace``, ``matchCase``, ``wholeWord``,
    ``regex``) and the ``support`` of the ``replace`` verb (record 0273)."""

class _Metadata(TypedDict, total=False):
    """``File.metadata()``: the document properties that are set
    (stub-only type). Dates are ISO 8601 text; ``*_local`` dates carry no
    time zone because the file stores none."""

    title: str
    subject: str
    author: str
    keywords: List[str]
    """One entry per keyword; absent when there are none."""
    description: str
    identifier: str
    language: str
    template: str
    last_modified_by: str
    revision: str
    created: str
    created_local: str
    modified: str
    modified_local: str
    page_count: int
    word_count: int
    character_count: int
    character_count_with_spaces: int
    editing_time_minutes: int
    application: str
    category: str
    company: str
    manager: str
    content_status: str
    content_type: str
    version: str
    last_printed_time: str
    last_printed_local: str
    last_backup_local: str
    hyperlink_base: str
    security: int
    codepage: int

#: A property value a write verb takes: None clears the property; bool,
#: int and float stay typed; text is coerced through the key's value type
#: ("12pt", "#1F4E79", "accent1+lumMod75"). The metadata ``keywords`` take
#: a list of str, or text separated by ";" or ",".
_PropValue = Union[bool, int, float, str, List[str], None]

#: Properties a write verb takes: a dict (dotted keys are fine) or
#: OfficeCLI's list of "key=value" strings.
_Props = Union[Mapping[str, _PropValue], Sequence[str]]

class _PartInfo(TypedDict):
    """One package part in ``File.parts()`` (stub-only type, record 0336)."""

    name: str
    """The part name, for example "/word/document.xml"."""
    contentType: str
    """The part's content type; empty when the package declares none."""

class _ValidationIssue(TypedDict):
    """One finding in ``File.validate()["issues"]`` (stub-only type)."""

    id: str
    check: str
    code: str
    """A stable finding code, for example "opc.relationship.target_missing"."""
    severity: Literal["info", "warning", "error", "fatal"]
    message: str
    locations: List[Dict[str, Union[str, int]]]
    """Where: ``part``, ``path``, ``byte_offset``, ``byte_length``,
    ``object_index`` (each only when known); never document content."""
    evidence: List[Dict[str, Any]]
    specification: NotRequired[Dict[str, str]]
    compatibility: str
    repair: Dict[str, str]

class _Validation(TypedDict):
    """``File.validate()``'s report (stub-only type, record 0336)."""

    format: str
    valid: bool
    """No finding of error or fatal severity."""
    complete: bool
    """Every check ran to completion."""
    checks: List[Dict[str, str]]
    """Each check's ``id`` and ``status`` ("complete", "not_applicable",
    "blocked" with a ``reason``, "stopped_by" with a ``check``)."""
    issues: List[_ValidationIssue]

class _ItemResult(TypedDict):
    """One batch item's result in ``Commit.results`` (stub-only type):
    OfficeCLI's ``BatchResult`` members, in input order."""

    index: int
    """The item's zero-based position in the batch."""
    success: bool
    changed: NotRequired[bool]
    """Successful items: whether the item changed the document."""
    path: NotRequired[str]
    """A successful add: the new element's canonical path."""
    matched: NotRequired[int]
    """A successful set with ``find``: the number of matches (record 0273)."""
    code: NotRequired[str]
    """A failed item (best-effort mode): the error code."""
    error: NotRequired[str]
    """A failed item: the error message."""
    suggestion: NotRequired[Optional[str]]
    """A failed item: how to recover, or None."""

class _ImportSummary(TypedDict):
    """``File.import_csv``'s ``Commit.value`` (stub-only type): content-free
    facts about the import."""

    sheet: int
    """The worksheet's zero-based tab position."""
    row: int
    """The zero-based origin row."""
    column: int
    """The zero-based origin column."""
    rows: int
    """Records imported, the header included."""
    columns: int
    """Fields in the longest record."""
    cells: int
    """Non-empty fields written."""
    texts: int
    numbers: int
    booleans: int
    dates: int
    """Fields written as date serials with a date format."""
    dates_as_text: int
    """Date fields kept as text (before the epoch, or XLS)."""
    formulas: int
    empty: int
    """Empty fields (each clears an existing cell's value)."""

class _MergeSummary(TypedDict):
    """``File.merge``'s ``Commit.value`` (stub-only type)."""

    replacements: int
    """Placeholders replaced."""
    fields: int
    """Word merge fields replaced."""
    skipped_fields: int
    """Word merge fields left in place (other content inside their markup)."""
    keys: List[str]
    """The record keys used, sorted."""
    unresolved: List[str]
    """The keys the template uses and the record lacks, sorted."""

class _RefreshField(TypedDict):
    """One entry of ``_RefreshSummary.fields`` (stub-only type)."""

    type: str
    """The field type in upper case: "TOC", "SEQ", "PAGE", "SLIDENUM", "FORMULA", ..."""
    path: Optional[str]
    """Where the field is ("/body/p[3]", "/slide[2]", "Sheet1!B4"), when it has an address."""
    status: str
    """"refreshed", "unchanged", "stale", "needs_layout", "needs_clock",
    "locked", "unresolved" or "unsupported"."""
    reason: Optional[str]
    """Why the field was not refreshed, when there is more to say."""

class _RefreshSummary(TypedDict):
    """``File.refresh``'s ``Commit.value`` (stub-only type)."""

    refreshed: int
    """Fields whose result was written."""
    bookmarks: int
    """Bookmarks added for table of contents entries."""
    missing_page_numbers: int
    """Table of contents entries left without a page number."""
    fields: List[_RefreshField]
    """Every field met, in document order."""

class _Evaluated(TypedDict):
    """``File.evaluate``'s result (stub-only type)."""

    value: Union[None, bool, float, str]
    """The value; an error value is its spelling ("#DIV/0!")."""
    origin: str
    """"literal", "cached" or "computed"."""

class _RecalculatedCell(TypedDict):
    """One formula of ``File.recalculate``'s ``cells`` (stub-only type)."""

    sheet: int
    row: int
    column: int
    formula: str
    """The formula as stored, without "="."""
    cached_value: Union[None, bool, float, str]
    computed_value: Union[None, bool, float, str]
    """None when the engine cannot evaluate the formula."""
    status: str
    """"match", "mismatch", "uncached", "unsupported" or "volatile"."""

class _Recalculation(TypedDict):
    """``File.recalculate``'s result (stub-only type)."""

    formulas: int
    match: int
    mismatch: int
    uncached: int
    unsupported: int
    volatile: int
    cells: List[_RecalculatedCell]

def open(path: Path | str) -> File:
    """Open the file at ``path`` through the unified API.

    The format is detected from the file's content, not from its name, in
    one pass. Any format File holds opens: .docx, .xlsx, .pptx, .xlsb, .doc,
    .xls, .ppt, .rtf, .odt, .ods and .odp. An encrypted file opens locked (its
    reads raise ProtectedError; ``capabilities()`` says so), and a signed one
    opens readable. The GIL is released while the file is opened.

    Args:
        path: Path to the file (str or os.PathLike)

    Returns:
        A File

    Raises:
        IoError: (an OSError) If the file cannot be read
        UnsupportedError: (a ValueError) If the file is not an Office file
            (an Apple iWork package directory opens as its single-file form,
            and save() writes it back as a directory)
        CorruptFileError: (a ValueError) If the file is malformed
    """
    ...

def open_bytes(data: bytes) -> File:
    """Open an in-memory file through the unified API.

    As ``open``, from the file's bytes, which are copied once. The GIL is
    released while the bytes are copied and opened.

    Args:
        data: The file's content, as bytes

    Returns:
        A File

    Raises:
        UnsupportedError: (a ValueError) If the bytes are not an Office file,
            or are in a format File does not hold
        CorruptFileError: (a ValueError) If the file is malformed
    """
    ...

def create(format: Union[FileFormat, str]) -> File:
    """Create a new, empty document of ``format`` through the unified API.

    The new File has no path: write it with ``save_as`` or ``to_bytes``.
    .docx, .xlsx, .pptx, .xlsb, .doc, .xls, .ppt, .rtf, the ODF formats,
    .pages and .key can be created.
    The GIL is released while the document is built.

    Args:
        format: A FileFormat, or its name ("docx", "Xlsx", ".pptx")

    Returns:
        A File

    Raises:
        PropertyError: (a ValueError) If format names no format
            ("invalid_argument")
        UnsupportedError: (a ValueError) If Litchi cannot create the format
    """
    ...

def schema(format: Union[FileFormat, str]) -> _Schema:
    """The property registry of ``format``: what ``set``, ``add`` and ``get``
    accept for each element, as data.

    Each element is in OfficeCLI's help-schema shape. Each property has its
    ``type``, the enumeration ``values`` and ``valueAliases``, the ``unit``
    a bare length number is read in, its ``aliases``, ``constraint`` and
    ``clearable`` flag, the ``add``/``set``/``get`` flags, and a ``writeGap``
    with its follow-up when it cannot be written yet. ``propertyGaps`` lists
    an element's keys Litchi does not answer yet; ``gaps`` the format's other
    gaps. A format without a reader in this build has no elements and a
    ``pending`` entry. The result equals the Rust ``litchi_ops::schema``
    JSON.

    Args:
        format: A FileFormat, or its name ("docx", "Xlsx", ".pptx")

    Returns:
        A dict

    Raises:
        PropertyError: (a ValueError) If format names no format
            ("invalid_argument")
    """
    ...

class File:
    """An open document, workbook or presentation: the unified API.

    Returned by ``litchi_py.open``, ``litchi_py.open_bytes`` and
    ``litchi_py.create``. A File is an immutable handle: reads never change
    it, every write verb returns a new File, and it can be shared between
    threads. Every call releases the GIL while Litchi works and holds it
    only to convert arguments and build the result.

    Reads address the document with OfficeCLI's paths (``/body/p[3]``,
    ``/Sheet1/B2``, ``Sheet1!B2``) and selectors
    (``paragraph[style=Heading1]``, ``row[Amount>5000]``), and return
    OfficeCLI's node view as dicts. ``capabilities()`` lists what the format
    answers; a format whose unified reader has not landed yet refuses every
    read with UnsupportedError (``unsupported_capability``).

    A File is a context manager: leaving the ``with`` block closes it. After
    ``close()``, every read raises IoError (``file_required``).

    Example:
        >>> import litchi_py
        >>> with litchi_py.open("report.docx") as doc:
        ...     paragraph = doc.get("/body/p[1]")
        ...     headings = doc.query("paragraph[style=Heading1]", limit=10)
        ...     doc = doc.set("/body/p[1]/r[1]", {"font.bold": True})
        ...     doc.save()
    """

    @property
    def format(self) -> Optional[FileFormat]:
        """The detected format, or None for an encrypted OOXML package whose
        format is hidden until it is unlocked. Still available after
        ``close()``."""
        ...

    @property
    def closed(self) -> bool:
        """Whether ``close()`` has been called."""
        ...

    def get(
        self, path: str, depth: int = 1, *, max_nodes: int = 100000, dense: Optional[int] = None
    ) -> _NodeView:
        """The node view of the element at ``path``.

        The view is OfficeCLI's node shape, as a dict: ``path`` (canonical),
        ``type``, ``text``, ``preview``, ``style`` and ``childCount`` (each
        left out when the node has none), ``format`` (the node's own
        properties, with JSON types) and ``children`` (child views, ``depth``
        levels down). Only the parts on the path are read.

        Args:
            path: An OfficeCLI path, such as "/body/p[3]", "/Sheet1/B2" or
                "Sheet1!B2"
            depth: Child levels to list below the node: 0 lists none; values
                above 64 count as 64
            max_nodes: The most nodes the result may hold, children included
            dense: For a cell range, list empty cells too, up to this many

        Returns:
            The node view

        Raises:
            PathError: (a LookupError) If the path is malformed
                ("invalid_path"), nothing is there ("path_not_found"), or the
                format has no such element ("unsupported_path")
            UnsupportedError: (a ValueError) If Litchi cannot read the
                element or the format yet ("unsupported_capability")
            ProtectedError: (a ValueError) If the file is encrypted
            PropertyError: (a ValueError) If an argument is negative
            LitchiError: If the view exceeds max_nodes ("output_too_large")
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    def find(
        self, path: str, depth: int = 1, *, max_nodes: int = 100000, dense: Optional[int] = None
    ) -> Optional[_NodeView]:
        """The node view of the element at ``path``, or None when nothing is
        there.

        As ``get``, except that a path that resolves to nothing returns None
        instead of raising PathError ("path_not_found"). A malformed path
        still raises.
        """
        ...

    def query(
        self,
        selector: str,
        limit: Optional[int] = None,
        depth: int = 0,
        *,
        max_nodes: int = 100000,
        dense: Optional[int] = None,
    ) -> List[_NodeView]:
        """The node views of the elements ``selector`` matches, in document
        order.

        The selector is compiled once and the document is walked in one
        pass, which stops after ``limit`` matches. Each match is a node view,
        as ``get`` returns, with ``depth`` levels of children (none by
        default). An unknown property key in the selector is not an error: it
        matches nothing and issues a UserWarning that suggests the nearest
        key.

        Args:
            selector: An OfficeCLI selector, such as "paragraph",
                "cell[type=Number]" or "Sheet1!row[Amount>5000]"
            limit: The most matches to return; None returns every match
            depth: Child levels to list below each match
            max_nodes: The most nodes the result may hold, matches and their
                children together
            dense: For cell ranges, list empty cells too, up to this many

        Returns:
            A list of node views

        Raises:
            PathError: (a LookupError) If the selector is malformed
                ("invalid_selector") or its scope path does not resolve
            UnsupportedError: (a ValueError) If Litchi cannot query the
                format yet ("unsupported_capability")
            ProtectedError: (a ValueError) If the file is encrypted
            PropertyError: (a ValueError) If an argument is negative
            LitchiError: If more than max_nodes nodes match
                ("output_too_large"); pass a limit
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    @overload
    def view(self, mode: str, *, json: Literal[False] = False) -> str: ...
    @overload
    def view(self, mode: str, *, json: Literal[True]) -> Dict[str, Any]: ...
    @overload
    def view(self, mode: str, *, json: bool = False) -> Union[str, Dict[str, Any]]:
        """One of OfficeCLI's semantic views of the whole document.

        The modes are "text" (the text, one line per element, with its
        path), "annotated" (text with formatting annotations), "outline" (the
        structure, with counts) and "stats" (counts and distributions), or
        their one-letter aliases. The other OfficeCLI modes ("issues",
        "forms", "html", "svg", "screenshot", "pdf") raise UnsupportedError.

        Args:
            mode: The view mode
            json: Return OfficeCLI's JSON form as Python objects instead of
                text

        Returns:
            The view as text, or as a dict when json is True

        Raises:
            PropertyError: (a ValueError) If the mode is unknown
                ("invalid_argument")
            UnsupportedError: (a ValueError) If the format or the mode is not
                supported ("unsupported_capability")
            ProtectedError: (a ValueError) If the file is encrypted
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    def capabilities(self) -> _Capabilities:
        """What this file can do through the unified API: OfficeCLI's help
        schema, as a dict.

        The report names the file's ``state`` ("open", "signed" or "locked"),
        whether each verb and view mode is supported (or the code its refusal
        raises, with the follow-up that adds it), every element the format's
        reader knows with its paths and properties, and every listed gap.

        Raises:
            IoError: (an OSError) If the file is closed
        """
        ...

    def text(self) -> str:
        """All text of the document, in reading order.

        Raises:
            ProtectedError: (a ValueError) If the file is encrypted
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    def metadata(self) -> Optional[_Metadata]:
        """The document properties (title, author, dates, counts, ...), as a
        dict with the keys that are set; dates are ISO 8601 text.

        Returns:
            The properties, or None when the format reports none (.pptx and
            .ppt without core properties)

        Raises:
            ProtectedError: (a ValueError) If the file is encrypted
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    @property
    def signed(self) -> bool:
        """Whether the File is signed: its write verbs refuse until
        ``unsign()``. False after ``close()``."""
        ...

    @property
    def locked(self) -> bool:
        """Whether the File is encrypted and locked: it reads and writes
        nothing. False after ``close()``."""
        ...

    def set(self, target: str, props: _Props) -> File:
        """Set properties on the element(s) ``target`` names; return the
        changed File.

        The File this is called on is unchanged. Properties use OfficeCLI's
        names, and dotted keys are fine in a dict. Values keep their Python
        types: ``None`` clears a property where it can be cleared, ``bool``,
        ``int`` and ``float`` stay typed, and text is coerced through the
        key's value type ("12pt", "#1F4E79", "accent1+lumMod75"). The GIL is
        released while the change is made.

        Args:
            target: An OfficeCLI path ("/body/p[3]/r[1]", "/Sheet1/B2"), or a
                selector scoped to a path
            props: A dict ({"font.bold": True, "fill": "1F4E79"}) or a list
                of "key=value" strings

        Returns:
            The new File

        Raises:
            PathError: (a LookupError) If the target is malformed or resolves
                to nothing
            PropertyError: (a ValueError) If a property is unknown (with a
                did-you-mean suggestion) or a value is invalid
            UnsupportedError: (a ValueError) If the format cannot set this
                property or element yet ("unsupported_capability")
            ProtectedError: (a ValueError) If the file is encrypted or signed
                (call ``unsign()`` first)
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    def add(
        self,
        parent: str,
        type: str,
        props: Optional[_Props] = None,
        *,
        index: Optional[int] = None,
        after: Optional[str] = None,
        before: Optional[str] = None,
    ) -> Tuple[File, str]:
        """Add an element under ``parent``; return the new File and the new
        element's path.

        Without ``index``, ``after`` or ``before`` (at most one), the element
        is added after the parent's last child.

        Args:
            parent: The path of the new element's parent ("/body", "/",
                "/slide[1]")
            type: The element type: "paragraph", "run", "sheet", "slide",
                "shape", ...
            props: The new element's properties, as for ``set``
            index: Insert at this zero-based position among the children
            after: Insert right after the element at this path
            before: Insert right before the element at this path

        Returns:
            A tuple (new File, the new element's canonical path)

        Raises:
            As ``set``
        """
        ...

    def remove(self, target: str) -> File:
        """Remove the element(s) ``target`` names; return the new File.

        Raises:
            As ``set``
        """
        ...

    def move(
        self,
        path: str,
        to: Optional[str] = None,
        *,
        index: Optional[int] = None,
        after: Optional[str] = None,
        before: Optional[str] = None,
    ) -> File:
        """Move the element at ``path``; return the new File.

        Without ``to``, the element stays under its parent and ``index``,
        ``after`` or ``before`` (at most one) says where it goes; without any
        of them it goes last. With ``to``, it moves under that parent, last
        or at ``index``. Positions count after the element is taken out.

        Raises:
            As ``set``
        """
        ...

    def swap(self, path: str, path2: str) -> File:
        """Swap the elements at ``path`` and ``path2``; return the new File.

        Raises:
            As ``set``
        """
        ...

    @overload
    def edit(self, fn: None = None) -> Edit: ...
    @overload
    def edit(self, fn: Callable[[Edit], Any]) -> Commit:
        """Start a transaction, or run ``fn`` in one.

        Without ``fn``, return an Edit: use it as a context manager (it
        commits when the block ends normally) or call its ``commit()``.
        With ``fn``, call ``fn(edit)`` and commit: the operations apply
        together or not at all, and the returned Commit's ``value`` is what
        ``fn`` returned. If ``fn`` raises, nothing is committed and the
        exception propagates.

        Raises:
            The first failed operation's error (its ``item`` names the
            operation), or what ``fn`` raised
        """
        ...

    def batch(
        self,
        items: Union[Sequence[Mapping[str, Any]], str, bytes],
        *,
        mode: Literal["atomic", "best_effort"] = "atomic",
    ) -> Commit:
        """Run OfficeCLI batch items; return the Commit.

        Each item is a dict in OfficeCLI's ``BatchItem`` shape: ``command``
        (or ``op``) and the fields it takes (``path``, ``path2``, ``parent``,
        ``type``, ``to``, ``index``, ``after``, ``before``, ``props``, ...).
        Keys match ASCII case-insensitively; an unknown or repeated key, or a
        value of the wrong type, is ``corrupt_batch``. ``items`` can also be
        JSON text (``str`` or ``bytes``) holding the array, as OfficeCLI's
        ``batch`` reads it. The items are converted to Rust once and run with
        the GIL released.

        In "atomic" mode (the default) the first failure raises, with the
        item's index as ``item``, and nothing changes. In "best_effort" mode
        a failed item is skipped and reported in ``results``.

        Returns:
            A Commit with one ``results`` entry per item

        Raises:
            PropertyError: (a ValueError) If an item is malformed
                ("corrupt_batch") or mode is unknown
            LitchiError: The first failed item's error in atomic mode, and in
                either mode a protection, I/O or resource-limit failure
        """
        ...

    def apply(self, item: Mapping[str, Any]) -> Commit:
        """Run one OfficeCLI batch item; return the Commit.

        As ``batch`` with one item in atomic mode.
        """
        ...

    def import_csv(
        self,
        target: str,
        data: Union[bytes, str],
        *,
        delimiter: Optional[str] = None,
        quote: Optional[str] = '"',
        header: bool = False,
        encoding: str = "auto",
        infer: bool = True,
        formulas: bool = False,
        max_cells: Optional[int] = None,
        max_bytes: Optional[int] = None,
    ) -> Commit:
        """Import CSV or TSV text into a worksheet; return the Commit.

        ``target`` is a sheet ("/Sheet1", origin A1) or an origin cell
        ("/Sheet1/B2") of an XLSX, XLSB, XLS or ODS workbook. Every field is
        written in one atomic commit; empty input changes nothing. Numbers,
        ISO dates and TRUE/FALSE are inferred unless ``infer`` is False;
        fields starting with "=" become formulas only with ``formulas``.
        ``Commit.value`` is the summary dict (``_ImportSummary``). The GIL is
        released while the data is parsed and written.

        Raises:
            PathError: (a LookupError) If the target does not resolve
            PropertyError: (a ValueError) If the text is malformed, an option
                is invalid or a field is too long
            UnsupportedError: (a ValueError) If the format has no worksheets
            LitchiError: If a limit is exceeded ("limit_exceeded")
        """
        ...

    def merge(
        self,
        data: Union[Mapping[str, Any], str, bytes],
        *,
        missing: str = "keep",
        fields: bool = True,
        max_replacements: Optional[int] = None,
        max_bytes: Optional[int] = None,
    ) -> Commit:
        """Fill this template's ``{{key}}`` placeholders (and, for DOCX, its
        Word ``MERGEFIELD`` fields) from one data record; return the Commit.

        ``data`` is a mapping (values rendered as JSON renders them; nested
        members are also reachable as "a.b" and "a[0]" keys) or a JSON
        object as str or bytes. DOCX, PPTX, XLSX, XLSB, DOC, PPT, XLS, RTF,
        ODT, OTH, ODM, ODS and ODP templates. The template is
        unchanged, so one opened template fills any number of records.
        Placeholders split across runs keep each run's formatting.
        ``missing`` is "keep" (leave a placeholder whose key the record
        lacks) or "empty" (remove it). ``Commit.value`` is the summary dict
        (``_MergeSummary``). The GIL is released while merging.

        Raises:
            PropertyError: (a ValueError) If the data is not one JSON
                object, an option is invalid, or a value holds a character
                the document cannot store
            UnsupportedError: (a ValueError) If the format has no merge
                adapter (the ODF object formats)
            LitchiError: If a limit is exceeded ("limit_exceeded")
        """
        ...

    def refresh(
        self,
        *,
        clock: Union[None, str, Any] = None,
        write_cache: bool = False,
        page_numbers: str = "keep",
        mark_dirty: bool = True,
        max_changes: Optional[int] = None,
        max_bytes: Optional[int] = None,
    ) -> Commit:
        """Recompute the document's derived field results; return the
        Commit.

        DOCX: tables of contents (from the headings), SEQ, REF, DATE/TIME,
        document-property and count fields; DOC: the same fields but REF;
        ODT, OTH, ODM: date, time and document-property fields; PPTX and
        ODP: slide-number and date fields; workbooks (XLSX, XLSB, XLS):
        formula results (written back as cached values only with
        ``write_cache``). Page-number fields keep their
        cached result and a table of contents keeps each entry's cached page
        number (``page_numbers="require"`` refuses unknown ones). ``clock``
        is "YYYY-MM-DD[THH:MM[:SS]]", a datetime.datetime or a
        datetime.date; without it date fields keep their cached result (the
        system clock is never read). ``Commit.value`` is the summary dict
        (``_RefreshSummary``). The GIL is released while refreshing.

        Raises:
            PropertyError: (a ValueError) If the clock or an option is invalid
            UnsupportedError: (a ValueError) If the format has no refresh
            LitchiError: If page numbers are required but unknown
                ("refresh_failed") or a limit is exceeded ("limit_exceeded")
        """
        ...

    def evaluate(self, path: str, *, recalculate: bool = False) -> _Evaluated:
        """Evaluate one workbook cell: a value cell's value, a formula's
        cached result, or (without a cache, or with ``recalculate``) its
        computed result. The file is unchanged.

        Raises:
            UnsupportedError: (a ValueError) If the formula uses syntax or a
                function the engine does not evaluate, or calls NOW, TODAY or
                RAND without a cached value
            PropertyError: (a ValueError) If the path is not a cell
        """
        ...

    def recalculate(self) -> _Recalculation:
        """Recalculate every formula and compare each result with its cached
        value. The file is unchanged."""
        ...

    def save(self, path: Optional[Union[Path, str]] = None, *, overwrite: bool = True) -> int:
        """Save the File atomically, in its own format; return the bytes
        written.

        Without ``path``, the File is written back to the path it was opened
        from; with one, this is ``save_as``. The file is written to a
        temporary sibling and renamed over the destination, so a failed save
        leaves the destination as it was. The GIL is released while the file
        is written.

        Args:
            path: Where to save; None saves to the path the File was opened
                from
            overwrite: Whether an existing file at the destination may be
                replaced

        Returns:
            The number of bytes written

        Raises:
            IoError: (an OSError) If the File has no path ("file_required";
                it was created or opened from bytes), the destination exists
                and overwrite is False ("file_exists"), the source changed,
                or writing fails
            ProtectedError: (a ValueError) If the File is encrypted and
                locked, or saving would remove its encryption
            UnsupportedError: (a ValueError) If the format cannot be saved
            PropertyError: (a ValueError) If path names another format's
                extension
        """
        ...

    def save_as(self, path: Union[Path, str], *, overwrite: bool = True) -> int:
        """Save the File atomically to ``path``, in its own format; return
        the bytes written. The File keeps the path it was opened from.

        Raises:
            As ``save``
        """
        ...

    def to_bytes(self) -> bytes:
        """The File's bytes, in its own format.

        The document is written with the GIL released, then copied once into
        the returned ``bytes``.

        Raises:
            As ``save``, without the path errors
        """
        ...

    def dump(self, *, strict: bool = False) -> List[Dict[str, Any]]:
        """The document as an OfficeCLI v2 batch: item dicts that ``batch``
        replays onto ``create(format)`` to rebuild it (record 0336).

        The first item is ``{"command": "meta", "dumpVersion": 2}``. Content
        the write verbs cannot rebuild is skipped, or with ``strict`` raises
        UnsupportedError ("unsupported_element").

        Raises:
            UnsupportedError: (a ValueError) If the format cannot be dumped
            ProtectedError: (a ValueError) If the file is encrypted
            IoError: (an OSError) If the file is closed or changed on disk
        """
        ...

    def parts(self) -> List[_PartInfo]:
        """The package parts in name order (design §12, record 0336).

        Raises:
            UnsupportedError: (a ValueError) If the format has no package parts
            IoError: (an OSError) If the file is closed
        """
        ...

    def raw(self, part: str) -> bytes:
        """One package part's bytes, as the File would publish them
        (OfficeCLI ``raw``).

        ``part`` is a part name ("/word/document.xml") or alias
        ("/document", "/styles", "/slide[1]", "/Sheet1", "/content").

        Raises:
            PathError: (a LookupError) If there is no such part
            UnsupportedError: (a ValueError) As ``parts``
        """
        ...

    def add_part(
        self,
        parent: str,
        type: Optional[str] = None,
        *,
        content_type: Optional[str] = None,
        relationship: Optional[str] = None,
        data: Optional[bytes] = None,
        name: Optional[str] = None,
    ) -> Commit:
        """Add a package part with its content type and a relationship from
        ``parent``; the Commit's ``value`` is ``{"part", "relationshipId"}``
        (OfficeCLI ``add-part``).

        ``type`` "chart", "header" or "footer" adds an empty part of that
        kind; otherwise pass ``content_type``, ``relationship``, ``data``
        and ``name``.

        Raises:
            PathError: (a LookupError) If the parent does not resolve
            PropertyError: (a ValueError) If an argument is missing or the
                type is unknown
            UnsupportedError: (a ValueError) Outside DOCX, XLSX, PPTX and XLSB
            ProtectedError: (a ValueError) If the File is signed or encrypted
        """
        ...

    def raw_set(
        self, part: str, xpath: str, action: str, xml: Optional[str] = None
    ) -> Commit:
        """Edit one part's XML at every element ``xpath`` matches; the
        Commit's ``value`` is ``{"matched": n}`` (OfficeCLI ``raw-set``).

        ``action`` is append, prepend, insertbefore, insertafter, replace,
        remove or setattr (``xml`` is then "name=value").

        Raises:
            PathError: (a LookupError) If the part does not resolve, the
                XPath is malformed or matches nothing
            PropertyError: (a ValueError) If the action or fragment is invalid
            UnsupportedError: (a ValueError) Outside DOCX, XLSX, PPTX and XLSB
            ProtectedError: (a ValueError) If the File is signed or encrypted
        """
        ...

    def validate(self) -> _Validation:
        """Check the package structure read-only (OfficeCLI ``validate``).

        Raises:
            UnsupportedError: (a ValueError) For iWork packages
            ProtectedError: (a ValueError) If the file is encrypted
            IoError: (an OSError) If the file is closed
        """
        ...

    def unsign(self) -> File:
        """A File that may edit this signed document.

        Editing a signed document invalidates its signature, so the write
        verbs refuse it (ProtectedError). The returned File reads the same
        content; its first write removes the signature and publishes the
        document without it. This File stays signed. A File that is not
        signed is returned as it is.

        Raises:
            IoError: (an OSError) If the file is closed
        """
        ...

    def close(self) -> None:
        """Close the file: drop this handle's reference to the document.
        Reads raise IoError ("file_required") afterwards; closing again does
        nothing. A read already running on another thread finishes first."""
        ...

    def __enter__(self) -> File:
        """Return the file itself; raises IoError if it is closed."""
        ...

    def __exit__(
        self,
        exc_type: Optional[Type[BaseException]],
        exc_value: Optional[BaseException],
        traceback: Optional[TracebackType],
    ) -> Literal[False]:
        """Close the file; exceptions propagate."""
        ...

class Commit:
    """The result of ``File.edit``, ``File.batch``, ``File.apply`` and
    ``Edit.commit``: the new File and what the change did."""

    @property
    def file(self) -> File:
        """The new File; the File the change started from is unchanged."""
        ...

    @property
    def changed(self) -> bool:
        """Whether the document changed."""
        ...

    @property
    def operations(self) -> int:
        """How many operations succeeded."""
        ...

    @property
    def results(self) -> Optional[List[_ItemResult]]:
        """One dict per batch item, in input order; None for ``edit``."""
        ...

    @property
    def value(self) -> Any:
        """What ``edit``'s callable returned; None otherwise."""
        ...

    @property
    def patch(self) -> Optional[Patch]:
        """The change's reversible Patch; None for ``File.add_part`` and
        ``File.raw_set``, which are not transaction steps."""
        ...

class Patch:
    """The reversible record of one change: ``Commit.patch``.

    ``apply(file)`` replays the change on the File it started from and
    returns the changed File; ``inverse()`` is the Patch that undoes it. The
    original File and clones use the fast identity path; a same-format
    independently opened File is accepted when its published base passes the
    bounded content proof. Changed content, protected state and stale sources
    raise their typed errors. A Patch lives only in this process.

    Example:
        >>> commit = doc.edit(lambda tx: tx.set("/body/p[1]", {"bold": True}))
        >>> undone = commit.patch.inverse().apply(commit.file)
        >>> assert undone.to_bytes() == doc.to_bytes()
    """

    def is_noop(self) -> bool:
        """Whether the change kept the document as it was."""
        ...

    def inverse(self) -> Patch:
        """The Patch that undoes this one."""
        ...

    def apply(self, file: File) -> File:
        """Replay the change on ``file`` and return the changed File. The
        original input is unchanged; a same-format independently opened copy
        is accepted after the bounded content proof.

        Raises:
            LitchiError: A typed patch-conflict, protection or source failure
        """
        ...

class Edit:
    """A transaction: operations that apply together or not at all.

    ``File.edit()`` returns one. Use it as a context manager, which commits
    when the block ends normally and discards the operations when it raises,
    or call ``commit()`` yourself. ``set``, ``remove``, ``move`` and ``swap``
    are staged and run together; ``add`` and the reads first run what is
    staged. An error of a staged operation is raised by the call that runs
    it (``add``, a read or ``commit``), with the operation's zero-based index
    as ``item``. The first failure ends the transaction.

    Example:
        >>> with doc.edit() as tx:
        ...     path = tx.add("/body", "paragraph", {"text": "Hello"})
        ...     tx.set(path, {"align": "center"})
        >>> doc = tx.commit().file
    """

    def set(self, target: str, props: _Props) -> None:
        """Stage setting ``props`` on the element(s) ``target`` names."""
        ...

    def add(
        self,
        parent: str,
        type: str,
        props: Optional[_Props] = None,
        *,
        index: Optional[int] = None,
        after: Optional[str] = None,
        before: Optional[str] = None,
    ) -> str:
        """Run what is staged, then add an element; return its path."""
        ...

    def remove(self, target: str) -> None:
        """Stage removing the element(s) ``target`` names."""
        ...

    def move(
        self,
        path: str,
        to: Optional[str] = None,
        *,
        index: Optional[int] = None,
        after: Optional[str] = None,
        before: Optional[str] = None,
    ) -> None:
        """Stage moving the element at ``path``."""
        ...

    def swap(self, path: str, path2: str) -> None:
        """Stage swapping the elements at ``path`` and ``path2``."""
        ...

    def get(
        self, path: str, depth: int = 1, *, max_nodes: int = 100000, dense: Optional[int] = None
    ) -> _NodeView:
        """The node view at ``path`` in the transaction's current state."""
        ...

    def find(
        self, path: str, depth: int = 1, *, max_nodes: int = 100000, dense: Optional[int] = None
    ) -> Optional[_NodeView]:
        """As ``get``, or None when nothing is there."""
        ...

    def query(
        self,
        selector: str,
        limit: Optional[int] = None,
        depth: int = 0,
        *,
        max_nodes: int = 100000,
        dense: Optional[int] = None,
    ) -> List[_NodeView]:
        """The node views ``selector`` matches in the transaction's current
        state."""
        ...

    def commit(self) -> Commit:
        """Run what is staged and finish the transaction; calling it again
        returns the same Commit."""
        ...

    def __enter__(self) -> Edit:
        """Return the transaction; raises if it is no longer open."""
        ...

    def __exit__(
        self,
        exc_type: Optional[Type[BaseException]],
        exc_value: Optional[BaseException],
        traceback: Optional[TracebackType],
    ) -> Literal[False]:
        """Commit when the block ended normally; discard when it raised."""
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
    "open",
    "open_bytes",
    "create",
    "schema",
    "File",
    "Commit",
    "Patch",
    "Edit",
]

