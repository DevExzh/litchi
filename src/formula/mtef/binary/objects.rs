//! MTEF object structures and types
//!
//! This module defines the data structures used to represent parsed MTEF equation objects.
//! Based on rtf2latex2e eqn_support.h structures.
//!
//! Each object type corresponds to a specific MTEF record tag and contains the data
//! needed to reconstruct the mathematical formula.

/// MTEF record types (as defined in rtf2latex2e)
///
/// These enum variants represent the different types of records that can appear
/// in an MTEF equation stream. Each record type has specific parsing and
/// conversion logic in the parser module.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MtefRecordType {
    /// Line object - contains a horizontal sequence of math elements
    Line = 1,
    /// Character object - represents a single character with typeface and embellishments
    Char = 2,
    /// Template object - represents structured formulas like fractions, roots, etc.
    Tmpl = 3,
    /// Pile object - represents a vertical stack of lines
    Pile = 4,
    /// Matrix object - represents a 2D array of cells
    Matrix = 5,
    /// Embellishment object - represents decorations like dots, hats, arrows
    Embell = 6,
    /// Ruler object - contains tab stops for alignment
    Ruler = 7,
    /// Size object - controls text size
    Size = 9,
    /// Full size - predefined size level
    Full = 10,
    /// Subscript size - predefined size level
    Sub = 11,
    /// Sub-subscript size - predefined size level
    Sub2 = 12,
    /// Symbol size - predefined size level
    Sym = 13,
    /// Sub-symbol size - predefined size level
    SubSym = 14,
    /// Color object - applies color to text
    Color = 15,
    /// Color definition - defines a custom color
    ColorDef = 16,
    /// Font definition - defines a custom font
    FontDef = 17,
    /// Font object - applies font to text
    Font = 18,
    /// Equation preferences - document-level settings
    EqnPrefs = 19,
    /// Encoding definition - character encoding information
    EncodingDef = 20,
    /// Future record type - for forward compatibility
    Future = 255,
}

/// Object list node
#[derive(Debug)]
pub struct MtefObjectList {
    pub tag: MtefRecordType,
    pub obj_ptr: Box<dyn MtefObject>,
    pub next: Option<Box<MtefObjectList>>,
}

/// Base trait for MTEF objects
pub trait MtefObject: std::fmt::Debug {
    fn as_any(&self) -> &dyn std::any::Any;
}

/// Character object (MT_CHAR)
///
/// Represents a single character in the equation with its typeface, encoding,
/// and optional embellishments (decorations like dots, arrows, etc.).
///
/// Note: Some fields like `nudge_x`, `nudge_y`, `atts`, and `bits16` are part of the
/// MTEF specification but not currently used in basic AST conversion. They are kept
/// for potential future enhancements (precise layout control, advanced character encoding).
#[derive(Debug)]
pub struct MtefChar {
    /// Horizontal positioning nudge (for fine-tuning layout)
    /// Part of MTEF spec, kept for future layout precision
    #[allow(dead_code)]
    pub nudge_x: i16,
    /// Vertical positioning nudge (for fine-tuning layout)
    /// Part of MTEF spec, kept for future layout precision
    #[allow(dead_code)]
    pub nudge_y: i16,
    /// Character attributes (flags indicating encoding type, embellishments, etc.)
    /// Part of MTEF spec, kept for advanced character handling
    #[allow(dead_code)]
    pub atts: u8,
    /// Typeface index (typically 128 + typeface_slot)
    pub typeface: u8,
    /// MathType character code
    pub character: u16,
    /// Additional 16-bit character encoding (for Unicode in MTEF v5)
    /// Part of MTEF spec, kept for extended Unicode support
    #[allow(dead_code)]
    pub bits16: u16,
    /// Optional embellishment list (decorations applied to this character)
    pub embellishment_list: Option<Box<MtefEmbell>>,
}

impl MtefObject for MtefChar {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Embellishment object (MT_EMBELL)
///
/// Represents decorations applied to characters (dots, hats, arrows, etc.).
/// Multiple embellishments can be chained via the `next` field.
///
/// Note: Nudge fields are part of MTEF spec for precise positioning,
/// kept for future layout enhancements.
#[derive(Debug)]
pub struct MtefEmbell {
    /// Horizontal positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_x: i16,
    /// Vertical positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_y: i16,
    /// Embellishment type (index into embellishment template table)
    pub embell: u8,
    /// Next embellishment in the chain (for stacked decorations)
    pub next: Option<Box<MtefEmbell>>,
}

impl MtefObject for MtefEmbell {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Template object (MT_TMPL)
///
/// Represents structured mathematical constructs like fractions, roots, integrals,
/// matrices with fences, subscripts/superscripts, etc. The specific template type
/// is determined by the selector and variation fields.
///
/// Note: Nudge and options fields are part of MTEF spec, kept for future enhancements.
#[derive(Debug)]
pub struct MtefTemplate {
    /// Horizontal positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_x: i16,
    /// Vertical positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_y: i16,
    /// Template selector (identifies the template type)
    pub selector: u8,
    /// Template variation (specific form within the template type)
    pub variation: u16,
    /// Template options (additional flags, part of MTEF spec)
    #[allow(dead_code)]
    pub options: u8,
    /// Subobjects that fill the template's slots (e.g., numerator/denominator for fractions)
    pub subobject_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefTemplate {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Line object (MT_LINE)
///
/// Represents a horizontal sequence of mathematical objects. Lines are the basic
/// building blocks of equations and can contain characters, templates, and other objects.
///
/// Note: Nudge, line_spacing, and ruler fields are part of MTEF spec,
/// kept for future advanced layout support.
#[derive(Debug)]
pub struct MtefLine {
    /// Horizontal positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_x: i16,
    /// Vertical positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_y: i16,
    /// Line spacing value (part of MTEF spec for vertical spacing control)
    #[allow(dead_code)]
    pub line_spacing: u8,
    /// Optional ruler defining tab stops (part of MTEF spec for alignment)
    #[allow(dead_code)]
    pub ruler: Option<Box<MtefRuler>>,
    /// Objects contained in this line
    pub object_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefLine {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Ruler object (MT_RULER)
///
/// Defines tab stops for aligning content within lines. Used primarily in
/// aligned equation environments. Part of MTEF spec, kept for future alignment support.
#[derive(Debug)]
pub struct MtefRuler {
    /// Number of tab stops defined (part of MTEF spec)
    #[allow(dead_code)]
    pub n_stops: i16,
    /// Linked list of tab stop definitions (part of MTEF spec)
    #[allow(dead_code)]
    pub tabstop_list: Option<Box<MtefTabstop>>,
}

impl MtefObject for MtefRuler {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Tabstop object (MT_TABSTOP)
///
/// Defines a single tab stop position and type for text alignment.
/// Part of MTEF spec, kept for future alignment support.
#[derive(Debug)]
pub struct MtefTabstop {
    /// Tab stop type (left, right, center, decimal alignment, part of MTEF spec)
    #[allow(dead_code)]
    pub r#type: i16,
    /// Offset position of the tab stop (part of MTEF spec)
    #[allow(dead_code)]
    pub offset: i16,
    /// Next tab stop in the list
    pub next: Option<Box<MtefTabstop>>,
}

impl MtefObject for MtefTabstop {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Pile object (MT_PILE)
///
/// Represents a vertical stack of lines, often used for column vectors,
/// cases constructs, or aligned equations.
///
/// Note: Alignment fields are part of MTEF spec, kept for future advanced formatting.
#[derive(Debug)]
pub struct MtefPile {
    /// Horizontal positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_x: i16,
    /// Vertical positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_y: i16,
    /// Horizontal alignment (left, center, right, part of MTEF spec)
    #[allow(dead_code)]
    pub halign: u8,
    /// Vertical alignment (top, center, bottom, part of MTEF spec)
    #[allow(dead_code)]
    pub valign: u8,
    /// Optional ruler for tab-aligned content (part of MTEF spec)
    #[allow(dead_code)]
    pub ruler: Option<Box<MtefRuler>>,
    /// Lines contained in this pile
    pub line_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefPile {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Matrix object (MT_MATRIX)
///
/// Represents a 2D array of cells (matrix, determinant, system of equations, etc.).
/// Cells are stored in row-major order in the element_list.
///
/// Note: Alignment and partition fields are part of MTEF spec,
/// kept for future advanced matrix formatting.
#[derive(Debug)]
pub struct MtefMatrix {
    /// Horizontal positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_x: i16,
    /// Vertical positioning nudge (part of MTEF spec)
    #[allow(dead_code)]
    pub nudge_y: i16,
    /// Vertical alignment relative to baseline (part of MTEF spec)
    #[allow(dead_code)]
    pub valign: u8,
    /// Horizontal justification within cells (part of MTEF spec)
    #[allow(dead_code)]
    pub h_just: u8,
    /// Vertical justification within cells (part of MTEF spec)
    #[allow(dead_code)]
    pub v_just: u8,
    /// Number of rows in the matrix
    pub rows: u8,
    /// Number of columns in the matrix
    pub cols: u8,
    /// Row partition information (part of MTEF spec for line spacing)
    #[allow(dead_code)]
    pub row_parts: [u8; 16],
    /// Column partition information (part of MTEF spec for spacing)
    #[allow(dead_code)]
    pub col_parts: [u8; 16],
    /// Matrix elements in row-major order
    pub element_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefMatrix {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Font object (MT_FONT)
///
/// Changes the current font for subsequent text rendering.
/// Part of MTEF spec, kept for future font and style support.
#[derive(Debug)]
pub struct MtefFont {
    /// Typeface index (part of MTEF spec)
    #[allow(dead_code)]
    pub tface: i32,
    /// Font style (bold, italic, etc., part of MTEF spec)
    #[allow(dead_code)]
    pub style: i32,
    /// Font name (part of MTEF spec)
    #[allow(dead_code)]
    pub zname: String,
}

impl MtefObject for MtefFont {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Size object (MT_SIZE)
///
/// Controls text size for subsequent content. Can specify absolute or relative sizes.
/// Part of MTEF spec, kept for future size and formatting support.
#[derive(Debug)]
pub struct MtefSize {
    /// Size type (predefined or custom, part of MTEF spec)
    #[allow(dead_code)]
    pub r#type: i32,
    /// Logical size level (part of MTEF spec)
    #[allow(dead_code)]
    pub lsize: i32,
    /// Delta size (relative adjustment, part of MTEF spec)
    #[allow(dead_code)]
    pub dsize: i32,
}

impl MtefObject for MtefSize {
    fn as_any(&self) -> &dyn std::any::Any { self }
}
