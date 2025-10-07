// MTEF object structures and types
//
// Based on rtf2latex2e eqn_support.h structures

/// MTEF record types (as defined in rtf2latex2e)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MtefRecordType {
    End = 0,
    Line = 1,
    Char = 2,
    Tmpl = 3,
    Pile = 4,
    Matrix = 5,
    Embell = 6,
    Ruler = 7,
    Font = 18,
    Size = 9,
    Full = 10,
    Sub = 11,
    Sub2 = 12,
    Sym = 13,
    SubSym = 14,
    Color = 15,
    ColorDef = 16,
    FontDef = 17,
    EqnPrefs = 19,
    EncodingDef = 20,
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
#[derive(Debug)]
pub struct MtefChar {
    pub nudge_x: i16,
    pub nudge_y: i16,
    pub atts: u8,           // Character attributes
    pub typeface: u8,       // Typeface (128 + index)
    pub character: u16,     // Character code
    pub bits16: u16,        // Additional bits for MTEF v5
    pub embellishment_list: Option<Box<MtefEmbell>>,
}

impl MtefObject for MtefChar {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Embellishment object (MT_EMBELL)
#[derive(Debug)]
pub struct MtefEmbell {
    pub nudge_x: i16,
    pub nudge_y: i16,
    pub embell: u8,         // Embellishment type
    pub next: Option<Box<MtefEmbell>>,
}

impl MtefObject for MtefEmbell {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Template object (MT_TMPL)
#[derive(Debug)]
pub struct MtefTemplate {
    pub nudge_x: i16,
    pub nudge_y: i16,
    pub selector: u8,       // Template selector
    pub variation: u16,     // Template variation
    pub options: u8,        // Template options
    pub subobject_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefTemplate {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Line object (MT_LINE)
#[derive(Debug)]
pub struct MtefLine {
    pub nudge_x: i16,
    pub nudge_y: i16,
    pub line_spacing: u8,
    pub ruler: Option<Box<MtefRuler>>,
    pub object_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefLine {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Ruler object (MT_RULER)
#[derive(Debug)]
pub struct MtefRuler {
    pub n_stops: i16,
    pub tabstop_list: Option<Box<MtefTabstop>>,
}

impl MtefObject for MtefRuler {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Tabstop object (MT_TABSTOP)
#[derive(Debug)]
pub struct MtefTabstop {
    pub r#type: i16,
    pub offset: i16,
    pub next: Option<Box<MtefTabstop>>,
}

impl MtefObject for MtefTabstop {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Pile object (MT_PILE)
#[derive(Debug)]
pub struct MtefPile {
    pub nudge_x: i16,
    pub nudge_y: i16,
    pub halign: u8,         // Horizontal alignment
    pub valign: u8,         // Vertical alignment
    pub ruler: Option<Box<MtefRuler>>,
    pub line_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefPile {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Matrix object (MT_MATRIX)
#[derive(Debug)]
pub struct MtefMatrix {
    pub nudge_x: i16,
    pub nudge_y: i16,
    pub valign: u8,         // Vertical alignment
    pub h_just: u8,         // Horizontal justification
    pub v_just: u8,         // Vertical justification
    pub rows: u8,           // Number of rows
    pub cols: u8,           // Number of columns
    pub row_parts: [u8; 16], // Row partition info
    pub col_parts: [u8; 16], // Column partition info
    pub element_list: Option<Box<MtefObjectList>>,
}

impl MtefObject for MtefMatrix {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Font object (MT_FONT)
#[derive(Debug)]
pub struct MtefFont {
    pub tface: i32,
    pub style: i32,
    pub zname: String,
}

impl MtefObject for MtefFont {
    fn as_any(&self) -> &dyn std::any::Any { self }
}

/// Size object (MT_SIZE)
#[derive(Debug)]
pub struct MtefSize {
    pub r#type: i32,
    pub lsize: i32,
    pub dsize: i32,
}

impl MtefObject for MtefSize {
    fn as_any(&self) -> &dyn std::any::Any { self }
}
