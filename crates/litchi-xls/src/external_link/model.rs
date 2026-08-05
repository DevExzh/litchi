//! Semantic BIFF8 external-link values.

use crate::{Error, Result};

/// Clipboard format cached by a DDE or OLE external-name item.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i16)]
pub enum ClipboardFormat {
    None = -1,
    Text = 0,
    EnhancedMetafile = 2,
    Csv = 5,
    Sylk = 6,
    RichText = 7,
    Biff8 = 8,
    Bitmap = 9,
    ApplicationTable = 16,
    Biff3 = 20,
    Biff4 = 30,
    MetafilePicture = 36,
    UnicodeText = 44,
    Biff12 = 63,
}

impl ClipboardFormat {
    /// Returns the BIFF clipboard-format code.
    pub const fn code(self) -> i16 {
        self as i16
    }
}

/// Current values cached for a DDE or OLE item, in row-major order.
#[derive(Debug, Clone, PartialEq)]
pub struct ValueMatrix {
    pub last_column: u8,
    pub last_row: u16,
    pub values: Vec<CachedValue>,
}

impl ValueMatrix {
    pub fn validate(&self) -> Result<()> {
        let count = (usize::from(self.last_column) + 1)
            .checked_mul(usize::from(self.last_row) + 1)
            .ok_or_else(|| Error::InvalidData("DDE/OLE matrix size overflows".to_string()))?;
        if count > super::MAX_DDE_OLE_VALUES || self.values.len() != count {
            return Err(Error::InvalidData(
                "DDE/OLE matrix dimensions do not match its bounded value array".to_string(),
            ));
        }
        for value in &self.values {
            if let CachedValue::Text(text) = value
                && text.encode_utf16().count() > 255
            {
                return Err(Error::InvalidData(
                    "DDE/OLE matrix text exceeds 255 UTF-16 code units".to_string(),
                ));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum NameBody {
    ExternalDefinedName {
        sheet_index: Option<u16>,
        name: String,
        formula_bytes: Vec<u8>,
    },
    AddInFunction {
        name: String,
        unused_data: Vec<u8>,
    },
    DdeOrOle {
        storage_id: u32,
        name: String,
        matrix: Option<ValueMatrix>,
    },
    DdeStandardDocumentName {
        name: String,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct Name {
    pub(super) supporting_book_index: u16,
    pub(super) built_in: bool,
    pub(super) automatic: bool,
    pub(super) picture: bool,
    pub(super) standard_document_name: bool,
    pub(super) ole_link: bool,
    pub(super) clipboard_format: ClipboardFormat,
    pub(super) displayed_as_icon: bool,
    pub(super) body: NameBody,
}

impl Name {
    pub fn supporting_book_index(&self) -> u16 {
        self.supporting_book_index
    }
    pub fn built_in(&self) -> bool {
        self.built_in
    }
    pub fn automatic(&self) -> bool {
        self.automatic
    }
    pub fn picture(&self) -> bool {
        self.picture
    }
    pub fn standard_document_name(&self) -> bool {
        self.standard_document_name
    }
    pub fn ole_link(&self) -> bool {
        self.ole_link
    }
    pub fn clipboard_format(&self) -> ClipboardFormat {
        self.clipboard_format
    }
    pub fn displayed_as_icon(&self) -> bool {
        self.displayed_as_icon
    }
    pub fn body(&self) -> &NameBody {
        &self.body
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorValue {
    Null,
    DivisionByZero,
    Value,
    Reference,
    Name,
    Number,
    NotAvailable,
    GettingData,
}

impl ErrorValue {
    pub(crate) const fn code(self) -> u8 {
        match self {
            Self::Null => 0x00,
            Self::DivisionByZero => 0x07,
            Self::Value => 0x0f,
            Self::Reference => 0x17,
            Self::Name => 0x1d,
            Self::Number => 0x24,
            Self::NotAvailable => 0x2a,
            Self::GettingData => 0x2b,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum CachedValue {
    Blank,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(ErrorValue),
}

#[derive(Debug, Clone, PartialEq)]
pub struct CacheRow {
    pub(super) row: u16,
    pub(super) first_column: u8,
    pub(super) values: Vec<CachedValue>,
}

impl CacheRow {
    pub fn row(&self) -> u16 {
        self.row
    }
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    pub fn values(&self) -> &[CachedValue] {
        &self.values
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Sheet {
    pub(super) name: String,
    pub(super) cache_valid: bool,
    pub(super) cache_rows: Vec<CacheRow>,
    pub(super) cache_declared: bool,
}

impl Sheet {
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn cache_valid(&self) -> bool {
        self.cache_valid
    }
    pub fn cache_rows(&self) -> &[CacheRow] {
        &self.cache_rows
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Workbook {
    pub(super) encoded_virtual_path: String,
    pub(super) sheets: Vec<Sheet>,
}

impl Workbook {
    /// Encoded BIFF virtual path. It is retained as inert metadata and never resolved.
    pub fn encoded_virtual_path(&self) -> &str {
        &self.encoded_virtual_path
    }
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum SupportingBook {
    SelfReference,
    AddIn,
    ExternalWorkbook(Workbook),
    SameSheet,
    Unused { sheet_names: Vec<String> },
    DdeOrOle { encoded_virtual_path: String },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SheetReference {
    pub(super) supporting_book_index: u16,
    pub(super) first_sheet_index: i16,
    pub(super) last_sheet_index: i16,
}

impl SheetReference {
    pub fn supporting_book_index(&self) -> u16 {
        self.supporting_book_index
    }
    pub fn first_sheet_index(&self) -> i16 {
        self.first_sheet_index
    }
    pub fn last_sheet_index(&self) -> i16 {
        self.last_sheet_index
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Links {
    pub(super) supporting_books: Vec<SupportingBook>,
    pub(super) external_names: Vec<Name>,
    pub(super) sheet_references: Vec<SheetReference>,
}

impl Links {
    pub fn supporting_books(&self) -> &[SupportingBook] {
        &self.supporting_books
    }
    pub fn external_names(&self) -> &[Name] {
        &self.external_names
    }
    pub fn sheet_references(&self) -> &[SheetReference] {
        &self.sheet_references
    }
    pub fn external_workbooks(&self) -> impl Iterator<Item = &Workbook> {
        self.supporting_books.iter().filter_map(|book| match book {
            SupportingBook::ExternalWorkbook(book) => Some(book),
            _ => None,
        })
    }
}
