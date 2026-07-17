//! Inert BIFF8 supporting-book links and external cell caches.

use std::collections::HashSet;

use super::{XlsError, XlsResult};

pub(crate) const EXTERN_SHEET_RECORD_TYPE: u16 = 0x0017;
pub(crate) const XCT_RECORD_TYPE: u16 = 0x0059;
pub(crate) const CRN_RECORD_TYPE: u16 = 0x005a;
pub(crate) const SUP_BOOK_RECORD_TYPE: u16 = 0x01ae;
pub(crate) const EXTERN_NAME_RECORD_TYPE: u16 = 0x0023;
const CONTINUE_RECORD_TYPE: u16 = 0x003c;
const MAX_SUPPORTING_BOOKS: usize = 1024;
const MAX_EXTERNAL_SHEETS: usize = 256;
const MAX_EXTERNAL_REFERENCES: usize = 1370;
const MAX_CACHED_CELLS: usize = 65_536;
const MAX_EXTERNAL_NAMES: usize = 4096;
const MAX_EXTERNAL_NAME_BYTES: usize = 1_048_576;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsExternalNameBody {
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
        opaque_data: Vec<u8>,
        continuation_chunks: Vec<Vec<u8>>,
    },
    DdeStandardDocumentName { name: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsExternalName {
    supporting_book_index: u16,
    built_in: bool,
    automatic: bool,
    picture: bool,
    standard_document_name: bool,
    ole_link: bool,
    clipboard_format: i16,
    displayed_as_icon: bool,
    body: XlsExternalNameBody,
}

impl XlsExternalName {
    pub fn supporting_book_index(&self) -> u16 { self.supporting_book_index }
    pub fn built_in(&self) -> bool { self.built_in }
    pub fn automatic(&self) -> bool { self.automatic }
    pub fn picture(&self) -> bool { self.picture }
    pub fn standard_document_name(&self) -> bool { self.standard_document_name }
    pub fn ole_link(&self) -> bool { self.ole_link }
    pub fn clipboard_format(&self) -> i16 { self.clipboard_format }
    pub fn displayed_as_icon(&self) -> bool { self.displayed_as_icon }
    pub fn body(&self) -> &XlsExternalNameBody { &self.body }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsExternalCachedError {
    Null,
    DivisionByZero,
    Value,
    Reference,
    Name,
    Number,
    NotAvailable,
    GettingData,
}

impl XlsExternalCachedError {
    pub(crate) fn code(self) -> u8 {
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

    fn parse(code: u8) -> XlsResult<Self> {
        match code {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0f => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1d => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2a => Ok(Self::NotAvailable),
            0x2b => Ok(Self::GettingData),
            _ => invalid(CRN_RECORD_TYPE, format!("invalid cached error code 0x{code:02X}")),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsExternalCachedValue {
    Blank,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(XlsExternalCachedError),
}

#[derive(Debug, Clone, PartialEq)]
pub struct XlsExternalCacheRow {
    row: u16,
    first_column: u8,
    values: Vec<XlsExternalCachedValue>,
}

impl XlsExternalCacheRow {
    pub fn row(&self) -> u16 { self.row }
    pub fn first_column(&self) -> u8 { self.first_column }
    pub fn values(&self) -> &[XlsExternalCachedValue] { &self.values }
}

#[derive(Debug, Clone, PartialEq)]
pub struct XlsExternalSheet {
    name: String,
    cache_valid: bool,
    cache_rows: Vec<XlsExternalCacheRow>,
    cache_declared: bool,
}

impl XlsExternalSheet {
    pub fn name(&self) -> &str { &self.name }
    pub fn cache_valid(&self) -> bool { self.cache_valid }
    pub fn cache_rows(&self) -> &[XlsExternalCacheRow] { &self.cache_rows }
}

#[derive(Debug, Clone, PartialEq)]
pub struct XlsExternalWorkbook {
    encoded_virtual_path: String,
    sheets: Vec<XlsExternalSheet>,
}

impl XlsExternalWorkbook {
    /// Encoded BIFF virtual path. It is retained as inert metadata and never resolved.
    pub fn encoded_virtual_path(&self) -> &str { &self.encoded_virtual_path }
    pub fn sheets(&self) -> &[XlsExternalSheet] { &self.sheets }
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsSupportingBook {
    SelfReference,
    AddIn,
    ExternalWorkbook(XlsExternalWorkbook),
    SameSheet,
    Unused { sheet_names: Vec<String> },
    DdeOrOle { encoded_virtual_path: String },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsExternalSheetReference {
    supporting_book_index: u16,
    first_sheet_index: i16,
    last_sheet_index: i16,
}

impl XlsExternalSheetReference {
    pub fn supporting_book_index(&self) -> u16 { self.supporting_book_index }
    pub fn first_sheet_index(&self) -> i16 { self.first_sheet_index }
    pub fn last_sheet_index(&self) -> i16 { self.last_sheet_index }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct XlsExternalLinks {
    supporting_books: Vec<XlsSupportingBook>,
    external_names: Vec<XlsExternalName>,
    sheet_references: Vec<XlsExternalSheetReference>,
}

impl XlsExternalLinks {
    pub fn supporting_books(&self) -> &[XlsSupportingBook] { &self.supporting_books }
    pub fn external_names(&self) -> &[XlsExternalName] { &self.external_names }
    pub fn sheet_references(&self) -> &[XlsExternalSheetReference] { &self.sheet_references }
    pub fn external_workbooks(&self) -> impl Iterator<Item = &XlsExternalWorkbook> {
        self.supporting_books.iter().filter_map(|book| match book {
            XlsSupportingBook::ExternalWorkbook(book) => Some(book),
            _ => None,
        })
    }
}

struct PendingCache {
    book: usize,
    sheet: usize,
    remaining: usize,
}

#[derive(Default)]
pub(crate) struct ExternalLinkCollector {
    books: Vec<XlsSupportingBook>,
    references: Vec<XlsExternalSheetReference>,
    external_names: Vec<XlsExternalName>,
    pending: Option<PendingCache>,
    continuable_name: Option<usize>,
    names_allowed: bool,
    extern_sheet_seen: bool,
    closed: bool,
    cached_cells: usize,
    external_name_bytes: usize,
}

impl ExternalLinkCollector {
    pub(crate) fn new() -> Self { Self::default() }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if self.pending.as_ref().is_some_and(|pending| pending.remaining > 0)
            && record_type != CRN_RECORD_TYPE
        {
            return invalid(XCT_RECORD_TYPE, "XCT must be followed immediately by its declared CRN records");
        }

        if record_type == CONTINUE_RECORD_TYPE {
            if let Some(index) = self.continuable_name {
                if data.is_empty() || data.len() > 8224 {
                    return invalid(CONTINUE_RECORD_TYPE, "ExternName Continue payload must be 1..=8224 bytes");
                }
                self.external_name_bytes = self.external_name_bytes.checked_add(data.len()).ok_or_else(|| {
                    XlsError::InvalidRecord {
                        record_type: CONTINUE_RECORD_TYPE,
                        message: "ExternName continuation size overflows".to_string(),
                    }
                })?;
                if self.external_name_bytes > MAX_EXTERNAL_NAME_BYTES {
                    return invalid(CONTINUE_RECORD_TYPE, "ExternName opaque data exceeds resource bound");
                }
                let XlsExternalNameBody::DdeOrOle { continuation_chunks, .. } =
                    &mut self.external_names[index].body
                else {
                    unreachable!()
                };
                continuation_chunks.push(data.to_vec());
                return Ok(());
            }
            if !self.books.is_empty() && !self.closed {
                return invalid(CONTINUE_RECORD_TYPE, "Continue is not associated with a DDE/OLE ExternName");
            }
            return Ok(());
        }
        self.continuable_name = None;

        let target = matches!(
            record_type,
            SUP_BOOK_RECORD_TYPE
                | EXTERN_NAME_RECORD_TYPE
                | XCT_RECORD_TYPE
                | CRN_RECORD_TYPE
                | EXTERN_SHEET_RECORD_TYPE
        );
        if !target {
            if !self.books.is_empty() {
                self.closed = true;
            }
            return Ok(());
        }
        if self.closed {
            return invalid(record_type, "external-link record is outside its contiguous SUPBOOK collection");
        }

        match record_type {
            SUP_BOOK_RECORD_TYPE => {
                if self.extern_sheet_seen {
                    return invalid(record_type, "SupBook must precede ExternSheet");
                }
                if self.books.len() >= MAX_SUPPORTING_BOOKS {
                    return invalid(record_type, "supporting-book count exceeds resource bound");
                }
                self.books.push(parse_sup_book(data)?);
                self.names_allowed = true;
            },
            EXTERN_NAME_RECORD_TYPE => {
                if !self.names_allowed {
                    return invalid(record_type, "ExternName must directly follow its active SupBook name collection");
                }
                if self.external_names.len() >= MAX_EXTERNAL_NAMES {
                    return invalid(record_type, "external name count exceeds resource bound");
                }
                let book_index = self.books.len().checked_sub(1).ok_or_else(|| XlsError::InvalidRecord {
                    record_type,
                    message: "ExternName appears without a preceding SupBook".to_string(),
                })?;
                let name = parse_external_name(data, book_index, &self.books[book_index])?;
                self.external_name_bytes = self.external_name_bytes.checked_add(data.len()).ok_or_else(|| {
                    XlsError::InvalidRecord {
                        record_type,
                        message: "external name size overflows".to_string(),
                    }
                })?;
                if self.external_name_bytes > MAX_EXTERNAL_NAME_BYTES {
                    return invalid(record_type, "external name data exceeds resource bound");
                }
                let continuable = matches!(name.body, XlsExternalNameBody::DdeOrOle { .. });
                self.external_names.push(name);
                if continuable { self.continuable_name = Some(self.external_names.len() - 1); }
            },
            XCT_RECORD_TYPE => {
                self.names_allowed = false;
                self.parse_xct(data)?;
            },
            CRN_RECORD_TYPE => self.parse_crn(data)?,
            EXTERN_SHEET_RECORD_TYPE => {
                self.names_allowed = false;
                self.extern_sheet_seen = true;
                let mut references = parse_extern_sheet(data)?;
                if self.references.len() + references.len() > MAX_EXTERNAL_REFERENCES {
                    return invalid(record_type, "external sheet reference count exceeds BIFF8 record bound");
                }
                self.references.append(&mut references);
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    fn parse_xct(&mut self, data: &[u8]) -> XlsResult<()> {
        if data.len() != 4 {
            return Err(XlsError::InvalidLength { expected: 4, found: data.len() });
        }
        let book_index = self.books.len().checked_sub(1).ok_or_else(|| XlsError::InvalidRecord {
            record_type: XCT_RECORD_TYPE,
            message: "XCT appears without a preceding SupBook".to_string(),
        })?;
        let declared = i16::from_le_bytes([data[0], data[1]]);
        if declared == i16::MIN {
            return invalid(XCT_RECORD_TYPE, "XCT CRN count absolute value overflows");
        }
        let sheet_index = usize::from(read_u16(data, 2));
        let XlsSupportingBook::ExternalWorkbook(book) = &mut self.books[book_index] else {
            return invalid(XCT_RECORD_TYPE, "XCT cache requires an external-workbook SupBook");
        };
        let sheet = book.sheets.get_mut(sheet_index).ok_or_else(|| XlsError::InvalidRecord {
            record_type: XCT_RECORD_TYPE,
            message: "XCT sheet index exceeds SupBook sheet count".to_string(),
        })?;
        if sheet.cache_declared {
            return invalid(XCT_RECORD_TYPE, "duplicate XCT cache for external sheet");
        }
        sheet.cache_declared = true;
        sheet.cache_valid = declared >= 0;
        let remaining = usize::from(declared.unsigned_abs());
        self.pending = Some(PendingCache { book: book_index, sheet: sheet_index, remaining });
        Ok(())
    }

    fn parse_crn(&mut self, data: &[u8]) -> XlsResult<()> {
        let pending = self.pending.as_mut().ok_or_else(|| XlsError::InvalidRecord {
            record_type: CRN_RECORD_TYPE,
            message: "CRN appears without XCT".to_string(),
        })?;
        if pending.remaining == 0 {
            return invalid(CRN_RECORD_TYPE, "more CRN records than declared by XCT");
        }
        let row = parse_cache_row(data)?;
        self.cached_cells = self.cached_cells.checked_add(row.values.len()).ok_or_else(|| {
            XlsError::InvalidRecord {
                record_type: CRN_RECORD_TYPE,
                message: "cached cell count overflows".to_string(),
            }
        })?;
        if self.cached_cells > MAX_CACHED_CELLS {
            return invalid(CRN_RECORD_TYPE, "external cached cell count exceeds resource bound");
        }
        let XlsSupportingBook::ExternalWorkbook(book) = &mut self.books[pending.book] else {
            unreachable!()
        };
        book.sheets[pending.sheet].cache_rows.push(row);
        pending.remaining -= 1;
        Ok(())
    }

    pub(crate) fn finish(self, internal_sheet_count: usize) -> XlsResult<XlsExternalLinks> {
        if self.pending.as_ref().is_some_and(|pending| pending.remaining > 0) {
            return invalid(XCT_RECORD_TYPE, "workbook ended before all CRN records declared by XCT");
        }
        for reference in &self.references {
            validate_reference(reference, &self.books, internal_sheet_count)?;
        }
        for book in &self.books {
            if let XlsSupportingBook::ExternalWorkbook(book) = book {
                for sheet in &book.sheets {
                    let mut cells = HashSet::new();
                    for row in &sheet.cache_rows {
                        for offset in 0..row.values.len() {
                            let column = usize::from(row.first_column) + offset;
                            if !cells.insert((row.row, column)) {
                                return invalid(CRN_RECORD_TYPE, "external cache contains duplicate cells");
                            }
                        }
                    }
                }
            }
        }
        Ok(XlsExternalLinks {
            supporting_books: self.books,
            external_names: self.external_names,
            sheet_references: self.references,
        })
    }
}

fn parse_external_name(
    data: &[u8],
    book_index: usize,
    book: &XlsSupportingBook,
) -> XlsResult<XlsExternalName> {
    if data.len() < 8 || data.len() > 8224 {
        return invalid(EXTERN_NAME_RECORD_TYPE, "ExternName payload length must be 8..=8224 bytes");
    }
    let flags = read_u16(data, 0);
    let built_in = flags & 0x0001 != 0;
    let automatic = flags & 0x0002 != 0;
    let picture = flags & 0x0004 != 0;
    let standard_document_name = flags & 0x0008 != 0;
    let ole_link = flags & 0x0010 != 0;
    if standard_document_name && ole_link {
        return invalid(EXTERN_NAME_RECORD_TYPE, "ExternName fOle and fOleLink are mutually exclusive");
    }
    let clipboard_bits = (flags >> 5) & 0x03ff;
    let clipboard_format = if clipboard_bits & 0x0200 != 0 {
        clipboard_bits as i16 - 1024
    } else {
        clipboard_bits as i16
    };
    if !matches!(clipboard_format, -1 | 0 | 2 | 5 | 6 | 7 | 8 | 9 | 16 | 20 | 30 | 36 | 44 | 63) {
        return invalid(EXTERN_NAME_RECORD_TYPE, "ExternName clipboard format is invalid");
    }
    let displayed_as_icon = flags & 0x8000 != 0;
    let (name, offset) = parse_short_unicode_string(data, 6, EXTERN_NAME_RECORD_TYPE)?;

    let body = match book {
        XlsSupportingBook::ExternalWorkbook(book) => {
            if flags & !0x0001 != 0 {
                return invalid(EXTERN_NAME_RECORD_TYPE, "external defined name has non-name flags");
            }
            if read_u16(data, 4) != 0 {
                return invalid(EXTERN_NAME_RECORD_TYPE, "ExternDocName reserved field must be zero");
            }
            let ixals = read_u16(data, 2);
            if usize::from(ixals) > book.sheets.len() {
                return invalid(EXTERN_NAME_RECORD_TYPE, "ExternDocName sheet scope exceeds SupBook sheet count");
            }
            let formula_len = usize::from(*data.get(offset).ok_or_else(|| XlsError::InvalidRecord {
                record_type: EXTERN_NAME_RECORD_TYPE,
                message: "ExtNameParsedFormula length is missing".to_string(),
            })?) | (usize::from(*data.get(offset + 1).ok_or_else(|| XlsError::InvalidRecord {
                record_type: EXTERN_NAME_RECORD_TYPE,
                message: "ExtNameParsedFormula length is truncated".to_string(),
            })?) << 8);
            let formula_start = offset + 2;
            if formula_start.checked_add(formula_len) != Some(data.len()) {
                return invalid(EXTERN_NAME_RECORD_TYPE, "ExtNameParsedFormula length does not consume ExternName");
            }
            let formula_bytes = data[formula_start..].to_vec();
            if formula_bytes.first().is_some_and(|token| !matches!(token, 0x1c | 0x3a | 0x3b | 0x3c | 0x3d)) {
                return invalid(EXTERN_NAME_RECORD_TYPE, "ExtNameParsedFormula token kind is invalid");
            }
            XlsExternalNameBody::ExternalDefinedName {
                sheet_index: ixals.checked_sub(1),
                name,
                formula_bytes,
            }
        },
        XlsSupportingBook::AddIn => {
            if flags != 0 || read_u16(data, 2) != 0 || read_u16(data, 4) != 0 {
                return invalid(EXTERN_NAME_RECORD_TYPE, "AddinUdf flags and reserved fields must be zero");
            }
            let unused_len = usize::from(read_u16_checked(data, offset, EXTERN_NAME_RECORD_TYPE)?);
            let unused_start = offset + 2;
            if unused_start.checked_add(unused_len) != Some(data.len()) {
                return invalid(EXTERN_NAME_RECORD_TYPE, "AddinUdf unused byte count does not consume ExternName");
            }
            XlsExternalNameBody::AddInFunction {
                name,
                unused_data: data[unused_start..].to_vec(),
            }
        },
        XlsSupportingBook::DdeOrOle { .. } => {
            let storage_id = u32::from_le_bytes(data[2..6].try_into().unwrap());
            if built_in {
                return invalid(EXTERN_NAME_RECORD_TYPE, "DDE/OLE ExternName has incompatible flags");
            }
            if standard_document_name {
                if displayed_as_icon
                    || storage_id != 0
                    || name != "StdDocumentName"
                    || offset != data.len()
                {
                    return invalid(EXTERN_NAME_RECORD_TYPE, "ExternDdeLinkNoOper body is invalid");
                }
                XlsExternalNameBody::DdeStandardDocumentName { name }
            } else {
                XlsExternalNameBody::DdeOrOle {
                    storage_id,
                    name,
                    opaque_data: data[offset..].to_vec(),
                    continuation_chunks: Vec::new(),
                }
            }
        },
        _ => return invalid(EXTERN_NAME_RECORD_TYPE, "SupBook variant cannot own ExternName records"),
    };
    Ok(XlsExternalName {
        supporting_book_index: u16::try_from(book_index).map_err(|_| XlsError::InvalidRecord {
            record_type: EXTERN_NAME_RECORD_TYPE,
            message: "supporting-book index exceeds u16".to_string(),
        })?,
        built_in,
        automatic,
        picture,
        standard_document_name,
        ole_link,
        clipboard_format,
        displayed_as_icon,
        body,
    })
}

fn parse_short_unicode_string(
    data: &[u8],
    offset: usize,
    record_type: u16,
) -> XlsResult<(String, usize)> {
    let count = usize::from(*data.get(offset).ok_or_else(|| XlsError::InvalidRecord {
        record_type,
        message: "short Unicode string length is missing".to_string(),
    })?);
    parse_unicode_no_cch(data, offset + 1, count, record_type)
}

fn read_u16_checked(data: &[u8], offset: usize, record_type: u16) -> XlsResult<u16> {
    let bytes = data.get(offset..offset + 2).ok_or_else(|| XlsError::InvalidRecord {
        record_type,
        message: "two-byte field is truncated".to_string(),
    })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn parse_sup_book(data: &[u8]) -> XlsResult<XlsSupportingBook> {
    if data.len() < 4 || data.len() > 8224 {
        return invalid(SUP_BOOK_RECORD_TYPE, "SupBook payload length is outside BIFF8 bounds");
    }
    let sheet_count = usize::from(read_u16(data, 0));
    let cch = read_u16(data, 2);
    if data.len() == 4 {
        return match cch {
            0x0401 => Ok(XlsSupportingBook::SelfReference),
            0x3a01 if sheet_count == 1 => Ok(XlsSupportingBook::AddIn),
            0x3a01 => invalid(SUP_BOOK_RECORD_TYPE, "add-in SupBook sheet count must be one"),
            _ => invalid(SUP_BOOK_RECORD_TYPE, "invalid four-byte SupBook marker"),
        };
    }
    if !(1..=255).contains(&cch) {
        return invalid(SUP_BOOK_RECORD_TYPE, "SupBook virtual path length must be 1..=255");
    }
    let (path, mut offset) = parse_unicode_no_cch(data, 4, usize::from(cch), SUP_BOOK_RECORD_TYPE)?;
    if path == "\0" {
        if sheet_count != 0 || offset != data.len() {
            return invalid(SUP_BOOK_RECORD_TYPE, "same-sheet SupBook must have zero sheets and no rgst");
        }
        return Ok(XlsSupportingBook::SameSheet);
    }
    if sheet_count == 0 {
        if offset != data.len() {
            return invalid(SUP_BOOK_RECORD_TYPE, "DDE/OLE SupBook has trailing sheet names");
        }
        return Ok(XlsSupportingBook::DdeOrOle { encoded_virtual_path: path });
    }
    if sheet_count > MAX_EXTERNAL_SHEETS {
        return invalid(SUP_BOOK_RECORD_TYPE, "external SupBook sheet count exceeds resource bound");
    }
    let mut names = Vec::with_capacity(sheet_count);
    for _ in 0..sheet_count {
        let (name, next) = parse_unicode_string(data, offset, 31, SUP_BOOK_RECORD_TYPE)?;
        names.push(name);
        offset = next;
    }
    if offset != data.len() {
        return invalid(SUP_BOOK_RECORD_TYPE, "SupBook sheet count does not consume payload exactly");
    }
    if path == " " {
        if names.iter().any(|name| name != " ") {
            return invalid(SUP_BOOK_RECORD_TYPE, "unused SupBook sheet placeholders must be spaces");
        }
        return Ok(XlsSupportingBook::Unused { sheet_names: names });
    }
    Ok(XlsSupportingBook::ExternalWorkbook(XlsExternalWorkbook {
        encoded_virtual_path: path,
        sheets: names.into_iter().map(|name| XlsExternalSheet {
            name,
            cache_valid: true,
            cache_rows: Vec::new(),
            cache_declared: false,
        }).collect(),
    }))
}

fn parse_extern_sheet(data: &[u8]) -> XlsResult<Vec<XlsExternalSheetReference>> {
    if data.len() < 2 {
        return Err(XlsError::InvalidLength { expected: 2, found: data.len() });
    }
    let count = usize::from(read_u16(data, 0));
    let expected = 2usize.checked_add(count.checked_mul(6).ok_or_else(|| XlsError::InvalidRecord {
        record_type: EXTERN_SHEET_RECORD_TYPE,
        message: "ExternSheet count overflows".to_string(),
    })?).ok_or_else(|| XlsError::InvalidRecord {
        record_type: EXTERN_SHEET_RECORD_TYPE,
        message: "ExternSheet size overflows".to_string(),
    })?;
    if data.len() != expected {
        return Err(XlsError::InvalidLength { expected, found: data.len() });
    }
    Ok(data[2..].chunks_exact(6).map(|entry| XlsExternalSheetReference {
        supporting_book_index: read_u16(entry, 0),
        first_sheet_index: i16::from_le_bytes([entry[2], entry[3]]),
        last_sheet_index: i16::from_le_bytes([entry[4], entry[5]]),
    }).collect())
}

fn validate_reference(
    reference: &XlsExternalSheetReference,
    books: &[XlsSupportingBook],
    internal_sheet_count: usize,
) -> XlsResult<()> {
    let book = books.get(usize::from(reference.supporting_book_index)).ok_or_else(|| {
        XlsError::InvalidRecord {
            record_type: EXTERN_SHEET_RECORD_TYPE,
            message: "XTI supporting-book index is out of range".to_string(),
        }
    })?;
    let (first, last) = (reference.first_sheet_index, reference.last_sheet_index);
    match book {
        XlsSupportingBook::AddIn | XlsSupportingBook::SameSheet | XlsSupportingBook::DdeOrOle { .. } => {
            if first != -2 || last != -2 {
                return invalid(EXTERN_SHEET_RECORD_TYPE, "unscoped supporting link requires -2/-2 XTI scope");
            }
        },
        XlsSupportingBook::SelfReference => validate_sheet_scope(first, last, internal_sheet_count)?,
        XlsSupportingBook::ExternalWorkbook(book) => validate_sheet_scope(first, last, book.sheets.len())?,
        XlsSupportingBook::Unused { sheet_names } => validate_sheet_scope(first, last, sheet_names.len())?,
    }
    Ok(())
}

fn validate_sheet_scope(first: i16, last: i16, sheet_count: usize) -> XlsResult<()> {
    if first == -2 {
        if last != -2 { return invalid(EXTERN_SHEET_RECORD_TYPE, "workbook XTI scope must be -2/-2"); }
        return Ok(());
    }
    if first == -1 || last == -1 {
        if first < -1 || last < -1 { return invalid(EXTERN_SHEET_RECORD_TYPE, "invalid missing-sheet XTI scope"); }
        return Ok(());
    }
    if first < 0 || last < first || usize::try_from(last).unwrap_or(usize::MAX) >= sheet_count {
        return invalid(EXTERN_SHEET_RECORD_TYPE, "XTI sheet scope is outside supporting book");
    }
    Ok(())
}

fn parse_cache_row(data: &[u8]) -> XlsResult<XlsExternalCacheRow> {
    if data.len() < 4 {
        return Err(XlsError::InvalidLength { expected: 4, found: data.len() });
    }
    let last_column = data[0];
    let first_column = data[1];
    if last_column < first_column {
        return invalid(CRN_RECORD_TYPE, "CRN column range is reversed");
    }
    let value_count = usize::from(last_column - first_column) + 1;
    let mut offset = 4;
    let mut values = Vec::with_capacity(value_count);
    for _ in 0..value_count {
        let tag = *data.get(offset).ok_or_else(|| XlsError::InvalidRecord {
            record_type: CRN_RECORD_TYPE,
            message: "CRN cached value is truncated".to_string(),
        })?;
        match tag {
            0x00 => {
                require_available(data, offset, 9)?;
                values.push(XlsExternalCachedValue::Blank);
                offset += 9;
            },
            0x01 => {
                require_available(data, offset, 9)?;
                let value = f64::from_le_bytes(data[offset + 1..offset + 9].try_into().unwrap());
                values.push(XlsExternalCachedValue::Number(value));
                offset += 9;
            },
            0x02 => {
                let (value, next) = parse_unicode_string(data, offset + 1, 255, CRN_RECORD_TYPE)?;
                values.push(XlsExternalCachedValue::Text(value));
                offset = next;
            },
            0x04 => {
                require_available(data, offset, 9)?;
                if data[offset + 2] != 0 {
                    return invalid(CRN_RECORD_TYPE, "SerBool reserved byte must be zero");
                }
                let value = match data[offset + 1] {
                    0 => false,
                    1 => true,
                    _ => return invalid(CRN_RECORD_TYPE, "SerBool value must be zero or one"),
                };
                values.push(XlsExternalCachedValue::Boolean(value));
                offset += 9;
            },
            0x10 => {
                require_available(data, offset, 9)?;
                if data[offset + 2] != 0 {
                    return invalid(CRN_RECORD_TYPE, "SerErr reserved byte must be zero");
                }
                values.push(XlsExternalCachedValue::Error(XlsExternalCachedError::parse(data[offset + 1])?));
                offset += 9;
            },
            _ => return invalid(CRN_RECORD_TYPE, format!("unknown SerAr tag 0x{tag:02X}")),
        }
    }
    if offset != data.len() {
        return invalid(CRN_RECORD_TYPE, "CRN cached values do not consume payload exactly");
    }
    Ok(XlsExternalCacheRow {
        row: read_u16(data, 2),
        first_column,
        values,
    })
}

fn parse_unicode_string(
    data: &[u8],
    offset: usize,
    max_characters: usize,
    record_type: u16,
) -> XlsResult<(String, usize)> {
    require_available(data, offset, 3)?;
    let count = usize::from(read_u16(data, offset));
    parse_unicode_no_cch(data, offset + 2, count, record_type).and_then(|(value, next)| {
        if count > max_characters {
            invalid(record_type, format!("string exceeds {max_characters} characters"))
        } else {
            Ok((value, next))
        }
    })
}

fn parse_unicode_no_cch(
    data: &[u8],
    offset: usize,
    count: usize,
    record_type: u16,
) -> XlsResult<(String, usize)> {
    let flags = *data.get(offset).ok_or_else(|| XlsError::InvalidRecord {
        record_type,
        message: "Unicode string options are missing".to_string(),
    })?;
    if flags & !1 != 0 {
        return invalid(record_type, "Unicode string reserved flags must be zero");
    }
    let wide = flags == 1;
    let bytes = count.checked_mul(if wide { 2 } else { 1 }).ok_or_else(|| XlsError::InvalidRecord {
        record_type,
        message: "Unicode string length overflows".to_string(),
    })?;
    let start = offset + 1;
    let end = start.checked_add(bytes).ok_or_else(|| XlsError::InvalidRecord {
        record_type,
        message: "Unicode string end overflows".to_string(),
    })?;
    let encoded = data.get(start..end).ok_or_else(|| XlsError::InvalidRecord {
        record_type,
        message: "Unicode string is truncated".to_string(),
    })?;
    let value = if wide {
        let units = encoded.chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units).map_err(|error| XlsError::InvalidRecord {
            record_type,
            message: format!("invalid UTF-16 string: {error}"),
        })?
    } else {
        encoded.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok((value, end))
}

fn require_available(data: &[u8], offset: usize, length: usize) -> XlsResult<()> {
    let end = offset.checked_add(length).ok_or_else(|| XlsError::InvalidRecord {
        record_type: CRN_RECORD_TYPE,
        message: "cached value length overflows".to_string(),
    })?;
    if end > data.len() {
        return Err(XlsError::InvalidLength { expected: end, found: data.len() });
    }
    Ok(())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord { record_type, message: message.into() })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn malformed_supbook_xct_and_crn_are_rejected() {
        assert!(parse_sup_book(&[1, 0, 2, 4]).is_err());
        assert!(parse_cache_row(&[0, 1, 0, 0]).is_err());
        assert!(parse_cache_row(&[0, 0, 0, 0, 4, 2, 0, 0, 0, 0, 0, 0, 0]).is_err());

        let mut collector = ExternalLinkCollector::new();
        assert!(collector.feed_record(XCT_RECORD_TYPE, &[1, 0, 0, 0]).is_err());
    }

    #[test]
    fn xct_cardinality_and_xti_bounds_are_strict() {
        let external = [1, 0, 2, 0, 0, 1, b'A', 1, 0, 0, b'S'];
        let mut collector = ExternalLinkCollector::new();
        collector.feed_record(SUP_BOOK_RECORD_TYPE, &external).unwrap();
        collector.feed_record(XCT_RECORD_TYPE, &[1, 0, 0, 0]).unwrap();
        assert!(collector.feed_record(EXTERN_SHEET_RECORD_TYPE, &[0, 0]).is_err());

        let mut collector = ExternalLinkCollector::new();
        collector.feed_record(SUP_BOOK_RECORD_TYPE, &external).unwrap();
        collector.feed_record(EXTERN_SHEET_RECORD_TYPE, &[1, 0, 1, 0, 0, 0, 0, 0]).unwrap();
        assert!(collector.finish(1).is_err());
    }

    #[test]
    fn extern_names_are_contextual_bounded_and_continuable_only_for_links() {
        let addin = [1, 0, 1, 0x3a];
        let addin_name = [0, 0, 0, 0, 0, 0, 1, 0, b'F', 2, 0, 0x1c, 0x17];
        let mut collector = ExternalLinkCollector::new();
        assert!(collector.feed_record(EXTERN_NAME_RECORD_TYPE, &addin_name).is_err());
        collector.feed_record(SUP_BOOK_RECORD_TYPE, &addin).unwrap();
        collector.feed_record(EXTERN_NAME_RECORD_TYPE, &addin_name).unwrap();
        assert!(collector.feed_record(CONTINUE_RECORD_TYPE, &[0]).is_err());

        let dde = [0, 0, 3, 0, 0, b'A', 3, b'B'];
        let dde_name = [2, 0, 0, 0, 0, 0, 1, 0, b'X', 0, 0, 0];
        let mut collector = ExternalLinkCollector::new();
        collector.feed_record(SUP_BOOK_RECORD_TYPE, &dde).unwrap();
        collector.feed_record(EXTERN_NAME_RECORD_TYPE, &dde_name).unwrap();
        collector.feed_record(CONTINUE_RECORD_TYPE, &[0; 9]).unwrap();
        let links = collector.finish(1).unwrap();
        let XlsExternalNameBody::DdeOrOle { continuation_chunks, .. } = links.external_names()[0].body() else {
            panic!("expected DDE/OLE body")
        };
        assert_eq!(continuation_chunks, &[vec![0; 9]]);
    }
}
