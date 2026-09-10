//! Typed `SpreadsheetML` external-link models.
//!
//! These values describe external-link markup without opening, fetching, or
//! activating any target. OPC relationship resolution belongs to `package`.

pub(crate) const MAX_SHEET_NAMES: usize = 65_536;
pub(crate) const MAX_DEFINED_NAMES: usize = 65_536;
pub(crate) const MAX_CACHED_SHEETS: usize = 65_536;
pub(crate) const MAX_CACHED_ROWS: usize = 1_048_576;
pub(crate) const MAX_CACHED_CELLS: usize = 1_000_000;
pub(crate) const MAX_LINK_ITEMS: usize = 65_536;
pub(crate) const MAX_CACHE_TEXT_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub(crate) const TRANSITIONAL_SML: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const MAX_EXTERNAL_TARGET_BYTES: usize = 32 * 1024;
/// Maximum bytes retained for either service-provided alternate-URL ID.
///
/// The normative type is `xsd:string` and does not impose a useful Office
/// lexical limit.  This package boundary therefore applies a deterministic
/// bound before retaining or emitting either inert identifier.
pub(crate) const MAX_ALTERNATE_URL_METADATA_BYTES: usize = 64 * 1024;
/// 2021 external-links extension namespace from [MS-XLSX] §5.34.
pub(crate) const ALTERNATE_URLS_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021";
/// Highest column index addressable by a `SpreadsheetML` cell reference (`XFD`).
pub(crate) const MAX_CELL_COLUMN: u32 = 16_384;
/// Highest row index addressable by a `SpreadsheetML` cell reference.
pub(crate) const MAX_CELL_ROW: u32 = 1_048_576;
/// Longest column prefix a valid reference can carry (`XFD` is three letters).
pub(crate) const MAX_COLUMN_LETTERS: usize = 3;

/// Namespace conformance used when authoring an external-link part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
    pub(super) fn sml(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_SML,
            Self::Strict => STRICT_SML,
        }
    }

    pub(super) fn rel(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_REL,
            Self::Strict => STRICT_REL,
        }
    }

    #[must_use]
    pub fn external_link_relationship(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(clippy::large_enum_variant)] // public API; boxing would break existing callers
pub enum Link {
    Workbook(Workbook),
    Dde(Dde),
    Ole(Ole),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dde {
    pub service: String,
    pub topic: String,
    pub items: Vec<DdeItem>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Ole {
    pub target: Target,
    pub program_id: String,
    pub items: Vec<OleItem>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Target {
    pub relationship_id: String,
    pub target: String,
    pub relationship_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DdeItem {
    pub name: Option<String>,
    pub use_ole: bool,
    pub advise: bool,
    pub prefer_picture: bool,
    pub values: Option<DdeValues>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ItemSource {
    SpreadsheetMl,
    Office2010,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleItem {
    pub source: ItemSource,
    pub name: String,
    pub icon: bool,
    pub advise: bool,
    pub prefer_picture: bool,
    pub values: Option<DdeValues>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DdeValues {
    pub rows: u32,
    pub columns: u32,
    pub values: Vec<DdeValue>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeValueType {
    Nil,
    Boolean,
    Number,
    Error,
    String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DdeValue {
    pub value_type: DdeValueType,
    pub raw_value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Workbook {
    pub target: Target,
    pub sheet_names: Vec<String>,
    pub defined_names: Vec<DefinedName>,
    pub cached_sheets: Vec<SheetData>,
    /// Optional service-provided alternate URLs for the external workbook.
    ///
    /// The URLs remain inert relationship metadata.  Loading and editing this
    /// value never opens, fetches, or refreshes an external target.
    pub alternate_urls: Option<AlternateUrls>,
}

impl Workbook {
    /// Borrow the optional inert alternate-URL extension.
    #[must_use]
    pub fn alternate_urls(&self) -> Option<&AlternateUrls> {
        self.alternate_urls.as_ref()
    }
}

/// One inert relationship-backed URL from `CT_AlternateUrl`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AlternateUrl {
    pub relationship_id: String,
    pub target: String,
    pub relationship_type: String,
}

/// The bounded `CT_ExternalBookAlternateUrls` semantic projection.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct AlternateUrls {
    pub drive_id: Option<String>,
    pub item_id: Option<String>,
    pub absolute_url: Option<AlternateUrl>,
    pub relative_url: Option<AlternateUrl>,
}

impl AlternateUrls {
    /// Whether this extension carries no service identifiers or URL entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.drive_id.is_none()
            && self.item_id.is_none()
            && self.absolute_url.is_none()
            && self.relative_url.is_none()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedName {
    pub name: String,
    pub refers_to: Option<String>,
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetData {
    pub sheet_id: u32,
    pub refresh_error: bool,
    pub rows: Vec<Row>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Row {
    pub row: u32,
    pub cells: Vec<Cell>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellType {
    Number,
    Boolean,
    Date,
    Error,
    InlineString,
    SharedString,
    String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cell {
    pub reference: Option<String>,
    pub cell_type: CellType,
    pub raw_value: Option<String>,
    pub value_metadata_index: u32,
}
