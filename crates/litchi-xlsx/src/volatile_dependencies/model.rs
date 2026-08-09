//! Typed semantic models for `SpreadsheetML` volatile dependencies.

pub(super) const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const NS_TEXT: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_NS_TEXT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";
pub(super) const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/volatileDependencies";
pub(super) const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml";
pub(super) const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_TYPES: usize = 64;
pub(super) const MAX_MAINS: usize = 16_384;
pub(super) const MAX_TOPICS: usize = 65_536;
pub(super) const MAX_SUBTOPICS: usize = 262_144;
pub(super) const MAX_REFERENCES: usize = 1_048_576;
pub(super) const MAX_TEXT_BYTES: usize = 1_048_576;

/// Namespace family used for the volatile-dependencies XML and workbook relationship.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum VolatileDependenciesConformance {
    #[default]
    Transitional,
    Strict,
}

impl VolatileDependenciesConformance {
    pub(super) const fn relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    /// Whether this conformance uses ISO/IEC 29500 Strict namespace URIs.
    #[must_use]
    pub const fn is_strict(self) -> bool {
        matches!(self, Self::Strict)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VolatileDependencyType {
    RealTimeData,
    OlapFunctions,
}

#[derive(Clone, Debug, PartialEq)]
pub enum VolatileValue {
    Unspecified(String),
    Boolean(bool),
    Number(f64),
    Error(String),
    String(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VolatileReference {
    pub cell_reference: String,
    pub sheet_id: u32,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileTopic {
    pub value: VolatileValue,
    pub subtopics: Vec<String>,
    pub references: Vec<VolatileReference>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileMain {
    pub first: String,
    pub topics: Vec<VolatileTopic>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileType {
    pub dependency_type: VolatileDependencyType,
    pub mains: Vec<VolatileMain>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileDependencies {
    pub types: Vec<VolatileType>,
    /// Raw `extLst` markup. Its payload is preserved, never interpreted or executed.
    pub extension_list_xml: Option<Vec<u8>>,
}
