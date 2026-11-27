//! Shared `ChartEx` resource and namespace policy.

pub(super) const CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
pub(super) const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const A_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(super) const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const R_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const PACKAGE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/package",
];
pub(super) const OLE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject",
];
pub(super) const IMAGE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
];
pub(super) const OLE_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.oleObject";
pub(super) const WORKBOOK_CONTENT_TYPES: [&str; 3] = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel.sheet.macroEnabled.12",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
];

pub(super) const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_NODES: usize = 250_000;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_ATTRIBUTES: usize = 64;
pub(super) const MAX_STRING_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_DATA_SETS: usize = 65_536;
pub(super) const MAX_FEATURES: usize = 256;
pub(super) const MAX_LEVELS_PER_DIMENSION: usize = 4096;
pub(super) const MAX_POINTS_PER_LEVEL: u32 = 1_000_000;
pub(super) const MAX_FORMULA_BYTES: usize = 32 * 1024;
pub(super) const MAX_SERIES: usize = 65_536;
pub(super) const MAX_AXES: usize = 4_096;
pub(super) const MAX_AXIS_REFS_PER_SERIES: usize = 64;
pub(super) const MAX_SUBTOTALS: usize = 100_000;
pub(super) const MAX_CULTURE_NAME_LEN: usize = 64;
pub(super) const MAX_ATTRIBUTION_LEN: usize = 4_096;
pub(super) const MAX_GEO_STRING_LEN: usize = 8_192;
pub(super) const MAX_GEO_POLYGON_DATA_LEN: usize = 1024 * 1024;
pub(super) const MAX_GEO_RESULTS: usize = 65_536;
pub(super) const MAX_GEO_CACHE_ENTRIES: usize = 1_024;
pub(super) const MAX_GEO_BINARY_BYTES: usize = 1024 * 1024;
pub(super) const MAX_SERIES_POINTS: usize = 100_000;
pub(super) const MAX_DATA_LABELS: usize = 100_000;
pub(super) const MAX_LABEL_TEXT_BYTES: usize = 32 * 1024;
pub(super) const MAX_FORMAT_OVERRIDES: usize = 65_536;
pub(super) const MAX_PRINT_TEXT_BYTES: usize = 32 * 1024;
