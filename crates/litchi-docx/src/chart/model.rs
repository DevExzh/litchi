//! Semantic DOCX chart-host snapshots.

use litchi_opc::constants::relationship_type as rt;

pub(crate) const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(crate) const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
pub(crate) const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(crate) const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(crate) const CS: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
pub(crate) const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const DOCUMENT_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
pub(crate) const CHART_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub(crate) const STYLE_CT: &str = "application/vnd.ms-office.chartstyle+xml";
pub(crate) const COLOR_STYLE_CT: &str = "application/vnd.ms-office.chartcolorstyle+xml";
pub(crate) const WORKBOOK_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
pub(crate) const STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
pub(crate) const COLOR_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
pub(crate) const MAX_DOCUMENT_XML: usize = 32 * 1024 * 1024;
pub(crate) const MAX_CHART_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_COMPANION_XML: usize = 4 * 1024 * 1024;
pub(crate) const MAX_WORKBOOK_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_TOTAL_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_CHARTS: usize = 256;
pub(crate) const MAX_COMPANIONS: usize = 64;
pub(crate) const MAX_RELATIONSHIPS: usize = 130;
pub(crate) const MAX_NODES: usize = 200_000;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_ATTRIBUTES: usize = 750_000;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = 16 * 1024 * 1024;

/// The OOXML conformance family used by the document and its chart host.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn w(self) -> &'static str {
        if self == Self::Strict { WS } else { W }
    }

    pub(crate) fn a(self) -> &'static str {
        if self == Self::Strict { AS } else { A }
    }

    pub(crate) fn c(self) -> &'static str {
        if self == Self::Strict { CS } else { C }
    }

    pub(crate) fn r(self) -> &'static str {
        if self == Self::Strict { RS } else { R }
    }

    pub(crate) fn chart_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_CHART
        } else {
            rt::CHART
        }
    }

    pub(crate) fn package_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_PACKAGE
        } else {
            rt::PACKAGE
        }
    }
}

/// Content type of an embedded workbook owned by a DOCX chart.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmbeddedWorkbookContentType {
    Xlsx,
}

impl EmbeddedWorkbookContentType {
    pub fn as_str(self) -> &'static str {
        WORKBOOK_CT
    }

    pub(crate) fn parse(value: &str) -> Option<Self> {
        (value == WORKBOOK_CT).then_some(Self::Xlsx)
    }

    pub(crate) fn validates_path(self, value: &str) -> bool {
        value.to_ascii_lowercase().ends_with(".xlsx")
    }
}

/// Opaque chart style or color-style companion part.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Companion {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
}

/// Opaque workbook embedded as a chart data source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedWorkbook {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: EmbeddedWorkbookContentType,
    pub data: Vec<u8>,
}

/// A chart part and all DOCX-owned companion resources.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    pub document_relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Original chart XML bytes, retained losslessly for unknown markup.
    pub data: Vec<u8>,
    pub styles: Vec<Companion>,
    pub color_styles: Vec<Companion>,
    pub workbook: Option<EmbeddedWorkbook>,
}

/// Immutable snapshot of the chart relationship subgraph owned by a DOCX.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Graph {
    pub conformance: Conformance,
    pub charts: Vec<Resource>,
}
