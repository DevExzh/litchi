//! Typed budgets of the chart family (record 0355a).

use super::model::{
    MAX_ATTRIBUTE_BYTES, MAX_ATTRIBUTES, MAX_CHART_XML, MAX_CHARTS, MAX_COMPANION_XML,
    MAX_COMPANIONS, MAX_DEPTH, MAX_DOCUMENT_XML, MAX_NODES, MAX_TOTAL_BYTES, MAX_WORKBOOK_BYTES,
};

/// Caller-selected admission limits for the raw chart graph codec.
///
/// The format ceilings remain hard invariants. A raw caller may select a
/// tighter finite policy, but the codec never widens it while reading or
/// writing a graph. The owner family uses `Limits::default()` after package
/// admission has already selected its semantic policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    document_xml_bytes: usize,
    chart_xml_bytes: usize,
    companion_xml_bytes: usize,
    workbook_bytes: usize,
    aggregate_bytes: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
    charts: usize,
    companions: usize,
}

impl Limits {
    /// The standard bounded raw chart policy.
    pub const DEFAULT: Self = Self {
        document_xml_bytes: MAX_DOCUMENT_XML,
        chart_xml_bytes: MAX_CHART_XML,
        companion_xml_bytes: MAX_COMPANION_XML,
        workbook_bytes: MAX_WORKBOOK_BYTES,
        aggregate_bytes: MAX_TOTAL_BYTES,
        depth: MAX_DEPTH,
        nodes: MAX_NODES,
        attributes: MAX_ATTRIBUTES,
        attribute_bytes: MAX_ATTRIBUTE_BYTES,
        charts: MAX_CHARTS,
        companions: MAX_COMPANIONS,
    };

    /// Construct a finite raw chart policy.
    #[must_use]
    pub const fn new(
        document_xml_bytes: usize,
        chart_xml_bytes: usize,
        companion_xml_bytes: usize,
        workbook_bytes: usize,
        aggregate_bytes: usize,
        depth: usize,
        nodes: usize,
        attributes: usize,
        attribute_bytes: usize,
        charts: usize,
        companions: usize,
    ) -> Option<Self> {
        if document_xml_bytes == 0
            || document_xml_bytes > MAX_DOCUMENT_XML
            || chart_xml_bytes == 0
            || chart_xml_bytes > MAX_CHART_XML
            || companion_xml_bytes == 0
            || companion_xml_bytes > MAX_COMPANION_XML
            || workbook_bytes == 0
            || workbook_bytes > MAX_WORKBOOK_BYTES
            || aggregate_bytes == 0
            || aggregate_bytes > MAX_TOTAL_BYTES
            || depth == 0
            || depth > MAX_DEPTH
            || nodes == 0
            || nodes > MAX_NODES
            || attributes == 0
            || attributes > MAX_ATTRIBUTES
            || attribute_bytes == 0
            || attribute_bytes > MAX_ATTRIBUTE_BYTES
            || charts == 0
            || charts > MAX_CHARTS
            || companions == 0
            || companions > MAX_COMPANIONS
        {
            None
        } else {
            Some(Self {
                document_xml_bytes,
                chart_xml_bytes,
                companion_xml_bytes,
                workbook_bytes,
                aggregate_bytes,
                depth,
                nodes,
                attributes,
                attribute_bytes,
                charts,
                companions,
            })
        }
    }

    #[must_use]
    pub const fn document_xml_bytes(self) -> usize {
        self.document_xml_bytes
    }
    #[must_use]
    pub const fn chart_xml_bytes(self) -> usize {
        self.chart_xml_bytes
    }
    #[must_use]
    pub const fn companion_xml_bytes(self) -> usize {
        self.companion_xml_bytes
    }
    #[must_use]
    pub const fn workbook_bytes(self) -> usize {
        self.workbook_bytes
    }
    #[must_use]
    pub const fn aggregate_bytes(self) -> usize {
        self.aggregate_bytes
    }
    #[must_use]
    pub const fn depth(self) -> usize {
        self.depth
    }
    #[must_use]
    pub const fn nodes(self) -> usize {
        self.nodes
    }
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }
    #[must_use]
    pub const fn attribute_bytes(self) -> usize {
        self.attribute_bytes
    }
    #[must_use]
    pub const fn charts(self) -> usize {
        self.charts
    }
    #[must_use]
    pub const fn companions(self) -> usize {
        self.companions
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The bounded resources of the chart family, carried by [`crate::Error::ChartLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Aggregate bytes.
    AggregateBytes,
    /// Chart bytes.
    Bytes,
    /// Chart-color-style count.
    ChartColorStyleCount,
    /// Chart-style count.
    ChartStyleCount,
    /// Chart companion bytes.
    CompanionBytes,
    /// Chart companion count.
    CompanionCount,
    /// Companion XML bytes.
    CompanionXmlBytes,
    /// Chart count.
    Count,
    /// Document XML bytes.
    DocumentXmlBytes,
    /// Embedded workbook bytes.
    EmbeddedWorkbookBytes,
    /// Processed chart XML bytes.
    ProcessedChartXmlBytes,
    /// Processed companion XML bytes.
    ProcessedCompanionXmlBytes,
    /// Processed document XML bytes.
    ProcessedDocumentXmlBytes,
    /// Chart relationship count.
    RelationshipCount,
    /// XML attribute bytes.
    XmlAttributeBytes,
    /// XML attribute count.
    XmlAttributeCount,
    /// Chart XML bytes.
    XmlBytes,
    /// XML structure.
    XmlStructure,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AggregateBytes => "aggregate bytes",
            Self::Bytes => "chart bytes",
            Self::ChartColorStyleCount => "chart-color-style count",
            Self::ChartStyleCount => "chart-style count",
            Self::CompanionBytes => "chart companion bytes",
            Self::CompanionCount => "chart companion count",
            Self::CompanionXmlBytes => "companion XML bytes",
            Self::Count => "chart count",
            Self::DocumentXmlBytes => "document XML bytes",
            Self::EmbeddedWorkbookBytes => "embedded workbook bytes",
            Self::ProcessedChartXmlBytes => "processed chart XML bytes",
            Self::ProcessedCompanionXmlBytes => "processed companion XML bytes",
            Self::ProcessedDocumentXmlBytes => "processed document XML bytes",
            Self::RelationshipCount => "chart relationship count",
            Self::XmlAttributeBytes => "XML attribute bytes",
            Self::XmlAttributeCount => "XML attribute count",
            Self::XmlBytes => "chart XML bytes",
            Self::XmlStructure => "XML structure",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
