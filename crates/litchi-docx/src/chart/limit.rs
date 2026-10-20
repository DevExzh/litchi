//! Typed budgets of the chart family (record 0355a).

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
