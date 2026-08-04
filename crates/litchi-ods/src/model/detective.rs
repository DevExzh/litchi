//! Inert spreadsheet formula-auditing ("detective") metadata.

use litchi_core::{Error, Result, xml::escape_xml};

use super::structure::validate_cell_range_addresses;

/// Direction of a highlighted formula dependency range.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DetectiveDirection {
    /// Dependency originates in another table.
    FromAnotherTable,
    /// Dependent cell is in another table.
    ToAnotherTable,
    /// Dependency originates in the same table.
    FromSameTable,
}

impl DetectiveDirection {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "from-another-table" => Ok(Self::FromAnotherTable),
            "to-another-table" => Ok(Self::ToAnotherTable),
            "from-same-table" => Ok(Self::FromSameTable),
            _ => Err(Error::InvalidFormat(format!(
                "invalid detective range direction '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::FromAnotherTable => "from-another-table",
            Self::ToAnotherTable => "to-another-table",
            Self::FromSameTable => "from-same-table",
        }
    }
}

/// Formula-auditing command preserved on a cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DetectiveOperationKind {
    /// Show cells that depend on this cell.
    TraceDependents,
    /// Remove displayed dependent arrows.
    RemoveDependents,
    /// Show cells referenced by this cell.
    TracePrecedents,
    /// Remove displayed precedent arrows.
    RemovePrecedents,
    /// Show the source of an error.
    TraceErrors,
}

impl DetectiveOperationKind {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "trace-dependents" => Ok(Self::TraceDependents),
            "remove-dependents" => Ok(Self::RemoveDependents),
            "trace-precedents" => Ok(Self::TracePrecedents),
            "remove-precedents" => Ok(Self::RemovePrecedents),
            "trace-errors" => Ok(Self::TraceErrors),
            _ => Err(Error::InvalidFormat(format!(
                "invalid detective operation '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::TraceDependents => "trace-dependents",
            Self::RemoveDependents => "remove-dependents",
            Self::TracePrecedents => "trace-precedents",
            Self::RemovePrecedents => "remove-precedents",
            Self::TraceErrors => "trace-errors",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum DetectiveHighlightedRangeKind {
    Valid {
        cell_range_address: Option<String>,
        direction: DetectiveDirection,
        contains_error: Option<bool>,
    },
    Invalid {
        marked_invalid: bool,
    },
}

/// One highlighted formula-auditing range.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DetectiveHighlightedRange {
    kind: DetectiveHighlightedRangeKind,
}

impl DetectiveHighlightedRange {
    /// Create and validate a directional dependency range.
    pub fn valid(
        cell_range_address: Option<String>,
        direction: DetectiveDirection,
        contains_error: Option<bool>,
    ) -> Result<Self> {
        if let Some(address) = &cell_range_address {
            validate_cell_range_addresses(std::slice::from_ref(address))?;
        }
        Ok(Self {
            kind: DetectiveHighlightedRangeKind::Valid {
                cell_range_address,
                direction,
                contains_error,
            },
        })
    }

    /// Create an invalid-range marker.
    pub fn invalid(marked_invalid: bool) -> Self {
        Self {
            kind: DetectiveHighlightedRangeKind::Invalid { marked_invalid },
        }
    }

    /// Optional ODF range address for a valid directional highlight.
    pub fn cell_range_address(&self) -> Option<&str> {
        match &self.kind {
            DetectiveHighlightedRangeKind::Valid {
                cell_range_address, ..
            } => cell_range_address.as_deref(),
            DetectiveHighlightedRangeKind::Invalid { .. } => None,
        }
    }

    /// Dependency direction, or `None` for an invalid-range marker.
    pub fn direction(&self) -> Option<DetectiveDirection> {
        match self.kind {
            DetectiveHighlightedRangeKind::Valid { direction, .. } => Some(direction),
            DetectiveHighlightedRangeKind::Invalid { .. } => None,
        }
    }

    /// Preserved error flag for a valid directional highlight.
    pub fn contains_error(&self) -> Option<bool> {
        match self.kind {
            DetectiveHighlightedRangeKind::Valid { contains_error, .. } => contains_error,
            DetectiveHighlightedRangeKind::Invalid { .. } => None,
        }
    }

    /// Preserved invalid-marker value, if this is an invalid-range marker.
    pub fn marked_invalid(&self) -> Option<bool> {
        match self.kind {
            DetectiveHighlightedRangeKind::Invalid { marked_invalid } => Some(marked_invalid),
            DetectiveHighlightedRangeKind::Valid { .. } => None,
        }
    }
}

/// One ordered formula-auditing operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DetectiveOperation {
    /// Requested auditing command.
    pub kind: DetectiveOperationKind,
    /// Non-negative sequence index from the ODF document.
    pub index: usize,
}

impl DetectiveOperation {
    /// Create an auditing operation.
    pub fn new(kind: DetectiveOperationKind, index: usize) -> Self {
        Self { kind, index }
    }
}

/// Formula-auditing highlights and commands attached to one cell.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CellDetective {
    highlighted_ranges: Vec<DetectiveHighlightedRange>,
    operations: Vec<DetectiveOperation>,
}

impl CellDetective {
    /// Create empty detective metadata.
    pub fn new() -> Self {
        Self::default()
    }

    /// Highlighted ranges, in document order.
    pub fn highlighted_ranges(&self) -> &[DetectiveHighlightedRange] {
        &self.highlighted_ranges
    }

    /// Auditing operations, in document order.
    pub fn operations(&self) -> &[DetectiveOperation] {
        &self.operations
    }

    /// Add a highlighted dependency range.
    pub fn add_highlighted_range(&mut self, range: DetectiveHighlightedRange) -> &mut Self {
        self.highlighted_ranges.push(range);
        self
    }

    /// Add an auditing operation.
    pub fn add_operation(&mut self, operation: DetectiveOperation) -> &mut Self {
        self.operations.push(operation);
        self
    }

    /// Whether this container has no ranges or operations.
    pub fn is_empty(&self) -> bool {
        self.highlighted_ranges.is_empty() && self.operations.is_empty()
    }
}

pub(crate) fn write_detective(out: &mut String, detective: &CellDetective) {
    out.push_str("<table:detective>");
    for range in detective.highlighted_ranges() {
        out.push_str("<table:highlighted-range");
        match &range.kind {
            DetectiveHighlightedRangeKind::Valid {
                cell_range_address,
                direction,
                contains_error,
            } => {
                if let Some(address) = cell_range_address {
                    write_attribute(out, "table:cell-range-address", address);
                }
                write_attribute(out, "table:direction", direction.as_str());
                if let Some(value) = contains_error {
                    write_attribute(
                        out,
                        "table:contains-error",
                        if *value { "true" } else { "false" },
                    );
                }
            },
            DetectiveHighlightedRangeKind::Invalid { marked_invalid } => write_attribute(
                out,
                "table:marked-invalid",
                if *marked_invalid { "true" } else { "false" },
            ),
        }
        out.push_str("/>");
    }
    for operation in detective.operations() {
        out.push_str("<table:operation");
        write_attribute(out, "table:name", operation.kind.as_str());
        write_attribute(out, "table:index", &operation.index.to_string());
        out.push_str("/>");
    }
    out.push_str("</table:detective>");
}

fn write_attribute(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape_xml(value));
    out.push('"');
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_ranges_before_operations_and_escapes_addresses() {
        let mut detective = CellDetective::new();
        detective
            .add_highlighted_range(
                DetectiveHighlightedRange::valid(
                    Some("'A&B'.$A$1:$B$2".to_string()),
                    DetectiveDirection::FromAnotherTable,
                    Some(true),
                )
                .unwrap(),
            )
            .add_highlighted_range(DetectiveHighlightedRange::invalid(true))
            .add_operation(DetectiveOperation::new(
                DetectiveOperationKind::TraceErrors,
                7,
            ));
        let mut xml = String::new();
        write_detective(&mut xml, &detective);
        assert!(xml.contains("&apos;A&amp;B&apos;.$A$1:$B$2"));
        assert!(xml.find("highlighted-range").unwrap() < xml.find("operation").unwrap());
    }

    #[test]
    fn rejects_multiple_addresses_in_one_highlight() {
        assert!(
            DetectiveHighlightedRange::valid(
                Some(".A1:.B2 .C1:.D2".to_string()),
                DetectiveDirection::FromSameTable,
                None,
            )
            .is_err()
        );
    }
}
