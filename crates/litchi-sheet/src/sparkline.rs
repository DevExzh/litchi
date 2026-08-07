//! Shared sparkline enumerations.
//!
//! These types express workbook-independent sparkline semantics. Format wire
//! adapters own formulas, colors, XML/BRT encoding, and lexical preservation.

/// The visual form of a sparkline group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SparklineType {
    /// A line sparkline.
    Line,
    /// A column sparkline.
    Column,
    /// A 100%-stacked (win/loss) sparkline.
    WinLoss,
}

impl SparklineType {
    /// Semantic name used by XLSB and the OOXML specifications.
    #[allow(non_upper_case_globals)]
    pub const Stacked: Self = Self::WinLoss;
}

/// How a sparkline handles empty source cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EmptyCells {
    /// Treat empty cells as zero.
    Zero,
    /// Leave a gap for empty cells.
    Gap,
    /// Span empty cells.
    Span,
}

/// The scale scope for one sparkline axis bound.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisType {
    /// Scale each sparkline independently.
    Individual,
    /// Share a scale across the group.
    Group,
    /// Use a caller-supplied scale.
    Custom,
}

#[cfg(test)]
mod tests {
    use super::{AxisType, EmptyCells, SparklineType};

    #[test]
    fn semantic_variants_are_distinct() {
        assert_eq!(SparklineType::Stacked, SparklineType::WinLoss);
        assert_ne!(SparklineType::Line, SparklineType::Column);
        assert_ne!(EmptyCells::Zero, EmptyCells::Gap);
        assert_ne!(AxisType::Individual, AxisType::Group);
    }
}
