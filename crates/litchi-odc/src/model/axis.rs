//! Chart axis semantics.

/// Axis dimension in a chart plot area.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Dimension {
    X,
    Y,
    Z,
}
