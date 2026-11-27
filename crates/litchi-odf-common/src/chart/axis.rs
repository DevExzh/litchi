//! Axis semantic values.

/// The coordinate dimension used by an ODF chart axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Dimension {
    X,
    Y,
    Z,
}
