//! Grid semantic values.

/// The standard ODF chart grid class.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Class {
    Major,
    Minor,
}
