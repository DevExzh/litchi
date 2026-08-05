//! Plot-area semantic values.

/// Which table rows or columns are labels for a plot area.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Labels {
    None,
    Row,
    Column,
    Both,
}
