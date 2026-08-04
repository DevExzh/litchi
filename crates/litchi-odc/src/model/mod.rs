//! Immutable semantic values for this document family.

/// Axis dimension in a chart plot area.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Dimension {
    X,
    Y,
    Z,
}
/// A semantic chart series selector.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Series {
    range: String,
}
impl Series {
    pub fn new(range: impl Into<String>) -> Self {
        Self {
            range: range.into(),
        }
    }
    pub fn range(&self) -> &str {
        &self.range
    }
}
/// A chart legend position.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LegendPosition {
    Start,
    End,
    Top,
    Bottom,
    TopStart,
    TopEnd,
    BottomStart,
    BottomEnd,
}
