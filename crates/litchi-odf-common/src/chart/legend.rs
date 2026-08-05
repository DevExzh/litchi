//! Legend semantic values.

/// The standard ODF chart legend placement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Position {
    Start,
    End,
    Top,
    Bottom,
    TopStart,
    TopEnd,
    BottomStart,
    BottomEnd,
}
