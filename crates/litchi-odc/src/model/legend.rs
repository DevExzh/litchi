//! Chart legend placement.

/// A chart legend position.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
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
