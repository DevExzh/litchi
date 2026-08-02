//! Typed cell-border controls shared by native iWork tables.

use crate::shapes::ShapeStroke;

/// One edge of a zero-based native iWork table cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TableCellBorderSide {
    Left,
    Right,
    Top,
    Bottom,
}

/// Effective explicit borders stored for one native iWork table cell.
///
/// `None` means the table style supplies the edge, or a later native stroke
/// run explicitly clears it.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct TableCellBorders {
    pub left: Option<ShapeStroke>,
    pub right: Option<ShapeStroke>,
    pub top: Option<ShapeStroke>,
    pub bottom: Option<ShapeStroke>,
}

impl TableCellBorders {
    pub const fn get(self, side: TableCellBorderSide) -> Option<ShapeStroke> {
        match side {
            TableCellBorderSide::Left => self.left,
            TableCellBorderSide::Right => self.right,
            TableCellBorderSide::Top => self.top,
            TableCellBorderSide::Bottom => self.bottom,
        }
    }

    pub(crate) fn set(&mut self, side: TableCellBorderSide, stroke: Option<ShapeStroke>) {
        match side {
            TableCellBorderSide::Left => self.left = stroke,
            TableCellBorderSide::Right => self.right = stroke,
            TableCellBorderSide::Top => self.top = stroke,
            TableCellBorderSide::Bottom => self.bottom = stroke,
        }
    }
}
