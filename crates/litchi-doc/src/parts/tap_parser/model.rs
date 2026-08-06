//! Layered semantic TAP parsing and table-property state transitions.

mod cells;
mod conditional;
mod parser;
mod prelude;
mod semantic;
mod validation;

#[derive(Debug, Clone, Copy)]
pub(super) enum CellBoolProperty {
    FitText,
    NoWrap,
    HideMark,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum WidthUsage {
    Table,
    TablePart,
    Indent,
}
