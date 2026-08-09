//! Layered `SpreadsheetML` worksheet auto-filter and sort-state owner.
//!
//! Typed criteria live in model, bounded fragment conversion in codec,
//! and worksheet XML extraction in package. Unknown XML is retained as
//! inert bounded data and replayed in its original child position.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{parse_auto_filter_fragment, write_auto_filter_fragment};
pub use model::{
    Calendar, Color, Column, Condition, Custom, Customs, DateGroup, Definition, Dynamic,
    DynamicType, Grouping, Icon, IconSet, Item, Operator, Payload, Range, State, Top10,
    UnknownAttribute, UnknownElement, Values,
};
pub use package::parse_auto_filter;
