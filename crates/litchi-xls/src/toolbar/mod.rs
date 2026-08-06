//! Bounded, inert Office Toolbars (`XCB`) metadata for BIFF workbooks.
//!
//! The facade models the `[MS-XLS]` `CTBWRAPPER`, `CTBS`, and `CTB` records
//! while reusing `[MS-OSHARED]` toolbar headers, names, and flags from
//! `litchi-ole-common`.  Variable `TBCData` boundaries are resolved only by
//! recursively validating the known control count and the next record
//! signatures; ambiguous or malformed streams are rejected, and no macro,
//! command, or UI behavior is ever executed.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, to_bytes};
pub use model::{
    APPLICATION_TOOLBAR_ID, Command, Control, Toolbar, ToolbarSet, VISUAL_DATA_LEN, VisualData,
    Wrapper,
};
