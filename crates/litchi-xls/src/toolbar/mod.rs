//! Bounded, inert Office Toolbars (`XCB`) metadata for BIFF workbooks.
//!
//! The facade models the `[MS-XLS]` `CTBWRAPPER`, `CTBS`, and `CTB` records
//! while reusing `[MS-OSHARED]` toolbar headers, names, and flags from
//! `litchi-ole-common`.  Controls whose `TBCData` is variable-length are
//! rejected before a boundary is guessed; no macro, command, or UI behavior
//! is ever executed.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, to_bytes};
pub use model::{
    APPLICATION_TOOLBAR_ID, Control, Toolbar, ToolbarSet, VISUAL_DATA_LEN, VisualData, Wrapper,
};
