//! Bounded, inert DOC command-bar metadata.
//!
//! This owner covers the `Tcg` table addressed by `fcCmds`/`lcbCmds`. It
//! exposes fixed-size macro, allocated-command, and key-map records, plus
//! bounded typed CTB/TBC metadata and lossless CTBWRAPPER snapshots. Macro
//! string-table records and unknown Tcg records remain explicitly unsupported.

mod codec;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_bytes, to_bytes, to_control_bytes};
pub use model::{
    Action, AllocatedCommand, AllocatedCommands, CommandBars, CommandId, Control, Customization,
    CustomizationData, Entry, KeyMap, KeyMapKind, KeyMaps, MacroCommand, MacroCommands, Operation,
    Toolbar, ToolbarDelta, ToolbarWrapper, XString,
};
pub use package::{FIB_INDEX_CMDS, parse, write};
