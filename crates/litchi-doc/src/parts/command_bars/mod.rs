//! Bounded, inert DOC command-bar metadata.
//!
//! This owner covers the `Tcg` table addressed by `fcCmds`/`lcbCmds`. It
//! exposes fixed-size macro, allocated-command, and key-map records, typed
//! `TcgSttbf`/`MacroNames` tables, plus bounded CTB/TBC metadata and lossless
//! CTBWRAPPER snapshots. Unknown Tcg records remain explicitly unsupported.

mod codec;
mod family;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_bytes, to_bytes, to_control_bytes};
pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use model::{
    Action, AllocatedCommand, AllocatedCommands, CommandBars, CommandId, CommandString,
    CommandStrings, Control, Customization, CustomizationData, Entry, KeyMap, KeyMapKind, KeyMaps,
    MacroCommand, MacroCommands, MacroName, MacroNames, Operation, Toolbar, ToolbarDelta,
    ToolbarWrapper, XString,
};
pub use package::{FIB_INDEX_CMDS, parse, write};

/// The part-level command-bar codec: the semantic [`record::Snapshot`] and its
/// [`record::Edit`] staging, which publishes a new semantic snapshot; package
/// edits go through [`Edit`].
pub mod record {
    pub use super::transaction::{Commit, Diagnostics, Edit, Error, Patch, Snapshot};
}
