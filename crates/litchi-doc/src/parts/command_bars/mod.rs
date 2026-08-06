//! Bounded, inert DOC command-bar metadata.
//!
//! This owner covers the `Tcg` table addressed by `fcCmds`/`lcbCmds`. It
//! exposes fixed-size macro, allocated-command, and key-map records, typed
//! `TcgSttbf`/`MacroNames` tables, plus bounded CTB/TBC metadata and lossless
//! CTBWRAPPER snapshots. Unknown Tcg records remain explicitly unsupported.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_bytes, to_bytes, to_control_bytes};
pub use model::{
    Action, AllocatedCommand, AllocatedCommands, CommandBars, CommandId, CommandString,
    CommandStrings, Control, Customization, CustomizationData, Entry, KeyMap, KeyMapKind, KeyMaps,
    MacroCommand, MacroCommands, MacroName, MacroNames, Operation, Toolbar, ToolbarDelta,
    ToolbarWrapper, XString,
};
pub use package::{Editor, FIB_INDEX_CMDS, PackageCommit, PackageSnapshot, parse, write};
pub use transaction::{Commit, Snapshot, Transaction, TransactionError};
