//! Bounded, inert DOC command-bar metadata.
//!
//! This owner covers the `Tcg` table addressed by `fcCmds`/`lcbCmds`. It
//! exposes fixed-size macro, allocated-command, and key-map records, plus a
//! lossless CTBWRAPPER shell. Variable TBCData, macro-name string tables, and
//! unknown Tcg records remain explicitly unsupported because their boundaries
//! cannot be recovered safely without interpreting UI or macro records.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{FIB_INDEX_CMDS, parse, parse_bytes, to_bytes, write};
pub use model::{
    Action, AllocatedCommand, AllocatedCommands, CommandBars, Customization, CustomizationData,
    Entry, KeyMap, KeyMapKind, KeyMaps, MacroCommand, MacroCommands, Operation, Toolbar,
    ToolbarDelta, ToolbarWrapper, XString,
};
