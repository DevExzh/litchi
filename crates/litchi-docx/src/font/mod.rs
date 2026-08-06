//! Typed WordprocessingML font tables and inert embedded-font resources.
//!
//! The owner is layered internally as semantic model, bounded XML codec, and
//! OPC relationship graph. The public facade keeps the contextual vocabulary
//! under the font module without compatibility aliases or format prefixes.

mod codec;
mod model;
pub mod open_type;
mod package;
#[cfg(test)]
mod tests;

pub use codec::{parse, write};
pub use model::{
    Charset, Conformance, Embed, Family, Font, FontKey, Key, License, Pitch, Resource, Signature,
    Style, Table, deobfuscate, obfuscate, raw,
};
pub use open_type::{
    Commit, Ligatures, NumForm, NumSpacing, OnOff, OpenType, Patch, Snapshot, StyleSet, StyleSetId,
    Transaction,
};
pub use package::{put, read, remove, validate_usage};
