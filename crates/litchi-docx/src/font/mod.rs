#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
//! Typed `WordprocessingML` font tables and inert embedded-font resources.
//!
//! The owner is layered internally as semantic model, bounded XML codec, and
//! OPC relationship graph. The public facade keeps the contextual vocabulary
//! under the font module without compatibility aliases or format prefixes.

pub(crate) mod codec;
mod edit;
mod limit;
mod model;
pub mod open_type;
pub(crate) mod package;
pub mod raw;
#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use limit::Limits;
pub use model::{
    Charset, Conformance, Embed, Family, Font, FontKey, Key, License, Pitch, Resource, Signature,
    Style, Table, deobfuscate, obfuscate,
};
pub use open_type::{
    Ligatures, NumForm, NumSpacing, OnOff, OpenType, Snapshot, StyleSet, StyleSetId,
};
pub use package::validate_usage;
