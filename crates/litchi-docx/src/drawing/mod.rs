//! Typed inventory of `WordprocessingML` drawing anchors.
//!
//! The owner is intentionally small: [`model`] contains the owned semantic
//! inventory, [`codec`] contains the streaming `<w:drawing>`, `w:object`, and
//! `w:pict` scanners, and [`validation`] owns checked Word 2010 extension
//! values. `DrawingML` preset tokens remain the closed domain supplied by
//! [`litchi_drawingml::geom::Preset`]; unknown drawing children are inert.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Anchor, AnchorId, Kind, LegacyAnchor, LegacyAnchorKind, Object};

pub(crate) use codec::{parse, parse_legacy};
pub(crate) use validation::append_word2010_anchor_id;
