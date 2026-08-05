//! Page-border semantics for one Word section.
//!
//! The public model is kept independent from the section-table parser and
//! writer. The private codec contains the Word 97 `Brc80` and `SPgbProp`
//! wire mappings used by both sides of the DOC implementation.

mod codec;
mod model;

pub(crate) use codec::{decode_brc80, decode_pgb_prop, encode_sepx};
pub use model::{ApplyTo, Art, Border, Borders, Color, Depth, Error, Offset, Style};
