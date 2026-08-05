//! Font discovery, metadata, subsetting, embedding, and OOXML obfuscation.
//!
//! The crate keeps format-independent font ownership in semantic modules:
//! [`discovery`] resolves system faces, [`subset`] maps and reduces glyph
//! programs, [`embedding`] prepares publishable font data, and
//! [`obfuscation`] implements the OOXML XOR transform.

#![forbid(unsafe_code)]

pub mod discovery;
pub mod embedding;
pub mod model;
pub mod obfuscation;
pub mod subset;

pub use discovery::Loader;
pub use embedding::{Prepared, prepare};
pub use model::{
    Charset, CollectGlyphs, Family, FontData, FontError, FontProperties, GlyphMap, Glyphs, License,
    LicenseError, Panose, Permission, Pitch, Request, Restrictions, Signature, Style,
};
pub use subset::{Allsorts, Subsetter, glyph_ids};
