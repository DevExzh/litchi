//! Portable font metadata, embedding, OOXML obfuscation, and optional native
//! discovery or OpenType subsetting.
//!
//! The crate keeps format-independent font ownership in semantic modules:
//! [`embedding`] prepares publishable font data from a caller-supplied
//! resolver, [`obfuscation`] implements the OOXML XOR transform, and the
//! optional [`discovery`] and [`subset`] modules provide native backends.

#![forbid(unsafe_code)]

#[cfg(feature = "discovery")]
pub mod discovery;
pub mod embedding;
pub mod model;
pub mod obfuscation;
#[cfg(feature = "subset")]
pub mod subset;

#[cfg(feature = "discovery")]
pub use discovery::Loader;
#[cfg(feature = "automatic")]
pub use embedding::prepare;
pub use embedding::{Prepared, Resolver, prepare_with};
pub use model::{
    Charset, CollectGlyphs, Family, FontData, FontError, FontProperties, GlyphMap, Glyphs, License,
    LicenseError, Panose, Permission, Pitch, Request, Restrictions, Signature, Style,
};
#[cfg(feature = "subset")]
pub use subset::{Allsorts, Subsetter, glyph_ids};
