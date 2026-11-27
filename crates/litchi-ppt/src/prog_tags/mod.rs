//! Document- and slide-level programmable tags.
//!
//! This module implements the `DocProgTagsContainer` family (MS-PPT sections
//! 2.4.23.1 through 2.4.23.4) and the `SlideProgTagsContainer` family (MS-PPT
//! sections 2.5.19 through 2.5.22), together with the shared
//! `ProgStringTagContainer`/`TagNameAtom`/`TagValueAtom` records (sections
//! 2.11.30 through 2.11.32) and `UnknownBinaryTag` (section 2.11.33).
//!
//! Parsing is inert: versioned binary payloads are validated as strict record
//! sequences and retained byte-for-byte, but they are never executed, loaded,
//! or resolved. Shape-scoped programmable tags (sections 2.7.14 through
//! 2.7.20) live in [`super::shape_programmable_tags`]; typed decoding of the
//! versioned extension payloads lives in [`super::prog_tag_extensions`].

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{
    ProgBinaryTag, ProgBinaryTagVersion, ProgStringTag, ProgTag, ProgTagLimits, ProgTagScope,
    ProgTags,
};
