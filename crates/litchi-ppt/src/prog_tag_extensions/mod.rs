//! Typed inner payloads of versioned document/slide binary programmable tags.
//!
//! This module implements the extension record grammars referenced by
//! [`super::prog_tags`]: `PP9DocBinaryTagExtension` through
//! `PP12DocBinaryTagExtension` (MS-PPT sections 2.4.23.5 through 2.4.23.8) and
//! `PP9SlideBinaryTagExtension`, `PP10SlideBinaryTagExtension`, and
//! `PP12SlideBinaryTagExtension` (sections 2.5.23, 2.5.24, and 2.5.34).
//!
//! Every grammar slot retains its raw [`crate::Record`]: parsing is strictly
//! ordered per the spec but completely inert, and serialization is byte-exact.
//! Deeper field decoding deliberately stays with the dedicated piecemeal
//! readers (`kinsoku.rs`, `broadcast.rs`, `html_publish.rs`,
//! `presentation_advisor.rs`, `envelope_data.rs`, and friends), which consume
//! the same records through `Record::versioned_binary_tag_records`; this
//! owner only assigns each record to its grammar slot.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{
    DocBinaryTagExtension, DocBinaryTagExtension9, DocBinaryTagExtension10,
    DocBinaryTagExtension11, DocBinaryTagExtension12, DocumentTagExtensions,
    SlideBinaryTagExtension, SlideBinaryTagExtension9, SlideBinaryTagExtension10,
    SlideBinaryTagExtension12, SlideTagExtensions,
};
