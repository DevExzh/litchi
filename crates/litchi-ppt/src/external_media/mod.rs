//! Typed, inert `PowerPoint` external-media metadata.
//!
//! The model is kept separate from the record codec so callers work with
//! contextual media objects while the binary layer remains strict and
//! lossless.  Paths are metadata only: this module never resolves or opens
//! them.

mod codec;
mod edit;
mod model;
mod package;
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use edit::Edit;
pub use model::{
    CdAudio, CdTime, Collection, EmbeddedWav, Limits, LinkedAudio, LinkedAudioKind, Media, Movie,
    MovieKind, Object, Playback, UnknownRecord, Video,
};
pub use transaction::{Change, Revision, Snapshot};

crate::scoped::family!("external-media");
