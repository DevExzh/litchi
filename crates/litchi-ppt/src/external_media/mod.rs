//! Typed, inert `PowerPoint` external-media metadata.
//!
//! The model is kept separate from the record codec so callers work with
//! contextual media objects while the binary layer remains strict and
//! lossless.  Paths are metadata only: this module never resolves or opens
//! them.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    CdAudio, CdTime, Collection, EmbeddedWav, Limits, LinkedAudio, LinkedAudioKind, Media, Movie,
    MovieKind, Object, Playback, UnknownRecord, Video,
};
pub use transaction::{Change, Commit, Patch, Revision, Snapshot, Transaction};
