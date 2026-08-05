//! Typed, inert PowerPoint external-media metadata.
//!
//! The model is kept separate from the record codec so callers work with
//! contextual media objects while the binary layer remains strict and
//! lossless.  Paths are metadata only: this module never resolves or opens
//! them.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{
    CdAudio, CdTime, Collection, EmbeddedWav, LinkedAudio, LinkedAudioKind, Media, Movie,
    MovieKind, Object, UnknownRecord, Video,
};
