//! Inert `[MS-OSHARED]` VBA digital-signature serialization.
//!
//! This owner validates the bounded `DigSigBlob`, `WordSigBlob`, and nested
//! `DigSigInfoSerialized` wire structures. Signature and certificate-store
//! payloads remain opaque: this module never verifies trust, executes VBA, or
//! interprets PKCS data. Parsed values retain their exact source allocation,
//! including producer-specific gaps and undefined padding bytes.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Blob, Error, Info, Kind, Limits};
