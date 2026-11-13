//! Inert [MS-OLEPS] alternate-stream binding metadata.
//!
//! This module models the fixed control stream packet and the `Docf_` name
//! used for non-simple property-set alternate streams.  These names belong to
//! a filesystem alternate-stream binding; they are not CFB directory paths.
//! Class identifiers and application state remain opaque values.  Nothing in
//! this module opens, resolves, or activates an OLE class or payload.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{
    ALTERNATE_STREAM_CONTROL_NAME, AlternateStreamControl, NonSimpleAlternateStreamName,
};
