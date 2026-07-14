//! Protobuf message support for iWork IWA files.

// Generated protobuf definitions are kept at this module's public root for
// compatibility with paths such as `protobuf::tsp::Reference`.
include!(concat!(env!("OUT_DIR"), "/iwa_protos.rs"));

mod decoder;
mod wrappers;

pub use decoder::decode;
pub use wrappers::*;
