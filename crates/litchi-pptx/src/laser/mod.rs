//! Layered, inert PowerPoint laser-trace support.
//!
//! The semantic point/trace values live in [`model`], XML scanning and
//! serialization live in [`codec`], and OPC slide mutation lives in [`package`].
//! No layer replays or interprets slide-show input.

mod codec;
mod model;
mod package;

pub use codec::{
    LASER_TRACE_EXTENSION_URI, read, read_with, validate, write, write_to,
};
pub use model::{Conformance, Limits, Trace, TracePoint};
pub use package::{load_slide_traces, store_slide_trace};
