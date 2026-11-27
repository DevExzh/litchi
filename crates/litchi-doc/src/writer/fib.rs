//! FIB (File Information Block) generation for DOC files.
//!
//! The facade keeps the historical writer API while the implementation is
//! divided by semantic responsibility:
//!
//! - header owns builder and story state;
//! - flags encodes typed FibBase flags;
//! - offsets owns FibRgFcLcb table references;
//! - codec writes the exact Word 2002 byte layout;
//! - validation guards the fixed layout before encoding.

mod codec;
mod flags;
mod header;
mod offsets;
mod validation;

#[cfg(test)]
mod tests;

/// I/O error returned while generating the FIB.
pub type IoError = std::io::Error;

pub use header::FibBuilder;
