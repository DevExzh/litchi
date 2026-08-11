//! ODF scalar data types and their lossless, checked codecs.
//!
//! The semantic values live in the private `model` module, while the
//! format-bound parsing and encoding rules live in the private `codec` module.
//! The separate [`lexical`] module contains
//! small contextual validators used by other ODF owners.
//!
//! Based on the reference implementation in
//! `3rdparty/odfdo/src/odfdo/datatype.py`.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub mod lexical;

pub use model::{Boolean, Date, DateTime, Duration, DurationValue};
