//! Bounded MathML parsing and canonical XML serialization.

mod parser;
pub(crate) mod serialize;

pub use parser::parse;
pub use serialize::serialize;
