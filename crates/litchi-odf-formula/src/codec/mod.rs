//! Bounded `MathML` parsing and canonical XML serialization.

mod limits;
mod parser;
pub(crate) mod serialize;

pub use limits::{
    HARD_MAX_ATTRIBUTE_BYTES, HARD_MAX_ATTRIBUTES, HARD_MAX_DEPTH, HARD_MAX_NODES,
    HARD_MAX_PACKAGE_BYTES, HARD_MAX_TEXT_BYTES, HARD_MAX_XML_BYTES, LimitError, LimitKind, Limits,
};
pub use parser::{parse, parse_with_limits};
pub use serialize::serialize;
