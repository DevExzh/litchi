//! Typed XML `NCName` and `QName` vocabulary shared by OOXML parts.
//!
//! The lexical grammar is the XML 1.0 Fifth Edition Name/NCName grammar used
//! by XML Schema `QName`. It is independent of any host vocabulary: namespace
//! resolution belongs to the owner that has the in-scope declarations, while
//! this module owns only the lexical value.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, write};
pub use model::{NameError, NcName, QualifiedName};
#[allow(
    clippy::module_name_repetitions,
    reason = "the public predicate names explicitly identify their XML grammar"
)]
pub use validation::{is_ncname, is_qualified_name, is_xml_name};
