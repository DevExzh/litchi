//! RTF parser owner.
//!
//! Token-to-document assembly lives in [`codec`], while focused parser
//! regression coverage lives in [`tests`].

mod codec;

#[cfg(test)]
mod tests;

pub(crate) use codec::{ParsedDocument, Parser};
