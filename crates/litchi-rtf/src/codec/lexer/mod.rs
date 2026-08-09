//! RTF lexer owner.
//!
//! RTF control-word tokenization lives in [`codec`], while focused lexer
//! regression coverage lives in [`tests`].

mod codec;

#[cfg(test)]
mod tests;

pub(crate) use codec::{ControlWord, Lexer, Token};
