//! Layered RTF lexer codec.
//!
//! The public lexer facade remains intentionally small. The control-word
//! vocabulary, bounded scanner state, and source-preserving token readers are
//! kept in focused modules so the tokenizer's ownership and safety rules are
//! easy to audit without changing its output model.

mod control_word;
mod lexer;
mod model;
mod scanner;

pub use control_word::ControlWord;
pub use lexer::Lexer;
pub use model::{CharacterSet, Token};
