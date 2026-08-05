//! Layered XLSB formula codec facade.
///
/// The wire record codecs and token parser are separated from semantic text
/// compilation and shared grammar validation. The public formula API remains
/// exposed by `formula/mod.rs`.
mod parser;
mod semantic;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

pub use parser::Parser;
pub use semantic::{Compiler, Resolution};
