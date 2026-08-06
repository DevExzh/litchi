//! Shared rich-text value models for the iWork format crates.
//!
//! This crate intentionally has no archive, protobuf, or application-format
//! dependencies. It owns only the allocation-bearing text values that Pages,
//! Numbers, and Keynote can exchange without importing one another's package
//! semantics.

#![forbid(unsafe_code)]

pub mod character;
pub mod columns;
pub mod font;
pub mod language;
pub mod paragraph;
pub mod position;
pub mod storage;

pub use character::{
    Error as CharacterError, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextDecorations, TextLigatures, TextPointSize, TextScript, TextStrikethrough, TextStyle,
    TextUnderline,
};
pub use font::{Font, Name, NameError};
pub use language::{Error as LanguageError, TextLanguage, TextLanguageRun, TextLanguageTag};
pub use position::{Error as PositionError, TextPosition, TextRange};
