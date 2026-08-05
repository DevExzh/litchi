use super::control_word::ControlWord;
use std::borrow::Cow;

/// Token types.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Token<'a> {
    /// Opening brace
    OpenBrace,
    /// Closing brace
    CloseBrace,
    /// Control word
    Control(ControlWord<'a>),
    /// Plain text
    Text(Cow<'a, str>),
    /// Exact payload consumed by a `binN` control word
    Binary(Cow<'a, [u8]>),
}

/// Character set encoding for RTF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CharacterSet {
    /// ANSI (Windows-1252 / CP1252)
    #[default]
    Ansi,
    /// Mac (Mac Roman)
    Mac,
    /// PC (DOS / CP437)
    Pc,
    /// PC (DOS / CP850)
    Pca,
}
