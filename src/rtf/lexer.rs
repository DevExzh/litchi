//! RTF lexer/tokenizer.
//!
//! This module implements a high-performance lexer that tokenizes RTF input
//! using arena allocation for temporary data structures.

use super::error::{RtfError, RtfResult};
use bumpalo::Bump;
use smallvec::SmallVec;
use std::borrow::Cow;

/// Control word with optional parameter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ControlWord<'a> {
    // Document structure
    Rtf(i32),
    Ansi,
    AnsiCodePage(i32),
    Mac,
    Pc,
    Pca,

    // Header groups
    FontTable,
    ColorTable,
    StyleSheet,
    Info,

    // Embedded content
    Picture,
    Object,
    Result,

    // Picture properties
    PictureWidth(i32),
    PictureHeight(i32),
    PictureGoalWidth(i32),
    PictureGoalHeight(i32),
    PictureScaleX(i32),
    PictureScaleY(i32),
    Emfblip,
    Pngblip,
    Jpegblip,
    Macpict,
    Pmmetafile(i32),
    Wmetafile(i32),
    Dibitmap(i32),
    Wbitmap(i32),

    // Field support
    Field,
    FieldInstruction,
    FieldResult,
    FieldLock,
    FieldDirty,
    FieldEdit,
    FieldPrivate,

    // Font properties
    FontNumber(i32),
    FontSize(i32),
    FontCharset(i32),
    FontFamily(&'a str),

    // Colors
    Red(i32),
    Green(i32),
    Blue(i32),
    ColorForeground(i32),
    ColorBackground(i32),

    // Character formatting
    Bold(bool),
    Italic(bool),
    Underline(bool),
    UnderlineNone,
    Strike(bool),
    Superscript(bool),
    Subscript(bool),
    SmallCaps(bool),
    Plain,

    // Paragraph formatting
    Par,
    Pard,
    LeftAlign,
    RightAlign,
    Center,
    Justify,

    // Paragraph spacing and indentation
    SpaceBefore(i32),
    SpaceAfter(i32),
    SpaceBetween(i32),
    LineMultiple(bool),
    LeftIndent(i32),
    RightIndent(i32),
    FirstLineIndent(i32),

    // Tables
    TableRowDefaults,
    TableRow,
    TableCell,
    CellX(i32),
    InTable,

    // Unicode
    Unicode(i32),
    UnicodeSkip(i32),

    // Special
    Tab,
    Line,
    Page,
    Section,
    SectionDefault,

    // Binary data
    Binary(i32),

    // Ignorable destination
    IgnorableDestination,

    // Unknown control word
    Unknown(&'a str, Option<i32>),
}

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
    /// Binary data (skipped for now)
    #[allow(dead_code)]
    Binary(usize),
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

/// RTF Lexer using arena allocation.
pub struct Lexer<'a> {
    /// Source input
    input: &'a str,
    /// Current position in bytes
    pos: usize,
    /// Arena allocator for temporary strings
    arena: &'a Bump,
    /// Current character set
    charset: CharacterSet,
}

impl<'a> Lexer<'a> {
    /// Create a new lexer.
    #[inline]
    pub fn new(input: &'a str, arena: &'a Bump) -> Self {
        Self {
            input,
            pos: 0,
            arena,
            charset: CharacterSet::default(),
        }
    }

    /// Set the character set for hex escape decoding.
    ///
    /// This can be used to properly decode hex escapes based on the document's
    /// declared character set (\ansi, \mac, \pc, \pca).
    #[inline]
    #[allow(dead_code)] // Reserved for future use when charset detection is implemented
    pub fn set_charset(&mut self, charset: CharacterSet) {
        self.charset = charset;
    }

    /// Tokenize the entire input.
    pub fn tokenize(&mut self) -> RtfResult<Vec<Token<'a>>> {
        let mut tokens = Vec::new();

        while self.pos < self.input.len() {
            let token = self.next_token()?;
            tokens.push(token);
        }

        Ok(tokens)
    }

    /// Get the next token.
    fn next_token(&mut self) -> RtfResult<Token<'a>> {
        self.skip_whitespace();

        if self.pos >= self.input.len() {
            return Err(RtfError::UnexpectedEof);
        }

        let ch = self.current_char();
        match ch {
            '{' => {
                self.advance();
                Ok(Token::OpenBrace)
            },
            '}' => {
                self.advance();
                Ok(Token::CloseBrace)
            },
            '\\' => self.parse_control_word(),
            _ => self.parse_text(),
        }
    }

    /// Parse a control word or control symbol.
    fn parse_control_word(&mut self) -> RtfResult<Token<'a>> {
        self.advance(); // Skip '\'

        if self.pos >= self.input.len() {
            return Err(RtfError::UnexpectedEof);
        }

        let ch = self.current_char();

        // Handle special control symbols
        match ch {
            '\\' | '{' | '}' => {
                let text = self.arena.alloc_str(&ch.to_string());
                self.advance();
                return Ok(Token::Text(Cow::Borrowed(text)));
            },
            '\'' => return self.parse_hex_char(),
            '*' => {
                self.advance();
                return Ok(Token::Control(ControlWord::IgnorableDestination));
            },
            '\n' | '\r' => {
                self.advance();
                return Ok(Token::Control(ControlWord::Par));
            },
            '~' => {
                self.advance();
                let text = self.arena.alloc_str("\u{00A0}"); // Non-breaking space
                return Ok(Token::Text(Cow::Borrowed(text)));
            },
            '-' => {
                self.advance();
                let text = self.arena.alloc_str("\u{00AD}"); // Optional hyphen
                return Ok(Token::Text(Cow::Borrowed(text)));
            },
            '_' => {
                self.advance();
                let text = self.arena.alloc_str("\u{2011}"); // Non-breaking hyphen
                return Ok(Token::Text(Cow::Borrowed(text)));
            },
            _ => {},
        }

        // Parse control word
        let start = self.pos;

        // Read alphabetic characters
        while self.pos < self.input.len() && self.current_char().is_ascii_alphabetic() {
            self.advance();
        }

        if start == self.pos {
            // No alphabetic characters, might be a control symbol
            return Err(RtfError::InvalidControlWord(format!(
                "Invalid control word at position {}",
                self.pos
            )));
        }

        let word = &self.input[start..self.pos];

        // Parse optional numeric parameter
        let param = self.parse_numeric_parameter()?;

        // Skip optional space delimiter after control word
        if self.pos < self.input.len() && self.current_char() == ' ' {
            self.advance();
        }

        // Match control word to enum variant
        let control = self.match_control_word(word, param)?;

        // Handle binary data immediately after \bin
        if let ControlWord::Binary(size) = control
            && size > 0
        {
            // Skip the binary data bytes
            let skip_bytes = size as usize;
            if self.pos + skip_bytes <= self.input.len() {
                self.pos += skip_bytes;
            }
            return Ok(Token::Binary(skip_bytes));
        }

        Ok(Token::Control(control))
    }

    /// Parse numeric parameter after control word.
    fn parse_numeric_parameter(&mut self) -> RtfResult<Option<i32>> {
        if self.pos >= self.input.len() {
            return Ok(None);
        }

        let ch = self.current_char();
        if !ch.is_ascii_digit() && ch != '-' {
            return Ok(None);
        }

        let start = self.pos;
        if ch == '-' {
            self.advance();
        }

        while self.pos < self.input.len() && self.current_char().is_ascii_digit() {
            self.advance();
        }

        let num_str = &self.input[start..self.pos];
        let num = num_str.parse::<i32>()?;
        Ok(Some(num))
    }

    /// Match control word string to enum variant.
    fn match_control_word(&self, word: &'a str, param: Option<i32>) -> RtfResult<ControlWord<'a>> {
        let param_value = param.unwrap_or(1);
        let param_bool = param.unwrap_or(1) != 0;

        #[allow(clippy::match_same_arms)]
        let control = match word {
            // Document
            "rtf" => ControlWord::Rtf(param_value),
            "ansi" => ControlWord::Ansi,
            "ansicpg" => ControlWord::AnsiCodePage(param_value),
            "mac" => ControlWord::Mac,
            "pc" => ControlWord::Pc,
            "pca" => ControlWord::Pca,

            // Headers
            "fonttbl" => ControlWord::FontTable,
            "colortbl" => ControlWord::ColorTable,
            "stylesheet" => ControlWord::StyleSheet,
            "info" => ControlWord::Info,

            // Embedded content
            "pict" => ControlWord::Picture,
            "object" => ControlWord::Object,
            "result" => ControlWord::Result,

            // Picture properties
            "picw" => ControlWord::PictureWidth(param_value),
            "pich" => ControlWord::PictureHeight(param_value),
            "picwgoal" => ControlWord::PictureGoalWidth(param_value),
            "pichgoal" => ControlWord::PictureGoalHeight(param_value),
            "picscalex" => ControlWord::PictureScaleX(param_value),
            "picscaley" => ControlWord::PictureScaleY(param_value),
            "emfblip" => ControlWord::Emfblip,
            "pngblip" => ControlWord::Pngblip,
            "jpegblip" => ControlWord::Jpegblip,
            "macpict" => ControlWord::Macpict,
            "pmmetafile" => ControlWord::Pmmetafile(param_value),
            "wmetafile" => ControlWord::Wmetafile(param_value),
            "dibitmap" => ControlWord::Dibitmap(param_value),
            "wbitmap" => ControlWord::Wbitmap(param_value),

            // Field support
            "field" => ControlWord::Field,
            "fldinst" => ControlWord::FieldInstruction,
            "fldrslt" => ControlWord::FieldResult,
            "fldlock" => ControlWord::FieldLock,
            "flddirty" => ControlWord::FieldDirty,
            "fldedit" => ControlWord::FieldEdit,
            "fldpriv" => ControlWord::FieldPrivate,

            // Fonts
            "f" => ControlWord::FontNumber(param_value),
            "fs" => ControlWord::FontSize(param_value),
            "fcharset" => ControlWord::FontCharset(param_value),
            "fnil" => ControlWord::FontFamily("nil"),
            "froman" => ControlWord::FontFamily("roman"),
            "fswiss" => ControlWord::FontFamily("swiss"),
            "fmodern" => ControlWord::FontFamily("modern"),
            "fscript" => ControlWord::FontFamily("script"),
            "fdecor" => ControlWord::FontFamily("decor"),
            "ftech" => ControlWord::FontFamily("tech"),

            // Colors
            "red" => ControlWord::Red(param_value),
            "green" => ControlWord::Green(param_value),
            "blue" => ControlWord::Blue(param_value),
            "cf" => ControlWord::ColorForeground(param_value),
            "cb" => ControlWord::ColorBackground(param_value),

            // Character formatting
            "b" => ControlWord::Bold(param_bool),
            "i" => ControlWord::Italic(param_bool),
            "ul" => ControlWord::Underline(param_bool),
            "ulnone" => ControlWord::UnderlineNone,
            "strike" => ControlWord::Strike(param_bool),
            "super" => ControlWord::Superscript(param_bool),
            "sub" => ControlWord::Subscript(param_bool),
            "scaps" => ControlWord::SmallCaps(param_bool),
            "plain" => ControlWord::Plain,

            // Paragraph
            "par" => ControlWord::Par,
            "pard" => ControlWord::Pard,
            "ql" => ControlWord::LeftAlign,
            "qr" => ControlWord::RightAlign,
            "qc" => ControlWord::Center,
            "qj" => ControlWord::Justify,

            // Paragraph spacing/indent
            "sb" => ControlWord::SpaceBefore(param_value),
            "sa" => ControlWord::SpaceAfter(param_value),
            "sl" => ControlWord::SpaceBetween(param_value),
            "slmult" => ControlWord::LineMultiple(param_bool),
            "li" => ControlWord::LeftIndent(param_value),
            "ri" => ControlWord::RightIndent(param_value),
            "fi" => ControlWord::FirstLineIndent(param_value),

            // Tables
            "trowd" => ControlWord::TableRowDefaults,
            "row" => ControlWord::TableRow,
            "cell" => ControlWord::TableCell,
            "cellx" => ControlWord::CellX(param_value),
            "intbl" => ControlWord::InTable,

            // Unicode
            "u" => ControlWord::Unicode(param_value),
            "uc" => ControlWord::UnicodeSkip(param_value),

            // Special
            "tab" => ControlWord::Tab,
            "line" => ControlWord::Line,
            "page" => ControlWord::Page,
            "sect" => ControlWord::Section,
            "sectd" => ControlWord::SectionDefault,

            // Binary data
            "bin" => ControlWord::Binary(param_value),

            // Unknown
            _ => ControlWord::Unknown(word, param),
        };

        Ok(control)
    }

    /// Parse hexadecimal character escape (\').
    fn parse_hex_char(&mut self) -> RtfResult<Token<'a>> {
        self.advance(); // Skip '\''

        if self.pos + 1 >= self.input.len() {
            return Err(RtfError::InvalidUnicode(
                "Incomplete hex escape".to_string(),
            ));
        }

        let hex = &self.input[self.pos..self.pos + 2];
        self.pos += 2;

        let byte = u8::from_str_radix(hex, 16)
            .map_err(|_| RtfError::InvalidUnicode(format!("Invalid hex escape: {}", hex)))?;

        // Decode based on character set
        let ch = self.decode_byte(byte);
        let text = self.arena.alloc_str(&ch.to_string());
        Ok(Token::Text(Cow::Borrowed(text)))
    }

    /// Decode a byte according to the current character set.
    ///
    /// This handles Windows-1252 (ANSI), Mac Roman, and DOS codepages.
    fn decode_byte(&self, byte: u8) -> char {
        match self.charset {
            CharacterSet::Ansi => {
                // Windows-1252 / CP1252
                // Bytes 0x00-0x7F are ASCII
                // Bytes 0x80-0x9F have special mappings
                // Bytes 0xA0-0xFF are mostly Latin-1 with some exceptions
                match byte {
                    0x80 => '€',
                    0x82 => '‚',
                    0x83 => 'ƒ',
                    0x84 => '„',
                    0x85 => '…',
                    0x86 => '†',
                    0x87 => '‡',
                    0x88 => 'ˆ',
                    0x89 => '‰',
                    0x8A => 'Š',
                    0x8B => '‹',
                    0x8C => 'Œ',
                    0x8E => 'Ž',
                    0x91 => '\u{2018}', // Left single quotation mark
                    0x92 => '\u{2019}', // Right single quotation mark
                    0x93 => '\u{201C}', // Left double quotation mark
                    0x94 => '\u{201D}', // Right double quotation mark
                    0x95 => '•',
                    0x96 => '–',
                    0x97 => '—',
                    0x98 => '˜',
                    0x99 => '™',
                    0x9A => 'š',
                    0x9B => '›',
                    0x9C => 'œ',
                    0x9E => 'ž',
                    0x9F => 'Ÿ',
                    // Others map to Latin-1
                    _ => byte as char,
                }
            },
            CharacterSet::Mac | CharacterSet::Pc | CharacterSet::Pca => {
                // For Mac Roman and DOS codepages, basic ASCII is the same
                // Extended characters would need full codepage tables
                // For now, use simple Latin-1 mapping as fallback
                if byte < 0x80 {
                    byte as char
                } else {
                    // Would need proper Mac Roman / CP437 / CP850 tables
                    // Fallback to Latin-1 for now
                    byte as char
                }
            },
        }
    }

    /// Parse plain text until special character.
    fn parse_text(&mut self) -> RtfResult<Token<'a>> {
        let _start = self.pos;

        // Use smallvec for efficient text accumulation
        let mut text = SmallVec::<[u8; 64]>::new();

        while self.pos < self.input.len() {
            let ch = self.current_char();
            match ch {
                '\\' | '{' | '}' => break,
                '\r' | '\n' => {
                    // Skip line breaks in plain text, but track them
                    self.advance();
                    // If we have accumulated text, break here
                    if !text.is_empty() {
                        break;
                    }
                },
                _ => {
                    text.push(ch as u8);
                    self.advance();
                },
            }
        }

        if text.is_empty() {
            // If we hit only whitespace/newlines, try to consume at least one whitespace
            // and return a space token, or skip to next token
            if self.pos >= self.input.len() {
                return Err(RtfError::UnexpectedEof);
            }
            // Return empty text for now - parser will handle it
            let allocated = self.arena.alloc_str("");
            return Ok(Token::Text(Cow::Borrowed(allocated)));
        }

        let text_str = std::str::from_utf8(&text)?;
        let allocated = self.arena.alloc_str(text_str);
        Ok(Token::Text(Cow::Borrowed(allocated)))
    }

    /// Get current character without advancing.
    #[inline]
    fn current_char(&self) -> char {
        self.input[self.pos..].chars().next().unwrap_or('\0')
    }

    /// Advance position by one character.
    #[inline]
    fn advance(&mut self) {
        if self.pos < self.input.len() {
            let ch = self.current_char();
            self.pos += ch.len_utf8();
        }
    }

    /// Skip whitespace (but not newlines, they might be significant).
    #[inline]
    fn skip_whitespace(&mut self) {
        while self.pos < self.input.len() {
            let ch = self.current_char();
            if ch == ' ' || ch == '\t' {
                self.advance();
            } else {
                break;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_tokenization() {
        let arena = Bump::new();
        let input = r"{\rtf1\ansi Hello}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();

        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[0], Token::OpenBrace));
        assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
        assert!(matches!(tokens[2], Token::Control(ControlWord::Ansi)));
    }
}
