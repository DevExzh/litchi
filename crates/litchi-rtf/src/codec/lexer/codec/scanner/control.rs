use super::super::control_word::{ControlWord, match_control_word};
use super::super::lexer::Lexer;
use super::super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use std::borrow::Cow;

impl<'a> Lexer<'a> {
    /// Get the next token.
    pub(in crate::codec::lexer) fn next_token(&mut self) -> RtfResult<Token<'a>> {
        if self.pos >= self.input.len() {
            return Err(RtfError::UnexpectedEof);
        }

        let ch = self.current_char()?;
        match ch {
            '{' => {
                self.advance()?;
                Ok(Token::OpenBrace)
            },
            '}' => {
                self.advance()?;
                Ok(Token::CloseBrace)
            },
            '\\' => self.parse_control_word(),
            _ => self.parse_text(),
        }
    }

    /// Parse a control word or control symbol.
    fn parse_control_word(&mut self) -> RtfResult<Token<'a>> {
        self.advance()?; // Skip '\'

        if self.pos >= self.input.len() {
            return Err(RtfError::UnexpectedEof);
        }

        let ch = self.current_char()?;

        // Handle special control symbols
        match ch {
            '\\' | '{' | '}' => {
                let start = self.pos;
                self.advance()?;
                return Ok(Token::Text(Cow::Borrowed(
                    self.source_range(start, self.pos)?,
                )));
            },
            '\'' => return self.parse_hex_char(),
            '*' => {
                self.advance()?;
                return Ok(Token::Control(ControlWord::IgnorableDestination));
            },
            '\n' | '\r' => {
                self.advance()?;
                return Ok(Token::Control(ControlWord::Par));
            },
            '~' => {
                self.advance()?;
                return Ok(Token::Control(ControlWord::NonBreakingSpace));
            },
            '-' => {
                self.advance()?;
                return Ok(Token::Control(ControlWord::OptionalHyphen));
            },
            '_' => {
                self.advance()?;
                return Ok(Token::Control(ControlWord::NonBreakingHyphen));
            },
            _ => {},
        }

        // Parse control word
        let start = self.pos;

        // Read alphabetic characters
        while self.pos < self.input.len() && self.current_char()?.is_ascii_alphabetic() {
            self.advance()?;
        }

        if start == self.pos {
            return self.parse_unhandled_control_symbol(start);
        }

        let word = self.source_range(start, self.pos)?;

        // Parse optional numeric parameter
        let param = self.parse_numeric_parameter()?;

        // Skip optional space delimiter after control word
        if self.pos < self.input.len() && self.current_char()? == ' ' {
            self.advance()?;
        }

        // Match control word to enum variant, applying the Word-compatible
        // interpretation of a malformed parameter shape.
        let mut repair = None;
        let control = match_control_word(word, param, &mut repair)?;
        if let Some(defect) = repair {
            self.tolerance.record(defect, start.saturating_sub(1));
        }

        // Handle binary data immediately after \bin. The input uses a one-byte
        // Latin-1 transport mapping, so each scalar maps back to its source byte.
        if let ControlWord::Binary(size) = control {
            return self.parse_binary_payload(size);
        }

        Ok(Token::Control(control))
    }

    /// Tokenize a control symbol the lexer does not translate itself.
    ///
    /// RTF 1.9.1 defines `\|` (formula character) and `\:` (index
    /// subentry), which the model does not interpret. Any other
    /// non-alphabetic character after a backslash is undefined; RTF 1.9.1
    /// tells readers to ignore control symbols they do not recognize, and
    /// Word and LibreOffice (`RTFTokenizer::dispatchKeyword`) do. Both kinds
    /// become [`ControlWord::Unknown`], which the parser retains as an inert
    /// opaque control exactly like an unknown control word; only the
    /// undefined symbols are recorded as repairs. Control symbols take no
    /// delimiter, so a following space stays text.
    fn parse_unhandled_control_symbol(&mut self, start: usize) -> RtfResult<Token<'a>> {
        self.advance()?;
        let symbol = self.source_range(start, self.pos)?;
        if !matches!(symbol, "|" | ":") {
            self.tolerance.record(
                crate::leniency::SyntaxDefect::UnknownControlSymbol,
                start.saturating_sub(1),
            );
        }
        Ok(Token::Control(ControlWord::Unknown(symbol, None)))
    }

    /// Parse numeric parameter after control word.
    fn parse_numeric_parameter(&mut self) -> RtfResult<Option<i32>> {
        if self.pos >= self.input.len() {
            return Ok(None);
        }

        let ch = self.current_char()?;
        if !ch.is_ascii_digit() && ch != '-' {
            return Ok(None);
        }

        let start = self.pos;
        if ch == '-' {
            self.advance()?;
        }

        while self.pos < self.input.len() && self.current_char()?.is_ascii_digit() {
            self.advance()?;
        }

        let num_str = self.source_range(start, self.pos)?;
        let num = num_str.parse::<i32>()?;
        Ok(Some(num))
    }
}
