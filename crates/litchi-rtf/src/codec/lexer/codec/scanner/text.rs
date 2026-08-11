use super::super::lexer::Lexer;
use super::super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use std::borrow::Cow;

impl<'a> Lexer<'a> {
    /// Parse hexadecimal character escape (\').
    pub(super) fn parse_hex_char(&mut self) -> RtfResult<Token<'a>> {
        let mut bytes = String::new();
        loop {
            self.advance()?; // Skip '\''
            let pair = self
                .input
                .as_bytes()
                .get(self.pos..self.pos + 2)
                .ok_or_else(|| RtfError::InvalidUnicode("Incomplete hex escape".to_string()))?;
            // A mutation or split multi-byte scalar yields non-ASCII bytes
            // here; treat them as an invalid escape rather than a slice panic.
            let hex = std::str::from_utf8(pair)
                .map_err(|_err| RtfError::InvalidUnicode("Invalid hex escape".to_string()))?;
            self.pos += 2;
            let byte = u8::from_str_radix(hex, 16)
                .map_err(|_err| RtfError::InvalidUnicode(format!("Invalid hex escape: {hex}")))?;
            bytes.push(char::from(byte));
            if !self.remaining()?.starts_with("\\'") {
                break;
            }
            self.advance()?; // Skip the next backslash; the loop skips its quote.
        }

        // Preserve source bytes. The parser applies the active code page after
        // interpreting document and group-level encoding controls. Consecutive
        // escapes stay together so multibyte encodings decode atomically.
        let text = self.arena.alloc_str(&bytes);
        Ok(Token::Text(Cow::Borrowed(text)))
    }

    /// Parse plain text until special character.
    pub(super) fn parse_text(&mut self) -> RtfResult<Token<'a>> {
        let mut start = self.pos;

        while self.pos < self.input.len() {
            let remaining = self.remaining()?;
            let Some((offset, delimiter)) = remaining
                .as_bytes()
                .iter()
                .copied()
                .enumerate()
                .find(|(_offset, byte)| matches!(*byte, b'\\' | b'{' | b'}' | b'\r' | b'\n'))
            else {
                self.pos = self.input.len();
                break;
            };
            self.pos += offset;
            if matches!(delimiter, b'\\' | b'{' | b'}') {
                break;
            }
            let end = self.pos;
            self.pos += 1;
            if end > start {
                return Ok(Token::Text(Cow::Borrowed(self.source_range(start, end)?)));
            }
            start = self.pos;
        }

        // Plain text already lives in the source and needs no arena copy. An
        // empty slice preserves the previous behavior for physical line breaks
        // immediately before EOF or a structural token.
        Ok(Token::Text(Cow::Borrowed(
            self.source_range(start, self.pos)?,
        )))
    }
}
