use super::super::lexer::Lexer;
use super::super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use std::borrow::Cow;

impl<'a> Lexer<'a> {
    /// Parse hexadecimal character escape (\').
    pub(super) fn parse_hex_char(&mut self) -> RtfResult<Token<'a>> {
        let mut bytes = String::new();
        let lenient = self.leniency.tolerates_malformed_hex_escapes();
        loop {
            self.advance()?; // Skip '\''
            if lenient {
                if let Some(byte) = self.lenient_hex_byte()? {
                    bytes.push(char::from(byte));
                }
                if !self.remaining()?.starts_with("\\'") {
                    break;
                }
                self.advance()?;
                continue;
            }
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

    /// Read the two digit slots after `\'` the way LibreOffice does.
    ///
    /// Hex digits accumulate into the byte, any other character in a slot is
    /// dropped, CR and LF are skipped without using a slot, and a structural
    /// `\\`, `{`, or `}` (or end of input) ends the escape early without being
    /// consumed. A well-formed escape is unchanged; any other shape is
    /// recorded as [`SyntaxDefect::MalformedHexEscape`] and yields the byte
    /// its digits spell, or `None` when it has none.
    ///
    /// [`SyntaxDefect::MalformedHexEscape`]: crate::leniency::SyntaxDefect::MalformedHexEscape
    fn lenient_hex_byte(&mut self) -> RtfResult<Option<u8>> {
        let backslash = self.pos.saturating_sub(2);
        let mut value = 0u8;
        let mut digits = 0u8;
        let mut slots = 0u8;
        let mut skipped_line_break = false;
        while slots < 2 {
            let Some(character) = self.remaining()?.chars().next() else {
                break;
            };
            if matches!(character, '\\' | '{' | '}') {
                break;
            }
            self.advance()?;
            if matches!(character, '\r' | '\n') {
                skipped_line_break = true;
                continue;
            }
            slots += 1;
            if let Some(digit) = character.to_digit(16) {
                // `to_digit(16)` is below 16, so the nibble always fits.
                value = (value << 4) | u8::try_from(digit).unwrap_or_default();
                digits += 1;
            }
        }
        if digits != 2 || skipped_line_break {
            self.tolerance
                .record(crate::leniency::SyntaxDefect::MalformedHexEscape, backslash);
        }
        Ok((digits > 0).then_some(value))
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
