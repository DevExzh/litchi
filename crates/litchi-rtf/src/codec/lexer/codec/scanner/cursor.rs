use super::super::lexer::Lexer;
use crate::codec::error::{RtfError, RtfResult};

impl<'a> Lexer<'a> {
    fn invalid_cursor() -> RtfError {
        RtfError::MalformedDocument("RTF lexer cursor is not on a UTF-8 boundary".to_string())
    }

    pub(super) fn source_range(&self, start: usize, end: usize) -> RtfResult<&'a str> {
        self.input.get(start..end).ok_or_else(Self::invalid_cursor)
    }

    pub(super) fn remaining(&self) -> RtfResult<&'a str> {
        self.input.get(self.pos..).ok_or_else(Self::invalid_cursor)
    }
    /// Get current character without advancing.
    #[inline]
    pub(super) fn current_char(&self) -> RtfResult<char> {
        self.remaining()?
            .chars()
            .next()
            .ok_or(RtfError::UnexpectedEof)
    }

    /// Advance position by one character.
    #[inline]
    pub(super) fn advance(&mut self) -> RtfResult<()> {
        let width = self.current_char()?.len_utf8();
        self.pos = self
            .pos
            .checked_add(width)
            .filter(|position| *position <= self.input.len())
            .ok_or_else(Self::invalid_cursor)?;
        Ok(())
    }
}
