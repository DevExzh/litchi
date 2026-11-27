use super::super::lexer::Lexer;

impl<'a> Lexer<'a> {
    fn invalid_cursor() -> crate::Error {
        crate::Error::MalformedDocument("RTF lexer cursor is not on a UTF-8 boundary".to_string())
    }

    pub(super) fn source_range(&self, start: usize, end: usize) -> crate::Result<&'a str> {
        self.input.get(start..end).ok_or_else(Self::invalid_cursor)
    }

    pub(super) fn remaining(&self) -> crate::Result<&'a str> {
        self.input.get(self.pos..).ok_or_else(Self::invalid_cursor)
    }
    /// Get current character without advancing.
    #[inline]
    pub(super) fn current_char(&self) -> crate::Result<char> {
        self.remaining()?
            .chars()
            .next()
            .ok_or(crate::Error::UnexpectedEof)
    }

    /// Advance position by one character.
    #[inline]
    pub(super) fn advance(&mut self) -> crate::Result<()> {
        let width = self.current_char()?.len_utf8();
        self.pos = self
            .pos
            .checked_add(width)
            .filter(|position| *position <= self.input.len())
            .ok_or_else(Self::invalid_cursor)?;
        Ok(())
    }
}
