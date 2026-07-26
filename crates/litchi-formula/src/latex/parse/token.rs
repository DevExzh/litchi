// Token type produced by the LaTeX lexer
//
// Tokens keep the byte range they occupy in the source string. That lets the
// parser rebuild verbatim slices (`\text{a b}`, environment names, multi-digit
// numbers) without ever allocating: every string in the resulting AST borrows
// from the original input.

/// The kind of lexeme a [`Token`] represents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TokenKind<'a> {
    /// A control sequence, stored without its leading backslash.
    ///
    /// Multi-letter sequences keep their full name (`frac` for `\frac`).
    /// Control symbols keep the single character that followed the backslash,
    /// so `\,` becomes `Command(",")` and `\\` becomes `Command("\\")`.
    Command(&'a str),

    /// An opening group brace `{`.
    GroupOpen,

    /// A closing group brace `}`.
    GroupClose,

    /// The subscript marker `_`.
    Subscript,

    /// The superscript marker `^`.
    Superscript,

    /// The alignment/cell separator `&`.
    Align,

    /// Any other single character of source text.
    Char(char),

    /// A collapsed run of whitespace.
    Space,
}

/// A single lexical unit together with its position in the source string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Token<'a> {
    /// What this token is.
    pub kind: TokenKind<'a>,
    /// Byte offset of the first byte of the token.
    pub start: usize,
    /// Byte offset one past the last byte of the token.
    pub end: usize,
}

impl<'a> Token<'a> {
    /// Build a token spanning `start..end`.
    #[inline]
    pub(crate) const fn new(kind: TokenKind<'a>, start: usize, end: usize) -> Self {
        Self { kind, start, end }
    }

    /// Return the control sequence name if this token is a command.
    #[inline]
    pub(crate) const fn command(&self) -> Option<&'a str> {
        match self.kind {
            TokenKind::Command(name) => Some(name),
            _ => None,
        }
    }

    /// Report whether this token is the control sequence `name`.
    #[inline]
    pub(crate) fn is_command(&self, name: &str) -> bool {
        matches!(self.kind, TokenKind::Command(found) if found == name)
    }

    /// Report whether this token is the literal character `ch`.
    #[inline]
    pub(crate) fn is_char(&self, ch: char) -> bool {
        matches!(self.kind, TokenKind::Char(found) if found == ch)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn command_accessors_recognise_control_sequences() {
        let token = Token::new(TokenKind::Command("frac"), 0, 5);
        assert_eq!(token.command(), Some("frac"));
        assert!(token.is_command("frac"));
        assert!(!token.is_command("sqrt"));
        assert!(!token.is_char('f'));
    }

    #[test]
    fn char_accessors_recognise_literal_characters() {
        let token = Token::new(TokenKind::Char('('), 3, 4);
        assert_eq!(token.command(), None);
        assert!(token.is_char('('));
        assert!(!token.is_char(')'));
        assert!(!token.is_command("("));
    }
}
