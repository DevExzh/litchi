// Delimiter tables for fenced expressions
//
// Two kinds of fencing are recognised: explicit `\left ... \right` pairs and
// bare pairs such as `(a+b)`. Both resolve to the same [`Fence`] vocabulary
// used by `latex::operators::fence_to_latex`, so a fenced node renders and
// re-parses identically.

use super::token::{Token, TokenKind};
use crate::ast::Fence;

/// Identifies a delimiter token: either a literal character or a control
/// sequence name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DelimKey {
    /// A single source character such as `(` or `|`.
    Char(char),
    /// A control sequence name such as `langle` (written `\langle`).
    Command(&'static str),
}

impl DelimKey {
    /// Report whether `token` is exactly this delimiter.
    #[inline]
    pub(crate) fn matches(self, token: &Token<'_>) -> bool {
        match self {
            DelimKey::Char(ch) => token.is_char(ch),
            DelimKey::Command(name) => token.is_command(name),
        }
    }
}

/// A delimiter pair that may be used without `\left` and `\right`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct BareFence {
    /// The opening delimiter.
    pub open: DelimKey,
    /// The closing delimiter.
    pub close: DelimKey,
    /// The AST fence the pair maps to.
    pub fence: Fence,
}

/// Delimiter pairs recognised without `\left`/`\right`.
///
/// `|` and `\|` are excluded on purpose: the same glyph opens and closes them,
/// so nesting cannot be resolved without a `\left`/`\right` annotation.
static BARE_FENCES: &[BareFence] = &[
    BareFence {
        open: DelimKey::Char('('),
        close: DelimKey::Char(')'),
        fence: Fence::Paren,
    },
    BareFence {
        open: DelimKey::Char('['),
        close: DelimKey::Char(']'),
        fence: Fence::Bracket,
    },
    BareFence {
        open: DelimKey::Command("{"),
        close: DelimKey::Command("}"),
        fence: Fence::Brace,
    },
    BareFence {
        open: DelimKey::Command("langle"),
        close: DelimKey::Command("rangle"),
        fence: Fence::Angle,
    },
    BareFence {
        open: DelimKey::Command("lfloor"),
        close: DelimKey::Command("rfloor"),
        fence: Fence::Floor,
    },
    BareFence {
        open: DelimKey::Command("lceil"),
        close: DelimKey::Command("rceil"),
        fence: Fence::Ceiling,
    },
    BareFence {
        open: DelimKey::Command("lbrack"),
        close: DelimKey::Command("rbrack"),
        fence: Fence::SquareBracket,
    },
    BareFence {
        open: DelimKey::Command("lbrace"),
        close: DelimKey::Command("rbrace"),
        fence: Fence::CurlyBrace,
    },
];

/// Return the bare fence pair opened by `token`, if any.
pub(crate) fn bare_fence_for(token: &Token<'_>) -> Option<BareFence> {
    BARE_FENCES
        .iter()
        .copied()
        .find(|pair| pair.open.matches(token))
}

/// Resolve the delimiter following `\left` or `\right` into a [`Fence`].
///
/// `.` yields [`Fence::None`], the null delimiter LaTeX uses for one-sided
/// fences.
pub(crate) fn fence_for_delimiter(token: &Token<'_>) -> Option<Fence> {
    match token.kind {
        TokenKind::Char('(') | TokenKind::Char(')') => Some(Fence::Paren),
        TokenKind::Char('[') | TokenKind::Char(']') => Some(Fence::Bracket),
        TokenKind::Char('<') | TokenKind::Char('>') => Some(Fence::Angle),
        TokenKind::Char('|') => Some(Fence::Pipe),
        TokenKind::Char('.') => Some(Fence::None),
        TokenKind::Command(name) => command_fence(name),
        _ => None,
    }
}

/// Resolve a delimiter control sequence into a [`Fence`].
fn command_fence(name: &str) -> Option<Fence> {
    match name {
        "{" | "}" | "lbrace" | "rbrace" => Some(match name {
            "lbrace" | "rbrace" => Fence::CurlyBrace,
            _ => Fence::Brace,
        }),
        "|" | "Vert" | "lVert" | "rVert" => Some(Fence::DoublePipe),
        "vert" | "lvert" | "rvert" => Some(Fence::Pipe),
        "langle" | "rangle" => Some(Fence::Angle),
        "lfloor" | "rfloor" => Some(Fence::Floor),
        "lceil" | "rceil" => Some(Fence::Ceiling),
        "lbrack" | "rbrack" => Some(Fence::SquareBracket),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn char_token(ch: char) -> Token<'static> {
        Token::new(TokenKind::Char(ch), 0, ch.len_utf8())
    }

    fn command_token(name: &'static str) -> Token<'static> {
        Token::new(TokenKind::Command(name), 0, name.len() + 1)
    }

    #[test]
    fn recognises_bare_parenthesis_pairs() {
        let pair = bare_fence_for(&char_token('(')).expect("`(` opens a bare fence");
        assert_eq!(pair.fence, Fence::Paren);
        assert!(pair.close.matches(&char_token(')')));
    }

    #[test]
    fn pipes_are_not_bare_fences() {
        assert!(bare_fence_for(&char_token('|')).is_none());
        assert!(bare_fence_for(&command_token("|")).is_none());
    }

    #[test]
    fn resolves_left_right_delimiters() {
        assert_eq!(fence_for_delimiter(&char_token('(')), Some(Fence::Paren));
        assert_eq!(fence_for_delimiter(&char_token('.')), Some(Fence::None));
        assert_eq!(fence_for_delimiter(&char_token('|')), Some(Fence::Pipe));
        assert_eq!(
            fence_for_delimiter(&command_token("|")),
            Some(Fence::DoublePipe)
        );
        assert_eq!(fence_for_delimiter(&command_token("{")), Some(Fence::Brace));
        assert_eq!(
            fence_for_delimiter(&command_token("lfloor")),
            Some(Fence::Floor)
        );
        assert_eq!(fence_for_delimiter(&command_token("frac")), None);
    }
}
