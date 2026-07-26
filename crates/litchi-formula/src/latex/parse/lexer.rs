// Tokenizer for LaTeX math source
//
// The lexer performs a single forward pass over the input and never allocates
// beyond the token vector itself. Comments are stripped, whitespace runs are
// collapsed into a single `Space` token, and control sequences are split into
// multi-letter names (`\frac`) and single-character control symbols (`\,`).

use super::error::LatexParseError;
use super::token::{Token, TokenKind};

/// Character that introduces a control sequence.
const ESCAPE: char = '\\';
/// Character that introduces a comment running to the end of the line.
const COMMENT: char = '%';
/// Character that opens a group.
const GROUP_OPEN: char = '{';
/// Character that closes a group.
const GROUP_CLOSE: char = '}';
/// Character that marks a subscript.
const SUBSCRIPT: char = '_';
/// Character that marks a superscript.
const SUPERSCRIPT: char = '^';
/// Character that separates cells inside an environment.
const ALIGN: char = '&';
/// Character that terminates a comment.
const NEWLINE: char = '\n';

/// Converts LaTeX source into a flat token stream.
pub(crate) struct Lexer<'a> {
    /// The source string being scanned.
    input: &'a str,
    /// Current byte offset into `input`.
    pos: usize,
}

impl<'a> Lexer<'a> {
    /// Create a lexer positioned at the start of `input`.
    pub(crate) const fn new(input: &'a str) -> Self {
        Self { input, pos: 0 }
    }

    /// Tokenize the whole input.
    ///
    /// Fails only when the input ends with a dangling backslash, which cannot
    /// be turned into any control sequence.
    pub(crate) fn tokenize(input: &'a str) -> Result<Vec<Token<'a>>, LatexParseError> {
        let mut lexer = Self::new(input);
        // One token per character is the worst case; reserving avoids repeated
        // growth for the common, mostly single-character math input.
        let mut tokens = Vec::with_capacity(input.len());

        while let Some(token) = lexer.next_token()? {
            tokens.push(token);
        }

        Ok(tokens)
    }

    /// Return the character at the cursor without consuming it.
    #[inline]
    fn peek(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    /// Advance the cursor past `ch`.
    #[inline]
    fn bump(&mut self, ch: char) {
        self.pos += ch.len_utf8();
    }

    /// Produce the next token, or `None` once the input is exhausted.
    fn next_token(&mut self) -> Result<Option<Token<'a>>, LatexParseError> {
        loop {
            let Some(ch) = self.peek() else {
                return Ok(None);
            };
            let start = self.pos;

            let kind = match ch {
                ESCAPE => return self.lex_command(start).map(Some),
                COMMENT => {
                    self.skip_comment();
                    continue;
                },
                GROUP_OPEN => {
                    self.bump(ch);
                    TokenKind::GroupOpen
                },
                GROUP_CLOSE => {
                    self.bump(ch);
                    TokenKind::GroupClose
                },
                SUBSCRIPT => {
                    self.bump(ch);
                    TokenKind::Subscript
                },
                SUPERSCRIPT => {
                    self.bump(ch);
                    TokenKind::Superscript
                },
                ALIGN => {
                    self.bump(ch);
                    TokenKind::Align
                },
                _ if ch.is_whitespace() => {
                    self.skip_whitespace();
                    TokenKind::Space
                },
                _ => {
                    self.bump(ch);
                    TokenKind::Char(ch)
                },
            };

            return Ok(Some(Token::new(kind, start, self.pos)));
        }
    }

    /// Lex a control sequence starting at the backslash located at `start`.
    fn lex_command(&mut self, start: usize) -> Result<Token<'a>, LatexParseError> {
        self.bump(ESCAPE);
        let name_start = self.pos;

        let Some(first) = self.peek() else {
            return Err(LatexParseError::IncompleteCommand { position: start });
        };

        if first.is_ascii_alphabetic() {
            while let Some(ch) = self.peek() {
                if !ch.is_ascii_alphabetic() {
                    break;
                }
                self.bump(ch);
            }
            let name = &self.input[name_start..self.pos];
            let end = self.pos;
            // TeX swallows the whitespace that terminates a multi-letter
            // control sequence, so `\alpha x` is two atoms and not three.
            self.skip_whitespace();
            return Ok(Token::new(TokenKind::Command(name), start, end));
        }

        // A control symbol: exactly one character follows the backslash.
        self.bump(first);
        let name = &self.input[name_start..self.pos];
        Ok(Token::new(TokenKind::Command(name), start, self.pos))
    }

    /// Consume a maximal run of whitespace.
    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.peek() {
            if !ch.is_whitespace() {
                break;
            }
            self.bump(ch);
        }
    }

    /// Consume a `%` comment together with the newline that terminates it.
    fn skip_comment(&mut self) {
        while let Some(ch) = self.peek() {
            self.bump(ch);
            if ch == NEWLINE {
                break;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(input: &str) -> Vec<TokenKind<'_>> {
        Lexer::tokenize(input)
            .expect("input lexes")
            .into_iter()
            .map(|token| token.kind)
            .collect()
    }

    #[test]
    fn lexes_multi_letter_control_sequences() {
        assert_eq!(kinds("\\frac"), vec![TokenKind::Command("frac")]);
    }

    #[test]
    fn swallows_whitespace_after_a_control_word() {
        assert_eq!(
            kinds("\\alpha x"),
            vec![TokenKind::Command("alpha"), TokenKind::Char('x')]
        );
    }

    #[test]
    fn lexes_control_symbols_as_single_character_names() {
        assert_eq!(
            kinds("\\,\\\\\\{"),
            vec![
                TokenKind::Command(","),
                TokenKind::Command("\\"),
                TokenKind::Command("{"),
            ]
        );
    }

    #[test]
    fn collapses_whitespace_runs_into_one_token() {
        assert_eq!(
            kinds("a \t\n b"),
            vec![TokenKind::Char('a'), TokenKind::Space, TokenKind::Char('b')]
        );
    }

    #[test]
    fn strips_comments_through_end_of_line() {
        assert_eq!(
            kinds("a% ignored \nb"),
            vec![TokenKind::Char('a'), TokenKind::Char('b')]
        );
    }

    #[test]
    fn records_byte_spans_for_every_token() {
        let tokens = Lexer::tokenize("\\pi^2").expect("input lexes");
        assert_eq!(tokens[0].start, 0);
        assert_eq!(tokens[0].end, 3);
        assert_eq!(tokens[1].kind, TokenKind::Superscript);
        assert_eq!(tokens[2].kind, TokenKind::Char('2'));
    }

    #[test]
    fn rejects_a_dangling_backslash() {
        assert_eq!(
            Lexer::tokenize("x \\"),
            Err(LatexParseError::IncompleteCommand { position: 2 })
        );
    }

    #[test]
    fn handles_multi_byte_characters() {
        assert_eq!(
            kinds("α+β"),
            vec![
                TokenKind::Char('α'),
                TokenKind::Char('+'),
                TokenKind::Char('β')
            ]
        );
    }
}
