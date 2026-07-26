// Core recursive-descent driver for the LaTeX parser
//
// This module owns the cursor over the token stream, the sequence loop and the
// script/prescript machinery. Command dispatch lives in `atoms`, environment
// bodies in `environments`.

use super::commands::{CMD_END, CMD_LINE_BREAK, CMD_RIGHT, ascii_operator};
use super::error::LatexParseError;
use super::fences::{BareFence, bare_fence_for};
use super::lexer::Lexer;
use super::token::{Token, TokenKind};
use crate::ast::{FractionType, MathNode, Operator, Symbol};
use std::borrow::Cow;

/// Character that toggles TeX math mode; carries no AST meaning here.
const MATH_MODE_DELIMITER: char = '$';
/// Character used for the prime symbol.
const PRIME: char = '\'';
/// Decimal separator accepted inside a number run.
const DECIMAL_POINT: char = '.';
/// Character that opens an optional argument such as `\sqrt[3]{x}`.
const OPTIONAL_ARG_OPEN: char = '[';
/// Smallest number of children a group must have to survive as a [`MathNode::Row`].
const MIN_ROW_CHILDREN: usize = 2;

/// Where a parsed token sequence stops.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Limit {
    /// Run to the end of the token stream.
    Eof,
    /// Run until the cursor reaches this absolute token index.
    Index(usize),
    /// Run until a `\right` is reached.
    Right,
    /// Run until a cell separator, a row separator or `\end` is reached.
    Cell,
}

/// How a digit token is turned into a [`MathNode::Number`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NumberMode {
    /// Merge the maximal run of adjacent digits into a single number.
    Run,
    /// Consume exactly one digit, as TeX does for `x^12`.
    Single,
}

/// Infix fraction commands, which split the surrounding list in two.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum InfixFraction {
    /// `\over`: a normal fraction bar.
    Over,
    /// `\atop`: stacked without a bar.
    Atop,
    /// `\choose`: stacked without a bar, wrapped in parentheses.
    Choose,
}

impl InfixFraction {
    /// Resolve a control sequence name into an infix fraction, if it is one.
    fn from_command(name: &str) -> Option<Self> {
        match name {
            "over" => Some(InfixFraction::Over),
            "atop" => Some(InfixFraction::Atop),
            "choose" => Some(InfixFraction::Choose),
            _ => None,
        }
    }
}

/// Cursor over a tokenized LaTeX expression.
pub(crate) struct Parser<'a> {
    /// The original source string; all borrowed AST strings point into it.
    input: &'a str,
    /// The full token stream.
    tokens: Vec<Token<'a>>,
    /// Index of the next token to consume.
    pos: usize,
    /// Current recursion depth.
    depth: usize,
    /// Maximum recursion depth before bailing out.
    max_depth: usize,
}

impl<'a> Parser<'a> {
    /// Tokenize `input` and build a parser over it.
    pub(crate) fn new(input: &'a str, max_depth: usize) -> Result<Self, LatexParseError> {
        Ok(Self {
            input,
            tokens: Lexer::tokenize(input)?,
            pos: 0,
            depth: 0,
            max_depth,
        })
    }

    /// Parse the entire input into a node list.
    pub(crate) fn parse(&mut self) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        self.parse_sequence(Limit::Eof)
    }

    // -- cursor primitives -------------------------------------------------

    /// Borrow a slice of the source with the input's lifetime.
    #[inline]
    pub(crate) fn slice(&self, start: usize, end: usize) -> &'a str {
        let input: &'a str = self.input;
        &input[start..end]
    }

    /// Peek at the token under the cursor.
    #[inline]
    pub(crate) fn peek(&self) -> Option<Token<'a>> {
        self.tokens.get(self.pos).copied()
    }

    /// Peek `offset` tokens past the cursor.
    #[inline]
    pub(crate) fn peek_at(&self, offset: usize) -> Option<Token<'a>> {
        self.tokens.get(self.pos.saturating_add(offset)).copied()
    }

    /// Advance the cursor by one token.
    #[inline]
    pub(crate) fn advance(&mut self) {
        self.pos = self.pos.saturating_add(1);
    }

    /// Move the cursor to an absolute token index.
    #[inline]
    pub(crate) fn seek(&mut self, index: usize) {
        self.pos = index.min(self.tokens.len());
    }

    /// Byte offset of the token under the cursor, or the end of the input.
    #[inline]
    pub(crate) fn offset(&self) -> usize {
        self.peek().map_or(self.input.len(), |token| token.start)
    }

    /// Skip over collapsed whitespace tokens.
    #[inline]
    pub(crate) fn skip_spaces(&mut self) {
        while matches!(self.peek().map(|token| token.kind), Some(TokenKind::Space)) {
            self.advance();
        }
    }

    /// Enter one level of recursion, refusing input that nests too deeply.
    pub(crate) fn enter(&mut self) -> Result<(), LatexParseError> {
        self.depth += 1;
        if self.depth > self.max_depth {
            return Err(LatexParseError::NestingTooDeep {
                position: self.offset(),
                limit: self.max_depth,
            });
        }
        Ok(())
    }

    /// Leave one level of recursion.
    #[inline]
    pub(crate) fn leave(&mut self) {
        self.depth = self.depth.saturating_sub(1);
    }

    /// Report whether the cursor has reached `limit`.
    fn at_limit(&self, limit: Limit) -> bool {
        let Some(token) = self.peek() else {
            return true;
        };
        match limit {
            Limit::Eof => false,
            Limit::Index(index) => self.pos >= index,
            Limit::Right => token.is_command(CMD_RIGHT),
            Limit::Cell => {
                matches!(token.kind, TokenKind::Align)
                    || token.is_command(CMD_LINE_BREAK)
                    || token.is_command(CMD_END)
            },
        }
    }

    // -- structural scanning ----------------------------------------------

    /// Locate the `}` matching the `{` at token index `open`.
    pub(crate) fn find_matching_brace(&self, open: usize) -> Result<usize, LatexParseError> {
        let position = self
            .tokens
            .get(open)
            .map_or(self.input.len(), |token| token.start);
        if !matches!(
            self.tokens.get(open).map(|token| token.kind),
            Some(TokenKind::GroupOpen)
        ) {
            return Err(LatexParseError::UnmatchedGroupOpen { position });
        }

        let mut depth = 0usize;
        for (index, token) in self.tokens.iter().enumerate().skip(open) {
            match token.kind {
                TokenKind::GroupOpen => depth += 1,
                TokenKind::GroupClose => {
                    // `depth` is at least one here because `open` is a `{`.
                    if depth <= 1 {
                        return Ok(index);
                    }
                    depth -= 1;
                },
                _ => {},
            }
        }

        Err(LatexParseError::UnmatchedGroupOpen { position })
    }

    /// Locate the token index closing the bare delimiter pair opened at `open`.
    ///
    /// Returns `None` when the pair is not closed within the current group, in
    /// which case the opening delimiter is treated as an ordinary symbol.
    pub(crate) fn find_matching_delimiter(&self, open: usize, pair: BareFence) -> Option<usize> {
        let mut brace_depth: i32 = 0;
        let mut delim_depth: i32 = 0;

        for (index, token) in self.tokens.iter().enumerate().skip(open) {
            match token.kind {
                TokenKind::GroupOpen => {
                    brace_depth += 1;
                    continue;
                },
                TokenKind::GroupClose => {
                    brace_depth -= 1;
                    if brace_depth < 0 {
                        return None;
                    }
                    continue;
                },
                // A bare pair never spans a row, cell or environment boundary.
                TokenKind::Align => {
                    if brace_depth == 0 {
                        return None;
                    }
                    continue;
                },
                _ => {},
            }

            if brace_depth != 0 {
                continue;
            }
            if token.is_command(CMD_LINE_BREAK) || token.is_command(CMD_END) {
                return None;
            }
            if pair.open.matches(token) {
                delim_depth += 1;
            } else if pair.close.matches(token) {
                delim_depth -= 1;
                if delim_depth == 0 {
                    return Some(index);
                }
            }
        }

        None
    }

    // -- sequences ---------------------------------------------------------

    /// Parse tokens until `limit` is reached.
    pub(crate) fn parse_sequence(
        &mut self,
        limit: Limit,
    ) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        self.enter()?;
        let mut nodes: Vec<MathNode<'a>> = Vec::new();

        loop {
            self.skip_spaces();
            if self.at_limit(limit) {
                break;
            }

            if let Some(kind) = self
                .peek()
                .and_then(|token| token.command())
                .and_then(InfixFraction::from_command)
            {
                self.advance();
                let numerator = flatten(std::mem::take(&mut nodes));
                let denominator = self.parse_sequence(limit)?;
                nodes.push(build_infix_fraction(kind, numerator, denominator));
                break;
            }

            if self.at_prescript() {
                let node = self.parse_prescript()?;
                nodes.push(node);
                continue;
            }

            if matches!(
                self.peek().map(|token| token.kind),
                Some(TokenKind::Subscript | TokenKind::Superscript)
            ) {
                let base = take_script_base(&mut nodes);
                let node = self.parse_scripts(base)?;
                nodes.push(node);
                continue;
            }

            if let Some(node) = self.parse_atom(NumberMode::Run)? {
                nodes.push(node);
            }
        }

        self.leave();
        Ok(flatten(nodes))
    }

    /// Parse the body of the group whose `{` is under the cursor.
    pub(crate) fn parse_group_contents(&mut self) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        let open = self.pos;
        let close = self.find_matching_brace(open)?;
        self.advance();
        let nodes = self.parse_sequence(Limit::Index(close))?;
        self.seek(close);
        self.advance();
        Ok(nodes)
    }

    /// Return the verbatim source between the braces of the group under the
    /// cursor, consuming the whole group.
    pub(crate) fn read_raw_group(&mut self) -> Result<&'a str, LatexParseError> {
        let open = self.pos;
        let close = self.find_matching_brace(open)?;
        let start = self.tokens[open].end;
        let end = self.tokens[close].start;
        self.seek(close);
        self.advance();
        Ok(self.slice(start, end.max(start)))
    }

    // -- arguments ---------------------------------------------------------

    /// Parse the mandatory argument of `command`.
    ///
    /// Accepts either a braced group or, following TeX, a single following atom.
    pub(crate) fn parse_required_argument(
        &mut self,
        command: &str,
    ) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        self.skip_spaces();
        let position = self.offset();

        match self.peek().map(|token| token.kind) {
            Some(TokenKind::GroupOpen) => self.parse_group_contents(),
            Some(TokenKind::GroupClose) | Some(TokenKind::Align) | None => {
                Err(LatexParseError::MissingArgument {
                    command: command.to_string(),
                    position,
                })
            },
            _ => match self.parse_atom(NumberMode::Single)? {
                Some(MathNode::Row(inner)) => Ok(inner),
                Some(node) => Ok(vec![node]),
                None => Ok(Vec::new()),
            },
        }
    }

    /// Parse an optional `[...]` argument, if one is present.
    ///
    /// An unterminated `[` is not an error: the bracket is left in place and
    /// parsed as ordinary content.
    pub(crate) fn parse_optional_argument(
        &mut self,
    ) -> Result<Option<Vec<MathNode<'a>>>, LatexParseError> {
        self.skip_spaces();
        let Some(token) = self.peek().filter(|token| token.is_char(OPTIONAL_ARG_OPEN)) else {
            return Ok(None);
        };
        let Some(pair) = bare_fence_for(&token) else {
            return Ok(None);
        };
        let Some(close) = self.find_matching_delimiter(self.pos, pair) else {
            return Ok(None);
        };

        self.advance();
        let nodes = self.parse_sequence(Limit::Index(close))?;
        self.seek(close);
        self.advance();
        Ok(Some(nodes))
    }

    // -- atoms -------------------------------------------------------------

    /// Parse a single atom.
    ///
    /// Returns `Ok(None)` for input that is consumed but carries no AST
    /// meaning, such as `$` or a presentation-only control sequence.
    pub(crate) fn parse_atom(
        &mut self,
        mode: NumberMode,
    ) -> Result<Option<MathNode<'a>>, LatexParseError> {
        self.skip_spaces();
        let Some(token) = self.peek() else {
            return Ok(None);
        };

        if let Some(node) = self.try_parse_bare_fence()? {
            return Ok(Some(node));
        }

        match token.kind {
            TokenKind::GroupOpen => {
                let nodes = self.parse_group_contents()?;
                Ok(Some(MathNode::Row(nodes)))
            },
            TokenKind::GroupClose => Err(LatexParseError::UnmatchedGroupClose {
                position: token.start,
            }),
            // A stray alignment marker outside an environment carries no
            // meaning; drop it rather than failing the whole parse.
            TokenKind::Align | TokenKind::Space => {
                self.advance();
                Ok(None)
            },
            TokenKind::Subscript | TokenKind::Superscript => {
                self.parse_scripts(Vec::new()).map(Some)
            },
            TokenKind::Command(name) => self.parse_command(name, token),
            TokenKind::Char(ch) => {
                self.advance();
                Ok(self.char_atom(ch, token, mode))
            },
        }
    }

    /// Turn a literal character into a node.
    fn char_atom(&mut self, ch: char, token: Token<'a>, mode: NumberMode) -> Option<MathNode<'a>> {
        if ch.is_ascii_digit() {
            let end = match mode {
                NumberMode::Run => self.extend_number(token.end),
                NumberMode::Single => token.end,
            };
            return Some(MathNode::Number(Cow::Borrowed(
                self.slice(token.start, end),
            )));
        }

        if ch == MATH_MODE_DELIMITER {
            return None;
        }
        if ch == PRIME {
            return Some(MathNode::Operator(Operator::Prime));
        }
        if let Some(operator) = ascii_operator(ch) {
            return Some(MathNode::Operator(operator));
        }

        let text = self.slice(token.start, token.end);
        if ch.is_ascii() {
            Some(MathNode::Text(Cow::Borrowed(text)))
        } else {
            Some(MathNode::Symbol(Symbol {
                name: Cow::Borrowed(text),
                unicode: Some(ch),
                variant: None,
            }))
        }
    }

    /// Extend a number literal over the following digits.
    ///
    /// Only source-adjacent tokens are merged, so a stripped comment or a space
    /// still separates two numbers.
    fn extend_number(&mut self, mut end: usize) -> usize {
        loop {
            let Some(token) = self.peek() else {
                return end;
            };
            if token.start != end {
                return end;
            }
            match token.kind {
                TokenKind::Char(ch) if ch.is_ascii_digit() => {
                    end = token.end;
                    self.advance();
                },
                TokenKind::Char(DECIMAL_POINT) => {
                    let Some(next) = self.peek_at(1) else {
                        return end;
                    };
                    let is_digit = matches!(next.kind, TokenKind::Char(ch) if ch.is_ascii_digit());
                    if !is_digit || next.start != token.end {
                        return end;
                    }
                    end = next.end;
                    self.advance();
                    self.advance();
                },
                _ => return end,
            }
        }
    }

    /// Parse `(...)`-style fencing that does not use `\left`/`\right`.
    fn try_parse_bare_fence(&mut self) -> Result<Option<MathNode<'a>>, LatexParseError> {
        let Some(token) = self.peek() else {
            return Ok(None);
        };
        let Some(pair) = bare_fence_for(&token) else {
            return Ok(None);
        };
        let Some(close) = self.find_matching_delimiter(self.pos, pair) else {
            return Ok(None);
        };

        self.advance();
        let content = self.parse_sequence(Limit::Index(close))?;
        self.seek(close);
        self.advance();

        Ok(Some(MathNode::Fenced {
            open: pair.fence,
            content,
            close: pair.fence,
            separator: None,
        }))
    }
}

/// Flatten groups that were never used as the base of a script.
///
/// A group with two or more children stays a [`MathNode::Row`] so the AST keeps
/// the author's grouping; shorter groups add nothing and are inlined.
fn flatten<'a>(nodes: Vec<MathNode<'a>>) -> Vec<MathNode<'a>> {
    if !nodes
        .iter()
        .any(|node| matches!(node, MathNode::Row(inner) if inner.len() < MIN_ROW_CHILDREN))
    {
        return nodes;
    }

    let mut flat = Vec::with_capacity(nodes.len());
    for node in nodes {
        match node {
            MathNode::Row(inner) if inner.len() < MIN_ROW_CHILDREN => flat.extend(inner),
            other => flat.push(other),
        }
    }
    flat
}

/// Detach the base a `_` or `^` should bind to from the nodes parsed so far.
fn take_script_base<'a>(nodes: &mut Vec<MathNode<'a>>) -> Vec<MathNode<'a>> {
    match nodes.pop() {
        Some(MathNode::Row(inner)) => inner,
        Some(node) => vec![node],
        None => Vec::new(),
    }
}

/// Build the node for an infix fraction command.
fn build_infix_fraction<'a>(
    kind: InfixFraction,
    numerator: Vec<MathNode<'a>>,
    denominator: Vec<MathNode<'a>>,
) -> MathNode<'a> {
    let frac_type = match kind {
        InfixFraction::Over => FractionType::Bar,
        InfixFraction::Atop | InfixFraction::Choose => FractionType::NoBar,
    };
    let fraction = MathNode::Frac {
        numerator,
        denominator,
        line_thickness: None,
        frac_type: Some(frac_type),
    };

    match kind {
        InfixFraction::Choose => MathNode::Fenced {
            open: crate::ast::Fence::Paren,
            content: vec![fraction],
            close: crate::ast::Fence::Paren,
            separator: None,
        },
        _ => fraction,
    }
}
