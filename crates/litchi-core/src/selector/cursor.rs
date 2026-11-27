//! Single-pass byte cursor and the bracket-predicate parser shared by paths and
//! queries.
//!
//! All delimiters in the grammar are ASCII, so the cursor walks bytes and only
//! slices the input at ASCII delimiter offsets, which are always UTF-8 char
//! boundaries. Offsets in errors are absolute byte offsets into the caller's
//! original string.

use std::borrow::Cow;

use super::Position;
use super::error::{ErrorKind, Expected, Found, Grammar, Limit, ParseError};
use super::limits;
use super::predicate::{CompareOp, Condition, Expr, Key, Predicate, Predicates, Value};
use super::text::{is_key_char, is_quote, is_ws, starts_with_ci, unescape};

pub(super) type Result<T> = std::result::Result<T, ParseError>;

pub(super) struct Cursor<'a> {
    src: &'a str,
    pos: usize,
    end: usize,
    grammar: Grammar,
    nesting: usize,
}

impl<'a> Cursor<'a> {
    /// Creates a cursor over the whole input after checking the length bound.
    pub(super) fn new(src: &'a str, grammar: Grammar) -> Result<Self> {
        if src.len() > limits::MAX_INPUT_BYTES {
            // Report the last char boundary within the bound.
            let mut offset = limits::MAX_INPUT_BYTES;
            while !src.is_char_boundary(offset) {
                offset -= 1;
            }
            return Err(ParseError::new(
                grammar,
                offset,
                ErrorKind::LimitExceeded(Limit::InputBytes),
            ));
        }
        Ok(Self {
            src,
            pos: 0,
            end: src.len(),
            grammar,
            nesting: 0,
        })
    }

    /// Creates a cursor restricted to `src[start..end]`, reporting absolute
    /// offsets.
    pub(super) fn window(src: &'a str, start: usize, end: usize, grammar: Grammar) -> Self {
        Self {
            src,
            pos: start,
            end,
            grammar,
            nesting: 0,
        }
    }

    #[inline]
    pub(super) fn src(&self) -> &'a str {
        self.src
    }

    #[inline]
    pub(super) fn pos(&self) -> usize {
        self.pos
    }

    #[inline]
    pub(super) fn end(&self) -> usize {
        self.end
    }

    #[inline]
    pub(super) fn set_pos(&mut self, pos: usize) {
        self.pos = pos;
    }

    #[inline]
    pub(super) fn set_end(&mut self, end: usize) {
        self.end = end;
    }

    #[inline]
    pub(super) fn at_end(&self) -> bool {
        self.pos >= self.end
    }

    #[inline]
    pub(super) fn peek(&self) -> Option<u8> {
        if self.pos < self.end {
            Some(self.src.as_bytes()[self.pos])
        } else {
            None
        }
    }

    #[inline]
    pub(super) fn peek_at(&self, ahead: usize) -> Option<u8> {
        let at = self.pos.checked_add(ahead)?;
        if at < self.end {
            Some(self.src.as_bytes()[at])
        } else {
            None
        }
    }

    #[inline]
    pub(super) fn rest(&self) -> &'a [u8] {
        &self.src.as_bytes()[self.pos.min(self.end)..self.end]
    }

    #[inline]
    pub(super) fn bump(&mut self) {
        self.pos += 1;
    }

    #[inline]
    pub(super) fn eat(&mut self, byte: u8) -> bool {
        if self.peek() == Some(byte) {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    /// Skips ASCII whitespace and reports whether any was skipped.
    pub(super) fn skip_ws(&mut self) -> bool {
        let start = self.pos;
        while matches!(self.peek(), Some(b) if is_ws(b)) {
            self.pos += 1;
        }
        self.pos != start
    }

    pub(super) fn found(&self) -> Found {
        match self.src[self.pos.min(self.end)..self.end].chars().next() {
            Some(c) => Found::Char(c),
            None => Found::End,
        }
    }

    #[inline]
    pub(super) fn error_at(&self, offset: usize, kind: ErrorKind) -> ParseError {
        ParseError::new(self.grammar, offset, kind)
    }

    #[inline]
    pub(super) fn error(&self, kind: ErrorKind) -> ParseError {
        self.error_at(self.pos, kind)
    }

    pub(super) fn unexpected(&self, expected: Expected) -> ParseError {
        self.error(ErrorKind::Unexpected {
            expected,
            found: self.found(),
        })
    }

    pub(super) fn expect(&mut self, byte: u8, expected: Expected) -> Result<()> {
        if self.eat(byte) {
            Ok(())
        } else {
            Err(self.unexpected(expected))
        }
    }

    pub(super) fn enter(&mut self) -> Result<()> {
        if self.nesting >= limits::MAX_NESTING {
            return Err(self.error(ErrorKind::LimitExceeded(Limit::Nesting)));
        }
        self.nesting += 1;
        Ok(())
    }

    pub(super) fn leave(&mut self) {
        self.nesting -= 1;
    }

    /// Scans a quoted span starting at the opening quote and returns the raw
    /// inner slice; a backslash escapes the next character.
    pub(super) fn quoted_raw(&mut self) -> Result<&'a str> {
        let open = self.pos;
        let quote = self.src.as_bytes()[open];
        self.pos += 1;
        let inner_start = self.pos;
        loop {
            match self.peek() {
                None => return Err(self.error_at(open, ErrorKind::UnterminatedQuote)),
                Some(b'\\') if self.pos + 1 < self.end => self.pos += 2,
                Some(b) if b == quote => {
                    let inner = &self.src[inner_start..self.pos];
                    self.pos += 1;
                    return Ok(inner);
                },
                Some(_) => self.pos += 1,
            }
        }
    }

    /// Reads a quoted token and reverses its escapes.
    pub(super) fn quoted(&mut self) -> Result<Cow<'a, str>> {
        let quote = char::from(self.src.as_bytes()[self.pos]);
        let inner = self.quoted_raw()?;
        Ok(unescape(inner, quote))
    }

    /// Parses one bracket predicate. The cursor is just past `[`; on success
    /// it is just past the closing `]`.
    pub(super) fn predicate(&mut self) -> Result<Predicate<'a>> {
        self.skip_ws();
        let start = self.pos;
        let digits = self
            .rest()
            .iter()
            .take_while(|b| b.is_ascii_digit())
            .count();
        if digits > 0 {
            self.pos += digits;
            let after = self.pos;
            self.skip_ws();
            if self.eat(b']') {
                return self.index(start, after).map(Predicate::Index);
            }
            self.pos = start;
        } else if self.peek() == Some(b'-')
            && matches!(self.peek_at(1), Some(b) if b.is_ascii_digit())
        {
            return Err(self.error(ErrorKind::IndexNegative));
        } else if starts_with_ci(self.rest(), b"last()") {
            self.pos += b"last()".len();
            self.skip_ws();
            if self.eat(b']') {
                return Ok(Predicate::Last);
            }
            self.pos = start;
        }
        if self.peek() == Some(b']') {
            return Err(self.unexpected(Expected::Predicate));
        }
        let mut terms = 0;
        let expr = self.or_expr(&mut terms)?;
        self.skip_ws();
        self.expect(b']', Expected::CloseBracket)?;
        Ok(Predicate::Filter(expr))
    }

    fn index(&self, start: usize, end: usize) -> Result<Position> {
        let mut value: usize = 0;
        for &b in &self.src.as_bytes()[start..end] {
            value = value
                .checked_mul(10)
                .and_then(|v| v.checked_add(usize::from(b - b'0')))
                .ok_or_else(|| self.error_at(start, ErrorKind::IndexOverflow))?;
        }
        Position::from_one_based(value).ok_or_else(|| self.error_at(start, ErrorKind::IndexZero))
    }

    fn or_expr(&mut self, terms: &mut usize) -> Result<Expr<'a>> {
        let first = self.and_expr(terms)?;
        if !self.connective(b"or") {
            return Ok(first);
        }
        let mut parts = Vec::with_capacity(4);
        parts.push(first);
        loop {
            parts.push(self.and_expr(terms)?);
            if !self.connective(b"or") {
                return Ok(Expr::Or(parts));
            }
        }
    }

    fn and_expr(&mut self, terms: &mut usize) -> Result<Expr<'a>> {
        let first = self.factor(terms)?;
        if !self.connective(b"and") {
            return Ok(first);
        }
        let mut parts = Vec::with_capacity(4);
        parts.push(first);
        loop {
            parts.push(self.factor(terms)?);
            if !self.connective(b"and") {
                return Ok(Expr::And(parts));
            }
        }
    }

    /// Consumes `keyword` when it follows (after optional whitespace) as a
    /// whole word, i.e. followed by whitespace or `(`.
    fn connective(&mut self, keyword: &[u8]) -> bool {
        let save = self.pos;
        self.skip_ws();
        let rest = self.rest();
        if starts_with_ci(rest, keyword)
            && matches!(rest.get(keyword.len()), Some(&b) if is_ws(b) || b == b'(')
        {
            self.pos += keyword.len();
            return true;
        }
        self.pos = save;
        false
    }

    /// Whether the whitespace at the cursor is followed by an `and`/`or`
    /// connective; an unquoted value ends there.
    fn connective_ahead(&self) -> bool {
        let rest = self.rest();
        let ws = rest.iter().take_while(|&&b| is_ws(b)).count();
        let rest = &rest[ws..];
        // `]`, `)`, and the end also bound the keyword here, so a dangling
        // `[a=1 and]` is reported instead of silently becoming the value
        // `1 and`.
        [&b"and"[..], &b"or"[..]].iter().any(|keyword| {
            starts_with_ci(rest, keyword)
                && match rest.get(keyword.len()) {
                    None => true,
                    Some(&b) => is_ws(b) || matches!(b, b'(' | b']' | b')'),
                }
        })
    }

    fn factor(&mut self, terms: &mut usize) -> Result<Expr<'a>> {
        self.skip_ws();
        if starts_with_ci(self.rest(), b"not") {
            let save = self.pos;
            self.pos += 3;
            self.skip_ws();
            if self.peek() == Some(b'(') {
                return self.group(terms).map(|inner| Expr::Not(Box::new(inner)));
            }
            self.pos = save;
        }
        if self.peek() == Some(b'(') {
            return self.group(terms);
        }
        *terms += 1;
        if *terms > limits::MAX_TERMS {
            return Err(self.error(ErrorKind::LimitExceeded(Limit::Terms)));
        }
        self.condition().map(Expr::Condition)
    }

    /// Parses `( or_expr )`; the cursor is at `(`.
    fn group(&mut self, terms: &mut usize) -> Result<Expr<'a>> {
        self.enter()?;
        self.bump();
        let inner = self.or_expr(terms)?;
        self.skip_ws();
        self.expect(b')', Expected::CloseParen)?;
        self.leave();
        Ok(inner)
    }

    fn condition(&mut self) -> Result<Condition<'a>> {
        let attribute = self.eat(b'@');
        let key = match self.peek() {
            Some(b) if is_quote(b) => self.quoted()?,
            _ => {
                let start = self.pos;
                while let Some(c) = self.src[self.pos..self.end].chars().next() {
                    if !is_key_char(c) {
                        break;
                    }
                    self.pos += c.len_utf8();
                }
                if self.pos == start {
                    return Err(self.unexpected(Expected::Predicate));
                }
                match self.peek() {
                    None => {},
                    Some(b) if is_ws(b) || is_op_start(b) || b == b']' || b == b')' => {},
                    Some(_) => return Err(self.unexpected(Expected::Operator)),
                }
                Cow::Borrowed(&self.src[start..self.pos])
            },
        };
        let key = Key::from_parts(key, attribute);
        self.skip_ws();
        match self.peek() {
            Some(b) if is_op_start(b) => {
                let op = self.operator()?;
                let value = self.value(op)?;
                Ok(Condition::Compare { key, op, value })
            },
            _ => Ok(Condition::Exists(key)),
        }
    }

    fn operator(&mut self) -> Result<CompareOp> {
        let rest = self.rest();
        let (op, len) = match rest {
            [b'\\', b'!', b'=', ..] => (CompareOp::Ne, 3),
            [b'>', b'=', ..] => (CompareOp::Ge, 2),
            [b'<', b'=', ..] => (CompareOp::Le, 2),
            [b'!', b'=', ..] => (CompareOp::Ne, 2),
            [b'~', b'=', ..] => (CompareOp::Contains, 2),
            [b'=', ..] => (CompareOp::Eq, 1),
            [b'>', ..] => (CompareOp::Gt, 1),
            [b'<', ..] => (CompareOp::Lt, 1),
            _ => return Err(self.unexpected(Expected::Operator)),
        };
        self.pos += len;
        Ok(op)
    }

    fn value(&mut self, op: CompareOp) -> Result<Value<'a>> {
        self.skip_ws();
        let start = self.pos;
        match (self.peek(), self.peek_at(1)) {
            (Some(b'r'), Some(q)) if is_quote(q) => {
                self.bump();
                let pattern = self.quoted_raw()?;
                if op != CompareOp::Contains {
                    return Err(self.error_at(start, ErrorKind::RegexRequiresContains));
                }
                return Ok(Value::Regex(Cow::Borrowed(pattern)));
            },
            (Some(q), _) if is_quote(q) => return self.quoted().map(Value::Text),
            _ => {},
        }
        loop {
            match self.peek() {
                None | Some(b']' | b')') => break,
                Some(b) if is_ws(b) => {
                    if self.connective_ahead() {
                        break;
                    }
                    self.bump();
                },
                Some(b) if is_quote(b) => return Err(self.error(ErrorKind::QuoteInBareValue)),
                Some(b'[') => return Err(self.unexpected(Expected::CloseBracket)),
                Some(b) if b.is_ascii_control() => {
                    return Err(self.error(ErrorKind::ControlCharacter));
                },
                Some(_) => self.bump(),
            }
        }
        let raw = &self.src[start..self.pos];
        let text = raw.trim_end_matches(|c: char| c.is_ascii() && is_ws(c as u8));
        match text.as_bytes().first() {
            None => {
                return Err(self.error_at(
                    start,
                    ErrorKind::Unexpected {
                        expected: Expected::Value,
                        found: self.found(),
                    },
                ));
            },
            Some(b'=' | b'~' | b'!') => {
                return Err(self.error_at(start, ErrorKind::MisparsedOperator));
            },
            Some(_) => {},
        }
        if let Some(star) = text.find('*') {
            return Err(self.error_at(start + star, ErrorKind::Wildcard));
        }
        Ok(Value::Text(Cow::Borrowed(text)))
    }

    /// Parses a run of `[...]` predicates at the cursor.
    pub(super) fn predicates(&mut self, out: &mut Predicates<'a>, already: usize) -> Result<()> {
        while self.peek() == Some(b'[') {
            if out.len() + already >= limits::MAX_PREDICATES {
                return Err(self.error(ErrorKind::LimitExceeded(Limit::Predicates)));
            }
            self.bump();
            out.push(self.predicate()?);
        }
        Ok(())
    }
}

#[inline]
const fn is_op_start(byte: u8) -> bool {
    matches!(byte, b'=' | b'!' | b'~' | b'>' | b'<' | b'\\')
}
