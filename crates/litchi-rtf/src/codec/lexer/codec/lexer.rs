use super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use crate::codec::limits::ParseLimits;
use bumpalo::Bump;
use std::mem::size_of;
use std::ops::Range;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum BinaryDestination {
    /// A `binN` payload whose destination has not been identified yet.
    Generic,
    /// A payload in a `pict` group.
    Picture,
    /// A payload in an `object`/`objdata` group.
    Object,
}

/// RTF Lexer using arena allocation.
pub(crate) struct Lexer<'a> {
    /// Source input
    pub(super) input: &'a str,
    /// Current position in bytes
    pub(in crate::codec::lexer) pos: usize,
    /// Arena allocator for temporary strings
    pub(super) arena: &'a Bump,
    /// Finite token and binary-payload ceilings.
    pub(super) limits: ParseLimits,
    /// Aggregate bytes claimed by accepted `binN` payloads.
    pub(super) total_binary_bytes: usize,
    /// Destination markers for open groups.  Keeping this small and bounded
    /// lets the lexer reject a destination payload before copying it into the
    /// arena while preserving the ordinary generic binary ceiling elsewhere.
    pub(super) binary_destinations: Vec<BinaryDestination>,
}

impl<'a> Lexer<'a> {
    /// Create a new lexer.
    #[cfg(test)]
    #[inline]
    pub(crate) fn new(input: &'a str, arena: &'a Bump) -> Self {
        Self::new_with_limits(input, arena, ParseLimits::default())
    }

    /// Create a lexer with an explicit finite resource profile.
    #[inline]
    pub(crate) fn new_with_limits(input: &'a str, arena: &'a Bump, limits: ParseLimits) -> Self {
        Self {
            input,
            pos: 0,
            arena,
            limits,
            total_binary_bytes: 0,
            binary_destinations: Vec::with_capacity(crate::codec::parser::MAX_GROUP_NESTING_DEPTH),
        }
    }

    /// Return the hard destination ceiling that applies before a `binN`
    /// payload is copied.  Generic binary destinations retain the caller's
    /// configured limit; pictures and objects use their model/writer ceiling.
    pub(super) fn binary_destination_limit(&self) -> usize {
        let mut index = self.binary_destinations.len();
        while index > 0 {
            index -= 1;
            match self.binary_destinations.get(index).copied() {
                Some(BinaryDestination::Picture) => {
                    return self
                        .limits
                        .max_binary_bytes()
                        .min(crate::picture::MAX_PICTURE_WRITE_BYTES);
                },
                Some(BinaryDestination::Object) => {
                    return self
                        .limits
                        .max_binary_bytes()
                        .min(crate::object::MAX_OBJECT_DATA_BYTES);
                },
                Some(BinaryDestination::Generic) => {},
                None => {},
            }
        }
        self.limits.max_binary_bytes()
    }

    pub(super) fn observe_group_token(&mut self, token: &Token<'a>) -> RtfResult<()> {
        match token {
            Token::OpenBrace => {
                if self.binary_destinations.len() >= crate::codec::parser::MAX_GROUP_NESTING_DEPTH {
                    return Err(RtfError::MalformedDocument(
                        "RTF group nesting depth exceeds the safety limit".to_string(),
                    ));
                }
                // The parser owns the format's 32-level structural-depth
                // diagnostic as well as the destination-aware `binN` ceilings.
                // Rejecting the next open brace here prevents malformed input
                // from growing the destination stack and token stream far
                // beyond the depth the recursive parser can safely consume.
                self.binary_destinations.try_reserve(1).map_err(|_err| {
                    RtfError::AllocationFailed {
                        resource: "lexer destination stack",
                        requested: self.binary_destinations.len().saturating_add(1),
                    }
                })?;
                self.binary_destinations.push(BinaryDestination::Generic);
            },
            Token::CloseBrace => {
                self.binary_destinations.pop();
            },
            Token::Control(crate::codec::lexer::ControlWord::Picture) => {
                if let Some(destination) = self.binary_destinations.last_mut() {
                    *destination = BinaryDestination::Picture;
                }
            },
            Token::Control(crate::codec::lexer::ControlWord::Object)
            | Token::Control(crate::codec::lexer::ControlWord::ObjectData) => {
                if let Some(destination) = self.binary_destinations.last_mut() {
                    *destination = BinaryDestination::Object;
                }
            },
            _ => {},
        }
        Ok(())
    }

    /// Tokenize the entire input.
    #[cfg(test)]
    pub(crate) fn tokenize(&mut self) -> RtfResult<Vec<Token<'a>>> {
        self.tokenize_with_spans().map(|(tokens, _)| tokens)
    }

    /// Tokenize while retaining exact source ranges for lossless syntax nodes.
    pub(crate) fn tokenize_with_spans(&mut self) -> RtfResult<(Vec<Token<'a>>, Vec<Range<usize>>)> {
        let mut tokens = Vec::new();
        let mut spans = Vec::new();

        while self.pos < self.input.len() {
            let observed = tokens.len().saturating_add(1);
            if observed > self.limits.max_tokens() {
                return Err(RtfError::LimitExceeded {
                    resource: "lexer tokens",
                    observed,
                    limit: self.limits.max_tokens(),
                });
            }
            // `Vec::try_reserve` grows geometrically, so calling it for every
            // token only repeats the capacity check on the hot path. Keep the
            // same fallible growth and diagnostics, but ask for capacity only
            // when the next push would actually need it.
            if tokens.len() == tokens.capacity() {
                tokens
                    .try_reserve(1)
                    .map_err(|_err| RtfError::AllocationFailed {
                        resource: "lexer tokens",
                        requested: observed.saturating_mul(size_of::<Token<'a>>()),
                    })?;
            }
            if spans.len() == spans.capacity() {
                spans
                    .try_reserve(1)
                    .map_err(|_err| RtfError::AllocationFailed {
                        resource: "lexer token spans",
                        requested: observed.saturating_mul(size_of::<Range<usize>>()),
                    })?;
            }
            let start = self.pos;
            let token = self.next_token()?;
            self.observe_group_token(&token)?;
            let end = self.pos;
            tokens.push(token);
            spans.push(start..end);
        }

        Ok((tokens, spans))
    }
}
