use super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use crate::codec::limits::ParseLimits;
use bumpalo::Bump;
use std::mem::size_of;
use std::ops::Range;

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
        }
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
            tokens
                .try_reserve(1)
                .map_err(|_err| RtfError::AllocationFailed {
                    resource: "lexer tokens",
                    requested: observed.saturating_mul(size_of::<Token<'a>>()),
                })?;
            spans
                .try_reserve(1)
                .map_err(|_err| RtfError::AllocationFailed {
                    resource: "lexer token spans",
                    requested: observed.saturating_mul(size_of::<Range<usize>>()),
                })?;
            let start = self.pos;
            let token = self.next_token()?;
            let end = self.pos;
            tokens.push(token);
            spans.push(start..end);
        }

        Ok((tokens, spans))
    }
}
