use super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use crate::codec::limits::ParseLimits;
use bumpalo::Bump;

/// RTF Lexer using arena allocation.
pub struct Lexer<'a> {
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
    pub fn new(input: &'a str, arena: &'a Bump) -> Self {
        Self::new_with_limits(input, arena, ParseLimits::default())
    }

    /// Create a lexer with an explicit finite resource profile.
    #[inline]
    pub fn new_with_limits(input: &'a str, arena: &'a Bump, limits: ParseLimits) -> Self {
        Self {
            input,
            pos: 0,
            arena,
            limits,
            total_binary_bytes: 0,
        }
    }

    /// Tokenize the entire input.
    pub fn tokenize(&mut self) -> RtfResult<Vec<Token<'a>>> {
        let mut tokens = Vec::new();

        while self.pos < self.input.len() {
            let observed = tokens.len().saturating_add(1);
            if observed > self.limits.max_tokens() {
                return Err(RtfError::LimitExceeded {
                    resource: "lexer tokens",
                    observed,
                    limit: self.limits.max_tokens(),
                });
            }
            let token = self.next_token()?;
            tokens.push(token);
        }

        Ok(tokens)
    }
}
