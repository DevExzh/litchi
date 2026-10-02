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
    /// Repairs applied while tokenizing.
    pub(super) tolerance: crate::leniency::ToleranceReport,
    /// Repair posture for opt-in lexical repairs.
    pub(super) leniency: crate::leniency::Leniency,
    /// Whether the input is a complete document whose root group ends the
    /// token stream, rather than a fragment of one.
    document: bool,
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
            tolerance: crate::leniency::ToleranceReport::default(),
            leniency: crate::leniency::Leniency::Strict,
            document: false,
        }
    }

    /// Apply `leniency`'s opt-in lexical repairs to a document fragment or a
    /// complete source that is not tokenized as one root group.
    pub(crate) fn with_leniency(mut self, leniency: crate::leniency::Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    /// Tokenize a complete document under `leniency`.
    ///
    /// Reading stops at the root group's closing brace, as Word and
    /// LibreOffice stop; anything but whitespace or NUL padding after it is
    /// ignored and recorded as
    /// [`SyntaxDefect::TrailingContentAfterRoot`](crate::leniency::SyntaxDefect::TrailingContentAfterRoot).
    /// Groups still open at end of input are closed when `leniency` allows.
    pub(crate) fn for_document(mut self, leniency: crate::leniency::Leniency) -> Self {
        self.leniency = leniency;
        self.document = true;
        self
    }

    /// Move out the repairs applied while tokenizing.
    pub(crate) fn take_tolerance_report(&mut self) -> crate::leniency::ToleranceReport {
        std::mem::take(&mut self.tolerance)
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
        let mut opened_root = false;

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
            let closes_root = opened_root
                && matches!(token, Token::CloseBrace)
                && self.binary_destinations.is_empty();
            opened_root |= matches!(token, Token::OpenBrace);
            tokens.push(token);
            spans.push(start..end);
            if self.document && closes_root {
                self.ignore_content_after_root();
                break;
            }
        }

        if self.document && self.leniency.tolerates_unclosed_groups() {
            self.close_unclosed_groups(&mut tokens, &mut spans)?;
        }

        Ok((tokens, spans))
    }

    /// Stop at the root group's closing brace, recording any content other
    /// than whitespace or NUL padding that follows it.
    fn ignore_content_after_root(&mut self) {
        let trailing = self
            .input
            .get(self.pos..)
            .unwrap_or_default()
            .char_indices()
            .find(|(_, character)| *character != '\0' && !character.is_whitespace());
        if let Some((offset, _)) = trailing {
            self.tolerance.record(
                crate::leniency::SyntaxDefect::TrailingContentAfterRoot,
                self.pos.saturating_add(offset),
            );
        }
        self.pos = self.input.len();
    }

    /// Close every group still open at end of input with a zero-width
    /// closing brace, recording one repair per group.
    fn close_unclosed_groups(
        &mut self,
        tokens: &mut Vec<Token<'a>>,
        spans: &mut Vec<Range<usize>>,
    ) -> RtfResult<()> {
        let open = self.binary_destinations.len();
        if open == 0 {
            return Ok(());
        }
        let observed = tokens.len().saturating_add(open);
        if observed > self.limits.max_tokens() {
            return Err(RtfError::LimitExceeded {
                resource: "lexer tokens",
                observed,
                limit: self.limits.max_tokens(),
            });
        }
        let allocation_failed = |_err| RtfError::AllocationFailed {
            resource: "lexer tokens",
            requested: observed.saturating_mul(size_of::<Token<'a>>()),
        };
        tokens.try_reserve(open).map_err(allocation_failed)?;
        spans.try_reserve(open).map_err(allocation_failed)?;
        let end = self.input.len();
        for _ in 0..open {
            tokens.push(Token::CloseBrace);
            spans.push(end..end);
            self.tolerance
                .record(crate::leniency::SyntaxDefect::UnclosedGroup, end);
        }
        self.binary_destinations.clear();
        Ok(())
    }
}
