//! Bounded `\\binN` payload consumption.

use super::super::lexer::Lexer;
use super::super::model::Token;
use crate::codec::error::{RtfError, RtfResult};
use std::borrow::Cow;

impl<'a> Lexer<'a> {
    /// Consume and arena-copy the exact byte payload declared by `\\binN`.
    pub(super) fn parse_binary_payload(&mut self, declared_size: i32) -> RtfResult<Token<'a>> {
        let size = usize::try_from(declared_size).map_err(|_err| {
            RtfError::MalformedDocument("RTF binary length cannot be negative".to_string())
        })?;
        if size > self.limits.max_binary_bytes() {
            return Err(RtfError::LimitExceeded {
                resource: "binary payload bytes",
                observed: size,
                limit: self.limits.max_binary_bytes(),
            });
        }
        let total_binary_bytes =
            self.total_binary_bytes
                .checked_add(size)
                .ok_or(RtfError::LimitExceeded {
                    resource: "aggregate binary payload bytes",
                    observed: usize::MAX,
                    limit: self.limits.max_total_binary_bytes(),
                })?;
        if total_binary_bytes > self.limits.max_total_binary_bytes() {
            return Err(RtfError::LimitExceeded {
                resource: "aggregate binary payload bytes",
                observed: total_binary_bytes,
                limit: self.limits.max_total_binary_bytes(),
            });
        }

        // Validate the entire declared payload before allocating it. A corrupt
        // `bin2147483647` near EOF must report truncation instead of first
        // requesting a multi-gigabyte allocation.
        let payload_start = self.pos;
        let mut payload_end = payload_start;
        for _ in 0..size {
            let ch = self
                .input
                .get(payload_end..)
                .and_then(|remaining| remaining.chars().next())
                .ok_or(RtfError::UnexpectedEof)?;
            u8::try_from(u32::from(ch)).map_err(|_err| {
                RtfError::MalformedDocument("RTF binary payload is not byte-preserving".to_string())
            })?;
            payload_end += ch.len_utf8();
        }

        let payload = self.source_range(payload_start, payload_end)?;
        let allocated = self.arena.alloc_slice_fill_copy(size, 0u8);
        for (slot, ch) in allocated.iter_mut().zip(payload.chars()) {
            *slot = u8::try_from(u32::from(ch)).map_err(|_err| {
                RtfError::MalformedDocument("RTF binary payload is not byte-preserving".to_string())
            })?;
        }
        self.pos = payload_end;
        self.total_binary_bytes = total_binary_bytes;
        Ok(Token::Binary(Cow::Borrowed(allocated)))
    }
}
