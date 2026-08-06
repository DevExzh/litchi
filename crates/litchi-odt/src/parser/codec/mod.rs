//! Layered XML codecs for ODT-specific parser owners.
//!
//! The facade keeps the parser package boundary small while the wire pass,
//! semantic state, validation policy, and focused codec tests remain in their
//! own owners.

mod semantic;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

pub(super) use wire::{
    correlate_change_ranges, parse_change_declarations, parse_comments, parse_sections,
};
