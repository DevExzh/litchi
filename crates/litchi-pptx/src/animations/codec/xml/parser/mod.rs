//! Bounded PresentationML timing parser facade.

mod semantic;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

use super::super::super::model::{Sequence, TimingTree};
use crate::Result;

pub(super) fn parse_recursive_timing_tree(xml: &str) -> Result<TimingTree> {
    wire::parse_recursive_timing_tree(xml)
}

pub(super) fn parse_processed_timing(xml: &[u8], require_valid_targets: bool) -> Result<Sequence> {
    wire::parse_processed_timing(xml, require_valid_targets)
}
