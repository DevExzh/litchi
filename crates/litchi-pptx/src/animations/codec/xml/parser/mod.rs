//! Bounded PresentationML timing parser facade.

mod semantic;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

use super::super::super::model::{Sequence, TimingTree};
use crate::Result;
use crate::animations::Limits;

pub(super) fn parse_recursive_timing_tree(xml: &str, limits: Limits) -> Result<TimingTree> {
    wire::parse_recursive_timing_tree(xml, limits)
}

pub(super) fn parse_processed_timing(
    xml: &[u8],
    require_valid_targets: bool,
    limits: Limits,
) -> Result<Sequence> {
    wire::parse_processed_timing(xml, require_valid_targets, limits)
}
