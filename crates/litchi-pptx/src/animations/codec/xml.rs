//! Namespace-aware PresentationML timing XML codec facade.
///
/// The implementation is layered by responsibility:
/// - [`parser`] owns bounded event-stream decoding and recursive timing-tree parsing.
/// - [`semantic`] owns private parser state and typed timing interpretation.
/// - [`writer`] owns compact, model-aware XML serialization.
mod parser;
mod semantic;
mod writer;

use super::super::model::{EffectInstance, Sequence, TimingChild, TimingTree};
use crate::Result;

pub(super) fn parse_processed_timing(xml: &[u8], require_valid_targets: bool) -> Result<Sequence> {
    parser::parse_processed_timing(xml, require_valid_targets)
}

pub(super) fn parse_recursive_timing_tree(xml: &str) -> Result<TimingTree> {
    parser::parse_recursive_timing_tree(xml)
}

pub(super) fn write_animation_xml(
    xml: &mut String,
    anim: &EffectInstance,
    tn_id: &mut u32,
    interactive_trigger: Option<u32>,
) {
    writer::write_animation_xml(xml, anim, tn_id, interactive_trigger);
}

pub(super) fn write_timing_child(xml: &mut String, child: &TimingChild) {
    writer::write_timing_child(xml, child);
}
