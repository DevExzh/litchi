//! Namespace-aware `PresentationML` timing XML codec facade.
///
/// The implementation is layered by responsibility:
/// - [`parser`] owns bounded event-stream decoding and recursive timing-tree parsing.
/// - [`semantic`] owns private parser state and typed timing interpretation.
/// - [`writer`] owns compact, model-aware XML serialization.
mod parser;
mod semantic;
mod writer;

use super::super::model::{EffectInstance, Sequence, TimingChild, TimingTree};
use crate::animations::Limits;
use crate::{Error, Result};

/// A fallible output sink that never grows past the caller's selected XML
/// ceiling.  The animation writer is intentionally kept allocation-light, but
/// every append still checks the logical size before reserving the next chunk.
pub(super) struct BoundedXml {
    value: String,
    limit: usize,
    failure: Option<Error>,
}

impl BoundedXml {
    pub(super) fn new(limit: usize) -> Self {
        Self {
            value: String::new(),
            limit,
            failure: None,
        }
    }

    pub(super) fn push_str(&mut self, value: &str) {
        if self.failure.is_some() {
            return;
        }
        let Some(next) = self.value.len().checked_add(value.len()) else {
            self.failure = Some(Error::Limit {
                kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlBytes),
                observed: usize::MAX,
                maximum: self.limit,
            });
            return;
        };
        if next > self.limit {
            self.failure = Some(Error::Limit {
                kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlBytes),
                observed: next,
                maximum: self.limit,
            });
            return;
        }
        let spare = self.value.capacity().saturating_sub(self.value.len());
        if value.len() > spare
            && let Err(source) = self.value.try_reserve_exact(value.len())
        {
            self.failure = Some(Error::Allocation {
                resource: "animation XML output",
                source,
            });
            return;
        }
        self.value.push_str(value);
    }

    pub(super) fn push(&mut self, value: char) {
        let mut encoded = [0; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }

    pub(super) fn finish(self) -> Result<String> {
        match self.failure {
            Some(error) => Err(error),
            None => Ok(self.value),
        }
    }
}

pub(super) fn parse_processed_timing(
    xml: &[u8],
    require_valid_targets: bool,
    limits: Limits,
) -> Result<Sequence> {
    parser::parse_processed_timing(xml, require_valid_targets, limits)
}

pub(super) fn parse_recursive_timing_tree(xml: &str, limits: Limits) -> Result<TimingTree> {
    parser::parse_recursive_timing_tree(xml, limits)
}

pub(super) fn write_animation_xml(
    xml: &mut BoundedXml,
    anim: &EffectInstance,
    tn_id: &mut u32,
    interactive_trigger: Option<u32>,
) {
    writer::write_animation_xml(xml, anim, tn_id, interactive_trigger);
}

pub(super) fn write_timing_child(xml: &mut BoundedXml, child: &TimingChild) {
    writer::write_timing_child(xml, child);
}
