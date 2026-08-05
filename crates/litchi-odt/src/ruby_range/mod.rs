// Structure-preserving range mechanics for ODT ruby annotations.
//
// The owner separates bounded range state from XML scanning while keeping the
// implementation private to the ruby-family codec and preserving its facade.

use super::*;

#[path = "codec.rs"]
mod codec;
#[path = "model.rs"]
mod model;

#[cfg(test)]
#[path = "tests.rs"]
mod tests;

use model::{RangeBoundary, RangeFrame};

pub(super) use model::PendingRubyRange;

pub(super) fn locate_balanced_ruby_ranges(
    xml: &str,
    paragraph_index: usize,
    range: &Range<usize>,
) -> Result<Vec<Span>> {
    codec::locate_balanced_ruby_ranges(xml, paragraph_index, range)
}

#[allow(clippy::too_many_arguments)]
pub(super) fn collect_ruby_text_node(
    xml: &str,
    span: Range<usize>,
    text: &str,
    text_offset: &mut usize,
    range: &Range<usize>,
    stack: &[(Ns, Vec<u8>)],
    value: &RubyAnnotation,
    fragment: &str,
    target_depth: Option<usize>,
    pending: &mut Option<PendingRubyRange>,
) -> Result<Option<String>> {
    codec::collect_ruby_text_node(
        xml,
        span,
        text,
        text_offset,
        range,
        stack,
        value,
        fragment,
        target_depth,
        pending,
    )
}
