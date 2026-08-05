//! Structure-preserving range mechanics for ODT ruby annotations.
//!
//! The owner separates bounded range state from XML scanning while keeping the
//! implementation private to the ruby-family codec and preserving its facade.

use super::*;

#[path = "ruby_range/codec.rs"]
mod codec;
#[path = "ruby_range/model.rs"]
mod model;

#[cfg(test)]
#[path = "ruby_range/tests.rs"]
mod tests;

use model::{RangeBoundary, RangeFrame};

pub(super) use codec::{collect_ruby_text_node, locate_balanced_ruby_ranges};
pub(super) use model::PendingRubyRange;
