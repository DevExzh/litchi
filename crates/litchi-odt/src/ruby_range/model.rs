//! Bounded structural state used by the ODT ruby-range codec.

use super::Ns;

pub(super) struct RangeFrame {
    pub(super) id: usize,
    pub(super) namespace: Ns,
    pub(super) local: Vec<u8>,
}

#[derive(Clone, Copy)]
pub(super) struct RangeBoundary {
    pub(super) offset: usize,
    pub(super) container_id: usize,
    pub(super) ruby_epoch: usize,
}

pub(super) struct PendingRubyRange {
    pub(super) xml_start: usize,
    pub(super) xml_end: usize,
    pub(super) prefix: String,
    pub(super) selected: String,
    pub(super) stack: Vec<(Ns, Vec<u8>)>,
}
