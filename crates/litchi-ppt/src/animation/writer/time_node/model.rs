//! Semantic timing-node views used by validation and record codecs.

use crate::animation::types::{ExtendedTimeNode, TimeNodeKind, TimeSubEffect};

/// A borrowed extended-node view with format defaults resolved once.
pub(super) struct NodeView<'a> {
    node: &'a ExtendedTimeNode,
    effective_kind: TimeNodeKind,
}

impl<'a> NodeView<'a> {
    pub(super) fn new(node: &'a ExtendedTimeNode) -> Self {
        Self {
            node,
            effective_kind: node.atom.node_type.unwrap_or(TimeNodeKind::Parallel),
        }
    }

    pub(super) fn source(&self) -> &'a ExtendedTimeNode {
        self.node
    }

    pub(super) fn effective_kind(&self) -> TimeNodeKind {
        self.effective_kind
    }
}

/// A borrowed subordinate-effect view retaining the explicit node kind.
pub(super) struct SubEffectView<'a> {
    sub_effect: &'a TimeSubEffect,
}

impl<'a> SubEffectView<'a> {
    pub(super) fn new(sub_effect: &'a TimeSubEffect) -> Self {
        Self { sub_effect }
    }

    pub(super) fn source(&self) -> &'a TimeSubEffect {
        self.sub_effect
    }

    pub(super) fn explicit_kind(&self) -> Option<TimeNodeKind> {
        self.sub_effect.atom.node_type
    }
}
