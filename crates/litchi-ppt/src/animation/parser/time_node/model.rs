//! Semantic state assembled while decoding timing records.

use crate::animation::types::{
    ExtendedTimeNode, TimeCondition, TimeIterateData, TimeModifier, TimeNodeAtom, TimeNodeBehavior,
    TimeNodeKind, TimeNodePropertyList, TimeSequenceData, TimeSubEffect, TimeSubEffectBehavior,
    TimeVisualElement,
};

/// The property-use flags in a `TimeNodeAtom`.
///
/// [MS-PPT] names the fourth bit `fGroupingTypeProperty`; keeping it typed
/// here prevents the record parser from scattering raw bit masks while the
/// public snapshot model continues to represent explicit values with `Option`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "each bool mirrors one independent fXxxProperty bit of the `TimeNodeAtom` flags dword, matching the MS-PPT bitfield layout"
)]
pub(super) struct AtomFlags {
    pub(super) fill_property: bool,
    pub(super) restart_property: bool,
    pub(super) grouping_type_property: bool,
    pub(super) duration_property: bool,
    /// Reserved bit 2 and the upper 27 bits, retained for spec accounting.
    pub(super) reserved: u32,
}

impl AtomFlags {
    pub(super) fn from_raw(raw: u32) -> Self {
        Self {
            fill_property: raw & (1 << 0) != 0,
            restart_property: raw & (1 << 1) != 0,
            grouping_type_property: raw & (1 << 3) != 0,
            duration_property: raw & (1 << 4) != 0,
            reserved: raw & !0x1B,
        }
    }
}

/// Fields collected from an `ExtTimeNodeContainer` after structural parsing.
#[derive(Debug, Default)]
pub(super) struct NodeParts {
    pub(super) behavior: Option<TimeNodeBehavior>,
    pub(super) visual_target: Option<TimeVisualElement>,
    pub(super) iterate_data: Option<TimeIterateData>,
    pub(super) sequence_data: Option<TimeSequenceData>,
    pub(super) begin_conditions: Vec<TimeCondition>,
    pub(super) end_conditions: Vec<TimeCondition>,
    pub(super) end_sync_condition: Option<TimeCondition>,
    pub(super) modifiers: Vec<TimeModifier>,
    pub(super) sub_effects: Vec<TimeSubEffect>,
    pub(super) children: Vec<ExtendedTimeNode>,
}

impl NodeParts {
    pub(super) fn finish(
        self,
        atom: TimeNodeAtom,
        properties: Option<TimeNodePropertyList>,
    ) -> ExtendedTimeNode {
        ExtendedTimeNode {
            atom,
            properties,
            behavior: self.behavior,
            visual_target: self.visual_target,
            iterate_data: self.iterate_data,
            sequence_data: self.sequence_data,
            begin_conditions: self.begin_conditions,
            end_conditions: self.end_conditions,
            end_sync_condition: self.end_sync_condition,
            modifiers: self.modifiers,
            sub_effects: self.sub_effects,
            children: self.children,
        }
    }
}

/// Fields collected from a `SubEffectContainer` after structural parsing.
#[derive(Debug, Default)]
pub(super) struct SubEffectParts {
    pub(super) behavior: Option<TimeSubEffectBehavior>,
    pub(super) visual_target: Option<TimeVisualElement>,
    pub(super) begin_conditions: Vec<TimeCondition>,
    pub(super) end_conditions: Vec<TimeCondition>,
    pub(super) modifiers: Vec<TimeModifier>,
}

impl SubEffectParts {
    pub(super) fn finish(
        self,
        atom: TimeNodeAtom,
        properties: Option<TimeNodePropertyList>,
    ) -> TimeSubEffect {
        TimeSubEffect {
            atom,
            properties,
            behavior: self.behavior,
            visual_target: self.visual_target,
            begin_conditions: self.begin_conditions,
            end_conditions: self.end_conditions,
            modifiers: self.modifiers,
        }
    }
}

/// The effective kind used by the container-level semantic checks.
pub(super) fn effective_kind(atom: &TimeNodeAtom) -> TimeNodeKind {
    atom.node_type.unwrap_or(TimeNodeKind::Parallel)
}
