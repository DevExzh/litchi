//! Semantic and ordering validation for timing containers.

use super::model::{NodeParts, SubEffectParts};
use crate::animation::types::{
    TimeNodeBehavior, TimeNodeKind, TimeNodeProperty, TimePropertyListContext,
    TimeSubEffectBehavior, has_valid_time_effect_properties,
};
use crate::package::{Error, Result};

pub(super) fn set_once<T>(slot: &mut Option<T>, value: T, field: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "ExtTimeNode contains multiple {field} records"
        )));
    }
    Ok(())
}

pub(super) fn set_time_node_behavior(
    slot: &mut Option<TimeNodeBehavior>,
    behavior: TimeNodeBehavior,
) -> Result<()> {
    set_once(slot, behavior, "animation behavior")
}

pub(super) fn set_subeffect_behavior(
    slot: &mut Option<TimeSubEffectBehavior>,
    behavior: TimeSubEffectBehavior,
) -> Result<()> {
    if slot.replace(behavior).is_some() {
        return Err(Error::InvalidFormat(
            "SubEffectContainer contains multiple animation behaviors".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn ensure_order(last_rank: &mut u8, rank: u8, container: &str) -> Result<()> {
    if rank < *last_rank {
        return Err(Error::InvalidFormat(format!(
            "{container} children are not in canonical order"
        )));
    }
    *last_rank = rank;
    Ok(())
}

pub(super) fn validate_extended_node(kind: TimeNodeKind, parts: &NodeParts) -> Result<()> {
    if parts.behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "animation behaviors require a behavior time node".to_string(),
        ));
    }
    if parts.visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "standalone visual targets require a media time node".to_string(),
        ));
    }
    if parts.sequence_data.is_some() && kind != TimeNodeKind::Sequential {
        return Err(Error::InvalidFormat(
            "sequence data requires a sequential time node".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_sub_effect(kind: TimeNodeKind, parts: &SubEffectParts) -> Result<()> {
    if parts.behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "subeffect behavior requires a behavior time node".to_string(),
        ));
    }
    if parts.visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "subeffect visual target requires a media time node".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_property_context(id: u16, context: TimePropertyListContext) -> Result<()> {
    if matches!(context, TimePropertyListContext::TimeNode) && matches!(id, 0x05 | 0x06) {
        return Err(Error::InvalidFormat(
            "subeffect-only property on time node".to_string(),
        ));
    }
    if matches!(context, TimePropertyListContext::SubEffect)
        && matches!(id, 0x09..=0x0B | 0x0F..=0x14 | 0x1A)
    {
        return Err(Error::InvalidFormat(
            "time-node-only property on subeffect".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_properties(
    properties: &[TimeNodeProperty],
    _context: TimePropertyListContext,
) -> Result<()> {
    if properties
        .iter()
        .any(|property| matches!(property, TimeNodeProperty::EventFilter(_)))
        && !properties.iter().any(|property| {
            matches!(
                property,
                TimeNodeProperty::EffectNodeType(
                    crate::animation::types::TimeEffectNodeType::InteractiveSequence
                )
            )
        })
    {
        return Err(Error::InvalidFormat(
            "event filter requires an interactive sequence".to_string(),
        ));
    }
    if !has_valid_time_effect_properties(properties) {
        return Err(Error::InvalidFormat(
            "invalid effect ID, type, or direction combination".to_string(),
        ));
    }
    Ok(())
}
