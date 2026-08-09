//! Semantic validation for timing-node containers and properties.

use super::model::{NodeView, SubEffectView};
use crate::animation::types::{
    TimeConditionType, TimeNodeKind, TimeNodeProperty, TimeNodePropertyList,
    TimePropertyListContext, has_valid_time_effect_properties, is_valid_time_filter,
};
use crate::package::{Error, Result};

pub(super) fn validate_extended_time_node(view: &NodeView<'_>) -> Result<()> {
    let node = view.source();
    let kind = view.effective_kind();
    if node.behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "animation behaviors require a behavior time node".to_string(),
        ));
    }
    if node.visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "standalone visual targets require a media time node".to_string(),
        ));
    }
    if node.sequence_data.is_some() && kind != TimeNodeKind::Sequential {
        return Err(Error::InvalidFormat(
            "sequence data requires a sequential time node".to_string(),
        ));
    }
    for condition in &node.begin_conditions {
        let valid = condition.condition_type == TimeConditionType::Begin
            || (kind == TimeNodeKind::Sequential
                && condition.condition_type == TimeConditionType::Next);
        if !valid {
            return Err(Error::InvalidFormat(
                "begin-condition arrays may contain only begin conditions, or next conditions on sequential nodes"
                    .to_string(),
            ));
        }
    }
    for condition in &node.end_conditions {
        let valid = condition.condition_type == TimeConditionType::End
            || (kind == TimeNodeKind::Sequential
                && condition.condition_type == TimeConditionType::Previous);
        if !valid {
            return Err(Error::InvalidFormat(
                "end-condition arrays may contain only end conditions, or previous conditions on sequential nodes"
                    .to_string(),
            ));
        }
    }
    if node
        .end_sync_condition
        .as_ref()
        .is_some_and(|condition| condition.condition_type != TimeConditionType::EndSync)
    {
        return Err(Error::InvalidFormat(
            "end-sync condition must use the EndSync condition type".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_sub_effect(view: &SubEffectView<'_>) -> Result<TimeNodeKind> {
    let sub_effect = view.source();
    let kind = match view.explicit_kind() {
        Some(TimeNodeKind::Behavior) => TimeNodeKind::Behavior,
        Some(TimeNodeKind::Media) => TimeNodeKind::Media,
        _ => {
            return Err(Error::InvalidFormat(
                "subeffect time-node type must explicitly be Behavior or Media".to_string(),
            ));
        },
    };
    if sub_effect.behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "subeffect behavior requires a behavior time node".to_string(),
        ));
    }
    if sub_effect.visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "subeffect visual target requires a media time node".to_string(),
        ));
    }
    if sub_effect
        .begin_conditions
        .iter()
        .any(|condition| condition.condition_type != TimeConditionType::Begin)
        || sub_effect
            .end_conditions
            .iter()
            .any(|condition| condition.condition_type != TimeConditionType::End)
    {
        return Err(Error::InvalidFormat(
            "subeffect condition arrays must contain only their matching Begin or End type"
                .to_string(),
        ));
    }
    Ok(kind)
}

pub(super) fn validate_property_list(properties: &TimeNodePropertyList) -> Result<()> {
    if !has_valid_time_effect_properties(&properties.properties) {
        return Err(Error::InvalidFormat(
            "invalid effect ID, type, or direction combination".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_time_property(property: &TimeNodeProperty) -> Result<()> {
    match property {
        TimeNodeProperty::TimeFilter(value) if !is_valid_time_filter(value) => {
            Err(Error::InvalidFormat("invalid time filter".to_string()))
        },
        TimeNodeProperty::EventFilter(value) if value != "cancelBubble" => Err(
            Error::InvalidFormat("event filter must be cancelBubble".to_string()),
        ),
        TimeNodeProperty::MediaVolume(value)
            if !value.is_finite() || !(0.0..=100_000.0).contains(value) =>
        {
            Err(Error::InvalidFormat(
                "media volume out of range".to_string(),
            ))
        },
        TimeNodeProperty::DisplayHidden(_)
        | TimeNodeProperty::MasterRelation(_)
        | TimeNodeProperty::SubType
        | TimeNodeProperty::EffectId(_)
        | TimeNodeProperty::EffectDirection(_)
        | TimeNodeProperty::EffectType(_)
        | TimeNodeProperty::AfterEffect(_)
        | TimeNodeProperty::SlideCount(_)
        | TimeNodeProperty::TimeFilter(_)
        | TimeNodeProperty::EventFilter(_)
        | TimeNodeProperty::HideWhenStopped(_)
        | TimeNodeProperty::GroupId(_)
        | TimeNodeProperty::EffectNodeType(_)
        | TimeNodeProperty::PlaceholderNode(_)
        | TimeNodeProperty::MediaVolume(_)
        | TimeNodeProperty::MediaMute(_)
        | TimeNodeProperty::ZoomToFullScreen(_) => Ok(()),
    }
}

pub(super) fn validate_event_filter(
    property: &TimeNodeProperty,
    has_interactive_sequence: bool,
) -> Result<()> {
    if matches!(property, TimeNodeProperty::EventFilter(_)) && !has_interactive_sequence {
        return Err(Error::InvalidFormat(
            "event filter requires an interactive sequence".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_time_property_context(
    id: u16,
    context: TimePropertyListContext,
) -> Result<()> {
    let invalid = match context {
        TimePropertyListContext::TimeNode => matches!(id, 0x05 | 0x06),
        TimePropertyListContext::SubEffect => {
            matches!(id, 0x09..=0x0B | 0x0F..=0x14 | 0x1A)
        },
    };
    if invalid {
        return Err(Error::InvalidFormat(format!(
            "time property {id:#X} is invalid in {context:?} context"
        )));
    }
    Ok(())
}
