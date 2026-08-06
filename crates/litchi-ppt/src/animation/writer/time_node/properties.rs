//! Time-node property-list semantic checks and wire encoding.

use super::validation::{
    validate_event_filter, validate_property_list, validate_time_property,
    validate_time_property_context,
};
use crate::animation::types::{
    TimeEffectNodeType, TimeEffectType, TimeMasterRelation, TimeNodeProperty, TimeNodePropertyList,
    TimePropertyListContext,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use std::collections::HashSet;

#[cfg(test)]
use crate::animation::types::{has_valid_time_effect_properties, is_valid_time_filter};

/// Serialize a typed `TimePropertyList4TimeNodeContainer`.
pub fn write_time_node_property_list(
    list: &TimeNodePropertyList,
    context: TimePropertyListContext,
) -> Result<Vec<u8>> {
    validate_property_list(list)?;
    let mut seen = HashSet::with_capacity(list.properties.len());
    let has_interactive_sequence = list.properties.iter().any(|property| {
        matches!(
            property,
            TimeNodeProperty::EffectNodeType(TimeEffectNodeType::InteractiveSequence)
        )
    });
    let mut children = Vec::new();
    for property in &list.properties {
        validate_time_property(property)?;
        let (id, data) = encode_time_node_property(property)?;
        if !seen.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate time property {id:#X}"
            )));
        }
        validate_time_property_context(id, context)?;
        validate_event_filter(property, has_interactive_sequence)?;
        let length = u32::try_from(data.len()).map_err(|_| {
            Error::InvalidFormat("time property exceeds 4 GiB record limit".to_string())
        })?;
        children.extend(super::super::support::create_record_header(
            RecordType::TimeVariant,
            0,
            id,
            length,
        ));
        children.extend(data);
    }
    super::super::support::wrap_record(RecordType::TimePropertyList, 0x0F, 0, children)
}

fn encode_time_node_property(property: &TimeNodeProperty) -> Result<(u16, Vec<u8>)> {
    let integer = |value: i32| {
        let mut data = vec![1];
        data.extend(value.to_le_bytes());
        data
    };
    let boolean = |value: bool| vec![0, u8::from(value)];
    let string = |value: &str| {
        let mut data = Vec::with_capacity(1 + value.len().saturating_mul(2));
        data.push(3);
        data.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
        data
    };
    Ok(match property {
        TimeNodeProperty::DisplayHidden(value) => (0x02, integer(i32::from(*value))),
        TimeNodeProperty::MasterRelation(value) => (
            0x05,
            integer(match value {
                TimeMasterRelation::DoNotStart => 0,
                TimeMasterRelation::StartWithMaster => 2,
            }),
        ),
        TimeNodeProperty::SubType => (0x06, integer(1)),
        TimeNodeProperty::EffectId(value) => (0x09, integer(*value)),
        TimeNodeProperty::EffectDirection(value) => (0x0A, integer(*value)),
        TimeNodeProperty::EffectType(value) => (
            0x0B,
            integer(match value {
                TimeEffectType::Entrance => 1,
                TimeEffectType::Exit => 2,
                TimeEffectType::Emphasis => 3,
                TimeEffectType::MotionPath => 4,
                TimeEffectType::ActionVerb => 5,
                TimeEffectType::MediaCommand => 6,
            }),
        ),
        TimeNodeProperty::AfterEffect(value) => (0x0D, boolean(*value)),
        TimeNodeProperty::SlideCount(value) => (0x0F, integer(*value)),
        TimeNodeProperty::TimeFilter(value) => (0x10, string(value)),
        TimeNodeProperty::EventFilter(value) => (0x11, string(value)),
        TimeNodeProperty::HideWhenStopped(value) => (0x12, boolean(*value)),
        TimeNodeProperty::GroupId(value) => (0x13, integer(*value)),
        TimeNodeProperty::EffectNodeType(value) => (
            0x14,
            integer(match value {
                TimeEffectNodeType::ClickEffect => 1,
                TimeEffectNodeType::WithPrevious => 2,
                TimeEffectNodeType::AfterPrevious => 3,
                TimeEffectNodeType::MainSequence => 4,
                TimeEffectNodeType::InteractiveSequence => 5,
                TimeEffectNodeType::ClickParallel => 6,
                TimeEffectNodeType::WithGroup => 7,
                TimeEffectNodeType::AfterGroup => 8,
                TimeEffectNodeType::TimingRoot => 9,
            }),
        ),
        TimeNodeProperty::PlaceholderNode(value) => (0x15, boolean(*value)),
        TimeNodeProperty::MediaVolume(value) => {
            let mut data = vec![2];
            data.extend(value.to_le_bytes());
            (0x16, data)
        },
        TimeNodeProperty::MediaMute(value) => (0x17, boolean(*value)),
        TimeNodeProperty::ZoomToFullScreen(value) => (0x1A, boolean(*value)),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn property_validation_remains_contextual() {
        let list = TimeNodePropertyList {
            properties: vec![TimeNodeProperty::MasterRelation(
                TimeMasterRelation::StartWithMaster,
            )],
        };
        assert!(write_time_node_property_list(&list, TimePropertyListContext::TimeNode).is_err());
        assert!(write_time_node_property_list(&list, TimePropertyListContext::SubEffect).is_ok());
    }

    #[test]
    fn effect_property_validation_is_preserved() {
        let list = TimeNodePropertyList {
            properties: vec![TimeNodeProperty::EffectDirection(1)],
        };
        assert!(!has_valid_time_effect_properties(&list.properties));
        assert!(write_time_node_property_list(&list, TimePropertyListContext::SubEffect).is_err());
    }

    #[test]
    fn time_filter_encoding_stays_utf16() {
        let list = TimeNodePropertyList {
            properties: vec![TimeNodeProperty::TimeFilter("0.5,1.0".to_string())],
        };
        let bytes = write_time_node_property_list(&list, TimePropertyListContext::TimeNode)
            .expect("valid time filter");
        assert!(bytes.windows(2).any(|pair| pair == [b'0', 0]));
        assert!(is_valid_time_filter("0.5,1.0"));
    }
}
