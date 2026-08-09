//! Time-node property-list decoding.

use super::super::behavior::require_time_variant_payload;
use super::super::support::{parse_bool1, read_f32, read_i32, require_container};
use super::validation::{validate_properties, validate_property_context};
use crate::animation::types::{
    TimeEffectNodeType, TimeEffectType, TimeMasterRelation, TimeNodeProperty, TimeNodePropertyList,
    TimePropertyListContext, is_valid_time_filter,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

/// Parse a time-node property list in its containing-node context.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_node_property_list(
    record: &Record,
    context: TimePropertyListContext,
) -> Result<TimeNodePropertyList> {
    require_container(record, RecordType::TimePropertyList, 0, "TimePropertyList")?;
    let mut seen = HashSet::with_capacity(record.children.len());
    let mut properties = Vec::with_capacity(record.children.len());
    for child in &record.children {
        if child.record_type != RecordType::TimeVariant || child.version != 0 {
            return Err(Error::InvalidFormat(
                "invalid TimePropertyList child".to_string(),
            ));
        }
        let id = child.instance;
        if !seen.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate time property {id:#X}"
            )));
        }
        validate_property_context(id, context)?;
        properties.push(parse_time_node_property(child)?);
    }
    validate_properties(&properties, context)?;
    Ok(TimeNodePropertyList { properties })
}

fn parse_time_node_property(record: &Record) -> Result<TimeNodeProperty> {
    require_time_variant_payload(record)?;
    let data = &record.data;
    let int = || -> Result<i32> {
        if data.len() != 5 || data[0] != 1 {
            return Err(Error::InvalidFormat(
                "invalid integer time variant".to_string(),
            ));
        }
        Ok(read_i32(data, 1))
    };
    let boolean = || -> Result<bool> {
        if data.len() != 2 || data[0] != 0 {
            return Err(Error::InvalidFormat(
                "invalid boolean time variant".to_string(),
            ));
        }
        parse_bool1(data[1], "TimeVariant.boolValue")
    };
    let string = || -> Result<String> {
        if data.len() < 3 || data.len() % 2 != 1 || data[0] != 3 {
            return Err(Error::InvalidFormat(
                "invalid string time variant".to_string(),
            ));
        }
        String::from_utf16(
            &data[1..]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .collect::<Vec<_>>(),
        )
        .map_err(|_err| Error::InvalidFormat("invalid UTF-16 time variant".to_string()))
    };
    Ok(match record.instance {
        0x02 => TimeNodeProperty::DisplayHidden(match int()? {
            0 => false,
            1 => true,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid display type {value}"
                )));
            },
        }),
        0x05 => TimeNodeProperty::MasterRelation(match int()? {
            0 => TimeMasterRelation::DoNotStart,
            2 => TimeMasterRelation::StartWithMaster,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid master relation {value}"
                )));
            },
        }),
        0x06 if int()? == 1 => TimeNodeProperty::SubType,
        0x06 => return Err(Error::InvalidFormat("invalid time subtype".to_string())),
        0x09 => TimeNodeProperty::EffectId(int()?),
        0x0A => TimeNodeProperty::EffectDirection(int()?),
        0x0B => TimeNodeProperty::EffectType(match int()? {
            1 => TimeEffectType::Entrance,
            2 => TimeEffectType::Exit,
            3 => TimeEffectType::Emphasis,
            4 => TimeEffectType::MotionPath,
            5 => TimeEffectType::ActionVerb,
            6 => TimeEffectType::MediaCommand,
            value => return Err(Error::InvalidFormat(format!("invalid effect type {value}"))),
        }),
        0x0D => TimeNodeProperty::AfterEffect(boolean()?),
        0x0F => TimeNodeProperty::SlideCount(int()?),
        0x10 => {
            let value = string()?;
            if !is_valid_time_filter(&value) {
                return Err(Error::InvalidFormat("invalid time filter".to_string()));
            }
            TimeNodeProperty::TimeFilter(value)
        },
        0x11 => {
            let value = string()?;
            if value != "cancelBubble" {
                return Err(Error::InvalidFormat("invalid event filter".to_string()));
            }
            TimeNodeProperty::EventFilter(value)
        },
        0x12 => TimeNodeProperty::HideWhenStopped(boolean()?),
        0x13 => TimeNodeProperty::GroupId(int()?),
        0x14 => TimeNodeProperty::EffectNodeType(match int()? {
            1 => TimeEffectNodeType::ClickEffect,
            2 => TimeEffectNodeType::WithPrevious,
            3 => TimeEffectNodeType::AfterPrevious,
            4 => TimeEffectNodeType::MainSequence,
            5 => TimeEffectNodeType::InteractiveSequence,
            6 => TimeEffectNodeType::ClickParallel,
            7 => TimeEffectNodeType::WithGroup,
            8 => TimeEffectNodeType::AfterGroup,
            9 => TimeEffectNodeType::TimingRoot,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid effect node type {value}"
                )));
            },
        }),
        0x15 => TimeNodeProperty::PlaceholderNode(boolean()?),
        0x16 => {
            if data.len() != 5 || data[0] != 2 {
                return Err(Error::InvalidFormat("invalid media volume".to_string()));
            }
            let value = read_f32(data, 1);
            if !value.is_finite() || !(0.0..=100_000.0).contains(&value) {
                return Err(Error::InvalidFormat(
                    "media volume out of range".to_string(),
                ));
            }
            TimeNodeProperty::MediaVolume(value)
        },
        0x17 => TimeNodeProperty::MediaMute(boolean()?),
        0x1A => TimeNodeProperty::ZoomToFullScreen(boolean()?),
        id => {
            return Err(Error::InvalidFormat(format!(
                "unknown time property {id:#X}"
            )));
        },
    })
}
