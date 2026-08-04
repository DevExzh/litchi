//! Shared lossless access to a native chart-legend style payload.
//!
//! Legend presentation controls are stored in the generated extension of a
//! chart's `TSCH.LegendStyleArchive`. This module resolves the one style object
//! referenced by a chart and provides guarded wire-level access for focused
//! legend feature modules.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, LEGEND_STYLE_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::protobuf::{tsch, tsp, tss};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::{
    register_private_style, register_style_reference, unregister_private_style,
};
use crate::wire::{
    parse_wire_fields, patch_length_delimited_field, patch_varint_field,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated legend-style properties.
pub(crate) const GENERATED_LEGEND_STYLE_EXTENSION_FIELD: u32 = 10_000;
const CHART_ARCHIVE_EXTENSION_FIELD: u32 = 10_000;
const CHART_LEGEND_STYLE_FIELD: u32 = 11;
const LEGEND_STYLE_SUPER_FIELD: u32 = 1;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_IS_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;

/// The single mutable native legend-style payload for one chart.
#[derive(Debug)]
pub(crate) struct LegendStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Resolve the native legend-style payload referenced by one chart.
pub(crate) fn legend_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<LegendStyleSlot> {
    let style_id = package.with_parsed_archive(chart_archive_name, |chart_archive| {
        let chart_object = chart_archive.object(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        let mut chart_messages = chart_object
            .messages
            .iter()
            .filter(|message| message.type_ == CHART_MESSAGE_TYPE);
        let Some(chart_message) = chart_messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
            )));
        };
        if chart_messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
            )));
        }
        let chart = IWorkChartArchive::decode(chart_message.data.as_slice())?;
        chart
            .chart
            .as_ref()
            .and_then(|payload| payload.legend_style.as_ref())
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{drawable_label} chart {drawable_object_id} has no legend style"
                ))
            })
    })?;
    let archive_name = unique_chart_object_archive_name(package, style_id, "legend style object")?;
    let message_index = package.with_parsed_archive(&archive_name, |archive| {
        let style_object = archive.object(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart legend style {style_id} is missing"
            ))
        })?;
        let mut messages = style_object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == LEGEND_STYLE_MESSAGE_TYPE);
        let Some((message_index, _)) = messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart legend style {style_id} must have exactly one legend-style payload"
            )));
        };
        if messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart legend style {style_id} must have exactly one legend-style payload"
            )));
        }
        Ok(message_index)
    })?;
    Ok(LegendStyleSlot {
        archive_name,
        object_id: style_id,
        message_index,
    })
}

impl LegendStyleSlot {
    pub(crate) fn archive_name(&self) -> &str {
        &self.archive_name
    }

    pub(crate) const fn object_id(&self) -> u64 {
        self.object_id
    }

    /// Read the resolved legend-style bytes without allocating a rewrite.
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
        package.with_parsed_archive(&self.archive_name, |archive| {
            let object = archive.object(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("legend style {} is missing", self.object_id))
            })?;
            let message = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "legend style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if message.type_ != LEGEND_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "legend style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            read(message.data.as_slice())
        })
    }

    /// Disconnect a shared or preset legend style before mutating it.
    pub(crate) fn ensure_exclusive(
        &mut self,
        package: &mut IWorkPackage,
        chart_archive_name: &str,
        drawable_object_id: u64,
        drawable_label: &str,
    ) -> Result<()> {
        let (style, stylesheet_id) = self.read(package, |data| {
            let style = tsch::LegendStyleArchive::decode(data)?;
            let stylesheet_id = style
                .super_
                .as_ref()
                .and_then(|style| style.stylesheet.as_ref())
                .map(|reference| reference.identifier)
                .filter(|identifier| *identifier != 0)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "{drawable_label} chart legend style {} has no stylesheet",
                        self.object_id
                    ))
                })?;
            Ok((style, stylesheet_id))
        })?;
        if style
            .super_
            .as_ref()
            .is_some_and(|style| style.is_variation == Some(true))
            && chart_owner_count(package, self.object_id)? == 1
        {
            return Ok(());
        }

        let parent_style_id = self.object_id;
        let style_id = next_object_identifier(package)?;
        let registry_archive_name =
            unique_chart_object_archive_name(package, stylesheet_id, "chart stylesheet")?;
        let archive = package.archive(&self.archive_name)?;
        let source = archive.object(parent_style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("legend style {parent_style_id} is missing"))
        })?;
        let mut archive_info = source.archive_info.clone();
        archive_info.identifier = Some(style_id);
        let mut variation = crate::archive::ArchiveObject::new(style_id, source.messages.clone())?;
        variation.archive_info = archive_info;
        let message = variation.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "legend style {parent_style_id} message index changed unexpectedly"
            ))
        })?;
        let data = make_style_variation(
            message.data.as_slice(),
            style.super_.as_ref().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{drawable_label} chart legend style {parent_style_id} has no style archive"
                ))
            })?,
            parent_style_id,
            stylesheet_id,
        )?;
        variation.replace_message(
            self.message_index,
            RawMessage {
                type_: LEGEND_STYLE_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut variation.archive_info.message_infos[self.message_index];
        for identifier in [parent_style_id, stylesheet_id] {
            if !info.object_references.contains(&identifier) {
                info.object_references.push(identifier);
            }
        }

        insert_style_variation(
            package,
            &registry_archive_name,
            stylesheet_id,
            parent_style_id,
            style_id,
            variation,
        )?;
        if registry_archive_name != self.archive_name {
            move_style_object(
                package,
                &registry_archive_name,
                &self.archive_name,
                style_id,
            )?;
        }
        register_private_style(
            package,
            &registry_archive_name,
            &self.archive_name,
            style_id,
        )?;
        register_style_reference(package, chart_archive_name, &self.archive_name, style_id)?;
        patch_chart_legend_style(
            package,
            chart_archive_name,
            drawable_object_id,
            parent_style_id,
            style_id,
        )?;
        set_package_last_object_identifier(package, style_id)?;
        self.object_id = style_id;
        Ok(())
    }

    /// Transactionally rewrite the resolved legend-style message.
    pub(crate) fn update(
        &self,
        package: &mut IWorkPackage,
        patch: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<()> {
        package.update_archive(&self.archive_name, |archive| {
            let object = archive.object_mut(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("legend style {} is missing", self.object_id))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "legend style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != LEGEND_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "legend style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: LEGEND_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }

    /// Reconnect an exact child to its parent and reclaim the disposable style.
    pub(crate) fn collapse_if_equivalent(
        &mut self,
        package: &mut IWorkPackage,
        chart_archive_name: &str,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let child_data = self.read(package, |data| Ok(data.to_vec()))?;
        let child = tsch::LegendStyleArchive::decode(child_data.as_slice())?;
        let Some(style) = child.super_.as_ref() else {
            return Ok(false);
        };
        if style.is_variation != Some(true) || chart_owner_count(package, self.object_id)? != 1 {
            return Ok(false);
        }
        let Some(parent_style_id) = style
            .parent
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
        else {
            return Ok(false);
        };
        let Some(stylesheet_id) = style
            .stylesheet
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
        else {
            return Ok(false);
        };
        let parent_archive_name =
            unique_chart_object_archive_name(package, parent_style_id, "legend style parent")?;
        let parent_archive = package.archive(&parent_archive_name)?;
        let parent = parent_archive.object(parent_style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("legend style parent {parent_style_id} is missing"))
        })?;
        let parent_data = parent
            .messages
            .iter()
            .filter(|message| message.type_ == LEGEND_STYLE_MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
            .collect::<Vec<_>>();
        let [parent_data] = parent_data.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "legend style parent {parent_style_id} must have exactly one payload"
            )));
        };
        if without_style_archive(&child_data)? != without_style_archive(parent_data)? {
            return Ok(false);
        }

        let style_id = self.object_id;
        let registry_archive_name =
            unique_chart_object_archive_name(package, stylesheet_id, "chart stylesheet")?;
        patch_chart_legend_style(
            package,
            chart_archive_name,
            drawable_object_id,
            style_id,
            parent_style_id,
        )?;
        unregister_private_style(
            package,
            &registry_archive_name,
            &self.archive_name,
            style_id,
            (parent_archive_name == self.archive_name).then_some(parent_style_id),
        )?;
        if registry_archive_name != self.archive_name {
            move_style_object(
                package,
                &self.archive_name,
                &registry_archive_name,
                style_id,
            )?;
        }
        remove_style_variation(
            package,
            &registry_archive_name,
            stylesheet_id,
            parent_style_id,
            style_id,
        )?;
        register_style_reference(
            package,
            chart_archive_name,
            &parent_archive_name,
            parent_style_id,
        )?;
        release_package_identifier_suffix(package, &[style_id])?;
        self.archive_name = parent_archive_name;
        self.object_id = parent_style_id;
        Ok(true)
    }
}

fn move_style_object(
    package: &mut IWorkPackage,
    source_archive_name: &str,
    destination_archive_name: &str,
    style_id: u64,
) -> Result<()> {
    let mut moved = None;
    package.update_archive(source_archive_name, |archive| {
        moved = archive.remove_object(style_id);
        if moved.is_none() {
            return Err(Error::InvalidFormat(format!(
                "disposable legend style {style_id} is missing from {source_archive_name}"
            )));
        }
        Ok(())
    })?;
    let style = moved.ok_or_else(|| {
        Error::InvalidFormat(format!("disposable legend style {style_id} disappeared"))
    })?;
    package.update_archive(destination_archive_name, |archive| {
        archive.insert_object(style)
    })
}

fn without_style_archive(data: &[u8]) -> Result<Vec<u8>> {
    patch_length_delimited_field(data, LEGEND_STYLE_SUPER_FIELD, true, None)
}

fn chart_owner_count(package: &IWorkPackage, style_id: u64) -> Result<usize> {
    let mut count = 0usize;
    for archive_name in package.iwa_entry_names() {
        for object in &package.archive(archive_name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
            {
                let chart = IWorkChartArchive::decode(message.data.as_slice())?;
                if chart
                    .chart
                    .as_ref()
                    .and_then(|payload| payload.legend_style.as_ref())
                    .is_some_and(|reference| reference.identifier == style_id)
                {
                    count = count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("legend style owner count overflow".to_owned())
                    })?;
                }
            }
        }
    }
    Ok(count)
}

fn make_style_variation(
    data: &[u8],
    style: &tss::StyleArchive,
    parent_style_id: u64,
    stylesheet_id: u64,
) -> Result<Vec<u8>> {
    let parent = tsp::Reference {
        identifier: parent_style_id,
        ..Default::default()
    }
    .encode_to_vec();
    let stylesheet = tsp::Reference {
        identifier: stylesheet_id,
        ..Default::default()
    }
    .encode_to_vec();
    transform_length_delimited_field(data, LEGEND_STYLE_SUPER_FIELD, |super_data| {
        let super_data = patch_length_delimited_field(
            super_data,
            STYLE_PARENT_FIELD,
            style.parent.is_some(),
            Some(&parent),
        )?;
        let super_data = patch_varint_field(
            &super_data,
            STYLE_IS_VARIATION_FIELD,
            style.is_variation.is_some(),
            Some(1),
        )?;
        patch_length_delimited_field(
            &super_data,
            STYLE_STYLESHEET_FIELD,
            style.stylesheet.is_some(),
            Some(&stylesheet),
        )
    })
}

fn patch_chart_legend_style(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(chart_archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart {drawable_object_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "chart {drawable_object_id} must have exactly one chart payload"
            )));
        };
        let original = &object.messages[*index];
        let data = transform_length_delimited_field(
            original.data.as_slice(),
            CHART_ARCHIVE_EXTENSION_FIELD,
            |chart| {
                transform_length_delimited_field(chart, CHART_LEGEND_STYLE_FIELD, |reference| {
                    let decoded = tsp::Reference::decode(reference)?;
                    if decoded.identifier != old_style_id {
                        return Err(Error::InvalidFormat(format!(
                            "chart {drawable_object_id} legend style changed unexpectedly"
                        )));
                    }
                    patch_varint_field(reference, 1, true, Some(new_style_id))
                })
            },
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        let mut replaced = 0usize;
        for reference in &mut info.object_references {
            if *reference == old_style_id {
                *reference = new_style_id;
                replaced += 1;
            }
        }
        for field in &mut info.field_infos {
            for reference in &mut field.object_references {
                if *reference == old_style_id {
                    *reference = new_style_id;
                }
            }
        }
        if replaced != 1 {
            return Err(Error::InvalidFormat(format!(
                "chart {drawable_object_id} metadata contains {replaced} legend-style references"
            )));
        }
        Ok(())
    })
}

/// Decode the outer legend-style payload and locate its generated extension.
pub(crate) fn generated_legend_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::LegendStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number == GENERATED_LEGEND_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "legend style extension {GENERATED_LEGEND_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "legend style extension {GENERATED_LEGEND_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    Ok(Some(&data[extension.payload_start..extension.end]))
}
