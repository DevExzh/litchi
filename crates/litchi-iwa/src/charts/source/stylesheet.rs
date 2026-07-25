//! Registration lifecycle for chart-private styles in the document stylesheet.

use std::collections::HashSet;

use prost::Message;

use super::{
    AXIS_NON_STYLE_MESSAGE_TYPE, AXIS_STYLE_MESSAGE_TYPE, CHART_NON_STYLE_MESSAGE_TYPE,
    CHART_STYLE_MESSAGE_TYPE, LEGEND_NON_STYLE_MESSAGE_TYPE, LEGEND_STYLE_MESSAGE_TYPE,
    SERIES_NON_STYLE_MESSAGE_TYPE, SERIES_STYLE_MESSAGE_TYPE,
};
use crate::archive::RawMessage;
use crate::charts::unique_chart_object_archive_name;
use crate::package_metadata::{
    add_component_external_reference, component_identifier_for_entry,
    remove_component_external_reference,
};
use crate::protobuf::{tsp, tss};
use crate::wire::{
    append_repeated_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const CHART_STYLE_MESSAGE_TYPES: [u32; 8] = [
    CHART_STYLE_MESSAGE_TYPE,
    CHART_NON_STYLE_MESSAGE_TYPE,
    LEGEND_STYLE_MESSAGE_TYPE,
    LEGEND_NON_STYLE_MESSAGE_TYPE,
    AXIS_STYLE_MESSAGE_TYPE,
    AXIS_NON_STYLE_MESSAGE_TYPE,
    SERIES_STYLE_MESSAGE_TYPE,
    SERIES_NON_STYLE_MESSAGE_TYPE,
];

/// Return chart-style objects that are local to one private chart graph.
pub(crate) fn local_chart_style_ids(
    package: &IWorkPackage,
    archive_name: &str,
    object_ids: &[u64],
) -> Result<Vec<u64>> {
    let archive = package.archive(archive_name)?;
    let mut style_ids = Vec::new();
    for &identifier in object_ids {
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart graph object {identifier} is missing from {archive_name}"
            ))
        })?;
        let style_payload_count = object
            .messages
            .iter()
            .filter(|message| CHART_STYLE_MESSAGE_TYPES.contains(&message.type_))
            .count();
        match style_payload_count {
            0 => {},
            1 => style_ids.push(identifier),
            count => {
                return Err(Error::InvalidFormat(format!(
                    "chart style object {identifier} has {count} style payloads"
                )));
            },
        }
    }
    Ok(style_ids)
}

/// Register local chart styles in the document stylesheet and component graph.
pub(crate) fn register_chart_styles(
    package: &mut IWorkPackage,
    stylesheet_id: u64,
    owner_archive_name: &str,
    style_ids: &[u64],
) -> Result<()> {
    validate_style_ids(package, owner_archive_name, style_ids)?;
    let stylesheet_archive_name =
        unique_chart_object_archive_name(package, stylesheet_id, "chart stylesheet")?;
    let owner_component = component_identifier_for_entry(package, owner_archive_name)?;
    let stylesheet_component = component_identifier_for_entry(package, &stylesheet_archive_name)?;

    package.update_archive(&stylesheet_archive_name, |archive| {
        let stylesheet = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart stylesheet {stylesheet_id} is missing"))
        })?;
        let message_index = unique_stylesheet_message_index(stylesheet_id, stylesheet)?;
        let original = &stylesheet.messages[message_index];
        let decoded = tss::StylesheetArchive::decode(original.data.as_slice())?;
        for &style_id in style_ids {
            if decoded
                .styles
                .iter()
                .any(|reference| reference.identifier == style_id)
            {
                return Err(Error::InvalidFormat(format!(
                    "chart stylesheet {stylesheet_id} already contains style {style_id}"
                )));
            }
        }

        let mut data = original.data.clone();
        for &style_id in style_ids {
            data = append_repeated_length_delimited_field(
                &data,
                STYLESHEET_STYLES_FIELD,
                &reference(style_id).encode_to_vec(),
            )?;
        }
        let verified = tss::StylesheetArchive::decode(data.as_slice())?;
        for &style_id in style_ids {
            if verified
                .styles
                .iter()
                .filter(|reference| reference.identifier == style_id)
                .count()
                != 1
            {
                return Err(Error::InvalidFormat(format!(
                    "chart stylesheet registration failed for style {style_id}"
                )));
            }
        }
        stylesheet.replace_message(
            message_index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut stylesheet.archive_info.message_infos[message_index];
        for &style_id in style_ids {
            if info.object_references.contains(&style_id) {
                return Err(Error::InvalidFormat(format!(
                    "chart stylesheet metadata already references style {style_id}"
                )));
            }
            info.object_references.push(style_id);
        }
        Ok(())
    })?;

    if let (Some(owner_component), Some(stylesheet_component)) =
        (owner_component, stylesheet_component)
        && owner_component != stylesheet_component
    {
        add_component_external_reference(
            package,
            owner_component,
            stylesheet_component,
            stylesheet_id,
        )?;
        for &style_id in style_ids {
            add_component_external_reference(
                package,
                stylesheet_component,
                owner_component,
                style_id,
            )?;
        }
    }
    Ok(())
}

/// Validate persisted stylesheet membership and reference metadata.
pub(crate) fn validate_chart_styles_registered(
    package: &IWorkPackage,
    stylesheet_id: u64,
    owner_archive_name: &str,
    style_ids: &[u64],
) -> Result<()> {
    validate_style_ids(package, owner_archive_name, style_ids)?;
    let stylesheet_archive_name =
        unique_chart_object_archive_name(package, stylesheet_id, "chart stylesheet")?;
    let archive = package.archive(&stylesheet_archive_name)?;
    let stylesheet = archive.object(stylesheet_id).ok_or_else(|| {
        Error::InvalidFormat(format!("chart stylesheet {stylesheet_id} is missing"))
    })?;
    let message_index = unique_stylesheet_message_index(stylesheet_id, stylesheet)?;
    let registered =
        tss::StylesheetArchive::decode(stylesheet.messages[message_index].data.as_slice())?;
    let info = &stylesheet.archive_info.message_infos[message_index];
    for &style_id in style_ids {
        if registered
            .styles
            .iter()
            .filter(|reference| reference.identifier == style_id)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "chart stylesheet {stylesheet_id} must contain style {style_id} exactly once"
            )));
        }
        if info
            .object_references
            .iter()
            .filter(|&&identifier| identifier == style_id)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "chart stylesheet metadata must reference style {style_id} exactly once"
            )));
        }
    }
    Ok(())
}

/// Remove local chart styles from the document stylesheet and component graph.
pub(crate) fn unregister_chart_styles(
    package: &mut IWorkPackage,
    stylesheet_id: u64,
    owner_archive_name: &str,
    style_ids: &[u64],
) -> Result<()> {
    validate_unique_nonzero_ids(style_ids)?;
    let stylesheet_archive_name =
        unique_chart_object_archive_name(package, stylesheet_id, "chart stylesheet")?;
    let owner_component = component_identifier_for_entry(package, owner_archive_name)?;
    let stylesheet_component = component_identifier_for_entry(package, &stylesheet_archive_name)?;

    package.update_archive(&stylesheet_archive_name, |archive| {
        let stylesheet = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart stylesheet {stylesheet_id} is missing"))
        })?;
        let message_index = unique_stylesheet_message_index(stylesheet_id, stylesheet)?;
        let original = &stylesheet.messages[message_index];
        let mut registered =
            repeated_length_delimited_payloads(original.data.as_slice(), STYLESHEET_STYLES_FIELD)?
                .into_iter()
                .map(|payload| {
                    Ok((
                        tsp::Reference::decode(payload)?.identifier,
                        payload.to_vec(),
                    ))
                })
                .collect::<Result<Vec<_>>>()?;
        for &style_id in style_ids {
            if registered
                .iter()
                .filter(|(identifier, _)| *identifier == style_id)
                .count()
                != 1
            {
                return Err(Error::InvalidFormat(format!(
                    "chart stylesheet {stylesheet_id} must contain style {style_id} exactly once"
                )));
            }
        }
        let removed = style_ids.iter().copied().collect::<HashSet<_>>();
        registered.retain(|(identifier, _)| !removed.contains(identifier));
        let payloads = registered
            .into_iter()
            .map(|(_, payload)| payload)
            .collect::<Vec<_>>();
        let data = rewrite_repeated_length_delimited_fields(
            original.data.as_slice(),
            STYLESHEET_STYLES_FIELD,
            &payloads,
        )?;
        stylesheet.replace_message(
            message_index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut stylesheet.archive_info.message_infos[message_index];
        for &style_id in style_ids {
            if info
                .object_references
                .iter()
                .filter(|&&identifier| identifier == style_id)
                .count()
                != 1
            {
                return Err(Error::InvalidFormat(format!(
                    "chart stylesheet metadata must reference style {style_id} exactly once"
                )));
            }
            info.object_references
                .retain(|&identifier| identifier != style_id);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|&identifier| identifier != style_id);
            }
        }
        Ok(())
    })?;
    if let (Some(owner_component), Some(stylesheet_component)) =
        (owner_component, stylesheet_component)
        && owner_component != stylesheet_component
    {
        for &style_id in style_ids {
            remove_component_external_reference(
                package,
                stylesheet_component,
                owner_component,
                style_id,
            )?;
        }
    }
    Ok(())
}

fn validate_style_ids(
    package: &IWorkPackage,
    owner_archive_name: &str,
    style_ids: &[u64],
) -> Result<()> {
    validate_unique_nonzero_ids(style_ids)?;
    let archive = package.archive(owner_archive_name)?;
    for &style_id in style_ids {
        let object = archive.object(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart style {style_id} is missing from {owner_archive_name}"
            ))
        })?;
        let count = object
            .messages
            .iter()
            .filter(|message| CHART_STYLE_MESSAGE_TYPES.contains(&message.type_))
            .count();
        if count != 1 {
            return Err(Error::InvalidFormat(format!(
                "chart style {style_id} must contain exactly one chart-style payload"
            )));
        }
    }
    Ok(())
}

fn validate_unique_nonzero_ids(style_ids: &[u64]) -> Result<()> {
    let unique = style_ids.iter().copied().collect::<HashSet<_>>();
    if unique.len() != style_ids.len() || unique.contains(&0) {
        return Err(Error::InvalidFormat(
            "chart style identifiers must be unique and nonzero".to_owned(),
        ));
    }
    Ok(())
}

fn unique_stylesheet_message_index(
    stylesheet_id: u64,
    stylesheet: &crate::archive::ArchiveObject,
) -> Result<usize> {
    let indexes = stylesheet
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "chart stylesheet {stylesheet_id} must contain exactly one stylesheet payload"
        )));
    };
    Ok(*message_index)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
