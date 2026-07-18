//! Numbers theme discovery and chart-preset registration.

use super::*;

pub(super) struct ChartThemeContext {
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) theme_id: u64,
    pub(super) paragraph_style_id: u64,
}

pub(super) fn chart_theme_context(package: &IWorkPackage) -> Result<ChartThemeContext> {
    let theme_id = numbers_document(package)?.theme.identifier;
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers theme {theme_id} is missing")))?
        .to_owned();
    let archive = package.archive(&archive_name)?;
    let object = archive.object(theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBERS_THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme {theme_id} must contain exactly one theme payload"
        )));
    };
    let theme = IWorkThemeArchive::decode(&message.data)?;
    let paragraph_style_id = theme
        .extensions
        .text
        .as_ref()
        .and_then(|presets| {
            presets
                .paragraph_style_presets
                .iter()
                .find(|reference| reference.identifier != 0)
        })
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers theme has no paragraph style preset".to_owned())
        })?;
    let component_id =
        component_identifier_for_entry(package, &archive_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers theme component {archive_name} is not registered"
            ))
        })?;
    Ok(ChartThemeContext {
        archive_name,
        component_id,
        theme_id,
        paragraph_style_id,
    })
}

pub(super) fn patch_theme_chart_preset(
    package: &mut IWorkPackage,
    context: &ChartThemeContext,
    previous: Option<u64>,
    replacement: Option<u64>,
) -> Result<()> {
    package.update_archive(&context.archive_name, |archive| {
        let object = archive.object_mut(context.theme_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers theme object {} is missing",
                context.theme_id
            ))
        })?;
        let message_indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == NUMBERS_THEME_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = message_indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers theme {} must contain exactly one theme payload",
                context.theme_id
            )));
        };
        let message_index = *message_index;
        let message_type = object.messages[message_index].type_;
        let mut theme = IWorkThemeArchive::decode(&object.messages[message_index].data)?;
        let presets = theme
            .extensions
            .chart
            .get_or_insert_with(tsch::ChartPresetsArchive::default);
        if let Some(previous) = previous {
            let count = presets
                .chart_presets
                .iter()
                .filter(|reference| reference.identifier == previous)
                .count();
            if count != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Numbers theme references chart preset {previous} {count} times"
                )));
            }
            presets
                .chart_presets
                .retain(|reference| reference.identifier != previous);
        }
        if let Some(replacement) = replacement {
            if presets
                .chart_presets
                .iter()
                .any(|reference| reference.identifier == replacement)
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers theme already references chart preset {replacement}"
                )));
            }
            presets.chart_presets.push(reference(replacement));
        }
        if presets.chart_presets.is_empty() {
            theme.extensions.chart = None;
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data: theme.encode()?,
            },
        )?;
        let references = &mut object.archive_info.message_infos[message_index].object_references;
        if let Some(previous) = previous {
            references.retain(|identifier| *identifier != previous);
        }
        if let Some(replacement) = replacement
            && !references.contains(&replacement)
        {
            references.push(replacement);
        }
        Ok(())
    })
}
