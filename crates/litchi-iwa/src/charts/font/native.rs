//! Paragraph-style graph operations backing chart font CRUD.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::charts::IWorkChartArchive;
use crate::charts::source::CHART_MESSAGE_TYPE;
use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, direct_overrides, inherited_text_font, inherited_text_style,
    locate_style, parent_style_id, stylesheet_id, variation_object,
};
use crate::text::style_registry::{
    register_private_style, register_style_reference, unregister_owner_reference_if_unused,
    unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

use super::{ChartFont, ChartFontSize};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;

#[derive(Debug, Clone, Copy)]
enum FontProperty<'a> {
    Identity(&'a ChartFont),
    Size(ChartFontSize),
}

#[derive(Debug, Clone, Copy)]
enum FontPropertyKind {
    Identity,
    Size,
}

/// Read the uniform effective font used across one chart's semantic text slots.
pub(crate) fn chart_font(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartFont> {
    let style_ids = chart_paragraph_style_ids(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    uniform_property(
        package,
        &style_ids,
        drawable_object_id,
        drawable_label,
        effective_font,
        "fonts",
    )
}

/// Read the uniform effective point size used across a chart's text slots.
pub(crate) fn chart_font_size(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartFontSize> {
    let style_ids = chart_paragraph_style_ids(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    uniform_property(
        package,
        &style_ids,
        drawable_object_id,
        drawable_label,
        effective_font_size,
        "font sizes",
    )
}

/// Set the uniform chart font using copy-on-write style variations.
pub(crate) fn set_chart_font(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    target: &ChartFont,
) -> Result<()> {
    if chart_font(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? == *target
    {
        return Ok(());
    }
    set_property(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        FontProperty::Identity(target),
    )?;
    if chart_font(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != *target
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart font update failed validation"
        )));
    }
    Ok(())
}

/// Set the uniform chart point size using copy-on-write style variations.
pub(crate) fn set_chart_font_size(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    target: ChartFontSize,
) -> Result<()> {
    if chart_font_size(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? == target
    {
        return Ok(());
    }
    set_property(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        FontProperty::Size(target),
    )?;
    if chart_font_size(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != target
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart font-size update failed validation"
        )));
    }
    Ok(())
}

/// Reset chart font identity and face traits while preserving point size.
pub(crate) fn reset_chart_font(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    reset_property(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        FontPropertyKind::Identity,
    )
}

/// Reset chart point size while preserving font identity and face traits.
pub(crate) fn reset_chart_font_size(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    reset_property(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        FontPropertyKind::Size,
    )
}

fn set_property(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    property: FontProperty<'_>,
) -> Result<()> {
    let style_ids = chart_paragraph_style_ids(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let unique = unique_ids(&style_ids);
    let mut next_identifier = next_object_identifier(package)?;
    let mut replacements = HashMap::with_capacity(unique.len());
    let mut additions = Vec::with_capacity(unique.len());

    for style_id in &unique {
        let location = locate_style(package, *style_id)?;
        let direct = direct_overrides(&location.style, &location.message.data)?;
        let (parent_id, stylesheet, style_archive_name, mut overrides) = if let Some(overrides) =
            direct.filter(|overrides| overrides.is_chart_font_format_only())
        {
            (
                parent_style_id(&location.style, *style_id)?,
                stylesheet_id(&location.style, *style_id)?,
                location.archive_name,
                overrides,
            )
        } else {
            (
                *style_id,
                stylesheet_id(&location.style, *style_id)?,
                location.archive_name,
                ParagraphStyleOverrides::default(),
            )
        };
        apply_property(package, parent_id, &mut overrides, property)?;
        allocate_replacement(
            &mut next_identifier,
            *style_id,
            parent_id,
            stylesheet,
            style_archive_name,
            overrides,
            &mut replacements,
            &mut additions,
        )?;
    }

    apply_replacements(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        &unique,
        replacements,
        additions,
    )
}

fn reset_property(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: FontPropertyKind,
) -> Result<bool> {
    let style_ids = chart_paragraph_style_ids(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let unique = unique_ids(&style_ids);
    let mut candidates = Vec::with_capacity(unique.len());
    for style_id in &unique {
        let location = locate_style(package, *style_id)?;
        let Some(mut overrides) = direct_overrides(&location.style, &location.message.data)? else {
            continue;
        };
        if !overrides.is_chart_font_format_only() || !clear_property(&mut overrides, kind) {
            continue;
        }
        candidates.push((
            *style_id,
            parent_style_id(&location.style, *style_id)?,
            stylesheet_id(&location.style, *style_id)?,
            location.archive_name,
            overrides,
        ));
    }
    if candidates.is_empty() {
        return Ok(false);
    }

    let mut next_identifier = next_object_identifier(package)?;
    let mut replacements = HashMap::with_capacity(candidates.len());
    let mut additions = Vec::with_capacity(candidates.len());
    for (style_id, parent_id, stylesheet, archive_name, overrides) in candidates {
        allocate_replacement(
            &mut next_identifier,
            style_id,
            parent_id,
            stylesheet,
            archive_name,
            overrides,
            &mut replacements,
            &mut additions,
        )?;
    }
    apply_replacements(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        &unique,
        replacements,
        additions,
    )?;
    Ok(true)
}

fn apply_property(
    package: &IWorkPackage,
    parent_style_id: u64,
    overrides: &mut ParagraphStyleOverrides,
    property: FontProperty<'_>,
) -> Result<()> {
    match property {
        FontProperty::Identity(target) => {
            let inherited = effective_font(package, parent_style_id)?;
            overrides.font = (inherited.font != target.font).then(|| target.font.clone());
            overrides.bold = (inherited.bold != target.bold).then_some(target.bold);
            overrides.italic = (inherited.italic != target.italic).then_some(target.italic);
        },
        FontProperty::Size(target) => {
            let inherited = effective_font_size(package, parent_style_id)?;
            overrides.point_size = (inherited != target).then_some(target.text_point_size());
        },
    }
    Ok(())
}

fn clear_property(overrides: &mut ParagraphStyleOverrides, kind: FontPropertyKind) -> bool {
    match kind {
        FontPropertyKind::Identity => {
            let present =
                overrides.font.is_some() || overrides.bold.is_some() || overrides.italic.is_some();
            overrides.font = None;
            overrides.bold = None;
            overrides.italic = None;
            present
        },
        FontPropertyKind::Size => overrides.point_size.take().is_some(),
    }
}

#[allow(clippy::too_many_arguments)]
fn allocate_replacement(
    next_identifier: &mut u64,
    current_style_id: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    archive_name: String,
    overrides: ParagraphStyleOverrides,
    replacements: &mut HashMap<u64, u64>,
    additions: &mut Vec<StyleAddition>,
) -> Result<()> {
    if overrides.is_empty() {
        replacements.insert(current_style_id, parent_style_id);
        return Ok(());
    }
    let identifier = *next_identifier;
    *next_identifier = next_identifier
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("chart font style identifier overflow".to_owned()))?;
    let object = variation_object(identifier, parent_style_id, stylesheet_id, overrides)?;
    replacements.insert(current_style_id, identifier);
    additions.push(StyleAddition {
        identifier,
        parent_id: parent_style_id,
        stylesheet_id,
        archive_name,
        object,
    });
    Ok(())
}

fn apply_replacements(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    previous_style_ids: &[u64],
    replacements: HashMap<u64, u64>,
    additions: Vec<StyleAddition>,
) -> Result<()> {
    let mut staged = package.clone();
    patch_chart_style_references(
        &mut staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        &replacements,
    )?;
    let last_added_identifier = additions.last().map(|addition| addition.identifier);
    for addition in additions {
        insert_style_variation(
            &mut staged,
            &addition.archive_name,
            addition.stylesheet_id,
            addition.parent_id,
            addition.identifier,
            addition.object,
        )?;
        register_private_style(
            &mut staged,
            chart_archive_name,
            &addition.archive_name,
            addition.identifier,
        )?;
    }
    if let Some(last) = last_added_identifier {
        set_package_last_object_identifier(&mut staged, last)?;
    }
    prune_replaced_variations(
        &mut staged,
        chart_archive_name,
        previous_style_ids.iter().copied(),
    )?;
    *package = staged;
    Ok(())
}

struct StyleAddition {
    identifier: u64,
    parent_id: u64,
    stylesheet_id: u64,
    archive_name: String,
    object: ArchiveObject,
}

fn effective_font(package: &IWorkPackage, style_id: u64) -> Result<ChartFont> {
    let font = inherited_text_font(package, style_id)?;
    let style = inherited_text_style(package, style_id)?;
    Ok(ChartFont::new(font)
        .with_bold(style.bold)
        .with_italic(style.italic))
}

fn effective_font_size(package: &IWorkPackage, style_id: u64) -> Result<ChartFontSize> {
    inherited_text_style(package, style_id).map(|style| style.point_size.into())
}

fn uniform_property<T>(
    package: &IWorkPackage,
    style_ids: &[u64],
    drawable_object_id: u64,
    drawable_label: &str,
    read: impl Fn(&IWorkPackage, u64) -> Result<T>,
    property_label: &str,
) -> Result<T>
where
    T: PartialEq,
{
    let first = read(package, style_ids[0])?;
    for style_id in &style_ids[1..] {
        if read(package, *style_id)? != first {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has mixed effective {property_label}"
            )));
        }
    }
    Ok(first)
}

fn chart_paragraph_style_ids(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Vec<u64>> {
    let (_, chart) = locate_chart(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    paragraph_style_ids(&chart, drawable_object_id, drawable_label)
}

fn paragraph_style_ids(
    chart: &IWorkChartArchive,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Vec<u64>> {
    let payload = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no chart payload"
        ))
    })?;
    if payload.paragraph_styles.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no paragraph styles"
        )));
    }
    payload
        .paragraph_styles
        .iter()
        .map(|reference| {
            (reference.identifier != 0)
                .then_some(reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "{drawable_label} chart {drawable_object_id} has a zero paragraph-style reference"
                    ))
                })
        })
        .collect()
}

fn unique_ids(ids: &[u64]) -> Vec<u64> {
    let mut seen = HashSet::with_capacity(ids.len());
    ids.iter()
        .copied()
        .filter(|identifier| seen.insert(*identifier))
        .collect()
}

fn locate_chart(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<(usize, IWorkChartArchive)> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let payloads = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE)
        .map(|(index, message)| {
            IWorkChartArchive::decode(&message.data).map(|chart| (index, chart))
        })
        .collect::<Result<Vec<_>>>()?;
    let [(index, chart)] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} drawable {drawable_object_id} must contain exactly one chart payload"
        )));
    };
    Ok((*index, chart.clone()))
}

fn patch_chart_style_references(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    replacements: &HashMap<u64, u64>,
) -> Result<()> {
    let (message_index, mut chart) =
        locate_chart(package, archive_name, drawable_object_id, drawable_label)?;
    let payload = chart.chart.as_mut().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has no chart payload"
        ))
    })?;
    for reference in &mut payload.paragraph_styles {
        if let Some(replacement) = replacements.get(&reference.identifier) {
            reference.identifier = *replacement;
        }
    }
    let previous = replacements.keys().copied().collect::<HashSet<_>>();
    let next = payload
        .paragraph_styles
        .iter()
        .map(|reference| reference.identifier)
        .collect::<HashSet<_>>();
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: chart.encode()?,
            },
        )?;
        let metadata = &mut object.archive_info.message_infos[message_index].object_references;
        metadata.retain(|identifier| !previous.contains(identifier));
        for identifier in next {
            if !metadata.contains(&identifier) {
                metadata.push(identifier);
            }
        }
        Ok(())
    })
}

fn prune_replaced_variations(
    package: &mut IWorkPackage,
    owner_archive_name: &str,
    identifiers: impl IntoIterator<Item = u64>,
) -> Result<()> {
    let mut removed = Vec::new();
    for identifier in identifiers {
        let location = locate_style(package, identifier)?;
        let Some(overrides) = direct_overrides(&location.style, &location.message.data)? else {
            continue;
        };
        if !overrides.is_chart_font_format_only() || style_has_consumers(package, identifier)? {
            unregister_owner_reference_if_unused(
                package,
                owner_archive_name,
                &location.archive_name,
                identifier,
            )?;
            continue;
        }
        let parent = parent_style_id(&location.style, identifier)?;
        let stylesheet = stylesheet_id(&location.style, identifier)?;
        remove_style_variation(
            package,
            &location.archive_name,
            stylesheet,
            parent,
            identifier,
        )?;
        unregister_private_style(
            package,
            owner_archive_name,
            &location.archive_name,
            identifier,
            Some(parent),
        )?;
        register_style_reference(package, owner_archive_name, &location.archive_name, parent)?;
        removed.push(identifier);
    }
    if !removed.is_empty() {
        release_package_identifier_suffix(package, &removed)?;
    }
    Ok(())
}

fn style_has_consumers(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    for archive_name in package.iwa_entry_names() {
        for object in &package.archive(archive_name)?.objects {
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if !info.object_references.contains(&style_id) {
                    continue;
                }
                if message.type_ == STYLESHEET_MESSAGE_TYPE {
                    continue;
                }
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE {
                    let style = crate::protobuf::tswp::ParagraphStyleArchive::decode(
                        message.data.as_slice(),
                    )?;
                    if style
                        .super_
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.identifier == style_id)
                    {
                        continue;
                    }
                }
                return Ok(true);
            }
        }
    }
    Ok(false)
}
