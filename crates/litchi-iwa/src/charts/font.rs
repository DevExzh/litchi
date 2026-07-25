//! Chart-wide font identity and face-trait CRUD.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::CHART_MESSAGE_TYPE;
use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::TextFont;
use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, direct_overrides, inherited_text_font, inherited_text_style,
    locate_style, parent_style_id, stylesheet_id, variation_object,
};
use crate::text::style_registry::{
    register_private_style, register_style_reference, unregister_owner_reference_if_unused,
    unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;

/// Effective chart-wide font identity and face traits.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartFont {
    font: TextFont,
    bold: bool,
    italic: bool,
}

impl ChartFont {
    /// Construct a chart font with regular face traits.
    pub const fn new(font: TextFont) -> Self {
        Self {
            font,
            bold: false,
            italic: false,
        }
    }

    /// Construct a named chart font from a validated PostScript identifier.
    pub fn named(name: impl Into<String>) -> Result<Self> {
        TextFont::named(name).map(Self::new)
    }

    /// Borrow the effective font identity.
    pub const fn font(&self) -> &TextFont {
        &self.font
    }

    /// Whether the chart uses bold face traits.
    pub const fn bold(&self) -> bool {
        self.bold
    }

    /// Whether the chart uses italic face traits.
    pub const fn italic(&self) -> bool {
        self.italic
    }

    /// Enable or disable bold face traits.
    pub const fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Enable or disable italic face traits.
    pub const fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }
}

/// Read the uniform effective font used across one chart's semantic text slots.
pub(crate) fn chart_font(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartFont> {
    let (_, chart) = locate_chart(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let style_ids = paragraph_style_ids(&chart, drawable_object_id, drawable_label)?;
    uniform_font(package, &style_ids, drawable_object_id, drawable_label)
}

/// Set the uniform chart font with copy-on-write paragraph-style variations.
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

    let (_, chart) = locate_chart(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let style_ids = paragraph_style_ids(&chart, drawable_object_id, drawable_label)?;
    let unique = unique_ids(&style_ids);
    let mut next_identifier = next_object_identifier(package)?;
    let mut replacements = HashMap::with_capacity(unique.len());
    let mut additions = Vec::with_capacity(unique.len());

    for style_id in &unique {
        let location = locate_style(package, *style_id)?;
        let owned = direct_overrides(&location.style, &location.message.data)?
            .is_some_and(|overrides| overrides.is_chart_font_only());
        let (parent_id, stylesheet, style_archive_name) = if owned {
            (
                parent_style_id(&location.style, *style_id)?,
                stylesheet_id(&location.style, *style_id)?,
                location.archive_name,
            )
        } else {
            (
                *style_id,
                stylesheet_id(&location.style, *style_id)?,
                location.archive_name,
            )
        };
        let inherited = effective_font(package, parent_id)?;
        let overrides = font_overrides(&inherited, target);
        if !overrides.is_chart_font_only() {
            replacements.insert(*style_id, parent_id);
            continue;
        }

        let identifier = next_identifier;
        next_identifier = next_identifier.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("chart font style identifier overflow".to_owned())
        })?;
        let object = variation_object(identifier, parent_id, stylesheet, overrides)?;
        replacements.insert(*style_id, identifier);
        additions.push(StyleAddition {
            identifier,
            parent_id,
            stylesheet_id: stylesheet,
            archive_name: style_archive_name,
            object,
        });
    }

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
    prune_replaced_variations(&mut staged, chart_archive_name, unique.iter().copied())?;
    let observed = chart_font(
        &staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if observed != *target {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart font update failed validation"
        )));
    }
    *package = staged;
    Ok(())
}

/// Collapse crate-owned font-only variations back to their parent styles.
pub(crate) fn reset_chart_font(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    let (_, chart) = locate_chart(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let style_ids = paragraph_style_ids(&chart, drawable_object_id, drawable_label)?;
    let unique = unique_ids(&style_ids);
    let mut replacements = HashMap::with_capacity(unique.len());
    for style_id in &unique {
        let location = locate_style(package, *style_id)?;
        let Some(overrides) = direct_overrides(&location.style, &location.message.data)? else {
            continue;
        };
        if overrides.is_chart_font_only() {
            replacements.insert(*style_id, parent_style_id(&location.style, *style_id)?);
        }
    }
    if replacements.is_empty() {
        return Ok(false);
    }

    let mut staged = package.clone();
    patch_chart_style_references(
        &mut staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        &replacements,
    )?;
    prune_replaced_variations(
        &mut staged,
        chart_archive_name,
        replacements.keys().copied(),
    )?;
    *package = staged;
    Ok(true)
}

struct StyleAddition {
    identifier: u64,
    parent_id: u64,
    stylesheet_id: u64,
    archive_name: String,
    object: crate::archive::ArchiveObject,
}

fn effective_font(package: &IWorkPackage, style_id: u64) -> Result<ChartFont> {
    let font = inherited_text_font(package, style_id)?;
    let style = inherited_text_style(package, style_id)?;
    Ok(ChartFont::new(font)
        .with_bold(style.bold)
        .with_italic(style.italic))
}

fn font_overrides(inherited: &ChartFont, target: &ChartFont) -> ParagraphStyleOverrides {
    let mut overrides = ParagraphStyleOverrides::default();
    overrides.font = (inherited.font != target.font).then(|| target.font.clone());
    overrides.bold = (inherited.bold != target.bold).then_some(target.bold);
    overrides.italic = (inherited.italic != target.italic).then_some(target.italic);
    overrides
}

fn uniform_font(
    package: &IWorkPackage,
    style_ids: &[u64],
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartFont> {
    let first = effective_font(package, style_ids[0])?;
    for style_id in &style_ids[1..] {
        if effective_font(package, *style_id)? != first {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has mixed effective fonts"
            )));
        }
    }
    Ok(first)
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
        if !overrides.is_chart_font_only() || style_has_consumers(package, identifier)? {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chart_font_builder_is_strict_and_typed() {
        let font = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        assert_eq!(font.font().name(), Some("AvenirNext-DemiBold"));
        assert!(font.bold());
        assert!(!font.italic());
        assert!(ChartFont::named(" AvenirNext-Regular").is_err());
    }
}
