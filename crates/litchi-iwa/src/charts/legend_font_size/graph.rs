//! Paragraph-style graph lifecycle for native chart-legend font sizes.
//!
//! iWork stores legend typography indirectly: field 2 of the generated legend
//! style selects one entry in the chart's paragraph-style table. A direct font
//! size is a private paragraph-style variation appended to that table.

use std::collections::HashSet;

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::font::ChartFontSize;
use crate::charts::legend_style::{LegendStyleSlot, legend_style_slot};
use crate::charts::source::CHART_MESSAGE_TYPE;
use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::protobuf::{tsp, tswp};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, direct_overrides, locate_style, parent_style_id, stylesheet_id,
    variation_object,
};
use crate::text::style_registry::{
    register_private_style, register_style_reference, unregister_owner_reference_if_unused,
    unregister_private_style,
};
use crate::wire::{rewrite_repeated_length_delimited_fields, transform_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

use super::ChartLegendFontSize;
use super::wire::{
    direct_paragraph_style_index, patch_direct_paragraph_style_index, patch_existing_size,
};

/// `TSCH.ChartArchive.paragraph_styles`.
const CHART_PARAGRAPH_STYLES_FIELD: u32 = 20;
/// `TSCH.ChartArchive` extension in `TSCH.ChartDrawableArchive`.
const CHART_ARCHIVE_EXTENSION_FIELD: u32 = 10_000;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;

/// Read the exact direct legend font-size state of one native chart.
pub(crate) fn chart_legend_font_size(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartLegendFontSize> {
    let graph = LegendFontGraph::locate(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    graph.read(package)
}

/// Set or remove the direct legend font-size override of one native chart.
pub(crate) fn set_chart_legend_font_size(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    target: ChartLegendFontSize,
) -> Result<()> {
    let graph = LegendFontGraph::locate(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if graph.read(package)? == target {
        return Ok(());
    }
    graph
        .legend_style
        .ensure_exclusive(package, drawable_object_id, drawable_label)?;
    let mut staged = package.clone();
    match target {
        ChartLegendFontSize::Inherited => graph.reset(&mut staged)?,
        ChartLegendFontSize::Size(size) => graph.set(&mut staged, size)?,
    }
    let verified = LegendFontGraph::locate(
        &staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if verified.read(&staged)? != target {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend font-size update failed validation"
        )));
    }
    *package = staged;
    Ok(())
}

struct LegendFontGraph {
    chart_archive_name: String,
    drawable_object_id: u64,
    drawable_label: String,
    chart_message_index: usize,
    paragraph_style_ids: Vec<u64>,
    direct_index: Option<usize>,
    legend_style: LegendStyleSlot,
}

impl LegendFontGraph {
    fn locate(
        package: &IWorkPackage,
        chart_archive_name: &str,
        drawable_object_id: u64,
        drawable_label: &str,
    ) -> Result<Self> {
        let archive = package.archive(chart_archive_name)?;
        let object = archive.object(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        let chart_messages = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        let [(chart_message_index, chart_message)] = chart_messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} drawable {drawable_object_id} must contain exactly one chart payload"
            )));
        };
        let chart = crate::charts::IWorkChartArchive::decode(chart_message.data.as_slice())?;
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
        let paragraph_style_ids = payload
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
            .collect::<Result<Vec<_>>>()?;
        let legend_style = legend_style_slot(
            package,
            chart_archive_name,
            drawable_object_id,
            drawable_label,
        )?;
        let direct_index = legend_style.read(package, direct_paragraph_style_index)?;
        if let Some(index) = direct_index
            && index >= paragraph_style_ids.len()
        {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} legend paragraph-style index {index} is out of bounds for {} styles",
                paragraph_style_ids.len()
            )));
        }
        Ok(Self {
            chart_archive_name: chart_archive_name.to_owned(),
            drawable_object_id,
            drawable_label: drawable_label.to_owned(),
            chart_message_index: *chart_message_index,
            paragraph_style_ids,
            direct_index,
            legend_style,
        })
    }

    fn read(&self, package: &IWorkPackage) -> Result<ChartLegendFontSize> {
        let Some(index) = self.direct_index else {
            return Ok(ChartLegendFontSize::Inherited);
        };
        let style_id = self.paragraph_style_ids[index];
        let location = locate_style(package, style_id)?;
        let overrides =
            direct_overrides(&location.style, &location.message.data)?.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "legend paragraph style {style_id} is not an exact native variation"
                ))
            })?;
        Ok(overrides
            .point_size
            .map(ChartFontSize::from)
            .map_or(ChartLegendFontSize::Inherited, ChartLegendFontSize::Size))
    }

    fn set(&self, package: &mut IWorkPackage, size: ChartFontSize) -> Result<()> {
        if let Some(index) = self.direct_index {
            let style_id = self.paragraph_style_ids[index];
            let exclusive = style_is_exclusive_to_chart(
                package,
                style_id,
                &self.chart_archive_name,
                self.drawable_object_id,
            )?;
            if exclusive {
                return patch_existing_size(package, style_id, Some(size));
            }
            let location = locate_style(package, style_id)?;
            let mut overrides = direct_overrides(&location.style, &location.message.data)?
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "legend paragraph style {style_id} is not an exact native variation"
                    ))
                })?;
            overrides.point_size = Some(size.text_point_size());
            return self.replace_shared_style(package, index, style_id, overrides);
        }

        let parent_id = self.paragraph_style_ids[0];
        let parent = locate_style(package, parent_id)?;
        let stylesheet = stylesheet_id(&parent.style, parent_id)?;
        let style_id = next_object_identifier(package)?;
        let overrides = ParagraphStyleOverrides {
            point_size: Some(size.text_point_size()),
            ..Default::default()
        };
        let variation = variation_object(style_id, parent_id, stylesheet, overrides)?;
        let new_index = self.paragraph_style_ids.len();
        let native_index = u64::try_from(new_index).map_err(|_| {
            Error::InvalidFormat("legend paragraph-style index exceeds u64".to_owned())
        })?;

        insert_style_variation(
            package,
            &parent.archive_name,
            stylesheet,
            parent_id,
            style_id,
            variation,
        )?;
        register_private_style(
            package,
            &self.chart_archive_name,
            &parent.archive_name,
            style_id,
        )?;
        let mut next_ids = self.paragraph_style_ids.clone();
        next_ids.push(style_id);
        self.patch_chart_paragraph_styles(package, &next_ids)?;
        self.patch_legend_index(package, Some(native_index))?;
        set_package_last_object_identifier(package, style_id)
    }

    fn reset(&self, package: &mut IWorkPackage) -> Result<()> {
        let Some(index) = self.direct_index else {
            return Ok(());
        };
        let style_id = self.paragraph_style_ids[index];
        let location = locate_style(package, style_id)?;
        let Some(mut overrides) = direct_overrides(&location.style, &location.message.data)? else {
            return Err(Error::InvalidFormat(format!(
                "legend paragraph style {style_id} is not an exact native variation"
            )));
        };
        if overrides.point_size.take().is_none() {
            return Ok(());
        }
        let exclusive = style_is_exclusive_to_chart(
            package,
            style_id,
            &self.chart_archive_name,
            self.drawable_object_id,
        )?;
        if !overrides.is_empty() {
            if exclusive {
                return patch_existing_size(package, style_id, None);
            }
            return self.replace_shared_style(package, index, style_id, overrides);
        }

        let parent_id = parent_style_id(&location.style, style_id)?;
        let stylesheet = stylesheet_id(&location.style, style_id)?;
        self.patch_legend_index(package, None)?;
        let mut next_ids = self.paragraph_style_ids.clone();
        if index + 1 == next_ids.len() {
            next_ids.pop();
        } else {
            next_ids[index] = parent_id;
        }
        self.patch_chart_paragraph_styles(package, &next_ids)?;
        if !exclusive {
            unregister_owner_reference_if_unused(
                package,
                &self.chart_archive_name,
                &location.archive_name,
                style_id,
            )?;
            return Ok(());
        }
        remove_style_variation(
            package,
            &location.archive_name,
            stylesheet,
            parent_id,
            style_id,
        )?;
        unregister_private_style(
            package,
            &self.chart_archive_name,
            &location.archive_name,
            style_id,
            Some(parent_id),
        )?;
        register_style_reference(
            package,
            &self.chart_archive_name,
            &location.archive_name,
            parent_id,
        )?;
        release_package_identifier_suffix(package, &[style_id])
    }

    fn replace_shared_style(
        &self,
        package: &mut IWorkPackage,
        index: usize,
        previous_style_id: u64,
        overrides: ParagraphStyleOverrides,
    ) -> Result<()> {
        if overrides.is_empty() {
            return Err(Error::InvalidFormat(
                "shared legend paragraph-style replacement has no overrides".to_owned(),
            ));
        }
        let previous = locate_style(package, previous_style_id)?;
        let parent_id = parent_style_id(&previous.style, previous_style_id)?;
        let stylesheet = stylesheet_id(&previous.style, previous_style_id)?;
        let style_id = next_object_identifier(package)?;
        let variation = variation_object(style_id, parent_id, stylesheet, overrides)?;
        insert_style_variation(
            package,
            &previous.archive_name,
            stylesheet,
            parent_id,
            style_id,
            variation,
        )?;
        register_private_style(
            package,
            &self.chart_archive_name,
            &previous.archive_name,
            style_id,
        )?;
        let mut next_ids = self.paragraph_style_ids.clone();
        next_ids[index] = style_id;
        self.patch_chart_paragraph_styles(package, &next_ids)?;
        unregister_owner_reference_if_unused(
            package,
            &self.chart_archive_name,
            &previous.archive_name,
            previous_style_id,
        )?;
        set_package_last_object_identifier(package, style_id)
    }

    fn patch_legend_index(&self, package: &mut IWorkPackage, index: Option<u64>) -> Result<()> {
        self.legend_style.update(package, |data| {
            patch_direct_paragraph_style_index(data, index)
        })
    }

    fn patch_chart_paragraph_styles(
        &self,
        package: &mut IWorkPackage,
        next_ids: &[u64],
    ) -> Result<()> {
        let encoded = next_ids
            .iter()
            .map(|identifier| {
                tsp::Reference {
                    identifier: *identifier,
                    ..Default::default()
                }
                .encode_to_vec()
            })
            .collect::<Vec<_>>();
        let previous = self
            .paragraph_style_ids
            .iter()
            .copied()
            .collect::<HashSet<_>>();
        let next = next_ids.iter().copied().collect::<HashSet<_>>();
        package.update_archive(&self.chart_archive_name, |archive| {
            let object = archive.object_mut(self.drawable_object_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{} chart {} is missing",
                    self.drawable_label, self.drawable_object_id
                ))
            })?;
            let message = object
                .messages
                .get(self.chart_message_index)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "{} chart {} message index changed unexpectedly",
                        self.drawable_label, self.drawable_object_id
                    ))
                })?;
            if message.type_ != CHART_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "{} chart {} message type changed unexpectedly",
                    self.drawable_label, self.drawable_object_id
                )));
            }
            let data = transform_length_delimited_field(
                &message.data,
                CHART_ARCHIVE_EXTENSION_FIELD,
                |chart| {
                    rewrite_repeated_length_delimited_fields(
                        chart,
                        CHART_PARAGRAPH_STYLES_FIELD,
                        &encoded,
                    )
                },
            )?;
            object.replace_message(
                self.chart_message_index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data,
                },
            )?;
            let metadata =
                &mut object.archive_info.message_infos[self.chart_message_index].object_references;
            metadata
                .retain(|identifier| !previous.contains(identifier) || next.contains(identifier));
            for identifier in next {
                if !metadata.contains(&identifier) {
                    metadata.push(identifier);
                }
            }
            Ok(())
        })
    }
}

fn style_is_exclusive_to_chart(
    package: &IWorkPackage,
    style_id: u64,
    chart_archive_name: &str,
    drawable_object_id: u64,
) -> Result<bool> {
    let mut matching_chart_consumers = 0usize;
    let mut other_consumers = 0usize;
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
                    let child = tswp::ParagraphStyleArchive::decode(message.data.as_slice())?;
                    if child
                        .super_
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.identifier == style_id)
                    {
                        continue;
                    }
                }
                if archive_name == chart_archive_name
                    && object.archive_info.identifier == Some(drawable_object_id)
                    && message.type_ == CHART_MESSAGE_TYPE
                {
                    matching_chart_consumers =
                        matching_chart_consumers.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "legend paragraph-style consumer count overflow".to_owned(),
                            )
                        })?;
                } else {
                    other_consumers = other_consumers.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat(
                            "legend paragraph-style consumer count overflow".to_owned(),
                        )
                    })?;
                }
            }
        }
    }
    if matching_chart_consumers != 1 {
        return Err(Error::InvalidFormat(format!(
            "legend paragraph style {style_id} has {matching_chart_consumers} references from the target chart"
        )));
    }
    Ok(other_consumers == 0)
}
