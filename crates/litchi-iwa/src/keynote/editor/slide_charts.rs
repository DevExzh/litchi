//! Standalone, inline-data chart CRUD for Keynote slides.

mod axis;
mod axis_bounds;
mod axis_gridlines;
mod axis_line;
mod axis_minimum_label;
mod axis_steps;
mod caption;
mod graph;
mod legend;
mod theme;
mod title;

use std::collections::HashMap;

use graph::chart_graph;
use theme::{chart_theme_context, patch_theme_chart_preset};

use super::*;
use crate::IWorkThemeArchive;
use crate::charts::source::{
    AXIS_NON_STYLE_MESSAGE_TYPE, AXIS_STYLE_MESSAGE_TYPE, CHART_MESSAGE_TYPE,
    CHART_NON_STYLE_MESSAGE_TYPE, CHART_PRESET_MESSAGE_TYPE, CHART_STYLE_MESSAGE_TYPE,
    ChartApplicationProfile, LEGEND_NON_STYLE_MESSAGE_TYPE, LEGEND_STYLE_MESSAGE_TYPE,
    SERIES_STYLE_MESSAGE_TYPE, STANDIN_MESSAGE_TYPE, SourceChartObjectIds, chart_data,
    chart_geometry, chart_grid, drawable_geometry, geometry_archive, reference,
    require_creatable_kind, source_chart_objects,
};
use crate::charts::{ChartData, ChartKind, ChartSeriesDirection, IWorkChartArchive};
use crate::protobuf::tsch;
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize, offset_drawable_geometry};

const KEYNOTE_THEME_MESSAGE_TYPE: u32 = 10;

/// One native chart drawable owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideChartInfo {
    pub slide_index: usize,
    pub slide_id: u64,
    pub drawable_object_id: u64,
    pub kind: ChartKind,
    pub direction: ChartSeriesDirection,
    pub data: ChartData,
    pub geometry: DrawableGeometry,
}

/// Result of removing a standalone Keynote chart and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideChart {
    pub chart: KeynoteSlideChartInfo,
}

impl KeynoteEditor {
    /// List charts owned directly by one slide in z-order.
    pub fn slide_charts(&self, slide_index: usize) -> Result<Vec<KeynoteSlideChartInfo>> {
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_create::text_box_context(&graph, slide_index)?;
        let mut charts = Vec::new();
        for reference in context.slide.drawables_z_order {
            let messages = graph.objects.get(&reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide {} drawable {} is missing",
                    context.slide_id, reference.identifier
                ))
            })?;
            if messages
                .iter()
                .any(|message| message.type_ == CHART_MESSAGE_TYPE)
            {
                charts.push(chart_graph(self, slide_index, reference.identifier)?.info);
            }
        }
        Ok(charts)
    }

    /// Build an independently editable chart directly from typed inline data.
    ///
    /// The chart is a native slide-owned drawable and does not depend on a
    /// source table, copied chart, or bundled presentation template.
    pub fn add_slide_chart(
        &mut self,
        slide_index: usize,
        kind: ChartKind,
        data: ChartData,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<KeynoteSlideChartInfo> {
        require_creatable_kind(kind)?;
        let geometry = chart_geometry("Keynote", position, size)?;
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_create::text_box_context(&graph, slide_index)?;
        let archive_name = graph.archive_name(context.slide_id)?.to_owned();
        let component_id = component_identifier_for_entry(self.package(), &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide component {archive_name} is not registered"
                ))
            })?;
        let theme = chart_theme_context(self.package(), &graph, context.theme_id)?;
        let paragraph_archive = graph.archive_name(theme.paragraph_style_id)?.to_owned();
        let paragraph_component_id =
            component_identifier_for_entry(self.package(), &paragraph_archive)?.ok_or_else(
                || {
                    Error::InvalidFormat(format!(
                        "Keynote paragraph-style component {paragraph_archive} is not registered"
                    ))
                },
            )?;
        let ids = SourceChartObjectIds::allocate(
            next_object_identifier(self.package())?,
            ChartApplicationProfile::Keynote,
        )?;
        let objects = source_chart_objects(
            ids,
            context.slide_id,
            kind,
            data.clone(),
            geometry,
            theme.paragraph_style_id,
            ChartApplicationProfile::Keynote,
        )?;

        let mut staged = self.package().clone();
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_slide_drawable_references(
            &mut staged,
            &archive_name,
            context.slide_id,
            None,
            Some(ids.drawable),
        )?;
        patch_theme_chart_preset(&mut staged, &theme, None, Some(ids.preset))?;
        if theme.component_id != component_id {
            add_component_external_reference(
                &mut staged,
                theme.component_id,
                component_id,
                ids.preset,
            )?;
        }
        if paragraph_component_id != component_id {
            add_component_external_reference(
                &mut staged,
                component_id,
                paragraph_component_id,
                theme.paragraph_style_id,
            )?;
        }
        let object_ids = ids.all();
        add_component_object_uuids(&mut staged, component_id, &object_ids)?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = chart_graph(&verified, slide_index, ids.drawable)?;
        if created.info.kind != kind
            || created.info.direction != ChartSeriesDirection::Rows
            || created.info.data != data
            || created.info.geometry != geometry
            || created.object_ids != object_ids
        {
            return Err(Error::InvalidFormat(
                "Keynote chart creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Change the native kind of one slide chart while preserving its data.
    pub fn set_slide_chart_kind(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        kind: ChartKind,
    ) -> Result<()> {
        require_creatable_kind(kind)?;
        self.update_slide_chart(slide_index, drawable_object_id, |chart| {
            chart
                .chart
                .as_mut()
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote chart {drawable_object_id} has no chart payload"
                    ))
                })?
                .chart_type = Some(kind.into_raw());
            Ok(())
        })?;
        if chart_graph(self, slide_index, drawable_object_id)?
            .info
            .kind
            != kind
        {
            return Err(Error::InvalidFormat(
                "Keynote chart kind update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Replace the complete inline data grid of one slide chart.
    pub fn set_slide_chart_data(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        data: ChartData,
    ) -> Result<()> {
        self.update_slide_chart(slide_index, drawable_object_id, |chart| {
            let payload = chart.chart.as_mut().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote chart {drawable_object_id} has no chart payload"
                ))
            })?;
            payload.grid = Some(chart_grid(drawable_object_id, data.clone())?);
            payload.is_dirty = Some(false);
            Ok(())
        })?;
        if chart_graph(self, slide_index, drawable_object_id)?
            .info
            .data
            != data
        {
            return Err(Error::InvalidFormat(
                "Keynote chart data update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Set whether rows or columns form one slide chart's series.
    pub fn set_slide_chart_direction(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        direction: ChartSeriesDirection,
    ) -> Result<()> {
        if matches!(direction, ChartSeriesDirection::Unsupported(_)) {
            return Err(Error::ParseError(
                "cannot assign an unsupported chart series direction".to_owned(),
            ));
        }
        self.update_slide_chart(slide_index, drawable_object_id, |chart| {
            chart
                .chart
                .as_mut()
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote chart {drawable_object_id} has no chart payload"
                    ))
                })?
                .series_direction = Some(direction.into_raw());
            Ok(())
        })?;
        if chart_graph(self, slide_index, drawable_object_id)?
            .info
            .direction
            != direction
        {
            return Err(Error::InvalidFormat(
                "Keynote chart direction update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Update one chart's slide-space geometry.
    pub fn set_slide_chart_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        geometry.validate()?;
        self.update_slide_chart(slide_index, drawable_object_id, |chart| {
            let drawable = chart.drawable.super_.as_mut().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote chart {drawable_object_id} has no drawable payload"
                ))
            })?;
            drawable.geometry = Some(geometry_archive(geometry)?);
            Ok(())
        })?;
        if chart_graph(self, slide_index, drawable_object_id)?
            .info
            .geometry
            != geometry
        {
            return Err(Error::InvalidFormat(
                "Keynote chart geometry update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Duplicate one slide chart using Keynote's native placement.
    ///
    /// The clone receives fresh drawable, title/caption graph, style, preset, and UUID
    /// identities while retaining editable inline data and opaque protobuf
    /// fields. It is owned by the same slide but has an independent chart grid
    /// and geometry.
    pub fn duplicate_slide_chart(
        &mut self,
        slide_index: usize,
        source_drawable_object_id: u64,
    ) -> Result<KeynoteSlideChartInfo> {
        let source = chart_graph(self, slide_index, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Keynote chart graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Keynote chart object {identifier} is missing"))
                })?;
                clone_slide_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Keynote chart clone has no drawable identifier".to_owned())
        })?;
        let geometry = offset_drawable_geometry(source.info.geometry, DRAWABLE_DUPLICATE_OFFSET)?;
        update_chart_payload(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            |chart| {
                let drawable = chart.drawable.super_.as_mut().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote chart {new_drawable_id} has no drawable payload"
                    ))
                })?;
                drawable.geometry = Some(geometry_archive(geometry)?);
                Ok(())
            },
        )?;
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.info.slide_id,
            None,
            Some(new_drawable_id),
        )?;
        if let Some(source_preset_id) = source.private_preset_id {
            let graph = ObjectGraph::read(&staged)?;
            let context = text_box_create::text_box_context(&graph, slide_index)?;
            let theme = chart_theme_context(&staged, &graph, context.theme_id)?;
            let new_preset_id = remap.get(&source_preset_id).copied().ok_or_else(|| {
                Error::InvalidFormat("Keynote chart clone has no preset identifier".to_owned())
            })?;
            patch_theme_chart_preset(&mut staged, &theme, None, Some(new_preset_id))?;
            if theme.component_id != source.component_id {
                add_component_external_reference(
                    &mut staged,
                    theme.component_id,
                    source.component_id,
                    new_preset_id,
                )?;
            }
        }
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Keynote chart graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote chart clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = chart_graph(&verified, slide_index, new_drawable_id)?;
        let expected_object_ids = source
            .object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote chart clone has no validated identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        if created.info.kind != source.info.kind
            || created.info.direction != source.info.direction
            || created.info.data != source.info.data
            || created.info.geometry != geometry
            || created.object_ids != expected_object_ids
        {
            return Err(Error::InvalidFormat(
                "Keynote chart duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Remove a standalone slide chart and its private title, caption, and styles.
    pub fn remove_slide_chart(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideChart> {
        let source = chart_graph(self, slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.info.slide_id,
            Some(drawable_object_id),
            None,
        )?;
        if let Some(preset_id) = source.private_preset_id {
            let graph = ObjectGraph::read(&staged)?;
            let context = text_box_create::text_box_context(&graph, slide_index)?;
            let theme = chart_theme_context(&staged, &graph, context.theme_id)?;
            patch_theme_chart_preset(&mut staged, &theme, Some(preset_id), None)?;
        }
        for identifier in &source.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                source.component_id,
                *identifier,
            )?;
        }
        staged.update_archive(&source.archive_name, |archive| {
            for identifier in &source.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Keynote chart object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Keynote chart object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_charts(slide_index)?
            .iter()
            .any(|chart| chart.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Keynote chart deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedKeynoteSlideChart { chart: source.info })
    }

    fn update_slide_chart(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        update: impl FnOnce(&mut IWorkChartArchive) -> Result<()>,
    ) -> Result<()> {
        let source = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        update_chart_payload(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            update,
        )?;
        *self = Self::from_bytes(&staged.to_bytes()?)?;
        Ok(())
    }
}

fn update_chart_payload(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    update: impl FnOnce(&mut IWorkChartArchive) -> Result<()>,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote chart {drawable_object_id} is missing"))
        })?;
        let message_indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = message_indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote chart {drawable_object_id} must contain exactly one chart payload"
            )));
        };
        let mut chart = IWorkChartArchive::decode(&object.messages[*message_index].data)?;
        update(&mut chart)?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: chart.encode()?,
            },
        )?;
        Ok(())
    })
}

#[cfg(test)]
mod tests;
