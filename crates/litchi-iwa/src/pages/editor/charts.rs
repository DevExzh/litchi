//! Standalone, inline-data chart CRUD for Pages body attachments.

mod axis;
mod caption;
mod graph;
mod legend;
mod theme;
mod title;

use graph::{
    body_chart_graph, body_chart_infos, chart_attachment_object, set_chart_attachment_position,
};
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

const PAGES_THEME_MESSAGE_TYPE: u32 = 10_001;

/// One native inline-data chart anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesBodyChartInfo {
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub drawable_object_id: u64,
    pub kind: ChartKind,
    pub direction: ChartSeriesDirection,
    pub data: ChartData,
    pub geometry: DrawableGeometry,
}

/// Result of removing a body chart and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedPagesBodyChart {
    pub chart: PagesBodyChartInfo,
}

impl PagesEditor {
    /// List native charts anchored to the body in text-flow order.
    pub fn body_charts(&self) -> Result<Vec<PagesBodyChartInfo>> {
        body_chart_infos(self)
    }

    /// Build an independently editable chart at a UTF-16 body position.
    ///
    /// The chart, private styles, stand-ins, attachment, object-replacement
    /// character, theme preset registration, and z-order are constructed from
    /// typed values. No source chart, table, or package is copied.
    pub fn add_body_chart(
        &mut self,
        anchor_character_index: usize,
        kind: ChartKind,
        data: ChartData,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<PagesBodyChartInfo> {
        require_creatable_kind(kind)?;
        let geometry = chart_geometry("Pages", position, size)?;
        let root = root_document(self.package())?;
        let theme_id = root
            .theme
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".into()))?
            .identifier;
        let theme = chart_theme_context(self.package(), theme_id)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;
        let component_id = component_identifier_for_entry(self.package(), &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages body component {archive_name} is not registered"
                ))
            })?;
        let paragraph_archive = find_object_archive(self.package(), theme.paragraph_style_id)?;
        let paragraph_component_id =
            component_identifier_for_entry(self.package(), &paragraph_archive)?.ok_or_else(
                || {
                    Error::InvalidFormat(format!(
                        "Pages paragraph-style component {paragraph_archive} is not registered"
                    ))
                },
            )?;

        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".into()))?;
        let ids =
            SourceChartObjectIds::allocate(graph_first_identifier, ChartApplicationProfile::Pages)?;
        let attachment_id = ids
            .last()
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".into()))?;
        let mut objects = source_chart_objects(
            ids,
            self.body_storage_id,
            kind,
            data.clone(),
            geometry,
            theme.paragraph_style_id,
            ChartApplicationProfile::Pages,
        )?;
        objects.push(chart_attachment_object(
            attachment_id,
            ids.drawable,
            position,
            root.left_margin.unwrap_or_default(),
        )?);

        let mut staged = self.package().clone();
        if creates_z_order {
            text_box_create::create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            attachment_id,
        )?;
        patch_pages_zorder(&mut staged, None, Some(ids.drawable))?;
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
        let chart_object_ids = ids.all();
        add_component_object_uuids(&mut staged, component_id, &chart_object_ids)?;
        set_package_last_object_identifier(&mut staged, attachment_id)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = body_chart_graph(&verified, ids.drawable)?;
        let mut expected_object_ids = chart_object_ids;
        expected_object_ids.push(attachment_id);
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".into()))?;
        if created.info.anchor_character_index != expected_anchor
            || created.info.kind != kind
            || created.info.direction != ChartSeriesDirection::Rows
            || created.info.data != data
            || created.info.geometry != geometry
            || created.object_ids != expected_object_ids
        {
            return Err(Error::InvalidFormat(
                "Pages chart creation produced an inconsistent graph".into(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Change one body chart's native kind while preserving its data.
    pub fn set_body_chart_kind(&mut self, drawable_object_id: u64, kind: ChartKind) -> Result<()> {
        require_creatable_kind(kind)?;
        self.update_body_chart(drawable_object_id, |chart| {
            chart
                .chart
                .as_mut()
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages chart {drawable_object_id} has no chart payload"
                    ))
                })?
                .chart_type = Some(kind.into_raw());
            Ok(())
        })?;
        if body_chart_graph(self, drawable_object_id)?.info.kind != kind {
            return Err(Error::InvalidFormat(
                "Pages chart kind update failed validation".into(),
            ));
        }
        Ok(())
    }

    /// Replace the complete inline data grid of one body chart.
    pub fn set_body_chart_data(&mut self, drawable_object_id: u64, data: ChartData) -> Result<()> {
        self.update_body_chart(drawable_object_id, |chart| {
            let payload = chart.chart.as_mut().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages chart {drawable_object_id} has no chart payload"
                ))
            })?;
            payload.grid = Some(chart_grid(drawable_object_id, data.clone())?);
            payload.is_dirty = Some(false);
            Ok(())
        })?;
        if body_chart_graph(self, drawable_object_id)?.info.data != data {
            return Err(Error::InvalidFormat(
                "Pages chart data update failed validation".into(),
            ));
        }
        Ok(())
    }

    /// Set whether rows or columns form one body chart's series.
    pub fn set_body_chart_direction(
        &mut self,
        drawable_object_id: u64,
        direction: ChartSeriesDirection,
    ) -> Result<()> {
        if matches!(direction, ChartSeriesDirection::Unsupported(_)) {
            return Err(Error::ParseError(
                "cannot assign an unsupported chart series direction".into(),
            ));
        }
        self.update_body_chart(drawable_object_id, |chart| {
            chart
                .chart
                .as_mut()
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages chart {drawable_object_id} has no chart payload"
                    ))
                })?
                .series_direction = Some(direction.into_raw());
            Ok(())
        })?;
        if body_chart_graph(self, drawable_object_id)?.info.direction != direction {
            return Err(Error::InvalidFormat(
                "Pages chart direction update failed validation".into(),
            ));
        }
        Ok(())
    }

    /// Update one chart's page-space geometry and body-anchor offsets.
    pub fn set_body_chart_geometry(
        &mut self,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        geometry.validate()?;
        let position = geometry.position.ok_or_else(|| {
            Error::ParseError("Pages body chart geometry requires a position".into())
        })?;
        let source = body_chart_graph(self, drawable_object_id)?;
        let left_margin = root_document(self.package())?
            .left_margin
            .unwrap_or_default();
        let mut staged = self.package().clone();
        update_chart_payload(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            |chart| {
                chart
                    .drawable
                    .super_
                    .as_mut()
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Pages chart {drawable_object_id} has no drawable payload"
                        ))
                    })?
                    .geometry = Some(geometry_archive(geometry)?);
                Ok(())
            },
        )?;
        set_chart_attachment_position(
            &mut staged,
            &source.archive_name,
            source.attachment_id,
            position,
            left_margin,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if body_chart_graph(&verified, drawable_object_id)?
            .info
            .geometry
            != geometry
        {
            return Err(Error::InvalidFormat(
                "Pages chart geometry update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one body chart at a UTF-16 body position.
    ///
    /// The clone receives fresh drawable, title/caption graph, style, preset, attachment,
    /// and UUID identities while retaining the source chart's editable inline
    /// data and opaque protobuf fields. Its geometry and attachment are offset
    /// using Pages' native duplicate placement, so both body charts remain
    /// independently positioned and editable.
    pub fn duplicate_body_chart(
        &mut self,
        source_drawable_object_id: u64,
        anchor_character_index: usize,
    ) -> Result<PagesBodyChartInfo> {
        let source = body_chart_graph(self, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Pages chart graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages chart object {identifier} is missing"))
                })?;
                clone_pages_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Pages chart clone has no drawable identifier".to_owned())
        })?;
        let new_attachment_id = *remap.get(&source.attachment_id).ok_or_else(|| {
            Error::InvalidFormat("Pages chart clone has no attachment identifier".to_owned())
        })?;
        let geometry =
            offset_drawable_geometry(source.info.geometry, BODY_DRAWABLE_DUPLICATE_OFFSET)?;
        let position = geometry.position.ok_or_else(|| {
            Error::InvalidFormat("Pages chart clone geometry has no position".to_owned())
        })?;
        update_chart_payload(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            |chart| {
                let drawable = chart.drawable.super_.as_mut().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages chart {new_drawable_id} has no drawable payload"
                    ))
                })?;
                drawable.geometry = Some(geometry_archive(geometry)?);
                Ok(())
            },
        )?;
        let left_margin = root_document(&staged)?.left_margin.unwrap_or_default();
        set_chart_attachment_position(
            &mut staged,
            &source.archive_name,
            new_attachment_id,
            position,
            left_margin,
        )?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            new_attachment_id,
        )?;
        patch_pages_zorder(&mut staged, None, Some(new_drawable_id))?;
        if let Some(source_preset_id) = source.private_preset_id {
            let root = root_document(&staged)?;
            let theme_id = root
                .theme
                .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".into()))?
                .identifier;
            let theme = chart_theme_context(&staged, theme_id)?;
            let new_preset_id = remap.get(&source_preset_id).copied().ok_or_else(|| {
                Error::InvalidFormat("Pages chart clone has no preset identifier".to_owned())
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
            Error::InvalidFormat("Pages chart graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages chart clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = body_chart_graph(&verified, new_drawable_id)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".into()))?;
        let expected_object_ids = source
            .object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages chart clone has no validated identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        if created.info.anchor_character_index != expected_anchor
            || created.info.kind != source.info.kind
            || created.info.direction != source.info.direction
            || created.info.data != source.info.data
            || created.info.geometry != geometry
            || created.object_ids != expected_object_ids
        {
            return Err(Error::InvalidFormat(
                "Pages chart duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Remove a body chart, its attachment, and any crate-owned private styles.
    pub fn remove_body_chart(&mut self, drawable_object_id: u64) -> Result<RemovedPagesBodyChart> {
        let source = body_chart_graph(self, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = source.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;
        if let Some(preset_id) = source.private_preset_id {
            let root = root_document(&staged)?;
            let theme_id = root
                .theme
                .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".into()))?
                .identifier;
            let theme = chart_theme_context(&staged, theme_id)?;
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
                    Error::InvalidFormat(format!(
                        "Pages chart object {identifier} is missing from {}",
                        source.archive_name
                    ))
                })?;
            }
            Ok(())
        })?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages chart object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .body_charts()?
            .iter()
            .any(|chart| chart.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Pages chart deletion failed validation".into(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesBodyChart { chart: source.info })
    }

    fn update_body_chart(
        &mut self,
        drawable_object_id: u64,
        update: impl FnOnce(&mut IWorkChartArchive) -> Result<()>,
    ) -> Result<()> {
        let source = body_chart_graph(self, drawable_object_id)?;
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
            Error::InvalidFormat(format!("Pages chart {drawable_object_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| (message.type_ == CHART_MESSAGE_TYPE).then_some(index))
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages chart {drawable_object_id} must contain exactly one chart payload"
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
