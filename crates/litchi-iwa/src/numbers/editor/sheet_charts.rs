//! Standalone, inline-data chart CRUD for Numbers sheets.

mod arrangement;
mod axis;
mod axis_bounds;
mod axis_gridlines;
mod axis_label_affixes;
mod axis_label_angle;
mod axis_label_position_3d;
mod axis_labels;
mod axis_line;
mod axis_minimum_label;
mod axis_number_format;
mod axis_scale;
mod axis_series_names;
mod axis_steps;
mod axis_tick_marks;
mod background_fill;
mod bar_shape_3d;
mod border;
mod border_stroke;
mod caption;
mod category_labels;
mod depth_3d;
mod donut_inner_radius;
mod font;
mod gaps;
mod graph;
mod hidden_data;
mod legend;
mod lighting_3d;
mod pie_label_distance;
mod pie_labels;
mod pie_leader_lines;
mod pie_start_angle;
mod pie_wedge_explosion;
mod radar_grid_shape;
mod radar_series_style;
mod radar_start_angle;
mod reference_line;
mod rounded_corners;
mod scene_3d;
mod series_connection_line;
mod series_error_bar_auto_fit;
mod series_error_bars;
mod series_fill;
mod series_gap_3d;
mod series_stroke;
mod series_symbol;
mod series_symbol_fill;
mod series_symbol_outline;
mod series_trendline;
mod series_value_label_affixes;
mod series_value_label_auto_fit;
mod series_value_label_number_format;
mod series_value_labels;
mod shadow;
mod theme;
mod title;

use graph::chart_graph;
use theme::{chart_theme_context, patch_theme_chart_preset};

use std::collections::{HashMap, HashSet};

use super::*;
use crate::IWorkThemeArchive;
use crate::charts::reference_line::chart_reference_line_objects;
use crate::charts::source::{
    AXIS_NON_STYLE_MESSAGE_TYPE, AXIS_STYLE_MESSAGE_TYPE, CHART_MEDIATOR_MESSAGE_TYPE,
    CHART_MESSAGE_TYPE, CHART_NON_STYLE_MESSAGE_TYPE, CHART_PRESET_MESSAGE_TYPE,
    CHART_STYLE_MESSAGE_TYPE, ChartApplicationProfile, LEGEND_NON_STYLE_MESSAGE_TYPE,
    LEGEND_STYLE_MESSAGE_TYPE, SERIES_NON_STYLE_MESSAGE_TYPE, SERIES_STYLE_MESSAGE_TYPE,
    STANDIN_MESSAGE_TYPE, SourceChartObjectIds, chart_data_from_source, chart_geometry, chart_grid,
    drawable_geometry, geometry_archive, local_chart_style_ids, reference, register_chart_styles,
    require_creatable_kind, single_message_index, source_chart_objects, unregister_chart_styles,
    validate_chart_styles_registered,
};
use crate::charts::{
    ChartArrangement, ChartData, Direction, DirectionKind, IWorkChartArchive, Kind,
};
use crate::data_reference_registry::{
    clone_component_data_references, remove_component_data_references_for_objects,
};
use crate::protobuf::tsch;
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableSize, offset_drawable_geometry,
    remove_orphaned_image_asset,
};
use litchi_numbers::{
    ChartSelector as FocusedChartSelector, Package as FocusedNumbersPackage,
    SheetSelector as FocusedSheetSelector,
};

const NUMBERS_THEME_MESSAGE_TYPE: u32 = 12_009;

/// One chart drawable owned by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetChartInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub kind: Kind,
    pub direction: Direction,
    pub data: ChartData,
    pub geometry: DrawableGeometry,
    pub arrangement: ChartArrangement,
}

/// Result of removing a standalone chart and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedNumbersSheetChart {
    pub chart: NumbersSheetChartInfo,
}

impl NumbersEditor {
    /// List charts owned directly by one reachable sheet.
    pub fn sheet_charts(&self, sheet_id: u64) -> Result<Vec<NumbersSheetChartInfo>> {
        let (_, _, sheet) = numbers_sheet(self.package(), sheet_id)?;
        let locations = object_locations(self.package())?;
        let mut charts = Vec::new();
        let mut chart_positions = Vec::new();
        for reference in sheet.drawable_infos {
            let Some(archive_name) = locations.get(&reference.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} drawable {} is missing",
                    reference.identifier
                )));
            };
            let archive = self.package().archive(archive_name)?;
            let Some(object) = archive.object(reference.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} drawable {} is missing",
                    reference.identifier
                )));
            };
            if object
                .messages
                .iter()
                .any(|message| message.type_ == CHART_MESSAGE_TYPE)
            {
                chart_positions.push(charts.len());
                charts.push(chart_graph(self, sheet_id, reference.identifier)?.info);
            }
        }
        arrangement::fill_focused_chart_arrangements(
            self,
            sheet_id,
            &mut charts,
            &chart_positions,
        )?;
        Ok(charts)
    }

    /// Build a standalone chart directly from typed inline data.
    ///
    /// The chart does not depend on a source table or copied template graph.
    pub fn add_sheet_chart(
        &mut self,
        sheet_id: u64,
        kind: Kind,
        data: ChartData,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<NumbersSheetChartInfo> {
        require_creatable_kind(kind)?;
        let geometry = chart_geometry("Numbers", position, size)?;
        let (sheet_archive_name, _, _) = numbers_sheet(self.package(), sheet_id)?;
        let sheet_component_id =
            component_identifier_for_entry(self.package(), &sheet_archive_name)?.ok_or_else(
                || {
                    Error::InvalidFormat(format!(
                        "Numbers sheet component {sheet_archive_name} is not registered"
                    ))
                },
            )?;
        let archive_name = self
            .package()
            .calculation_engine_entry_name()?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers chart creation requires a CalculationEngine component".to_owned(),
                )
            })?
            .to_owned();
        let component_id = component_identifier_for_entry(self.package(), &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers chart component {archive_name} is not registered"
                ))
            })?;
        let theme = chart_theme_context(self.package())?;
        let ids = SourceChartObjectIds::allocate(
            next_object_identifier(self.package())?,
            ChartApplicationProfile::Numbers,
        )?;
        let objects = source_chart_objects(
            ids,
            sheet_id,
            kind,
            data.clone(),
            geometry,
            theme.stylesheet_id,
            theme.paragraph_style_id,
            ChartApplicationProfile::Numbers,
        )?;

        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        register_chart_styles(
            &mut staged,
            theme.stylesheet_id,
            &archive_name,
            &ids.style_ids(),
        )?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &sheet_archive_name,
            sheet_id,
            None,
            Some(ids.drawable),
        )?;
        if sheet_component_id != component_id {
            add_component_external_reference(
                &mut staged,
                sheet_component_id,
                component_id,
                ids.drawable,
            )?;
        }
        patch_theme_chart_preset(&mut staged, &theme, None, Some(ids.preset))?;
        if theme.component_id != component_id {
            add_component_external_reference(
                &mut staged,
                theme.component_id,
                component_id,
                ids.preset,
            )?;
        }
        if theme.stylesheet_component_id != component_id {
            add_component_external_reference(
                &mut staged,
                component_id,
                theme.stylesheet_component_id,
                theme.paragraph_style_id,
            )?;
        }
        add_component_object_uuids(&mut staged, component_id, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        validate_chart_styles_registered(
            verified.package(),
            theme.stylesheet_id,
            &archive_name,
            &ids.style_ids(),
        )?;
        let mut created = chart_graph(&verified, sheet_id, ids.drawable)?;
        created.info.arrangement =
            arrangement::focused_chart_arrangement_for_drawable(&verified, sheet_id, ids.drawable)?;
        if created.info.kind != kind
            || created.info.direction != Direction::Rows
            || created.info.data != data
            || created.info.geometry != geometry
            || created.object_ids != ids.all()
        {
            return Err(Error::InvalidFormat(
                "Numbers chart creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Change the native kind of one chart while preserving its data and graph.
    pub fn set_sheet_chart_kind(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        kind: Kind,
    ) -> Result<()> {
        require_creatable_kind(kind)?;
        self.update_sheet_chart(sheet_id, drawable_object_id, |chart| {
            chart
                .chart
                .as_mut()
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers chart {drawable_object_id} has no chart payload"
                    ))
                })?
                .chart_type = Some(kind.native_value());
            Ok(())
        })?;
        if chart_graph(self, sheet_id, drawable_object_id)?.info.kind != kind {
            return Err(Error::InvalidFormat(
                "Numbers chart kind update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Replace the complete inline data grid of one standalone chart.
    pub fn set_sheet_chart_data(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        data: ChartData,
    ) -> Result<()> {
        let source = chart_graph(self, sheet_id, drawable_object_id)?;
        if chart_data_shape_and_labels_match(&source.info.data, &data) {
            let (sheet_index, chart_position) =
                arrangement::focused_chart_data_target(self, sheet_id, drawable_object_id)?;
            match set_focused_sheet_chart_data(self, sheet_index, chart_position, data) {
                Ok(()) => return Ok(()),
                Err(FocusedSheetChartDataError::UnsupportedDependency(data)) => {
                    return self.set_sheet_chart_data_full_replace(
                        sheet_id,
                        drawable_object_id,
                        data,
                    );
                },
                Err(FocusedSheetChartDataError::Host(error)) => return Err(error),
            }
        }
        self.set_sheet_chart_data_full_replace(sheet_id, drawable_object_id, data)
    }

    /// Replace the complete inline data grid through the legacy generated
    /// archive path when labels or dimensions change.
    fn set_sheet_chart_data_full_replace(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        data: ChartData,
    ) -> Result<()> {
        self.update_sheet_chart(sheet_id, drawable_object_id, |chart| {
            let payload = chart.chart.as_mut().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers chart {drawable_object_id} has no chart payload"
                ))
            })?;
            payload.grid = Some(chart_grid(drawable_object_id, data.clone())?);
            payload.is_dirty = Some(false);
            Ok(())
        })?;
        if chart_graph(self, sheet_id, drawable_object_id)?.info.data != data {
            return Err(Error::InvalidFormat(
                "Numbers chart data update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Set whether rows or columns form the chart's series.
    pub fn set_sheet_chart_direction(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        direction: Direction,
    ) -> Result<()> {
        if direction.is_unsupported() {
            return Err(Error::ParseError(
                "cannot assign an unsupported chart series direction".to_owned(),
            ));
        }
        self.update_sheet_chart(sheet_id, drawable_object_id, |chart| {
            chart
                .chart
                .as_mut()
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers chart {drawable_object_id} has no chart payload"
                    ))
                })?
                .series_direction = Some(direction.native_value());
            Ok(())
        })?;
        if chart_graph(self, sheet_id, drawable_object_id)?
            .info
            .direction
            != direction
        {
            return Err(Error::InvalidFormat(
                "Numbers chart direction update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Update one chart's sheet-space geometry.
    pub fn set_sheet_chart_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        geometry.validate()?;
        self.update_sheet_chart(sheet_id, drawable_object_id, |chart| {
            let drawable = chart.drawable.super_.as_mut().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers chart {drawable_object_id} has no drawable payload"
                ))
            })?;
            drawable.geometry = Some(geometry_archive(geometry)?);
            Ok(())
        })?;
        if chart_graph(self, sheet_id, drawable_object_id)?
            .info
            .geometry
            != geometry
        {
            return Err(Error::InvalidFormat(
                "Numbers chart geometry update failed validation".to_owned(),
            ));
        }
        Ok(())
    }

    /// Duplicate one sheet chart using Numbers' native placement.
    ///
    /// The clone receives fresh drawable, title/caption graph, mediator, style, preset,
    /// and UUID identities while retaining editable inline data and opaque
    /// protobuf fields. The source and clone have independent chart grids and
    /// are both owned directly by the same sheet.
    pub fn duplicate_sheet_chart(
        &mut self,
        sheet_id: u64,
        source_drawable_object_id: u64,
    ) -> Result<NumbersSheetChartInfo> {
        let source_arrangement = arrangement::focused_chart_arrangement_for_drawable(
            self,
            sheet_id,
            source_drawable_object_id,
        )?;
        let mut source = chart_graph(self, sheet_id, source_drawable_object_id)?;
        source.info.arrangement = source_arrangement;
        let source_style_ids =
            local_chart_style_ids(self.package(), &source.archive_name, &source.object_ids)?;
        let mut staged = self.package.clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Numbers chart graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers chart object {identifier} is missing"))
                })?;
                clone_numbers_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
        }
        let new_style_ids = source_style_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers chart clone has no style identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let theme = chart_theme_context(&staged)?;
        register_chart_styles(
            &mut staged,
            theme.stylesheet_id,
            &source.archive_name,
            &new_style_ids,
        )?;

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers chart clone has no drawable identifier".to_owned())
        })?;
        let geometry = offset_drawable_geometry(source.info.geometry, DRAWABLE_DUPLICATE_OFFSET)?;
        update_chart_payload(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            |chart| {
                let drawable = chart.drawable.super_.as_mut().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers chart {new_drawable_id} has no drawable payload"
                    ))
                })?;
                drawable.geometry = Some(geometry_archive(geometry)?);
                Ok(())
            },
        )?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.sheet_archive_name,
            sheet_id,
            None,
            Some(new_drawable_id),
        )?;
        if source.sheet_component_id != source.component_id {
            add_component_external_reference(
                &mut staged,
                source.sheet_component_id,
                source.component_id,
                new_drawable_id,
            )?;
        }
        if let Some(source_preset_id) = source.private_preset_id {
            let new_preset_id = remap.get(&source_preset_id).copied().ok_or_else(|| {
                Error::InvalidFormat("Numbers chart clone has no preset identifier".to_owned())
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
            Error::InvalidFormat("Numbers chart graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers chart clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;
        clone_component_data_references(&mut staged, source.component_id, &remap)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        validate_chart_styles_registered(
            verified.package(),
            theme.stylesheet_id,
            &source.archive_name,
            &new_style_ids,
        )?;
        let mut created = chart_graph(&verified, sheet_id, new_drawable_id)?;
        created.info.arrangement = arrangement::focused_chart_arrangement_for_drawable(
            &verified,
            sheet_id,
            new_drawable_id,
        )?;
        let expected_object_ids = source
            .object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers chart clone has no validated identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        if created.info.kind != source.info.kind
            || created.info.direction != source.info.direction
            || created.info.data != source.info.data
            || created.info.geometry != geometry
            || created.info.arrangement != source.info.arrangement
            || created.object_ids != expected_object_ids
        {
            return Err(Error::InvalidFormat(
                "Numbers chart duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Remove a standalone chart and its private caption, title, and mediator objects.
    #[allow(deprecated)]
    pub fn remove_sheet_chart(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetChart> {
        let source_arrangement = arrangement::focused_chart_arrangement_for_drawable(
            self,
            sheet_id,
            drawable_object_id,
        )?;
        let mut source = chart_graph(self, sheet_id, drawable_object_id)?;
        source.info.arrangement = source_arrangement;
        let style_ids =
            local_chart_style_ids(self.package(), &source.archive_name, &source.object_ids)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(litchi_iwa_common::comment::DrawableId::from_raw(
            drawable_object_id,
        )?)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.sheet_archive_name,
            sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        let theme = chart_theme_context(&staged)?;
        unregister_chart_styles(
            &mut staged,
            theme.stylesheet_id,
            &source.archive_name,
            &style_ids,
        )?;
        if let Some(preset_id) = source.private_preset_id {
            patch_theme_chart_preset(&mut staged, &theme, Some(preset_id), None)?;
        }
        for identifier in &source.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                source.component_id,
                *identifier,
            )?;
        }
        let affected_data_identifiers = remove_component_data_references_for_objects(
            &mut staged,
            source.component_id,
            &source.object_ids,
        )?;
        staged.update_archive(&source.archive_name, |archive| {
            for identifier in &source.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers chart object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let chart_component_still_used =
            staged
                .archive(&source.archive_name)?
                .objects
                .iter()
                .any(|object| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == CHART_MESSAGE_TYPE)
                });
        if !chart_component_still_used && source.component_id != theme.stylesheet_component_id {
            remove_component_external_reference(
                &mut staged,
                source.component_id,
                theme.stylesheet_component_id,
                theme.paragraph_style_id,
            )?;
            remove_component_external_reference(
                &mut staged,
                source.component_id,
                theme.stylesheet_component_id,
                theme.stylesheet_id,
            )?;
        }
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers chart object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
        for data_identifier in affected_data_identifiers {
            remove_orphaned_image_asset(&mut staged, Some(data_identifier))?;
        }
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .sheet_charts(sheet_id)?
            .iter()
            .any(|chart| chart.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers chart deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersSheetChart { chart: source.info })
    }

    fn update_sheet_chart(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        update: impl FnOnce(&mut IWorkChartArchive) -> Result<()>,
    ) -> Result<()> {
        let source = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
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

fn chart_data_shape_and_labels_match(before: &ChartData, after: &ChartData) -> bool {
    before.row_names() == after.row_names()
        && before.column_names() == after.column_names()
        && before.values().len() == after.values().len()
        && before
            .values()
            .iter()
            .zip(after.values())
            .all(|(before_row, after_row)| before_row.len() == after_row.len())
}

fn set_focused_sheet_chart_data(
    editor: &mut NumbersEditor,
    sheet_index: usize,
    chart_position: usize,
    data: ChartData,
) -> std::result::Result<(), FocusedSheetChartDataError> {
    let source_bytes = editor
        .to_bytes()
        .map_err(FocusedSheetChartDataError::Host)?;
    let focused = FocusedNumbersPackage::from_bytes(&source_bytes).map_err(|error| {
        FocusedSheetChartDataError::Host(Error::InvalidFormat(format!(
            "focused Numbers chart data source failed: {error}"
        )))
    })?;
    let edit = match focused.edit_sheet_chart_data(
        FocusedSheetSelector::index(sheet_index),
        FocusedChartSelector::index(chart_position),
    ) {
        Ok(edit) => edit,
        Err(error) => {
            if matches!(
                &error,
                litchi_numbers::ChartDataError::UnsupportedDependency
            ) {
                return Err(FocusedSheetChartDataError::UnsupportedDependency(data));
            }
            return Err(FocusedSheetChartDataError::Host(Error::InvalidFormat(
                format!("focused Numbers chart data edit failed: {error}"),
            )));
        },
    };
    let committed = edit.set(data).commit().map_err(|error| {
        if matches!(
            &error,
            litchi_numbers::ChartDataError::UnsupportedDependency
        ) {
            FocusedSheetChartDataError::Host(Error::InvalidFormat(
                "focused Numbers chart data commit rejected a previously admitted dependency"
                    .to_owned(),
            ))
        } else {
            FocusedSheetChartDataError::Host(Error::InvalidFormat(format!(
                "focused Numbers chart data commit failed: {error}"
            )))
        }
    })?;
    let mut target_bytes = Vec::new();
    committed
        .package()
        .write_to(&mut target_bytes)
        .map_err(|error| {
            FocusedSheetChartDataError::Host(Error::InvalidFormat(format!(
                "focused Numbers chart data package write failed: {error}"
            )))
        })?;
    let verified =
        NumbersEditor::from_bytes(&target_bytes).map_err(FocusedSheetChartDataError::Host)?;
    *editor = verified;
    Ok(())
}

enum FocusedSheetChartDataError {
    UnsupportedDependency(ChartData),
    Host(Error),
}

fn update_chart_payload(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    update: impl FnOnce(&mut IWorkChartArchive) -> Result<()>,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {drawable_object_id} is missing"))
        })?;
        let Some(message_index) = single_message_index(&object.messages, CHART_MESSAGE_TYPE) else {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {drawable_object_id} must contain exactly one chart payload"
            )));
        };
        let mut chart = IWorkChartArchive::decode(&object.messages[message_index].data)?;
        update(&mut chart)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: chart.encode()?,
            },
        )?;
        Ok(())
    })
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::PathBuf;

    use super::*;
    use litchi_iwa_common::chart::axis::style::Visibility as AxisVisibility;
    use litchi_iwa_common::chart::error_bar::{
        CustomValues as ErrorBarCustomValues, Direction as ErrorBarDirection,
        FixedValue as ErrorBarFixedValue, Series,
    };
    use litchi_iwa_common::chart::gaps::{Percentage, Spacing};

    use crate::charts::object_container::is_object_container_archive;
    use crate::charts::source::SERIES_NON_STYLE_MESSAGE_TYPE;
    use crate::charts::unique_chart_object_archive_name;
    use crate::charts::{
        Axis, Bound, Bounds, ChartCornerRadius, ChartDonutInnerRadius, ChartFont, ChartFontSize,
        ChartLegendFill, ChartLegendFont, ChartLegendFontSize, ChartLegendFrame, ChartLegendRect,
        ChartLegendShadow, ChartLegendStroke, ChartPieLabelDistance, ChartPieStartAngle,
        ChartPieWedgeExplosion, ChartPieWedgeIndex, ChartRoundedCorners,
        ChartSeriesErrorBarAutoFit, ChartSeriesStroke, ChartSeriesStrokePattern,
        ChartSeriesTrendline, ChartSeriesTrendlineMovingAveragePeriod,
        ChartSeriesTrendlinePolynomialOrder, ChartSeriesValueLabelAutoFit,
        ChartSeriesValueLabelLocation, ChartShadow, DecimalPlaces, Index, LabelAffixes,
        LabelVisibility, LeaderLineVisibility, MajorStepCount, MinorStepCount, NegativeStyle,
        NumberFormat, Scale, Steps, TickMarkLocation, Visibility,
    };
    use crate::numbers::NumbersDocumentBuilder;
    use crate::package_metadata::{component_identifier_for_entry, component_uuid_identifiers};
    use crate::protobuf::tsch;
    use crate::shapes::{
        Appearance, BlurRadius, Drop, Offset, Pattern, RgbColorSpace, RgbaColor, ShapeFill,
        ShapeImageFillTechnique, Stroke, Width,
    };
    use litchi_iwa_common::shape::shadow::{Angle, Opacity};

    const POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 120.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 420.0,
        height: 280.0,
    };

    fn sample_data() -> ChartData {
        ChartData::new(
            vec!["North".to_owned(), "South".to_owned()],
            vec!["Q1".to_owned(), "Q2".to_owned()],
            vec![vec![Some(12.0), Some(18.0)], vec![Some(9.0), Some(21.0)]],
        )
        .unwrap()
    }

    fn pie_data() -> ChartData {
        ChartData::new(
            vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
            vec!["Revenue".to_owned()],
            vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
        )
        .unwrap()
    }

    const UNKNOWN_ARCHIVE_INFO_FIELD: u32 = 4_095;

    fn test_encode_varint(mut value: u64) -> Vec<u8> {
        let mut encoded = Vec::new();
        while value >= 0x80 {
            encoded.push((value as u8) | 0x80);
            value >>= 7;
        }
        encoded.push(value as u8);
        encoded
    }

    fn test_varint_width(value: u64) -> usize {
        test_encode_varint(value).len()
    }

    fn test_decode_varint(source: &[u8]) -> (usize, usize) {
        let mut value = 0usize;
        let mut shift = 0usize;
        for (index, byte) in source.iter().copied().enumerate() {
            value |= usize::from(byte & 0x7f) << shift;
            if byte & 0x80 == 0 {
                return (value, index + 1);
            }
            shift += 7;
        }
        panic!("truncated test varint")
    }

    fn numbers_raw_fields(data: &[u8], number: u32) -> Vec<Vec<u8>> {
        crate::wire::parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number() == number)
            .map(|field| data[field.start()..field.end()].to_vec())
            .collect()
    }

    fn numbers_chart_caption_data(
        editor: &NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Vec<u8> {
        let graph = chart_graph(editor, sheet_id, drawable_object_id).unwrap();
        editor
            .package
            .archive(&graph.archive_name)
            .unwrap()
            .object(drawable_object_id)
            .unwrap()
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .unwrap()
            .data
            .clone()
    }

    fn numbers_chart_caption_reference(
        editor: &NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> u64 {
        let data = numbers_chart_caption_data(editor, sheet_id, drawable_object_id);
        numbers_chart_caption_reference_from_data(&data)
    }

    fn numbers_chart_caption_reference_from_data(data: &[u8]) -> u64 {
        IWorkChartArchive::decode(&data)
            .unwrap()
            .drawable
            .super_
            .unwrap()
            .caption
            .unwrap()
            .identifier
    }

    fn numbers_chart_caption_reference_in_archive(
        editor: &NumbersEditor,
        archive_name: &str,
        drawable_object_id: u64,
    ) -> u64 {
        let archive = editor.package.archive(archive_name).unwrap();
        let data = archive
            .object(drawable_object_id)
            .unwrap()
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .unwrap()
            .data
            .as_slice();
        numbers_chart_caption_reference_from_data(data)
    }

    fn numbers_chart_message_info(
        editor: &NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> crate::archive::MessageInfo {
        let graph = chart_graph(editor, sheet_id, drawable_object_id).unwrap();
        let archive = editor.package.archive(&graph.archive_name).unwrap();
        let object = archive.object(drawable_object_id).unwrap();
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == CHART_MESSAGE_TYPE)
            .unwrap();
        object.archive_info.message_infos[message_index].clone()
    }

    fn numbers_archive_header(
        editor: &NumbersEditor,
        archive_name: &str,
        object_identifier: u64,
    ) -> Vec<u8> {
        let compressed = editor.package.entry(archive_name).unwrap();
        let source = crate::snappy::SnappyStream::decompress(compressed)
            .unwrap()
            .into_bytes();
        let archive = crate::archive::Archive::parse(&source).unwrap();
        let object = archive.object(object_identifier).unwrap();
        let object_start = usize::try_from(object.header_offset).unwrap();
        let (_, prefix_length) = test_decode_varint(&source[object_start..]);
        let header_start = object_start + prefix_length;
        let header_end = usize::try_from(object.data_offset).unwrap();
        source[header_start..header_end].to_vec()
    }

    fn inject_numbers_archive_unknown_header(
        editor: &mut NumbersEditor,
        archive_name: &str,
        object_identifier: u64,
    ) -> Vec<u8> {
        let mut package = editor.package.clone();
        let compressed = package.entry(archive_name).unwrap().to_vec();
        let source = crate::snappy::SnappyStream::decompress(&compressed)
            .unwrap()
            .into_bytes();
        let archive = crate::archive::Archive::parse(&source).unwrap();
        let object = archive.object(object_identifier).unwrap();
        let object_start = usize::try_from(object.header_offset).unwrap();
        let (header_length, prefix_length) = test_decode_varint(&source[object_start..]);
        let header_start = object_start + prefix_length;
        let header_end = usize::try_from(object.data_offset).unwrap();
        assert_eq!(header_end - header_start, header_length);
        let payload_end = header_end + usize::try_from(object.data_length).unwrap();

        let mut unknown = Vec::new();
        crate::wire::append_varint_field(&mut unknown, UNKNOWN_ARCHIVE_INFO_FIELD, 42).unwrap();
        let rewritten_header_length = header_length + unknown.len();
        let mut rewritten = Vec::with_capacity(source.len() + unknown.len());
        rewritten.extend_from_slice(&source[..object_start]);
        rewritten.extend_from_slice(&test_encode_varint(rewritten_header_length as u64));
        rewritten.extend_from_slice(&source[header_start..header_end]);
        rewritten.extend_from_slice(&unknown);
        rewritten.extend_from_slice(&source[header_end..payload_end]);
        rewritten.extend_from_slice(&source[payload_end..]);
        package
            .insert_entry(
                archive_name,
                crate::snappy::SnappyStream::compress(&rewritten).unwrap(),
            )
            .unwrap();
        *editor = NumbersEditor::from_bytes(&package.to_bytes().unwrap()).unwrap();
        unknown
    }

    fn reserve_numbers_three_byte_caption_identifier(
        editor: &mut NumbersEditor,
        archive_name: &str,
    ) {
        let next = crate::package_metadata::next_object_identifier(&editor.package).unwrap();
        if next < 16_383 {
            let mut package = editor.package.clone();
            package
                .update_archive(archive_name, |archive| {
                    archive.insert_object(crate::archive::ArchiveObject::new(
                        16_383,
                        vec![RawMessage {
                            type_: 65_534,
                            data: vec![0],
                        }],
                    )?)?;
                    Ok(())
                })
                .unwrap();
            *editor = NumbersEditor::from_bytes(&package.to_bytes().unwrap()).unwrap();
        }
    }

    #[derive(Debug, Clone, Copy)]
    enum NumbersCaptionMetadataMode {
        AggregateDuplicate,
        FieldOnly,
        WrongFieldPath,
        DuplicateField,
        StaleField,
        NewField,
        DataReference,
        AuthorizedField,
    }

    fn mutate_numbers_caption_metadata(
        editor: &mut NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
        old_reference_id: u64,
        replacement_id: u64,
        mode: NumbersCaptionMetadataMode,
    ) {
        let graph = chart_graph(editor, sheet_id, drawable_object_id).unwrap();
        let mut package = editor.package.clone();
        package
            .update_archive(&graph.archive_name, |archive| {
                let object = archive.object_mut(drawable_object_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                    .unwrap();
                let info = &mut object.archive_info.message_infos[message_index];
                match mode {
                    NumbersCaptionMetadataMode::AggregateDuplicate => {
                        info.object_references.push(old_reference_id);
                    },
                    NumbersCaptionMetadataMode::FieldOnly => {
                        info.object_references.retain(|id| *id != old_reference_id);
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                            object_references: vec![old_reference_id],
                            ..Default::default()
                        });
                    },
                    NumbersCaptionMetadataMode::WrongFieldPath => {
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![9, 9]),
                            object_references: vec![old_reference_id],
                            ..Default::default()
                        });
                    },
                    NumbersCaptionMetadataMode::DuplicateField => {
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                            object_references: vec![old_reference_id, old_reference_id],
                            ..Default::default()
                        });
                    },
                    NumbersCaptionMetadataMode::StaleField => {
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                            object_references: vec![old_reference_id, 999_999],
                            ..Default::default()
                        });
                    },
                    NumbersCaptionMetadataMode::NewField => {
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                            object_references: vec![replacement_id],
                            ..Default::default()
                        });
                    },
                    NumbersCaptionMetadataMode::DataReference => {
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                            data_references: vec![old_reference_id],
                            ..Default::default()
                        });
                    },
                    NumbersCaptionMetadataMode::AuthorizedField => {
                        info.field_infos.push(crate::archive::FieldInfo {
                            path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                            object_references: vec![old_reference_id],
                            ..Default::default()
                        });
                    },
                }
                Ok(())
            })
            .unwrap();
        *editor = NumbersEditor::from_bytes(&package.to_bytes().unwrap()).unwrap();
    }

    fn numbers_caption_storage_id(
        editor: &NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> u64 {
        let reference_id = numbers_chart_caption_reference(editor, sheet_id, drawable_object_id);
        let graph = chart_graph(editor, sheet_id, drawable_object_id).unwrap();
        let archive = editor.package.archive(&graph.archive_name).unwrap();
        let object = archive.object(reference_id).unwrap();
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == crate::image_caption::CAPTION_INFO_MESSAGE_TYPE)
            .unwrap();
        litchi_iwa_protos::pages_movie_caption_codec::decode_caption_info(
            &message.data,
            litchi_iwa_protos::pages_movie_caption_codec::DecodeOptions::new(
                message.data.len(),
                message.data.len().saturating_mul(4),
                message.data.len().saturating_mul(64),
                8,
            ),
        )
        .unwrap()
        .owned_storage_identifier()
        .unwrap()
    }

    fn make_numbers_caption_storage_shared(
        editor: &mut NumbersEditor,
        sheet_id: u64,
        source_drawable_object_id: u64,
        shared_drawable_object_id: u64,
    ) {
        let source_storage_id =
            numbers_caption_storage_id(editor, sheet_id, source_drawable_object_id);
        let shared_reference_id =
            numbers_chart_caption_reference(editor, sheet_id, shared_drawable_object_id);
        let old_storage_id =
            numbers_caption_storage_id(editor, sheet_id, shared_drawable_object_id);
        let graph = chart_graph(editor, sheet_id, shared_drawable_object_id).unwrap();
        let mut package = editor.package.clone();
        package
            .update_archive(&graph.archive_name, |archive| {
                let object = archive.object_mut(shared_reference_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| {
                        message.type_ == crate::image_caption::CAPTION_INFO_MESSAGE_TYPE
                    })
                    .unwrap();
                let original = object.messages[message_index].data.clone();
                let data = litchi_iwa_protos::pages_movie_caption_codec::rewrite_caption_info(
                    &original,
                    litchi_iwa_protos::pages_movie_caption_codec::CaptionInfoWrite::new(&[(
                        old_storage_id,
                        source_storage_id,
                    )]),
                    litchi_iwa_protos::pages_movie_caption_codec::DecodeOptions::new(
                        original.len(),
                        original.len().saturating_mul(4),
                        original.len().saturating_mul(64),
                        8,
                    ),
                )
                .unwrap();
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: crate::image_caption::CAPTION_INFO_MESSAGE_TYPE,
                        data,
                    },
                )?;
                for identifier in
                    &mut object.archive_info.message_infos[message_index].object_references
                {
                    if *identifier == old_storage_id {
                        *identifier = source_storage_id;
                    }
                }
                Ok(())
            })
            .unwrap();
        *editor = NumbersEditor::from_bytes(&package.to_bytes().unwrap()).unwrap();
    }

    fn make_numbers_caption_info_shared(
        editor: &mut NumbersEditor,
        sheet_id: u64,
        source_drawable_object_id: u64,
        shared_drawable_object_id: u64,
    ) {
        let source_caption_id =
            numbers_chart_caption_reference(editor, sheet_id, source_drawable_object_id);
        let shared_caption_id =
            numbers_chart_caption_reference(editor, sheet_id, shared_drawable_object_id);
        let graph = chart_graph(editor, sheet_id, shared_drawable_object_id).unwrap();
        let limits = editor.package.limits();
        let mut package = editor.package.clone();
        package
            .update_archive(&graph.archive_name, |archive| {
                let object = archive.object_mut(shared_drawable_object_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                    .unwrap();
                let original = object.messages[message_index].data.clone();
                let data = crate::charts::caption_edge::rewrite_chart_caption_identifier(
                    limits,
                    &original,
                    source_caption_id,
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: CHART_MESSAGE_TYPE,
                        data,
                    },
                )?;
                let info = &mut object.archive_info.message_infos[message_index];
                for identifier in &mut info.object_references {
                    if *identifier == shared_caption_id {
                        *identifier = source_caption_id;
                    }
                }
                for field in &mut info.field_infos {
                    for identifier in &mut field.object_references {
                        if *identifier == shared_caption_id {
                            *identifier = source_caption_id;
                        }
                    }
                }
                Ok(())
            })
            .unwrap();
        *editor = NumbersEditor::from_bytes(&package.to_bytes().unwrap()).unwrap();
    }

    fn move_numbers_caption_aggregate_owner_to_unrelated_object(
        editor: &mut NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
        unrelated_object_id: u64,
    ) {
        let caption_id = numbers_chart_caption_reference(editor, sheet_id, drawable_object_id);
        let graph = chart_graph(editor, sheet_id, unrelated_object_id).unwrap();
        let mut package = editor.package.clone();
        package
            .update_archive(&graph.archive_name, |archive| {
                let selected = archive.object_mut(drawable_object_id).unwrap();
                let selected_message_index = selected
                    .messages
                    .iter()
                    .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                    .unwrap();
                selected.archive_info.message_infos[selected_message_index]
                    .object_references
                    .retain(|identifier| *identifier != caption_id);

                let unrelated = archive.object_mut(unrelated_object_id).unwrap();
                let unrelated_message_index = unrelated
                    .messages
                    .iter()
                    .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                    .unwrap();
                unrelated.archive_info.message_infos[unrelated_message_index]
                    .object_references
                    .push(caption_id);
                Ok(())
            })
            .unwrap();
        *editor = NumbersEditor::from_bytes(&package.to_bytes().unwrap()).unwrap();
    }

    fn assert_series_non_styles_are_unstyled(
        editor: &NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
        expected_count: usize,
    ) {
        let graph = chart_graph(editor, sheet_id, drawable_object_id).unwrap();
        let archive = editor.package().archive(&graph.archive_name).unwrap();
        let chart = archive.object(drawable_object_id).unwrap();
        let chart = chart
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .and_then(|message| IWorkChartArchive::decode(message.data.as_slice()).ok())
            .and_then(|archive| archive.chart)
            .unwrap();
        let entries = chart.series_non_styles.unwrap().entries;
        assert_eq!(entries.len(), expected_count);
        for entry in entries {
            let non_style_archive_name = unique_chart_object_archive_name(
                editor.package(),
                entry.reference.identifier,
                "test series non-style",
            )
            .unwrap();
            assert_ne!(non_style_archive_name, graph.archive_name);
            assert!(
                is_object_container_archive(editor.package(), &non_style_archive_name).unwrap()
            );
            let non_style_archive = editor.package().archive(&non_style_archive_name).unwrap();
            let non_style = non_style_archive
                .object(entry.reference.identifier)
                .unwrap();
            let message = non_style
                .messages
                .iter()
                .find(|message| message.type_ == SERIES_NON_STYLE_MESSAGE_TYPE)
                .unwrap();
            let parent = tsch::ChartSeriesNonStyleArchive::decode(message.data.as_slice())
                .unwrap()
                .super_
                .unwrap();
            assert_eq!(parent.stylesheet, None);
            let component_id =
                component_identifier_for_entry(editor.package(), &non_style_archive_name)
                    .unwrap()
                    .unwrap();
            assert!(
                component_uuid_identifiers(editor.package(), component_id)
                    .unwrap()
                    .unwrap()
                    .contains(&entry.reference.identifier)
            );
        }
    }

    fn assert_series_non_styles_are_styled(
        editor: &NumbersEditor,
        sheet_id: u64,
        drawable_object_id: u64,
        expected_count: usize,
    ) {
        let graph = chart_graph(editor, sheet_id, drawable_object_id).unwrap();
        let archive = editor.package().archive(&graph.archive_name).unwrap();
        let chart = archive.object(drawable_object_id).unwrap();
        let chart = chart
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .and_then(|message| IWorkChartArchive::decode(message.data.as_slice()).ok())
            .and_then(|archive| archive.chart)
            .unwrap();
        let entries = chart.series_non_styles.unwrap().entries;
        assert_eq!(entries.len(), expected_count);

        let mut stylesheet_id = None;
        let style_ids = entries
            .iter()
            .map(|entry| {
                let non_style_archive_name = unique_chart_object_archive_name(
                    editor.package(),
                    entry.reference.identifier,
                    "test series non-style",
                )
                .unwrap();
                assert_eq!(non_style_archive_name, graph.archive_name);
                assert!(
                    !is_object_container_archive(editor.package(), &non_style_archive_name)
                        .unwrap()
                );
                let non_style_archive = editor.package().archive(&non_style_archive_name).unwrap();
                let non_style = non_style_archive
                    .object(entry.reference.identifier)
                    .unwrap();
                let message = non_style
                    .messages
                    .iter()
                    .find(|message| message.type_ == SERIES_NON_STYLE_MESSAGE_TYPE)
                    .unwrap();
                let parent = tsch::ChartSeriesNonStyleArchive::decode(message.data.as_slice())
                    .unwrap()
                    .super_
                    .unwrap();
                let parent_stylesheet_id = parent.stylesheet.unwrap().identifier;
                match stylesheet_id {
                    Some(expected) => assert_eq!(expected, parent_stylesheet_id),
                    None => stylesheet_id = Some(parent_stylesheet_id),
                }
                let component_id =
                    component_identifier_for_entry(editor.package(), &non_style_archive_name)
                        .unwrap()
                        .unwrap();
                assert!(
                    component_uuid_identifiers(editor.package(), component_id)
                        .unwrap()
                        .unwrap()
                        .contains(&entry.reference.identifier)
                );
                entry.reference.identifier
            })
            .collect::<Vec<_>>();

        validate_chart_styles_registered(
            editor.package(),
            stylesheet_id.unwrap(),
            &graph.archive_name,
            &style_ids,
        )
        .unwrap();
    }

    fn gap_spacing(between_items: f32, between_sets: f32) -> Spacing {
        Spacing::new(
            Percentage::new(between_items).unwrap(),
            Percentage::new(between_sets).unwrap(),
        )
    }

    fn chart_stroke(pattern: Pattern, width: f32) -> Stroke {
        Stroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(width).unwrap(),
            pattern,
        )
    }

    fn chart_series_stroke(pattern: ChartSeriesStrokePattern, width: f32) -> ChartSeriesStroke {
        ChartSeriesStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(width).unwrap(),
            pattern,
        )
    }

    fn chart_background_fill() -> ShapeFill {
        ShapeFill::Solid(RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap())
    }

    fn chart_shadow() -> ChartShadow {
        ChartShadow::Grouped(Drop::new(
            Appearance::new(
                RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
                BlurRadius::from_points(15).unwrap(),
                Offset::from_points(8.0).unwrap(),
                Opacity::new(0.6).unwrap(),
            ),
            Angle::from_degrees(60.0).unwrap(),
        ))
    }

    fn fixture(relative: &str) -> Vec<u8> {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
        fs::read(root.join(relative)).unwrap()
    }

    #[test]
    fn scratch_spreadsheet_supports_standalone_chart_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Scratch Chart")
            .table_name("Source Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();

        let created = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, Kind::Column2d);
        assert_eq!(created.direction, Direction::Rows);
        assert_eq!(created.data, sample_data());

        let replacement = ChartData::new(
            vec!["Revenue".to_owned()],
            vec!["2026".to_owned(), "2027".to_owned(), "2028".to_owned()],
            vec![vec![Some(30.0), Some(45.0), None]],
        )
        .unwrap();
        editor
            .set_sheet_chart_kind(sheet_id, created.drawable_object_id, Kind::Bar2d)
            .unwrap();

        let same_shape = ChartData::new(
            vec!["North".to_owned(), "South".to_owned()],
            vec!["Q1".to_owned(), "Q2".to_owned()],
            vec![vec![Some(15.0), None], vec![Some(11.0), Some(23.0)]],
        )
        .unwrap();
        editor
            .set_sheet_chart_data(sheet_id, created.drawable_object_id, same_shape.clone())
            .unwrap();
        let focused = litchi_numbers::Package::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            focused.sheet_chart_data(0usize, 0usize).unwrap(),
            same_shape
        );
        assert_eq!(editor.sheet_charts(sheet_id).unwrap()[0].data, same_shape);

        // Changed labels and dimensions intentionally retain the legacy full-
        // grid replacement path until grid authoring owns those transitions.
        editor
            .set_sheet_chart_data(sheet_id, created.drawable_object_id, replacement.clone())
            .unwrap();
        editor
            .set_sheet_chart_direction(sheet_id, created.drawable_object_id, Direction::Columns)
            .unwrap();
        let changed_geometry = chart_geometry(
            "Numbers",
            DrawablePoint { x: 72.0, y: 360.0 },
            DrawableSize {
                width: 500.0,
                height: 300.0,
            },
        )
        .unwrap();
        editor
            .set_sheet_chart_geometry(sheet_id, created.drawable_object_id, changed_geometry)
            .unwrap();

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let chart = &reopened.sheet_charts(sheet_id).unwrap()[0];
        assert_eq!(chart.kind, Kind::Bar2d);
        assert_eq!(chart.direction, Direction::Columns);
        assert_eq!(chart.data, replacement);
        assert_eq!(chart.geometry, changed_geometry);

        let removed = editor
            .remove_sheet_chart(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.chart.drawable_object_id, created.drawable_object_id);
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn chart_creation_rejects_invalid_kind_and_geometry_transactionally() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();

        assert!(
            editor
                .add_sheet_chart(sheet_id, Kind::Undefined, sample_data(), POSITION, SIZE)
                .is_err()
        );
        assert!(
            editor
                .add_sheet_chart(
                    sheet_id,
                    Kind::Column2d,
                    sample_data(),
                    POSITION,
                    DrawableSize {
                        width: 0.0,
                        height: SIZE.height,
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn multiple_chart_theme_registrations_are_removed_independently() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        let first = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let second = editor
            .add_sheet_chart(
                sheet_id,
                Kind::Column2d,
                sample_data(),
                DrawablePoint {
                    x: POSITION.x + SIZE.width,
                    y: POSITION.y,
                },
                SIZE,
            )
            .unwrap();

        editor
            .remove_sheet_chart(sheet_id, first.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_charts(sheet_id)
                .unwrap()
                .iter()
                .map(|chart| chart.drawable_object_id)
                .collect::<Vec<_>>(),
            vec![second.drawable_object_id]
        );
        editor
            .remove_sheet_chart(sheet_id, second.drawable_object_id)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn duplicate_sheet_chart_clones_the_private_graph_and_inline_data() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let source_graph = chart_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(editor.duplicate_sheet_chart(sheet_id, u64::MAX).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        let duplicate_graph = chart_graph(&editor, sheet_id, duplicate.drawable_object_id).unwrap();
        let expected_geometry =
            offset_drawable_geometry(source.geometry, DRAWABLE_DUPLICATE_OFFSET).unwrap();

        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        assert_eq!(duplicate.kind, source.kind);
        assert_eq!(duplicate.direction, source.direction);
        assert_eq!(duplicate.data, source.data);
        assert_eq!(duplicate.geometry, expected_geometry);
        assert_eq!(
            duplicate_graph.object_ids.len(),
            source_graph.object_ids.len()
        );
        assert!(
            source_graph
                .object_ids
                .iter()
                .all(|identifier| !duplicate_graph.object_ids.contains(identifier))
        );

        let replacement = ChartData::new(
            vec!["Revenue".to_owned()],
            vec!["2026".to_owned(), "2027".to_owned()],
            vec![vec![Some(30.0), Some(45.0)]],
        )
        .unwrap();
        editor
            .set_sheet_chart_data(sheet_id, duplicate.drawable_object_id, replacement.clone())
            .unwrap();
        assert_eq!(
            chart_graph(&editor, sheet_id, source.drawable_object_id)
                .unwrap()
                .info
                .data,
            source.data
        );
        assert_eq!(
            chart_graph(&editor, sheet_id, duplicate.drawable_object_id)
                .unwrap()
                .info
                .data,
            replacement
        );

        editor
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_charts(sheet_id)
                .unwrap()
                .iter()
                .map(|chart| chart.drawable_object_id)
                .collect::<Vec<_>>(),
            vec![duplicate.drawable_object_id]
        );
        editor
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(editor.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_caption_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert_eq!(
            editor
                .sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap(),
            None
        );
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "Revenue by region")
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_caption(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );

        editor
            .set_sheet_chart_caption(
                sheet_id,
                source.drawable_object_id,
                "Updated source caption",
            )
            .unwrap();
        assert!(
            editor
                .remove_sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            !editor
                .remove_sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap(),
            None
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_caption(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            reopened
                .sheet_charts(sheet_id)
                .unwrap()
                .iter()
                .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
        );
    }

    #[test]
    fn native_chart_caption_rewrite_preserves_unknown_chart_fields() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "Revenue by region")
            .unwrap();

        let graph = chart_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        let unknown = [0xc0, 0x0c, 0x07];
        editor
            .package
            .update_archive(&graph.archive_name, |archive| {
                let object = archive.object_mut(source.drawable_object_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                    .unwrap();
                let mut data = object.messages[message_index].data.clone();
                data.extend_from_slice(&unknown);
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: CHART_MESSAGE_TYPE,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        assert!(
            editor
                .remove_sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap(),
            None
        );
        let archive = editor.package.archive(&graph.archive_name).unwrap();
        let chart_payload = archive
            .object(source.drawable_object_id)
            .unwrap()
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .unwrap()
            .data
            .as_slice();
        assert!(
            chart_payload
                .windows(unknown.len())
                .any(|window| window == unknown.as_slice())
        );
    }

    #[test]
    fn native_chart_caption_retarget_preserves_unknown_archive_header_and_metadata_on_width_growth()
    {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let graph = chart_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        reserve_numbers_three_byte_caption_identifier(&mut editor, &graph.archive_name);

        let old_reference_id =
            numbers_chart_caption_reference(&editor, sheet_id, source.drawable_object_id);
        let next_identifier =
            crate::package_metadata::next_object_identifier(&editor.package).unwrap();
        let replacement_id = next_identifier + 1;
        assert_ne!(
            test_varint_width(old_reference_id),
            test_varint_width(replacement_id),
            "fixture must exercise a caption-reference varint-width change (old={old_reference_id}, replacement={replacement_id})"
        );
        let before_info = numbers_chart_message_info(&editor, sheet_id, source.drawable_object_id);
        let unknown = inject_numbers_archive_unknown_header(
            &mut editor,
            &graph.archive_name,
            source.drawable_object_id,
        );
        let before_header =
            numbers_archive_header(&editor, &graph.archive_name, source.drawable_object_id);
        let before_unknown = numbers_raw_fields(&before_header, UNKNOWN_ARCHIVE_INFO_FIELD);
        assert_eq!(before_unknown, vec![unknown.clone()]);

        editor
            .set_sheet_chart_caption(
                sheet_id,
                source.drawable_object_id,
                "Caption after width growth",
            )
            .unwrap();
        let after_data = numbers_chart_caption_data(&editor, sheet_id, source.drawable_object_id);
        let after_info = numbers_chart_message_info(&editor, sheet_id, source.drawable_object_id);
        let mut expected_info = before_info;
        expected_info.length = u32::try_from(after_data.len()).unwrap();
        expected_info
            .object_references
            .retain(|identifier| *identifier != old_reference_id);
        expected_info.object_references.push(replacement_id);
        for field in &mut expected_info.field_infos {
            let had_old_reference = field
                .object_references
                .iter()
                .any(|identifier| *identifier == old_reference_id);
            field
                .object_references
                .retain(|identifier| *identifier != old_reference_id);
            if had_old_reference {
                field.object_references.push(replacement_id);
            }
        }
        assert_eq!(after_info, expected_info);
        let after_header =
            numbers_archive_header(&editor, &graph.archive_name, source.drawable_object_id);
        assert_eq!(
            numbers_raw_fields(&after_header, UNKNOWN_ARCHIVE_INFO_FIELD),
            before_unknown,
            "retarget must retain the complete unknown ArchiveInfo field"
        );
        assert_eq!(
            numbers_chart_caption_reference(&editor, sheet_id, source.drawable_object_id),
            replacement_id
        );
    }

    #[test]
    fn native_chart_caption_reference_transition_requires_exact_aggregate_and_field_metadata() {
        let modes = [
            NumbersCaptionMetadataMode::AggregateDuplicate,
            NumbersCaptionMetadataMode::FieldOnly,
            NumbersCaptionMetadataMode::WrongFieldPath,
            NumbersCaptionMetadataMode::DuplicateField,
            NumbersCaptionMetadataMode::StaleField,
            NumbersCaptionMetadataMode::NewField,
            NumbersCaptionMetadataMode::DataReference,
        ];
        for mode in modes {
            let mut editor = NumbersDocumentBuilder::new().build().unwrap();
            let sheet_id = editor.sheets().unwrap()[0].object_id;
            let source = editor
                .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
                .unwrap();
            let old_reference_id =
                numbers_chart_caption_reference(&editor, sheet_id, source.drawable_object_id);
            let replacement_id = crate::package_metadata::next_object_identifier(&editor.package)
                .unwrap()
                .saturating_add(1);
            mutate_numbers_caption_metadata(
                &mut editor,
                sheet_id,
                source.drawable_object_id,
                old_reference_id,
                replacement_id,
                mode,
            );
            let before = editor.to_bytes().unwrap();
            assert!(
                editor
                    .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "must reject")
                    .is_err(),
                "malformed caption metadata mode {mode:?} was accepted"
            );
            assert_eq!(editor.to_bytes().unwrap(), before);
        }

        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let old_reference_id =
            numbers_chart_caption_reference(&editor, sheet_id, source.drawable_object_id);
        let replacement_id = crate::package_metadata::next_object_identifier(&editor.package)
            .unwrap()
            .saturating_add(1);
        mutate_numbers_caption_metadata(
            &mut editor,
            sheet_id,
            source.drawable_object_id,
            old_reference_id,
            replacement_id,
            NumbersCaptionMetadataMode::AuthorizedField,
        );
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "authorized")
            .unwrap();
        let info = numbers_chart_message_info(&editor, sheet_id, source.drawable_object_id);
        assert_eq!(
            info.object_references
                .iter()
                .filter(|id| **id == old_reference_id)
                .count(),
            0
        );
        assert_eq!(
            info.object_references
                .iter()
                .filter(|id| **id == replacement_id)
                .count(),
            1
        );
        assert_eq!(info.field_infos.len(), 1);
        assert_eq!(info.field_infos[0].object_references, [replacement_id]);
    }

    #[test]
    fn native_chart_caption_rejects_two_caption_infos_sharing_one_storage_atomically() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "original")
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        make_numbers_caption_storage_shared(
            &mut editor,
            sheet_id,
            source.drawable_object_id,
            duplicate.drawable_object_id,
        );
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_chart_caption(
                    sheet_id,
                    source.drawable_object_id,
                    "must reject shared storage"
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn native_chart_caption_rejects_two_chart_payloads_sharing_one_caption_info_atomically() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "original")
            .unwrap();
        let graph = chart_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        make_numbers_caption_info_shared(
            &mut editor,
            sheet_id,
            source.drawable_object_id,
            duplicate.drawable_object_id,
        );
        assert_eq!(
            numbers_chart_caption_reference_in_archive(
                &editor,
                &graph.archive_name,
                source.drawable_object_id,
            ),
            numbers_chart_caption_reference_in_archive(
                &editor,
                &graph.archive_name,
                duplicate.drawable_object_id,
            )
        );

        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_chart_caption(
                    sheet_id,
                    source.drawable_object_id,
                    "must reject shared info"
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .set_sheet_chart_caption(
                    sheet_id,
                    duplicate.drawable_object_id,
                    "must reject shared info"
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn native_chart_caption_rejects_caption_info_owner_on_unrelated_object_atomically() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "original")
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        move_numbers_caption_aggregate_owner_to_unrelated_object(
            &mut editor,
            sheet_id,
            source.drawable_object_id,
            duplicate.drawable_object_id,
        );

        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_chart_caption(
                    sheet_id,
                    source.drawable_object_id,
                    "must reject stale owner"
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_title_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert_eq!(
            editor
                .sheet_chart_title(sheet_id, source.drawable_object_id)
                .unwrap(),
            None
        );
        editor
            .set_sheet_chart_title(sheet_id, source.drawable_object_id, "Revenue by region")
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_title(sheet_id, source.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_title(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );

        editor
            .set_sheet_chart_title(sheet_id, source.drawable_object_id, "Updated source title")
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_title(sheet_id, source.drawable_object_id)
                .unwrap(),
            Some("Updated source title".to_owned())
        );
        assert_eq!(
            editor
                .sheet_chart_title(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );
        assert!(
            editor
                .remove_sheet_chart_title(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            !editor
                .remove_sheet_chart_title(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .sheet_chart_title(sheet_id, source.drawable_object_id)
                .unwrap(),
            None
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_title(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            reopened
                .sheet_charts(sheet_id)
                .unwrap()
                .iter()
                .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_title_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        for axis in [Axis::Category, Axis::Value] {
            assert_eq!(
                editor
                    .sheet_chart_axis_title(sheet_id, source.drawable_object_id, axis)
                    .unwrap(),
                None
            );
        }
        editor
            .set_sheet_chart_axis_title(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                "Month",
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_title(sheet_id, source.drawable_object_id, Axis::Value, "Revenue")
            .unwrap();

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for (axis, title) in [(Axis::Category, "Month"), (Axis::Value, "Revenue")] {
            assert_eq!(
                editor
                    .sheet_chart_axis_title(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
                    .as_deref(),
                Some(title)
            );
            assert_eq!(
                editor
                    .sheet_chart_axis_title(sheet_id, duplicate.drawable_object_id, axis)
                    .unwrap()
                    .as_deref(),
                Some(title)
            );
        }

        editor
            .set_sheet_chart_axis_title(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                "Updated month",
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_title(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                "Updated revenue",
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_axis_title(sheet_id, source.drawable_object_id, Axis::Category)
                .unwrap()
                .as_deref(),
            Some("Updated month")
        );
        assert_eq!(
            editor
                .sheet_chart_axis_title(sheet_id, source.drawable_object_id, Axis::Value)
                .unwrap()
                .as_deref(),
            Some("Updated revenue")
        );
        assert_eq!(
            editor
                .sheet_chart_axis_title(sheet_id, duplicate.drawable_object_id, Axis::Category)
                .unwrap()
                .as_deref(),
            Some("Month")
        );
        assert_eq!(
            editor
                .sheet_chart_axis_title(sheet_id, duplicate.drawable_object_id, Axis::Value)
                .unwrap()
                .as_deref(),
            Some("Revenue")
        );

        for axis in [Axis::Category, Axis::Value] {
            assert!(
                editor
                    .remove_sheet_chart_axis_title(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
            assert!(
                !editor
                    .remove_sheet_chart_axis_title(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_axis_title(sheet_id, duplicate.drawable_object_id, Axis::Category)
                .unwrap()
                .as_deref(),
            Some("Month")
        );
        assert_eq!(
            reopened
                .sheet_chart_axis_title(sheet_id, duplicate.drawable_object_id, Axis::Value)
                .unwrap()
                .as_deref(),
            Some("Revenue")
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            reopened
                .sheet_charts(sheet_id)
                .unwrap()
                .iter()
                .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_value_axis_bounds_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let automatic = Bounds::automatic();
        let fixed = Bounds::fixed(Bound::new(-10.0).unwrap(), Bound::new(40.0).unwrap()).unwrap();
        let minimum_only = Bounds::new(Some(Bound::new(-5.0).unwrap()), None).unwrap();

        assert_eq!(
            editor
                .sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id)
                .unwrap(),
            automatic
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id, automatic)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id, fixed)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_value_axis_bounds(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            fixed
        );

        editor
            .set_sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id, minimum_only)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id)
                .unwrap(),
            minimum_only
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id)
                .unwrap(),
            minimum_only
        );
        assert_eq!(
            reopened
                .sheet_chart_value_axis_bounds(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            fixed
        );
        reopened
            .set_sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id, automatic)
            .unwrap();
        assert_eq!(
            reopened
                .sheet_chart_value_axis_bounds(sheet_id, source.drawable_object_id)
                .unwrap(),
            automatic
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_value_axis_steps_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = Steps::fixed(
            MajorStepCount::new(5).unwrap(),
            MinorStepCount::new(1).unwrap(),
        );
        let fixed = Steps::fixed(
            MajorStepCount::new(6).unwrap(),
            MinorStepCount::new(2).unwrap(),
        );
        let major_only = Steps::new(Some(MajorStepCount::new(4).unwrap()), None);

        assert_eq!(
            editor
                .sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id, defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id, fixed)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_value_axis_steps(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            fixed
        );

        editor
            .set_sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id, major_only)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id)
                .unwrap(),
            major_only
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id)
                .unwrap(),
            major_only
        );
        assert_eq!(
            reopened
                .sheet_chart_value_axis_steps(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            fixed
        );
        reopened
            .set_sheet_chart_value_axis_steps(
                sheet_id,
                source.drawable_object_id,
                Steps::automatic(),
            )
            .unwrap();
        assert_eq!(
            reopened
                .sheet_chart_value_axis_steps(sheet_id, source.drawable_object_id)
                .unwrap(),
            Steps::automatic()
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_value_axis_minimum_label_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert!(
            editor
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
                .is_visible()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                AxisVisibility::Visible,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert!(
            !editor
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
                .is_visible()
        );
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(
            !editor
                .sheet_chart_value_axis_minimum_label_visible(
                    sheet_id,
                    duplicate.drawable_object_id
                )
                .unwrap()
                .is_visible()
        );

        editor
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                AxisVisibility::Visible,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
                .is_visible()
        );
        assert!(
            !reopened
                .sheet_chart_value_axis_minimum_label_visible(
                    sheet_id,
                    duplicate.drawable_object_id
                )
                .unwrap()
                .is_visible()
        );
        reopened
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert!(
            !reopened
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
                .is_visible()
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_category_axis_series_names_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert!(
            !editor
                .sheet_chart_category_axis_series_names_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_category_axis_series_names_visible(
                sheet_id,
                source.drawable_object_id,
                false,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_category_axis_series_names_visible(
                sheet_id,
                source.drawable_object_id,
                true,
            )
            .unwrap();
        assert!(
            editor
                .sheet_chart_category_axis_series_names_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(
            editor
                .sheet_chart_category_axis_series_names_visible(
                    sheet_id,
                    duplicate.drawable_object_id
                )
                .unwrap()
        );

        editor
            .set_sheet_chart_category_axis_series_names_visible(
                sheet_id,
                source.drawable_object_id,
                false,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            !reopened
                .sheet_chart_category_axis_series_names_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            reopened
                .sheet_chart_category_axis_series_names_visible(
                    sheet_id,
                    duplicate.drawable_object_id
                )
                .unwrap()
        );
        reopened
            .set_sheet_chart_category_axis_series_names_visible(
                sheet_id,
                source.drawable_object_id,
                true,
            )
            .unwrap();
        assert!(
            reopened
                .sheet_chart_category_axis_series_names_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_label_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        for axis in [Axis::Category, Axis::Value] {
            assert!(
                editor
                    .sheet_chart_axis_labels_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
        }
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_labels_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                true,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        for axis in [Axis::Category, Axis::Value] {
            editor
                .set_sheet_chart_axis_labels_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    false,
                )
                .unwrap();
            assert!(
                !editor
                    .sheet_chart_axis_labels_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
        }

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                !editor
                    .sheet_chart_axis_labels_visible(sheet_id, duplicate.drawable_object_id, axis)
                    .unwrap()
            );
            editor
                .set_sheet_chart_axis_labels_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    true,
                )
                .unwrap();
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                reopened
                    .sheet_chart_axis_labels_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
            assert!(
                !reopened
                    .sheet_chart_axis_labels_visible(sheet_id, duplicate.drawable_object_id, axis,)
                    .unwrap()
            );
            reopened
                .set_sheet_chart_axis_labels_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    false,
                )
                .unwrap();
            assert!(
                !reopened
                    .sheet_chart_axis_labels_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
        }
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_line_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        for axis in [Axis::Category, Axis::Value] {
            assert!(
                editor
                    .sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
                    .is_visible()
            );
            editor
                .set_sheet_chart_axis_line_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Hidden,
                )
                .unwrap();
            assert!(
                !editor
                    .sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
                    .is_visible()
            );
        }

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                !editor
                    .sheet_chart_axis_line_visible(sheet_id, duplicate.drawable_object_id, axis)
                    .unwrap()
                    .is_visible()
            );
            editor
                .set_sheet_chart_axis_line_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Visible,
                )
                .unwrap();
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                reopened
                    .sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
                    .is_visible()
            );
            assert!(
                !reopened
                    .sheet_chart_axis_line_visible(sheet_id, duplicate.drawable_object_id, axis)
                    .unwrap()
                    .is_visible()
            );
        }
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_major_gridline_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert!(
            !editor
                .sheet_chart_axis_major_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Category,
                )
                .unwrap()
                .is_visible()
        );
        assert!(
            editor
                .sheet_chart_axis_major_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Value,
                )
                .unwrap()
                .is_visible()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                AxisVisibility::Visible,
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                AxisVisibility::Hidden,
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert_eq!(
                editor
                    .sheet_chart_axis_major_gridlines_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible(),
                axis == Axis::Category
            );
        }

        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                AxisVisibility::Hidden,
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                AxisVisibility::Visible,
            )
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert_eq!(
                reopened
                    .sheet_chart_axis_major_gridlines_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible(),
                axis == Axis::Value
            );
            assert_eq!(
                reopened
                    .sheet_chart_axis_major_gridlines_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible(),
                axis == Axis::Category
            );
        }
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_minor_gridline_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        for axis in [Axis::Category, Axis::Value] {
            assert!(
                !editor
                    .sheet_chart_axis_minor_gridlines_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
        }
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_minor_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        for axis in [Axis::Category, Axis::Value] {
            editor
                .set_sheet_chart_axis_minor_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Visible,
                )
                .unwrap();
        }
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                editor
                    .sheet_chart_axis_minor_gridlines_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
            editor
                .set_sheet_chart_axis_minor_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Hidden,
                )
                .unwrap();
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                !reopened
                    .sheet_chart_axis_minor_gridlines_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
            assert!(
                reopened
                    .sheet_chart_axis_minor_gridlines_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
        }
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_minor_tick_mark_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        for axis in [Axis::Category, Axis::Value] {
            assert!(
                editor
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
        }
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_minor_tick_marks_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                AxisVisibility::Visible,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        for axis in [Axis::Category, Axis::Value] {
            editor
                .set_sheet_chart_axis_minor_tick_marks_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Hidden,
                )
                .unwrap();
            assert!(
                !editor
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
        }

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                !editor
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
            editor
                .set_sheet_chart_axis_minor_tick_marks_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Visible,
                )
                .unwrap();
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                reopened
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
            assert!(
                !reopened
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
            reopened
                .set_sheet_chart_axis_minor_tick_marks_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    AxisVisibility::Hidden,
                )
                .unwrap();
            assert!(
                !reopened
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
                    .is_visible()
            );
        }
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_axis_tick_mark_location_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        for axis in [Axis::Category, Axis::Value] {
            assert_eq!(
                editor
                    .sheet_chart_axis_tick_mark_location(sheet_id, source.drawable_object_id, axis,)
                    .unwrap(),
                TickMarkLocation::Centered
            );
        }
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_tick_mark_location(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                TickMarkLocation::Centered,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_axis_tick_mark_location(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                TickMarkLocation::None,
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_tick_mark_location(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                TickMarkLocation::Outside,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Category,
                )
                .unwrap(),
            TickMarkLocation::None
        );
        assert_eq!(
            editor
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Value,
                )
                .unwrap(),
            TickMarkLocation::Outside
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    duplicate.drawable_object_id,
                    Axis::Category,
                )
                .unwrap(),
            TickMarkLocation::None
        );
        assert_eq!(
            editor
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    duplicate.drawable_object_id,
                    Axis::Value,
                )
                .unwrap(),
            TickMarkLocation::Outside
        );

        editor
            .set_sheet_chart_axis_tick_mark_location(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                TickMarkLocation::Inside,
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_tick_mark_location(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                TickMarkLocation::Centered,
            )
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Category,
                )
                .unwrap(),
            TickMarkLocation::Inside
        );
        assert_eq!(
            reopened
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Value,
                )
                .unwrap(),
            TickMarkLocation::Centered
        );
        assert_eq!(
            reopened
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    duplicate.drawable_object_id,
                    Axis::Category,
                )
                .unwrap(),
            TickMarkLocation::None
        );
        assert_eq!(
            reopened
                .sheet_chart_axis_tick_mark_location(
                    sheet_id,
                    duplicate.drawable_object_id,
                    Axis::Value,
                )
                .unwrap(),
            TickMarkLocation::Outside
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_legend_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert!(
            editor
                .sheet_chart_legend_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_legend_visible(sheet_id, source.drawable_object_id, true)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_sheet_chart_legend_visible(sheet_id, source.drawable_object_id, false)
            .unwrap();
        assert!(
            !editor
                .sheet_chart_legend_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(
            !editor
                .sheet_chart_legend_visible(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );

        editor
            .set_sheet_chart_legend_visible(sheet_id, source.drawable_object_id, true)
            .unwrap();
        assert!(
            editor
                .sheet_chart_legend_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            !editor
                .sheet_chart_legend_visible(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .sheet_chart_legend_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            !reopened
                .sheet_chart_legend_visible(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_exact_chart_legend_fill_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let object_id = chart.drawable_object_id;
        let baseline = editor.to_bytes().unwrap();

        assert_eq!(
            editor.sheet_chart_legend_fill(sheet_id, object_id).unwrap(),
            ChartLegendFill::Inherited
        );
        let solid = ChartLegendFill::Fill(ShapeFill::Solid(
            RgbaColor::new(0.2, 0.75, 0.4, 1.0, RgbColorSpace::Srgb).unwrap(),
        ));
        editor
            .set_sheet_chart_legend_fill(sheet_id, object_id, &solid)
            .unwrap();
        assert_eq!(
            editor.sheet_chart_legend_fill(sheet_id, object_id).unwrap(),
            solid
        );
        assert!(
            editor
                .sheet_chart_legend_visible(sheet_id, object_id)
                .unwrap()
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_legend_fill(sheet_id, object_id)
                .unwrap(),
            solid
        );
        reopened
            .set_sheet_chart_legend_fill(sheet_id, object_id, &ChartLegendFill::Inherited)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_spreadsheet_supports_exact_chart_legend_frame_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let object_id = chart.drawable_object_id;
        let baseline = editor.to_bytes().unwrap();

        assert_eq!(
            editor
                .sheet_chart_legend_frame(sheet_id, object_id)
                .unwrap(),
            ChartLegendFrame::Automatic
        );
        let frame =
            ChartLegendFrame::Frame(ChartLegendRect::from_points(43.5, 12.0, 0.0, 0.0).unwrap());
        editor
            .set_sheet_chart_legend_frame(sheet_id, object_id, frame)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_legend_frame(sheet_id, object_id)
                .unwrap(),
            frame
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_legend_frame(sheet_id, object_id)
                .unwrap(),
            frame
        );
        reopened
            .set_sheet_chart_legend_frame(sheet_id, object_id, ChartLegendFrame::Automatic)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn legend_style_copy_on_write_is_shared_across_properties() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let object_id = chart.drawable_object_id;
        let graph = chart_graph(&editor, sheet_id, object_id).unwrap();
        let original = crate::charts::legend_style::legend_style_slot(
            editor.package(),
            &graph.archive_name,
            object_id,
            "Numbers",
        )
        .unwrap();
        let original_id = original.object_id();
        let original_bytes = original
            .read(editor.package(), |data| Ok(data.to_vec()))
            .unwrap();

        let fill = ChartLegendFill::Fill(ShapeFill::Solid(
            RgbaColor::new(0.3, 0.6, 0.9, 1.0, RgbColorSpace::Srgb).unwrap(),
        ));
        editor
            .set_sheet_chart_legend_fill(sheet_id, object_id, &fill)
            .unwrap();
        let private = crate::charts::legend_style::legend_style_slot(
            editor.package(),
            &graph.archive_name,
            object_id,
            "Numbers",
        )
        .unwrap();
        assert_ne!(private.object_id(), original_id);
        assert_eq!(
            original
                .read(editor.package(), |data| Ok(data.to_vec()))
                .unwrap(),
            original_bytes
        );

        let stroke = ChartLegendStroke::Stroke(Stroke::new(
            RgbaColor::new(0.8, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(2.0).unwrap(),
            Pattern::Solid,
        ));
        editor
            .set_sheet_chart_legend_stroke(sheet_id, object_id, stroke)
            .unwrap();
        let composed = crate::charts::legend_style::legend_style_slot(
            editor.package(),
            &graph.archive_name,
            object_id,
            "Numbers",
        )
        .unwrap();
        assert_eq!(composed.object_id(), private.object_id());
        assert_eq!(
            editor.sheet_chart_legend_fill(sheet_id, object_id).unwrap(),
            fill
        );

        editor
            .set_sheet_chart_legend_fill(sheet_id, object_id, &ChartLegendFill::Inherited)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_legend_stroke(sheet_id, object_id)
                .unwrap(),
            stroke
        );
        editor
            .set_sheet_chart_legend_stroke(sheet_id, object_id, ChartLegendStroke::Inherited)
            .unwrap();
        let collapsed = crate::charts::legend_style::legend_style_slot(
            editor.package(),
            &graph.archive_name,
            object_id,
            "Numbers",
        )
        .unwrap();
        assert_eq!(collapsed.object_id(), original_id);
    }

    #[test]
    fn scratch_spreadsheet_supports_exact_chart_legend_typography_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Legend Font Size")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let object_id = chart.drawable_object_id;
        let baseline = editor.to_bytes().unwrap();

        assert_eq!(
            editor.sheet_chart_legend_font(sheet_id, object_id).unwrap(),
            ChartLegendFont::Inherited
        );
        let bold =
            ChartLegendFont::Font(ChartFont::named("AvenirNext-Bold").unwrap().with_bold(true));
        editor
            .set_sheet_chart_legend_font(sheet_id, object_id, &bold)
            .unwrap();

        assert_eq!(
            editor
                .sheet_chart_legend_font_size(sheet_id, object_id)
                .unwrap(),
            ChartLegendFontSize::Inherited
        );
        let eighteen = ChartLegendFontSize::Size(ChartFontSize::from_points(18.0).unwrap());
        editor
            .set_sheet_chart_legend_font_size(sheet_id, object_id, eighteen)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_legend_font_size(sheet_id, object_id)
                .unwrap(),
            eighteen
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_legend_font(sheet_id, object_id)
                .unwrap(),
            bold
        );
        let italic = ChartLegendFont::Font(
            ChartFont::named("AvenirNext-Italic")
                .unwrap()
                .with_italic(true),
        );
        reopened
            .set_sheet_chart_legend_font(sheet_id, object_id, &italic)
            .unwrap();
        let fourteen = ChartLegendFontSize::Size(ChartFontSize::from_points(14.0).unwrap());
        reopened
            .set_sheet_chart_legend_font_size(sheet_id, object_id, fourteen)
            .unwrap();
        assert_eq!(
            reopened
                .sheet_chart_legend_font_size(sheet_id, object_id)
                .unwrap(),
            fourteen
        );
        reopened
            .set_sheet_chart_legend_font(sheet_id, object_id, &ChartLegendFont::Inherited)
            .unwrap();
        assert_eq!(
            reopened
                .sheet_chart_legend_font_size(sheet_id, object_id)
                .unwrap(),
            fourteen
        );
        reopened
            .set_sheet_chart_legend_font_size(sheet_id, object_id, ChartLegendFontSize::Inherited)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_spreadsheet_supports_exact_chart_legend_stroke_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let object_id = chart.drawable_object_id;
        let baseline = editor.to_bytes().unwrap();

        assert_eq!(
            editor
                .sheet_chart_legend_stroke(sheet_id, object_id)
                .unwrap(),
            ChartLegendStroke::Inherited
        );
        let stroke = ChartLegendStroke::Stroke(Stroke::new(
            RgbaColor::new(0.15, 0.7, 0.35, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(3.0).unwrap(),
            Pattern::Solid,
        ));
        editor
            .set_sheet_chart_legend_stroke(sheet_id, object_id, stroke)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_legend_stroke(sheet_id, object_id)
                .unwrap(),
            stroke
        );
        assert!(
            editor
                .sheet_chart_legend_visible(sheet_id, object_id)
                .unwrap()
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        reopened
            .set_sheet_chart_legend_stroke(sheet_id, object_id, ChartLegendStroke::Inherited)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_spreadsheet_supports_exact_chart_legend_shadow_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let object_id = chart.drawable_object_id;
        let baseline = editor.to_bytes().unwrap();

        assert_eq!(
            editor
                .sheet_chart_legend_shadow(sheet_id, object_id)
                .unwrap(),
            ChartLegendShadow::Inherited
        );
        let shadow = ChartLegendShadow::Shadow(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(9).unwrap(),
                Offset::from_points(5.0).unwrap(),
                Opacity::new(0.5).unwrap(),
            ),
            Angle::from_degrees(60.0).unwrap(),
        ));
        editor
            .set_sheet_chart_legend_shadow(sheet_id, object_id, shadow)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_legend_shadow(sheet_id, object_id)
                .unwrap(),
            shadow
        );
        assert!(
            editor
                .sheet_chart_legend_visible(sheet_id, object_id)
                .unwrap()
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        reopened
            .set_sheet_chart_legend_shadow(sheet_id, object_id, ChartLegendShadow::Inherited)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_hidden_data_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert!(
            editor
                .sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id, true)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id, false)
            .unwrap();
        assert!(
            !editor
                .sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        editor
            .set_sheet_chart_legend_visible(sheet_id, source.drawable_object_id, false)
            .unwrap();
        editor
            .set_sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id, true)
            .unwrap();
        assert!(
            !editor
                .sheet_chart_legend_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        editor
            .set_sheet_chart_legend_visible(sheet_id, source.drawable_object_id, true)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id, false)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id, true)
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .sheet_chart_includes_hidden_data(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            !reopened
                .sheet_chart_includes_hidden_data(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_includes_hidden_data(sheet_id, u64::MAX, false)
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn native_chart_caption_graphs_tolerate_partial_uuid_registration() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        editor
            .set_sheet_chart_caption(sheet_id, source.drawable_object_id, "Revenue by region")
            .unwrap();

        let graph = chart_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        let caption =
            caption::sheet_chart_caption_slot(&editor, sheet_id, source.drawable_object_id)
                .unwrap();
        remove_component_object_uuids(&mut editor.package, graph.component_id, &caption.object_ids)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_caption(sheet_id, source.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );
        let duplicate = reopened
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            reopened
                .sheet_chart_caption(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some("Revenue by region".to_owned())
        );

        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_value_axis_scale_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert_eq!(
            editor
                .sheet_chart_value_axis_scale(sheet_id, source.drawable_object_id)
                .unwrap(),
            Scale::Linear
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_value_axis_scale(sheet_id, source.drawable_object_id, Scale::Linear)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_value_axis_scale(
                sheet_id,
                source.drawable_object_id,
                Scale::Logarithmic,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_value_axis_scale(sheet_id, source.drawable_object_id)
                .unwrap(),
            Scale::Logarithmic
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_value_axis_scale(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Scale::Logarithmic
        );
        editor
            .set_sheet_chart_value_axis_scale(sheet_id, source.drawable_object_id, Scale::Linear)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_value_axis_scale(sheet_id, source.drawable_object_id)
                .unwrap(),
            Scale::Linear
        );
        assert_eq!(
            reopened
                .sheet_chart_value_axis_scale(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Scale::Logarithmic
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_border_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();

        assert!(
            !editor
                .sheet_chart_border_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_border_visible(sheet_id, source.drawable_object_id, false)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_border_visible(sheet_id, source.drawable_object_id, true)
            .unwrap();
        assert!(
            editor
                .sheet_chart_border_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(
            editor
                .sheet_chart_border_visible(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );
        editor
            .set_sheet_chart_border_visible(sheet_id, source.drawable_object_id, false)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            !reopened
                .sheet_chart_border_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            reopened
                .sheet_chart_border_visible(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_rounded_corner_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let rounded = ChartRoundedCorners::new(ChartCornerRadius::new(20.0).unwrap(), true);
        let changed = ChartRoundedCorners::new(ChartCornerRadius::new(35.0).unwrap(), false);

        assert_eq!(
            editor
                .sheet_chart_rounded_corners(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartRoundedCorners::NONE
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_rounded_corners(
                sheet_id,
                source.drawable_object_id,
                ChartRoundedCorners::NONE,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_rounded_corners(sheet_id, source.drawable_object_id, rounded)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_rounded_corners(sheet_id, source.drawable_object_id)
                .unwrap(),
            rounded
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_rounded_corners(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            rounded
        );
        editor
            .set_sheet_chart_rounded_corners(sheet_id, source.drawable_object_id, changed)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_rounded_corners(sheet_id, source.drawable_object_id)
                .unwrap(),
            changed
        );
        assert_eq!(
            reopened
                .sheet_chart_rounded_corners(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            rounded
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_gap_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let customized = gap_spacing(25.0, 70.0);
        let changed = gap_spacing(30.0, 60.0);

        assert_eq!(
            editor
                .sheet_chart_gap_spacing(sheet_id, source.drawable_object_id)
                .unwrap(),
            Spacing::DEFAULT
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_gap_spacing(sheet_id, source.drawable_object_id, Spacing::DEFAULT)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_gap_spacing(sheet_id, source.drawable_object_id, customized)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_gap_spacing(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        editor
            .set_sheet_chart_gap_spacing(sheet_id, source.drawable_object_id, changed)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_gap_spacing(sheet_id, source.drawable_object_id)
                .unwrap(),
            changed
        );
        assert_eq!(
            reopened
                .sheet_chart_gap_spacing(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_border_stroke_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let default = Stroke::new(RgbaColor::black(), Width::ONE, Pattern::Solid);
        let customized = chart_stroke(Pattern::MediumDash, 3.0);
        let changed = chart_stroke(Pattern::RoundedDash, 2.0);

        assert_eq!(
            editor
                .sheet_chart_border_stroke(sheet_id, source.drawable_object_id)
                .unwrap(),
            Some(default)
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_border_stroke(sheet_id, source.drawable_object_id, Some(default))
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_border_stroke(sheet_id, source.drawable_object_id, Some(customized))
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_border_stroke(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Some(customized)
        );
        editor
            .set_sheet_chart_border_stroke(sheet_id, source.drawable_object_id, Some(changed))
            .unwrap();
        editor
            .set_sheet_chart_border_stroke(sheet_id, duplicate.drawable_object_id, None)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_border_stroke(sheet_id, source.drawable_object_id)
                .unwrap(),
            Some(changed)
        );
        assert_eq!(
            reopened
                .sheet_chart_border_stroke(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            None
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_background_fill_crud() {
        let image_bytes = fixture("test-data/images/png/lena.png");
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let native_default = editor
            .sheet_chart_background_fill(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(matches!(native_default, ShapeFill::Gradient(_)));
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_background_fill(sheet_id, source.drawable_object_id, &native_default)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let customized = chart_background_fill();
        let image = editor
            .set_sheet_chart_background_image_fill(
                sheet_id,
                source.drawable_object_id,
                "lena.png",
                &image_bytes,
                ShapeImageFillTechnique::ScaleToFit,
                None,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_background_fill(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            ShapeFill::Image(image.clone())
        );
        editor
            .set_sheet_chart_background_fill(sheet_id, source.drawable_object_id, &customized)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_background_fill(sheet_id, source.drawable_object_id)
                .unwrap(),
            customized
        );
        assert_eq!(
            reopened
                .sheet_chart_background_fill(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            ShapeFill::Image(image.clone())
        );
        assert_eq!(
            reopened
                .extract_media(
                    crate::MediaAssetId::try_from(image.data_identifier().unwrap().get()).unwrap(),
                )
                .unwrap(),
            image_bytes
        );
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.media_assets().unwrap().is_empty());
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_inherited_series_fill_crud() {
        let image_bytes = fixture("test-data/images/png/lena.png");
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = editor
            .sheet_chart_series_fills(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(defaults.len(), 2);
        assert!(
            defaults
                .iter()
                .all(|fill| matches!(fill, ShapeFill::Solid(_)))
        );
        assert_ne!(defaults[0], defaults[1]);
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_fills(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let first = Index::from_zero_based(0);
        let second = Index::from_zero_based(1);
        editor
            .set_sheet_chart_series_fill(
                sheet_id,
                source.drawable_object_id,
                first,
                &ShapeFill::None,
            )
            .unwrap();
        let image = editor
            .set_sheet_chart_series_image_fill(
                sheet_id,
                source.drawable_object_id,
                second,
                "lena.png",
                &image_bytes,
                ShapeImageFillTechnique::ScaleToFit,
                None,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_fills(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            vec![ShapeFill::None, ShapeFill::Image(image.clone())]
        );
        assert_eq!(
            editor
                .reset_sheet_chart_series_fill(sheet_id, source.drawable_object_id, first)
                .unwrap(),
            defaults[0]
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_fill(sheet_id, source.drawable_object_id, first)
                .unwrap(),
            defaults[0]
        );
        assert_eq!(
            reopened
                .sheet_chart_series_fill(sheet_id, source.drawable_object_id, second)
                .unwrap(),
            ShapeFill::Image(image.clone())
        );
        assert_eq!(
            reopened
                .extract_media(
                    crate::MediaAssetId::try_from(image.data_identifier().unwrap().get()).unwrap(),
                )
                .unwrap(),
            image_bytes
        );
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(reopened.media_assets().unwrap().len(), 1);
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.media_assets().unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_inherited_series_stroke_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![None, None];
        assert_eq!(
            editor
                .sheet_chart_series_strokes(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_strokes(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let first = Index::from_zero_based(0);
        let second = Index::from_zero_based(1);
        let rounded = chart_series_stroke(ChartSeriesStrokePattern::RoundedDash, 3.5);
        let medium = chart_series_stroke(ChartSeriesStrokePattern::MediumDash, 2.0);
        editor
            .set_sheet_chart_series_strokes(
                sheet_id,
                source.drawable_object_id,
                &[Some(rounded), Some(medium)],
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_strokes(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            vec![Some(rounded), Some(medium)]
        );
        editor
            .set_sheet_chart_series_stroke(sheet_id, source.drawable_object_id, first, None)
            .unwrap();
        assert_eq!(
            editor
                .reset_sheet_chart_series_stroke(sheet_id, source.drawable_object_id, first,)
                .unwrap(),
            None
        );

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_stroke(sheet_id, source.drawable_object_id, first)
                .unwrap(),
            None
        );
        assert_eq!(
            reopened
                .sheet_chart_series_stroke(sheet_id, source.drawable_object_id, second)
                .unwrap(),
            Some(medium)
        );
        assert_eq!(
            reopened
                .sheet_chart_series_strokes(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            vec![Some(rounded), Some(medium)]
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_native_chart_shadow_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let native_default = ChartShadow::native_default();
        assert_eq!(
            editor
                .sheet_chart_shadow(sheet_id, source.drawable_object_id)
                .unwrap(),
            native_default
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_shadow(sheet_id, source.drawable_object_id, native_default)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let customized = chart_shadow();
        editor
            .set_sheet_chart_shadow(sheet_id, source.drawable_object_id, customized)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_shadow(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        editor
            .set_sheet_chart_shadow(sheet_id, source.drawable_object_id, ChartShadow::None)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_shadow(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartShadow::None
        );
        assert_eq!(
            reopened
                .sheet_chart_shadow(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        reopened
            .set_sheet_chart_shadow(sheet_id, duplicate.drawable_object_id, native_default)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_pie_start_angle_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Pie2d, pie_data(), POSITION, SIZE)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_pie_start_angle(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartPieStartAngle::ZERO
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_pie_start_angle(
                sheet_id,
                source.drawable_object_id,
                ChartPieStartAngle::ZERO,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let customized = ChartPieStartAngle::from_degrees(123.0).unwrap();
        editor
            .set_sheet_chart_pie_start_angle(sheet_id, source.drawable_object_id, customized)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_kind(sheet_id, duplicate.drawable_object_id, Kind::Donut2d)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_pie_start_angle(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        editor
            .set_sheet_chart_pie_start_angle(
                sheet_id,
                source.drawable_object_id,
                ChartPieStartAngle::HALF_TURN,
            )
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_pie_start_angle(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartPieStartAngle::HALF_TURN
        );
        assert_eq!(
            reopened
                .sheet_chart_pie_start_angle(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        reopened
            .set_sheet_chart_pie_start_angle(
                sheet_id,
                duplicate.drawable_object_id,
                ChartPieStartAngle::ZERO,
            )
            .unwrap();

        let column = reopened
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let before_rejected_update = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .sheet_chart_pie_start_angle(sheet_id, column.drawable_object_id)
                .is_err()
        );
        assert!(
            reopened
                .set_sheet_chart_pie_start_angle(
                    sheet_id,
                    column.drawable_object_id,
                    ChartPieStartAngle::QUARTER_TURN,
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected_update);

        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, column.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_donut_inner_radius_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Donut2d, pie_data(), POSITION, SIZE)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_donut_inner_radius(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartDonutInnerRadius::DEFAULT
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_donut_inner_radius(
                sheet_id,
                source.drawable_object_id,
                ChartDonutInnerRadius::DEFAULT,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let customized = ChartDonutInnerRadius::from_percent(42.0).unwrap();
        editor
            .set_sheet_chart_donut_inner_radius(sheet_id, source.drawable_object_id, customized)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_donut_inner_radius(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        editor
            .set_sheet_chart_kind(sheet_id, duplicate.drawable_object_id, Kind::Pie2d)
            .unwrap();
        let before_rejected_update = editor.to_bytes().unwrap();
        assert!(
            editor
                .sheet_chart_donut_inner_radius(sheet_id, duplicate.drawable_object_id)
                .is_err()
        );
        assert!(
            editor
                .set_sheet_chart_donut_inner_radius(
                    sheet_id,
                    duplicate.drawable_object_id,
                    ChartDonutInnerRadius::MAXIMUM,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_rejected_update);
        editor
            .set_sheet_chart_kind(sheet_id, duplicate.drawable_object_id, Kind::Donut3d)
            .unwrap();

        editor
            .set_sheet_chart_donut_inner_radius(
                sheet_id,
                source.drawable_object_id,
                ChartDonutInnerRadius::MINIMUM,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_donut_inner_radius(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartDonutInnerRadius::MINIMUM
        );
        assert_eq!(
            reopened
                .sheet_chart_donut_inner_radius(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        reopened
            .set_sheet_chart_donut_inner_radius(
                sheet_id,
                duplicate.drawable_object_id,
                ChartDonutInnerRadius::DEFAULT,
            )
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_pie_wedge_explosion_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Pie2d, pie_data(), POSITION, SIZE)
            .unwrap();
        let zeros = vec![ChartPieWedgeExplosion::ZERO; 3];
        assert_eq!(
            editor
                .sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id)
                .unwrap(),
            zeros
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id, &zeros)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let customized = [
            ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
            ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
            ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
        ];
        editor
            .set_sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id, &customized)
            .unwrap();
        assert_series_non_styles_are_unstyled(
            &editor,
            sheet_id,
            source.drawable_object_id,
            customized.len(),
        );
        assert_eq!(
            editor
                .sheet_chart_pie_wedge_explosion(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(1),
                )
                .unwrap(),
            customized[1]
        );
        editor
            .set_sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id, &zeros)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id, &customized)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_kind(sheet_id, duplicate.drawable_object_id, Kind::Donut2d)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_pie_wedge_explosions(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        let isolated = ChartPieWedgeExplosion::from_percent(55.0).unwrap();
        editor
            .set_sheet_chart_pie_wedge_explosion(
                sheet_id,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
                isolated,
            )
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_pie_wedge_explosion(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(0),
                )
                .unwrap(),
            isolated
        );
        assert_eq!(
            reopened
                .sheet_chart_pie_wedge_explosions(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );

        let before_rejected_updates = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_pie_wedge_explosions(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..2],
                )
                .is_err()
        );
        assert!(
            reopened
                .set_sheet_chart_pie_wedge_explosion(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(3),
                    isolated,
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected_updates);

        let column = reopened
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let before_wrong_kind = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .sheet_chart_pie_wedge_explosions(sheet_id, column.drawable_object_id)
                .is_err()
        );
        assert!(
            reopened
                .set_sheet_chart_pie_wedge_explosions(
                    sheet_id,
                    column.drawable_object_id,
                    &customized,
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_wrong_kind);

        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, column.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_pie_label_visibility_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Pie2d, pie_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![LabelVisibility::DEFAULT; 3];
        let customized = [
            LabelVisibility::DATA_POINT_NAMES_ONLY,
            LabelVisibility::ALL,
            LabelVisibility::HIDDEN,
        ];
        assert_eq!(
            editor
                .sheet_chart_pie_label_visibilities(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_pie_label_visibilities(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_pie_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        assert_series_non_styles_are_styled(
            &editor,
            sheet_id,
            source.drawable_object_id,
            customized.len(),
        );
        assert_eq!(
            editor
                .sheet_chart_pie_label_visibility(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(1),
                )
                .unwrap(),
            LabelVisibility::ALL
        );
        let explosions = [
            ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
            ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
            ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
        ];
        editor
            .set_sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id, &explosions)
            .unwrap();
        editor
            .set_sheet_chart_pie_label_visibilities(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_pie_wedge_explosions(sheet_id, source.drawable_object_id)
                .unwrap(),
            explosions
        );
        editor
            .set_sheet_chart_pie_wedge_explosions(
                sheet_id,
                source.drawable_object_id,
                &[ChartPieWedgeExplosion::ZERO; 3],
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_pie_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_kind(sheet_id, duplicate.drawable_object_id, Kind::Donut2d)
            .unwrap();
        editor
            .set_sheet_chart_pie_label_visibility(
                sheet_id,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
                LabelVisibility::VALUES_ONLY,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_pie_label_visibilities(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_pie_label_visibilities(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..2],
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_pie_label_distance_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Pie2d, pie_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![ChartPieLabelDistance::DEFAULT; 3];
        let customized = [
            ChartPieLabelDistance::MINIMUM,
            ChartPieLabelDistance::from_percent(100.0).unwrap(),
            ChartPieLabelDistance::MAXIMUM,
        ];
        assert_eq!(
            editor
                .sheet_chart_pie_label_distances(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_pie_label_distances(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        let leader_line_defaults = [LeaderLineVisibility::Visible; 3];
        let leader_line_customized = [
            LeaderLineVisibility::Hidden,
            LeaderLineVisibility::Visible,
            LeaderLineVisibility::Hidden,
        ];
        assert_eq!(
            editor
                .sheet_chart_pie_leader_line_visibilities(sheet_id, source.drawable_object_id,)
                .unwrap(),
            leader_line_defaults
        );
        editor
            .set_sheet_chart_pie_leader_line_visibilities(
                sheet_id,
                source.drawable_object_id,
                &leader_line_defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_sheet_chart_pie_leader_line_visibilities(
                sheet_id,
                source.drawable_object_id,
                &leader_line_customized,
            )
            .unwrap();
        assert_series_non_styles_are_unstyled(
            &editor,
            sheet_id,
            source.drawable_object_id,
            leader_line_customized
                .iter()
                .filter(|visibility| **visibility == LeaderLineVisibility::Hidden)
                .count(),
        );
        assert_eq!(
            editor
                .sheet_chart_pie_leader_line_visibility(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(0),
                )
                .unwrap(),
            LeaderLineVisibility::Hidden
        );
        editor
            .set_sheet_chart_pie_leader_line_visibilities(
                sheet_id,
                source.drawable_object_id,
                &leader_line_defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_pie_label_distances(sheet_id, source.drawable_object_id, &customized)
            .unwrap();
        assert_series_non_styles_are_unstyled(
            &editor,
            sheet_id,
            source.drawable_object_id,
            customized.len(),
        );
        assert_eq!(
            editor
                .sheet_chart_pie_label_distance(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(1),
                )
                .unwrap(),
            customized[1]
        );
        let visibilities = [
            LabelVisibility::DATA_POINT_NAMES_ONLY,
            LabelVisibility::ALL,
            LabelVisibility::VALUES_ONLY,
        ];
        editor
            .set_sheet_chart_pie_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &visibilities,
            )
            .unwrap();
        editor
            .set_sheet_chart_pie_label_distances(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_pie_label_visibilities(sheet_id, source.drawable_object_id)
                .unwrap(),
            visibilities
        );
        editor
            .set_sheet_chart_pie_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &[LabelVisibility::DEFAULT; 3],
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_pie_label_distances(sheet_id, source.drawable_object_id, &customized)
            .unwrap();
        editor
            .set_sheet_chart_pie_leader_line_visibilities(
                sheet_id,
                source.drawable_object_id,
                &leader_line_customized,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_kind(sheet_id, duplicate.drawable_object_id, Kind::Donut2d)
            .unwrap();
        editor
            .set_sheet_chart_pie_label_distance(
                sheet_id,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
                ChartPieLabelDistance::DEFAULT,
            )
            .unwrap();
        editor
            .set_sheet_chart_pie_leader_line_visibility(
                sheet_id,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
                LeaderLineVisibility::Visible,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_pie_label_distances(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        assert_eq!(
            reopened
                .sheet_chart_pie_leader_line_visibilities(sheet_id, duplicate.drawable_object_id,)
                .unwrap(),
            leader_line_customized
        );
        assert_eq!(
            reopened
                .sheet_chart_pie_leader_line_visibilities(sheet_id, source.drawable_object_id,)
                .unwrap(),
            [
                LeaderLineVisibility::Visible,
                LeaderLineVisibility::Visible,
                LeaderLineVisibility::Hidden,
            ]
        );
        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_pie_label_distances(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..2],
                )
                .is_err()
        );
        assert!(
            reopened
                .set_sheet_chart_pie_leader_line_visibilities(
                    sheet_id,
                    source.drawable_object_id,
                    &leader_line_customized[..2],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_pie_leader_line_visibility(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(3),
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_pie_label_distance(
                    sheet_id,
                    source.drawable_object_id,
                    ChartPieWedgeIndex::from_zero_based(3),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_value_label_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = [Visibility::Hidden; 2];
        let customized = [Visibility::Visible, Visibility::Hidden];

        assert_eq!(
            editor
                .sheet_chart_series_value_label_visibilities(sheet_id, source.drawable_object_id,)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_value_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_series_value_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_value_label_visibility(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(0),
                )
                .unwrap(),
            Visibility::Visible
        );
        editor
            .set_sheet_chart_series_value_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_series_value_label_visibilities(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_series_value_label_visibility(
                sheet_id,
                source.drawable_object_id,
                Index::from_zero_based(0),
                Visibility::Hidden,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_visibilities(sheet_id, source.drawable_object_id,)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_visibilities(
                    sheet_id,
                    duplicate.drawable_object_id,
                )
                .unwrap(),
            customized
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_value_label_visibilities(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_value_label_visibility(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_value_label_location_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = [ChartSeriesValueLabelLocation::Top; 2];
        let customized = [
            ChartSeriesValueLabelLocation::Outside,
            ChartSeriesValueLabelLocation::Top,
        ];

        assert_eq!(
            editor
                .sheet_chart_series_value_label_locations(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_value_label_locations(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_series_value_label_locations(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_value_label_location(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(0),
                )
                .unwrap(),
            ChartSeriesValueLabelLocation::Outside
        );
        editor
            .set_sheet_chart_series_value_label_locations(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_series_value_label_locations(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_series_value_label_location(
                sheet_id,
                source.drawable_object_id,
                Index::from_zero_based(0),
                ChartSeriesValueLabelLocation::Top,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_locations(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_locations(sheet_id, duplicate.drawable_object_id,)
                .unwrap(),
            customized
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_value_label_locations(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_value_label_location(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_value_label_affix_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![LabelAffixes::default(); 2];
        let customized = vec![
            LabelAffixes::new("$", " USD").unwrap(),
            LabelAffixes::new("€", " net").unwrap(),
        ];

        assert_eq!(
            editor
                .sheet_chart_series_value_label_affixes(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_value_label_affixes(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_series_value_label_affixes(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_value_label_affix(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(1),
                )
                .unwrap()
                .prefix(),
            "€"
        );
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for series in 0..2 {
            editor
                .set_sheet_chart_series_value_label_affix(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(series),
                    LabelAffixes::default(),
                )
                .unwrap();
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_affixes(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_affixes(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_value_label_affixes(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_value_label_affix(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_value_label_number_format_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT; 2];
        let fixed_two = NumberFormat::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            false,
        );
        let customized = vec![fixed_two, NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT];

        assert_eq!(
            editor
                .sheet_chart_series_value_label_number_formats(sheet_id, source.drawable_object_id,)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_value_label_number_formats(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_sheet_chart_series_value_label_number_formats(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_value_label_number_format(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(0),
                )
                .unwrap(),
            fixed_two
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_series_value_label_number_format(
                sheet_id,
                source.drawable_object_id,
                Index::from_zero_based(0),
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_number_formats(sheet_id, source.drawable_object_id,)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_number_formats(
                    sheet_id,
                    duplicate.drawable_object_id,
                )
                .unwrap(),
            customized
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_value_label_number_formats(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_value_label_number_format(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_value_label_auto_fit_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![ChartSeriesValueLabelAutoFit::Enabled; 2];
        let customized = vec![
            ChartSeriesValueLabelAutoFit::Disabled,
            ChartSeriesValueLabelAutoFit::Enabled,
        ];

        assert_eq!(
            editor
                .sheet_chart_series_value_label_auto_fits(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_value_label_auto_fits(
                sheet_id,
                source.drawable_object_id,
                &defaults,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_sheet_chart_series_value_label_auto_fits(
                sheet_id,
                source.drawable_object_id,
                &customized,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_value_label_auto_fit(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(0),
                )
                .unwrap(),
            ChartSeriesValueLabelAutoFit::Disabled
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_series_value_label_auto_fit(
                sheet_id,
                source.drawable_object_id,
                Index::from_zero_based(0),
                ChartSeriesValueLabelAutoFit::Enabled,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_auto_fits(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_value_label_auto_fits(sheet_id, duplicate.drawable_object_id,)
                .unwrap(),
            customized
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_value_label_auto_fits(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_value_label_auto_fit(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_trendline_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![ChartSeriesTrendline::none(); 2];
        let customized = vec![
            ChartSeriesTrendline::linear()
                .with_legend_name("Revenue fit")
                .unwrap()
                .with_equation_visibility(true)
                .unwrap()
                .with_r_squared_visibility(true)
                .unwrap(),
            ChartSeriesTrendline::moving_average(
                ChartSeriesTrendlineMovingAveragePeriod::new(3).unwrap(),
            )
            .with_legend_visibility(true)
            .unwrap(),
        ];

        assert_eq!(
            editor
                .sheet_chart_series_trendlines(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_trendlines(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_sheet_chart_series_trendlines(sheet_id, source.drawable_object_id, &customized)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_trendline(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(1),
                )
                .unwrap(),
            customized[1]
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for series in 0..2 {
            editor
                .set_sheet_chart_series_trendline(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(series),
                    ChartSeriesTrendline::none(),
                )
                .unwrap();
        }
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_trendlines(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_trendlines(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_trendlines(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_trendline(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert!(ChartSeriesTrendline::unsupported(1).is_err());
        assert!(ChartSeriesTrendlinePolynomialOrder::new(7).is_err());
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_native_series_error_bar_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_chart(sheet_id, Kind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let defaults = vec![Series::None; 2];
        let customized = vec![
            Series::FixedValue {
                direction: ErrorBarDirection::PositiveAndNegative,
                value: ErrorBarFixedValue::new(12.5).unwrap(),
            },
            Series::CustomValues {
                direction: ErrorBarDirection::PositiveOnly,
                values: ErrorBarCustomValues::new([1.0, 2.0, 3.0], []).unwrap(),
            },
        ];
        let default_auto_fits = vec![ChartSeriesErrorBarAutoFit::Enabled; 2];
        let customized_auto_fits = vec![
            ChartSeriesErrorBarAutoFit::Disabled,
            ChartSeriesErrorBarAutoFit::Enabled,
        ];

        assert_eq!(
            editor
                .sheet_chart_series_error_bars(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        assert_eq!(
            editor
                .sheet_chart_series_error_bar_auto_fits(sheet_id, source.drawable_object_id,)
                .unwrap(),
            default_auto_fits
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_series_error_bars(sheet_id, source.drawable_object_id, &defaults)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_sheet_chart_series_error_bars(sheet_id, source.drawable_object_id, &customized)
            .unwrap();
        editor
            .set_sheet_chart_series_error_bar_auto_fits(
                sheet_id,
                source.drawable_object_id,
                &customized_auto_fits,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_series_error_bar(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(1),
                )
                .unwrap(),
            customized[1]
        );
        assert_eq!(
            editor
                .sheet_chart_series_error_bar_auto_fit(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(0),
                )
                .unwrap(),
            ChartSeriesErrorBarAutoFit::Disabled
        );

        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        for series in 0..2 {
            editor
                .set_sheet_chart_series_error_bar(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(series),
                    Series::None,
                )
                .unwrap();
        }
        editor
            .set_sheet_chart_series_error_bar_auto_fits(
                sheet_id,
                source.drawable_object_id,
                &default_auto_fits,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_series_error_bars(sheet_id, source.drawable_object_id)
                .unwrap(),
            defaults
        );
        assert_eq!(
            reopened
                .sheet_chart_series_error_bars(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            customized
        );
        assert_eq!(
            reopened
                .sheet_chart_series_error_bar_auto_fits(sheet_id, source.drawable_object_id,)
                .unwrap(),
            default_auto_fits
        );
        assert_eq!(
            reopened
                .sheet_chart_series_error_bar_auto_fits(sheet_id, duplicate.drawable_object_id,)
                .unwrap(),
            customized_auto_fits
        );

        let before_rejected = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_sheet_chart_series_error_bars(
                    sheet_id,
                    source.drawable_object_id,
                    &customized[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .set_sheet_chart_series_error_bar_auto_fits(
                    sheet_id,
                    source.drawable_object_id,
                    &customized_auto_fits[..1],
                )
                .is_err()
        );
        assert!(
            reopened
                .sheet_chart_series_error_bar(
                    sheet_id,
                    source.drawable_object_id,
                    Index::from_zero_based(2),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
        reopened
            .remove_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        reopened
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(reopened.sheet_charts(sheet_id).unwrap().is_empty());
    }
}
