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
    STANDIN_MESSAGE_TYPE, SourceChartObjectIds, chart_data, chart_geometry, chart_grid,
    drawable_geometry, geometry_archive, local_chart_style_ids, reference, register_chart_styles,
    require_creatable_kind, single_message_index, source_chart_objects, unregister_chart_styles,
    validate_chart_styles_registered,
};
use crate::charts::{
    ChartArrangement, ChartData, Kind, Direction, DirectionKind, IWorkChartArchive,
};
use crate::data_reference_registry::{
    clone_component_data_references, remove_component_data_references_for_objects,
};
use crate::protobuf::tsch;
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableSize, offset_drawable_geometry,
    remove_orphaned_image_asset,
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
                charts.push(chart_graph(self, sheet_id, reference.identifier)?.info);
            }
        }
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
        let created = chart_graph(&verified, sheet_id, ids.drawable)?;
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
        let source = chart_graph(self, sheet_id, source_drawable_object_id)?;
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
        let created = chart_graph(&verified, sheet_id, new_drawable_id)?;
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
    pub fn remove_sheet_chart(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetChart> {
        let source = chart_graph(self, sheet_id, drawable_object_id)?;
        let style_ids =
            local_chart_style_ids(self.package(), &source.archive_name, &source.object_ids)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(crate::comments::DrawableObjectId::from_object_id(
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
    use crate::charts::object_container::is_object_container_archive;
    use crate::charts::source::SERIES_NON_STYLE_MESSAGE_TYPE;
    use crate::charts::unique_chart_object_archive_name;
    use crate::charts::{
        Axis, Bound, Bounds, ChartCornerRadius, ChartDonutInnerRadius, ChartErrorBarCustomValues,
        ChartErrorBarDirection, ChartErrorBarFixedValue, ChartFont, ChartFontSize,
        ChartGapPercentage, ChartGapSpacing, ChartLegendFill, ChartLegendFont, ChartLegendFontSize,
        ChartLegendFrame, ChartLegendRect, ChartLegendShadow, ChartLegendStroke,
        ChartPieLabelDistance, ChartPieStartAngle, ChartPieWedgeExplosion, ChartPieWedgeIndex,
        ChartRoundedCorners, ChartSeriesErrorBarAutoFit, ChartSeriesErrorBars, ChartSeriesIndex,
        ChartSeriesStroke, ChartSeriesStrokePattern, ChartSeriesTrendline,
        ChartSeriesTrendlineMovingAveragePeriod, ChartSeriesTrendlinePolynomialOrder,
        ChartSeriesValueLabelAutoFit, ChartSeriesValueLabelLocation,
        ChartSeriesValueLabelVisibility, ChartShadow, DecimalPlaces, LabelAffixes, LabelVisibility,
        LeaderLineVisibility, MajorStepCount, MinorStepCount, NegativeStyle, NumberFormat, Scale,
        Steps, TickMarkLocation,
    };
    use crate::numbers::NumbersDocumentBuilder;
    use crate::package_metadata::{component_identifier_for_entry, component_uuid_identifiers};
    use crate::protobuf::tsch;
    use crate::shapes::{
        RgbColorSpace, RgbaColor, ShapeDropShadow, ShapeFill, ShapeImageFillTechnique,
        ShapeShadowAngle, ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowOffset,
        ShapeShadowOpacity, ShapeStroke, StrokePattern, StrokeWidth,
    };

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

    fn gap_spacing(between_items: f32, between_sets: f32) -> ChartGapSpacing {
        ChartGapSpacing::new(
            ChartGapPercentage::new(between_items).unwrap(),
            ChartGapPercentage::new(between_sets).unwrap(),
        )
    }

    fn chart_stroke(pattern: StrokePattern, width: f32) -> ShapeStroke {
        ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(width).unwrap(),
            pattern,
        )
    }

    fn chart_series_stroke(pattern: ChartSeriesStrokePattern, width: f32) -> ChartSeriesStroke {
        ChartSeriesStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(width).unwrap(),
            pattern,
        )
    }

    fn chart_background_fill() -> ShapeFill {
        ShapeFill::Solid(RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap())
    }

    fn chart_shadow() -> ChartShadow {
        ChartShadow::Grouped(ShapeDropShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
                ShapeShadowBlurRadius::from_points(15).unwrap(),
                ShapeShadowOffset::from_points(8.0).unwrap(),
                ShapeShadowOpacity::new(0.6).unwrap(),
            ),
            ShapeShadowAngle::from_degrees(60.0).unwrap(),
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
                .add_sheet_chart(
                    sheet_id,
                    Kind::Undefined,
                    sample_data(),
                    POSITION,
                    SIZE
                )
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
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                true,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                false,
            )
            .unwrap();
        assert!(
            !editor
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
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
        );

        editor
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                true,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
        );
        assert!(
            !reopened
                .sheet_chart_value_axis_minimum_label_visible(
                    sheet_id,
                    duplicate.drawable_object_id
                )
                .unwrap()
        );
        reopened
            .set_sheet_chart_value_axis_minimum_label_visible(
                sheet_id,
                source.drawable_object_id,
                false,
            )
            .unwrap();
        assert!(
            !reopened
                .sheet_chart_value_axis_minimum_label_visible(sheet_id, source.drawable_object_id)
                .unwrap()
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
            );
            editor
                .set_sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis, false)
                .unwrap();
            assert!(
                !editor
                    .sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
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
            );
            editor
                .set_sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis, true)
                .unwrap();
        }

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        for axis in [Axis::Category, Axis::Value] {
            assert!(
                reopened
                    .sheet_chart_axis_line_visible(sheet_id, source.drawable_object_id, axis)
                    .unwrap()
            );
            assert!(
                !reopened
                    .sheet_chart_axis_line_visible(sheet_id, duplicate.drawable_object_id, axis)
                    .unwrap()
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
        );
        assert!(
            editor
                .sheet_chart_axis_major_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    Axis::Value,
                )
                .unwrap()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                false,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                true,
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                false,
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
                    .unwrap(),
                axis == Axis::Category
            );
        }

        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                false,
            )
            .unwrap();
        editor
            .set_sheet_chart_axis_major_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Value,
                true,
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
                    .unwrap(),
                axis == Axis::Value
            );
            assert_eq!(
                reopened
                    .sheet_chart_axis_major_gridlines_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap(),
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
            );
        }
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_minor_gridlines_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                false,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        for axis in [Axis::Category, Axis::Value] {
            editor
                .set_sheet_chart_axis_minor_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    true,
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
            );
            editor
                .set_sheet_chart_axis_minor_gridlines_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    false,
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
            );
            assert!(
                reopened
                    .sheet_chart_axis_minor_gridlines_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
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
            );
        }
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_minor_tick_marks_visible(
                sheet_id,
                source.drawable_object_id,
                Axis::Category,
                true,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        for axis in [Axis::Category, Axis::Value] {
            editor
                .set_sheet_chart_axis_minor_tick_marks_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    false,
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
            );
            editor
                .set_sheet_chart_axis_minor_tick_marks_visible(
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
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        source.drawable_object_id,
                        axis,
                    )
                    .unwrap()
            );
            assert!(
                !reopened
                    .sheet_chart_axis_minor_tick_marks_visible(
                        sheet_id,
                        duplicate.drawable_object_id,
                        axis,
                    )
                    .unwrap()
            );
            reopened
                .set_sheet_chart_axis_minor_tick_marks_visible(
                    sheet_id,
                    source.drawable_object_id,
                    axis,
                    false,
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

        let stroke = ChartLegendStroke::Stroke(ShapeStroke::new(
            RgbaColor::new(0.8, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(2.0).unwrap(),
            StrokePattern::Solid,
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
        let stroke = ChartLegendStroke::Stroke(ShapeStroke::new(
            RgbaColor::new(0.15, 0.7, 0.35, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(3.0).unwrap(),
            StrokePattern::Solid,
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
        let shadow = ChartLegendShadow::Shadow(ShapeDropShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::black(),
                ShapeShadowBlurRadius::from_points(9).unwrap(),
                ShapeShadowOffset::from_points(5.0).unwrap(),
                ShapeShadowOpacity::new(0.5).unwrap(),
            ),
            ShapeShadowAngle::from_degrees(60.0).unwrap(),
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
            ChartGapSpacing::NATIVE_DEFAULT
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_gap_spacing(
                sheet_id,
                source.drawable_object_id,
                ChartGapSpacing::NATIVE_DEFAULT,
            )
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
        let default = ShapeStroke::new(RgbaColor::black(), StrokeWidth::ONE, StrokePattern::Solid);
        let customized = chart_stroke(StrokePattern::MediumDash, 3.0);
        let changed = chart_stroke(StrokePattern::RoundedDash, 2.0);

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
                .extract_media(image.data_identifier().unwrap().get())
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

        let first = ChartSeriesIndex::from_zero_based(0);
        let second = ChartSeriesIndex::from_zero_based(1);
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
                .extract_media(image.data_identifier().unwrap().get())
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

        let first = ChartSeriesIndex::from_zero_based(0);
        let second = ChartSeriesIndex::from_zero_based(1);
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
        let defaults = [ChartSeriesValueLabelVisibility::Hidden; 2];
        let customized = [
            ChartSeriesValueLabelVisibility::Visible,
            ChartSeriesValueLabelVisibility::Hidden,
        ];

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
                    ChartSeriesIndex::from_zero_based(0),
                )
                .unwrap(),
            ChartSeriesValueLabelVisibility::Visible
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
                ChartSeriesIndex::from_zero_based(0),
                ChartSeriesValueLabelVisibility::Hidden,
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
                    ChartSeriesIndex::from_zero_based(2),
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
                    ChartSeriesIndex::from_zero_based(0),
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
                ChartSeriesIndex::from_zero_based(0),
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
                    ChartSeriesIndex::from_zero_based(2),
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
                    ChartSeriesIndex::from_zero_based(1),
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
                    ChartSeriesIndex::from_zero_based(series),
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
                    ChartSeriesIndex::from_zero_based(2),
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
                    ChartSeriesIndex::from_zero_based(0),
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
                ChartSeriesIndex::from_zero_based(0),
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
                    ChartSeriesIndex::from_zero_based(2),
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
                    ChartSeriesIndex::from_zero_based(0),
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
                ChartSeriesIndex::from_zero_based(0),
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
                    ChartSeriesIndex::from_zero_based(2),
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
                    ChartSeriesIndex::from_zero_based(1),
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
                    ChartSeriesIndex::from_zero_based(series),
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
                    ChartSeriesIndex::from_zero_based(2),
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
        let defaults = vec![ChartSeriesErrorBars::None; 2];
        let customized = vec![
            ChartSeriesErrorBars::FixedValue {
                direction: ChartErrorBarDirection::PositiveAndNegative,
                value: ChartErrorBarFixedValue::new(12.5).unwrap(),
            },
            ChartSeriesErrorBars::CustomValues {
                direction: ChartErrorBarDirection::PositiveOnly,
                values: ChartErrorBarCustomValues::new([1.0, 2.0, 3.0], []).unwrap(),
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
                    ChartSeriesIndex::from_zero_based(1),
                )
                .unwrap(),
            customized[1]
        );
        assert_eq!(
            editor
                .sheet_chart_series_error_bar_auto_fit(
                    sheet_id,
                    source.drawable_object_id,
                    ChartSeriesIndex::from_zero_based(0),
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
                    ChartSeriesIndex::from_zero_based(series),
                    ChartSeriesErrorBars::None,
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
                    ChartSeriesIndex::from_zero_based(2),
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
