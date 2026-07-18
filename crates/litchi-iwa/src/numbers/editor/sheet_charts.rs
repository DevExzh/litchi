//! Standalone, inline-data chart CRUD for Numbers sheets.

mod graph;
mod theme;

use graph::chart_graph;
use theme::{chart_theme_context, patch_theme_chart_preset};

use std::collections::HashSet;

use super::*;
use crate::IWorkThemeArchive;
use crate::charts::source::{
    AXIS_NON_STYLE_MESSAGE_TYPE, AXIS_STYLE_MESSAGE_TYPE, CHART_MEDIATOR_MESSAGE_TYPE,
    CHART_MESSAGE_TYPE, CHART_NON_STYLE_MESSAGE_TYPE, CHART_PRESET_MESSAGE_TYPE,
    CHART_STYLE_MESSAGE_TYPE, ChartApplicationProfile, LEGEND_NON_STYLE_MESSAGE_TYPE,
    LEGEND_STYLE_MESSAGE_TYPE, SERIES_STYLE_MESSAGE_TYPE, STANDIN_MESSAGE_TYPE,
    SourceChartObjectIds, chart_data, chart_geometry, chart_grid, drawable_geometry,
    geometry_archive, reference, require_creatable_kind, source_chart_objects,
};
use crate::charts::{ChartData, ChartKind, ChartSeriesDirection, IWorkChartArchive};
use crate::protobuf::tsch;
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};

const NUMBERS_THEME_MESSAGE_TYPE: u32 = 12_009;

/// One chart drawable owned by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetChartInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub kind: ChartKind,
    pub direction: ChartSeriesDirection,
    pub data: ChartData,
    pub geometry: DrawableGeometry,
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
        kind: ChartKind,
        data: ChartData,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<NumbersSheetChartInfo> {
        require_creatable_kind(kind)?;
        let geometry = chart_geometry("Numbers", position, size)?;
        let (archive_name, _, _) = numbers_sheet(self.package(), sheet_id)?;
        let component_id = component_identifier_for_entry(self.package(), &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet component {archive_name} is not registered"
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
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &archive_name,
            sheet_id,
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
            add_component_external_reference(
                &mut staged,
                component_id,
                theme.component_id,
                theme.paragraph_style_id,
            )?;
        }
        add_component_object_uuids(&mut staged, component_id, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = chart_graph(&verified, sheet_id, ids.drawable)?;
        if created.info.kind != kind
            || created.info.direction != ChartSeriesDirection::Rows
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
        kind: ChartKind,
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
                .chart_type = Some(kind.into_raw());
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
        direction: ChartSeriesDirection,
    ) -> Result<()> {
        if matches!(direction, ChartSeriesDirection::Unsupported(_)) {
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
                .series_direction = Some(direction.into_raw());
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

    /// Remove a standalone chart and its private caption, title, and mediator objects.
    pub fn remove_sheet_chart(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetChart> {
        let source = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        if let Some(preset_id) = source.private_preset_id {
            let theme = chart_theme_context(&staged)?;
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
                    Error::InvalidFormat(format!("Numbers chart object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers chart object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
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
        staged.update_archive(&source.archive_name, |archive| {
            let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers chart {drawable_object_id} is missing"))
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
                    "Numbers chart {drawable_object_id} must contain exactly one chart payload"
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
        })?;
        *self = Self::from_bytes(&staged.to_bytes()?)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

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
            .add_sheet_chart(sheet_id, ChartKind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, ChartKind::Column2d);
        assert_eq!(created.direction, ChartSeriesDirection::Rows);
        assert_eq!(created.data, sample_data());

        let replacement = ChartData::new(
            vec!["Revenue".to_owned()],
            vec!["2026".to_owned(), "2027".to_owned(), "2028".to_owned()],
            vec![vec![Some(30.0), Some(45.0), None]],
        )
        .unwrap();
        editor
            .set_sheet_chart_kind(sheet_id, created.drawable_object_id, ChartKind::Bar2d)
            .unwrap();
        editor
            .set_sheet_chart_data(sheet_id, created.drawable_object_id, replacement.clone())
            .unwrap();
        editor
            .set_sheet_chart_direction(
                sheet_id,
                created.drawable_object_id,
                ChartSeriesDirection::Columns,
            )
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
        assert_eq!(chart.kind, ChartKind::Bar2d);
        assert_eq!(chart.direction, ChartSeriesDirection::Columns);
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
                    ChartKind::Undefined,
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
                    ChartKind::Column2d,
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
            .add_sheet_chart(sheet_id, ChartKind::Column2d, sample_data(), POSITION, SIZE)
            .unwrap();
        let second = editor
            .add_sheet_chart(
                sheet_id,
                ChartKind::Column2d,
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
}
