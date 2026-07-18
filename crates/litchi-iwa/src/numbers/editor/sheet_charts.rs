//! Standalone, inline-data chart CRUD for Numbers sheets.

use std::collections::HashSet;

use super::*;
use crate::IWorkThemeArchive;
use crate::charts::{ChartData, ChartKind, IWorkChartArchive};
use crate::protobuf::{tsch, tsk, tss};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};
use crate::wire::append_length_delimited_field;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_MEDIATOR_MESSAGE_TYPE: u32 = 12_006;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_PRESET_MESSAGE_TYPE: u32 = 5_020;
const CHART_STYLE_MESSAGE_TYPE: u32 = 5_022;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const LEGEND_STYLE_MESSAGE_TYPE: u32 = 5_024;
const LEGEND_NON_STYLE_MESSAGE_TYPE: u32 = 5_025;
const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
const NUMBERS_THEME_MESSAGE_TYPE: u32 = 12_009;
const CURRENT_STYLE_EXTENSION_FIELD: u32 = 10_000;
const CHART_SCENE_DEPTH_EXTENSION_FIELD: u32 = 10_002;
const CHART_APPEARANCE_PRESERVED_EXTENSION_FIELD: u32 = 10_023;
const CHART_PROPORTIONAL_CALLOUTS_EXTENSION_FIELD: u32 = 10_024;
const CHART_ROUNDED_CORNERS_EXTENSION_FIELD: u32 = 10_026;
const CHART_VALUE_LABEL_SPACING_EXTENSION_FIELD: u32 = 10_027;
const CHART_ERROR_BAR_SPACING_EXTENSION_FIELD: u32 = 10_028;
const CHART_STACKED_SUMMARY_LABELS_EXTENSION_FIELD: u32 = 10_029;
const CHART_CACHED_FORMATTERS_EXTENSION_FIELD: u32 = 10_030;
const STANDARD_MESSAGE_VERSION: &[u32] = &[1, 0, 5];
const STANDIN_MESSAGE_VERSION: &[u32] = &[10, 1, 0];
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const MEDIATOR_LOCAL_SERIES_SENTINEL: u32 = u32::MAX;
const MEDIATOR_REMOTE_SERIES_INDEX: u32 = 0;
const MEDIATOR_FORMULA_DIRECTION: i32 = 0;
const MEDIATOR_FORMULA_SCHEME: i32 = 0;
const DEFAULT_CHART_DATASET_INDEX: u32 = 0;
const DEFAULT_CHART_NUMBER_FORMAT_TYPE: u32 = 256;
const AUTOMATIC_CHART_DECIMAL_PLACES: u32 = 253;
const DEFAULT_CHART_NEGATIVE_STYLE: u32 = 0;
const CHART_PARAGRAPH_STYLE_COUNT: usize = 30;

/// Whether chart series are stored in rows or columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartSeriesDirection {
    Rows,
    Columns,
    Unsupported(i32),
}

impl ChartSeriesDirection {
    const fn from_raw(value: i32) -> Self {
        match value {
            x if x == tsch::SeriesDirection::ByRow as i32 => Self::Rows,
            x if x == tsch::SeriesDirection::ByColumn as i32 => Self::Columns,
            value => Self::Unsupported(value),
        }
    }

    const fn into_raw(self) -> i32 {
        match self {
            Self::Rows => tsch::SeriesDirection::ByRow as i32,
            Self::Columns => tsch::SeriesDirection::ByColumn as i32,
            Self::Unsupported(value) => value,
        }
    }
}

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

#[derive(Debug, Clone, Copy)]
struct ChartObjectIds {
    drawable: u64,
    caption: u64,
    title: u64,
    mediator: u64,
    preset: u64,
    chart_style: u64,
    chart_non_style: u64,
    legend_style: u64,
    legend_non_style: u64,
    value_axis_styles: [u64; 2],
    value_axis_non_styles: [u64; 2],
    category_axis_style: u64,
    category_axis_non_style: u64,
    series_styles: [u64; 6],
}

impl ChartObjectIds {
    const COUNT: usize = 21;

    fn allocate(package: &IWorkPackage) -> Result<Self> {
        let mut next = next_object_identifier(package)?;
        Ok(Self {
            drawable: take_identifier(&mut next)?,
            caption: take_identifier(&mut next)?,
            title: take_identifier(&mut next)?,
            mediator: take_identifier(&mut next)?,
            preset: take_identifier(&mut next)?,
            chart_style: take_identifier(&mut next)?,
            chart_non_style: take_identifier(&mut next)?,
            legend_style: take_identifier(&mut next)?,
            legend_non_style: take_identifier(&mut next)?,
            value_axis_styles: [take_identifier(&mut next)?, take_identifier(&mut next)?],
            value_axis_non_styles: [take_identifier(&mut next)?, take_identifier(&mut next)?],
            category_axis_style: take_identifier(&mut next)?,
            category_axis_non_style: take_identifier(&mut next)?,
            series_styles: [
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
            ],
        })
    }

    const fn all(self) -> [u64; Self::COUNT] {
        [
            self.drawable,
            self.caption,
            self.title,
            self.mediator,
            self.preset,
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
            self.value_axis_styles[0],
            self.value_axis_styles[1],
            self.value_axis_non_styles[0],
            self.value_axis_non_styles[1],
            self.category_axis_style,
            self.category_axis_non_style,
            self.series_styles[0],
            self.series_styles[1],
            self.series_styles[2],
            self.series_styles[3],
            self.series_styles[4],
            self.series_styles[5],
        ]
    }

    const fn last(self) -> u64 {
        self.series_styles[5]
    }

    fn chart_references(self) -> Vec<u64> {
        let mut references = Vec::with_capacity(Self::COUNT - 1);
        references.extend([self.caption, self.title, self.mediator]);
        references.extend([
            self.preset,
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ]);
        references.extend(self.value_axis_styles);
        references.extend(self.value_axis_non_styles);
        references.extend([self.category_axis_style, self.category_axis_non_style]);
        references.extend(self.series_styles);
        references
    }
}

struct SheetChartGraph {
    archive_name: String,
    component_id: u64,
    info: NumbersSheetChartInfo,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
    private_preset_id: Option<u64>,
}

struct ChartThemeContext {
    archive_name: String,
    component_id: u64,
    theme_id: u64,
    paragraph_style_id: u64,
}

#[derive(Debug, Clone, Copy)]
enum GridAxis {
    Row,
    Column,
}

impl GridAxis {
    const fn label(self) -> &'static str {
        match self {
            Self::Row => "row",
            Self::Column => "column",
        }
    }

    const fn identifier_offset(self) -> u64 {
        match self {
            Self::Row => 0,
            Self::Column => 1_u64 << 47,
        }
    }
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
        let geometry = chart_geometry(position, size)?;
        let (archive_name, _, _) = numbers_sheet(self.package(), sheet_id)?;
        let component_id = component_identifier_for_entry(self.package(), &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet component {archive_name} is not registered"
                ))
            })?;
        let theme = chart_theme_context(self.package())?;
        let ids = ChartObjectIds::allocate(self.package())?;
        let objects = chart_objects(
            ids,
            sheet_id,
            kind,
            data.clone(),
            geometry,
            theme.paragraph_style_id,
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

#[allow(deprecated)]
fn chart_objects(
    ids: ChartObjectIds,
    sheet_id: u64,
    kind: ChartKind,
    data: ChartData,
    geometry: DrawableGeometry,
    paragraph_style_id: u64,
) -> Result<[ArchiveObject; ChartObjectIds::COUNT]> {
    let paragraph_styles = repeated_references(CHART_PARAGRAPH_STYLE_COUNT, paragraph_style_id);
    let series_count = data.row_names().len();
    let mut chart = IWorkChartArchive::new(
        tsch::ChartDrawableArchive {
            super_: Some(tsd::DrawableArchive {
                geometry: Some(geometry_archive(geometry)?),
                parent: Some(reference(sheet_id)),
                exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                    r#type: Some(4),
                    direction: Some(2),
                    fit_type: Some(1),
                    margin: Some(DEFAULT_TEXT_WRAP_MARGIN_POINTS),
                    alpha_threshold: Some(DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD),
                    is_html_wrap: Some(false),
                }),
                locked: Some(false),
                aspect_ratio_locked: Some(false),
                title: Some(reference(ids.title)),
                caption: Some(reference(ids.caption)),
                title_hidden: Some(false),
                caption_hidden: Some(false),
                ..Default::default()
            }),
        },
        tsch::ChartArchive {
            chart_type: Some(kind.into_raw()),
            scatter_format: Some(tsch::ScatterFormat::SharedX as i32),
            preset: Some(reference(ids.preset)),
            series_direction: Some(tsch::SeriesDirection::ByRow as i32),
            contains_default_data: None,
            grid: Some(chart_grid(ids.drawable, data)?),
            mediator: Some(reference(ids.mediator)),
            chart_style: Some(reference(ids.chart_style)),
            chart_non_style: Some(reference(ids.chart_non_style)),
            legend_style: Some(reference(ids.legend_style)),
            legend_non_style: Some(reference(ids.legend_non_style)),
            value_axis_styles: ids.value_axis_styles.map(reference).to_vec(),
            value_axis_nonstyles: ids.value_axis_non_styles.map(reference).to_vec(),
            category_axis_styles: vec![reference(ids.category_axis_style)],
            category_axis_nonstyles: vec![reference(ids.category_axis_non_style)],
            series_theme_styles: ids.series_styles.map(reference).to_vec(),
            series_private_styles: Some(tsp::SparseReferenceArray {
                count: 0,
                entries: Vec::new(),
            }),
            series_non_styles: Some(tsp::SparseReferenceArray {
                count: 0,
                entries: Vec::new(),
            }),
            paragraph_styles: paragraph_styles.clone(),
            multidataset_index: Some(DEFAULT_CHART_DATASET_INDEX),
            needs_calc_engine_deferred_import_action: Some(false),
            is_dirty: Some(false),
            ..Default::default()
        },
    );
    chart.append_chart_bool_extension(CHART_SCENE_DEPTH_EXTENSION_FIELD, true)?;
    chart.append_chart_bool_extension(CHART_APPEARANCE_PRESERVED_EXTENSION_FIELD, true)?;
    chart.append_chart_bool_extension(CHART_PROPORTIONAL_CALLOUTS_EXTENSION_FIELD, true)?;
    chart.append_chart_bool_extension(CHART_ROUNDED_CORNERS_EXTENSION_FIELD, true)?;
    chart.append_chart_bool_extension(CHART_VALUE_LABEL_SPACING_EXTENSION_FIELD, true)?;
    chart.append_chart_bool_extension(CHART_ERROR_BAR_SPACING_EXTENSION_FIELD, true)?;
    chart.append_chart_bool_extension(CHART_STACKED_SUMMARY_LABELS_EXTENSION_FIELD, true)?;
    chart.append_chart_message_extension(
        CHART_CACHED_FORMATTERS_EXTENSION_FIELD,
        &default_cached_formatters(series_count)?,
    )?;
    let mediator = tn::ChartMediatorArchive {
        super_: tsch::ChartMediatorArchive {
            info: None,
            local_series_indexes: vec![MEDIATOR_LOCAL_SERIES_SENTINEL],
            remote_series_indexes: vec![MEDIATOR_REMOTE_SERIES_INDEX],
        },
        entity_id: allocate_table_uuid(ids.mediator, &HashSet::new()),
        formulas: Some(tn::ChartMediatorFormulaStorage {
            direction: Some(MEDIATOR_FORMULA_DIRECTION),
            scheme: Some(MEDIATOR_FORMULA_SCHEME),
            ..Default::default()
        }),
        columns_are_series: None,
        is_registered_with_calc_engine: None,
    };
    let preset = tsch::ChartStylePreset {
        chart_style: Some(reference(ids.chart_style)),
        legend_style: Some(reference(ids.legend_style)),
        value_axis_styles: ids.value_axis_styles.map(reference).to_vec(),
        category_axis_styles: vec![reference(ids.category_axis_style)],
        series_styles: ids.series_styles.map(reference).to_vec(),
        paragraph_styles,
        uuid: None,
    };
    let mut chart_references = ids.chart_references();
    chart_references.push(paragraph_style_id);
    let mut preset_references = vec![
        ids.chart_style,
        ids.legend_style,
        ids.value_axis_styles[0],
        ids.value_axis_styles[1],
        ids.category_axis_style,
        ids.series_styles[0],
        ids.series_styles[1],
        ids.series_styles[2],
        ids.series_styles[3],
        ids.series_styles[4],
        ids.series_styles[5],
    ];
    preset_references.push(paragraph_style_id);
    Ok([
        chart_object(
            ids.drawable,
            CHART_MESSAGE_TYPE,
            chart.encode()?,
            STANDARD_MESSAGE_VERSION,
            &chart_references,
        )?,
        message_object(
            ids.caption,
            STANDIN_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            STANDIN_MESSAGE_VERSION,
            &[],
        )?,
        message_object(
            ids.title,
            STANDIN_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            STANDIN_MESSAGE_VERSION,
            &[],
        )?,
        message_object(
            ids.mediator,
            CHART_MEDIATOR_MESSAGE_TYPE,
            mediator,
            STANDARD_MESSAGE_VERSION,
            &[],
        )?,
        message_object(
            ids.preset,
            CHART_PRESET_MESSAGE_TYPE,
            preset,
            STANDARD_MESSAGE_VERSION,
            &preset_references,
        )?,
        extension_style_object(
            ids.chart_style,
            CHART_STYLE_MESSAGE_TYPE,
            tsch::ChartStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            tsch::generated::ChartStyleArchive {
                tschchartinfodefaultshowborder: Some(false),
                tschchartinfodefaultgridbackgroundopacity: Some(1.0),
                tschchartinfodefaultinterbargap: Some(0.2),
                tschchartinfodefaultintersetgap: Some(0.4),
                ..Default::default()
            },
        )?,
        extension_style_object(
            ids.chart_non_style,
            CHART_NON_STYLE_MESSAGE_TYPE,
            tsch::ChartNonStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            tsch::generated::ChartNonStyleArchive {
                tschchartinfodefaultshowlegend: Some(true),
                tschchartinfodefaultshowtitle: Some(false),
                tschchartinfodefaultskiphiddendata: Some(false),
                ..Default::default()
            },
        )?,
        extension_style_object(
            ids.legend_style,
            LEGEND_STYLE_MESSAGE_TYPE,
            tsch::LegendStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            tsch::generated::LegendStyleArchive {
                tschlegendmodeldefaultopacity: Some(1.0),
                ..Default::default()
            },
        )?,
        extension_style_object(
            ids.legend_non_style,
            LEGEND_NON_STYLE_MESSAGE_TYPE,
            tsch::LegendNonStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            tsch::generated::LegendNonStyleArchive::default(),
        )?,
        extension_style_object(
            ids.value_axis_styles[0],
            AXIS_STYLE_MESSAGE_TYPE,
            tsch::ChartAxisStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_axis_style(),
        )?,
        extension_style_object(
            ids.value_axis_styles[1],
            AXIS_STYLE_MESSAGE_TYPE,
            tsch::ChartAxisStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_axis_style(),
        )?,
        extension_style_object(
            ids.value_axis_non_styles[0],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            tsch::ChartAxisNonStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_axis_non_style(),
        )?,
        extension_style_object(
            ids.value_axis_non_styles[1],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            tsch::ChartAxisNonStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_axis_non_style(),
        )?,
        extension_style_object(
            ids.category_axis_style,
            AXIS_STYLE_MESSAGE_TYPE,
            tsch::ChartAxisStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_axis_style(),
        )?,
        extension_style_object(
            ids.category_axis_non_style,
            AXIS_NON_STYLE_MESSAGE_TYPE,
            tsch::ChartAxisNonStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_axis_non_style(),
        )?,
        extension_style_object(
            ids.series_styles[0],
            SERIES_STYLE_MESSAGE_TYPE,
            tsch::ChartSeriesStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_series_style(0),
        )?,
        extension_style_object(
            ids.series_styles[1],
            SERIES_STYLE_MESSAGE_TYPE,
            tsch::ChartSeriesStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_series_style(1),
        )?,
        extension_style_object(
            ids.series_styles[2],
            SERIES_STYLE_MESSAGE_TYPE,
            tsch::ChartSeriesStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_series_style(2),
        )?,
        extension_style_object(
            ids.series_styles[3],
            SERIES_STYLE_MESSAGE_TYPE,
            tsch::ChartSeriesStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_series_style(3),
        )?,
        extension_style_object(
            ids.series_styles[4],
            SERIES_STYLE_MESSAGE_TYPE,
            tsch::ChartSeriesStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_series_style(4),
        )?,
        extension_style_object(
            ids.series_styles[5],
            SERIES_STYLE_MESSAGE_TYPE,
            tsch::ChartSeriesStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            },
            default_series_style(5),
        )?,
    ])
}

fn chart_grid(seed: u64, data: ChartData) -> Result<tsch::ChartGridArchive> {
    let (row_names, column_names, values) = data.into_parts();
    let row_id_map = (0..row_names.len())
        .map(|index| grid_id_entry(seed, index, GridAxis::Row))
        .collect::<Result<_>>()?;
    let column_id_map = (0..column_names.len())
        .map(|index| grid_id_entry(seed, index, GridAxis::Column))
        .collect::<Result<_>>()?;
    Ok(tsch::ChartGridArchive {
        row_name: row_names,
        column_name: column_names,
        grid_row: values
            .into_iter()
            .map(|row| tsch::GridRow {
                value: row
                    .into_iter()
                    .map(|numeric_value| tsch::GridValue {
                        numeric_value,
                        ..Default::default()
                    })
                    .collect(),
            })
            .collect(),
        id_map: Some(tsch::chart_grid_archive::ChartGridRowColumnIdMap {
            row_id_map,
            column_id_map,
        }),
    })
}

fn grid_id_entry(
    seed: u64,
    index: usize,
    axis: GridAxis,
) -> Result<tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry> {
    let index_u32 = u32::try_from(index)
        .map_err(|_| Error::ParseError(format!("chart {} index exceeds u32", axis.label())))?;
    let offset = u64::try_from(index)
        .map_err(|_| Error::ParseError(format!("chart {} index exceeds u64", axis.label())))?;
    Ok(
        tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
            unique_id: allocate_table_uuid(
                seed.wrapping_add(axis.identifier_offset())
                    .wrapping_add(offset),
                &HashSet::new(),
            ),
            index: index_u32,
        },
    )
}

fn chart_object(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    versions: &[u32],
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = versions.to_vec();
    info.object_references = references.to_vec();
    Ok(object)
}

fn message_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    versions: &[u32],
    references: &[u64],
) -> Result<ArchiveObject> {
    chart_object(
        identifier,
        message_type,
        message.encode_to_vec(),
        versions,
        references,
    )
}

fn extension_style_object(
    identifier: u64,
    message_type: u32,
    base: impl Message,
    extension: impl Message,
) -> Result<ArchiveObject> {
    let mut data = base.encode_to_vec();
    append_length_delimited_field(
        &mut data,
        CURRENT_STYLE_EXTENSION_FIELD,
        &extension.encode_to_vec(),
    )?;
    chart_object(
        identifier,
        message_type,
        data,
        STANDARD_MESSAGE_VERSION,
        &[],
    )
}

fn default_axis_style() -> tsch::generated::ChartAxisStyleArchive {
    tsch::generated::ChartAxisStyleArchive {
        tschchartaxiscategoryshowaxis: Some(true),
        tschchartaxisvalueshowaxis: Some(true),
        tschchartaxiscategoryshowlastlabel: Some(true),
        tschchartaxisvalueshowmajorgridlines: Some(true),
        tschchartaxiscategoryshowmajorgridlines: Some(false),
        ..Default::default()
    }
}

fn default_axis_non_style() -> tsch::generated::ChartAxisNonStyleArchive {
    tsch::generated::ChartAxisNonStyleArchive {
        tschchartaxiscategoryshowlabels: Some(true),
        tschchartaxisdefaultshowlabels: Some(true),
        tschchartaxisvalueshowlabels: Some(true),
        tschchartaxisvaluenumberofmajorgridlines: Some(5),
        tschchartaxisvaluenumberofminorgridlines: Some(1),
        ..Default::default()
    }
}

fn default_series_style(index: usize) -> tsch::generated::ChartSeriesStyleArchive {
    const COLORS: [(f32, f32, f32); 6] = [
        (0.16, 0.55, 0.88),
        (0.29, 0.70, 0.39),
        (0.57, 0.57, 0.60),
        (0.95, 0.65, 0.16),
        (0.72, 0.25, 0.23),
        (0.62, 0.25, 0.55),
    ];
    let (red, green, blue) = COLORS[index % COLORS.len()];
    let fill = tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(red),
            g: Some(green),
            b: Some(blue),
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
            a: Some(1.0),
            ..Default::default()
        }),
        ..Default::default()
    };
    tsch::generated::ChartSeriesStyleArchive {
        tschchartseriesdefaultfill: Some(fill.clone()),
        tschchartseriescolumnfill: Some(fill.clone()),
        tschchartseriesbarfill: Some(fill.clone()),
        tschchartseriesareafill: Some(fill.clone()),
        tschchartseriespiefill: Some(fill),
        ..Default::default()
    }
}

fn default_cached_formatters(
    series_count: usize,
) -> Result<tsch::CachedDataFormatterPersistableStyleObjects> {
    let series_count = i32::try_from(series_count)
        .map_err(|_| Error::ParseError("chart series count exceeds i32".to_owned()))?;
    Ok(tsch::CachedDataFormatterPersistableStyleObjects {
        axis_data_formatter_list: [tsch::AxisType::X, tsch::AxisType::Y]
            .into_iter()
            .map(
                |axis_type| tsch::CachedAxisDataFormatterPersistableStyleObject {
                    axis_id: Some(tsch::ChartAxisIdArchive {
                        axis_type: Some(axis_type as i32),
                        ordinal: Some(0),
                    }),
                    style_object: Some(default_number_formatter()),
                },
            )
            .collect(),
        series_data_formatter_list: (0..series_count)
            .map(
                |series_index| tsch::CachedSeriesDataFormatterPersistableStyleObject {
                    series_index: Some(series_index),
                    style_object: Some(default_number_formatter()),
                },
            )
            .collect(),
        summary_label_style_object: Some(default_number_formatter()),
    })
}

fn default_number_formatter() -> tsk::FormatStructArchive {
    tsk::FormatStructArchive {
        format_type: Some(DEFAULT_CHART_NUMBER_FORMAT_TYPE),
        decimal_places: Some(AUTOMATIC_CHART_DECIMAL_PLACES),
        negative_style: Some(DEFAULT_CHART_NEGATIVE_STYLE),
        show_thousands_separator: Some(true),
        ..Default::default()
    }
}

fn chart_theme_context(package: &IWorkPackage) -> Result<ChartThemeContext> {
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

fn patch_theme_chart_preset(
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

fn repeated_references(count: usize, identifier: u64) -> Vec<tsp::Reference> {
    std::iter::repeat_with(|| reference(identifier))
        .take(count)
        .collect()
}

fn chart_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetChartGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own chart {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers chart {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not exactly one chart"
        )));
    };
    let chart = IWorkChartArchive::decode(&message.data)?;
    let drawable = chart.drawable.super_.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no drawable payload"
        ))
    })?;
    if drawable
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(sheet_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} does not name sheet {sheet_id} as its parent"
        )));
    }
    let payload = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no chart payload"
        ))
    })?;
    let caption_id = required_chart_reference(
        drawable_object_id,
        drawable.caption.as_ref(),
        "caption stand-in",
    )?;
    let title_id = required_chart_reference(
        drawable_object_id,
        drawable.title.as_ref(),
        "title stand-in",
    )?;
    let mediator_id =
        required_chart_reference(drawable_object_id, payload.mediator.as_ref(), "mediator")?;
    let preset_id = payload
        .preset
        .as_ref()
        .map(|reference| reference.identifier);
    let mut object_ids = vec![drawable_object_id, caption_id, title_id, mediator_id];
    let mut local_styles = Vec::new();
    local_styles.extend(
        payload
            .preset
            .map(|reference| (reference.identifier, CHART_PRESET_MESSAGE_TYPE, "preset")),
    );
    local_styles.extend(payload.chart_style.map(|reference| {
        (
            reference.identifier,
            CHART_STYLE_MESSAGE_TYPE,
            "chart style",
        )
    }));
    local_styles.extend(payload.chart_non_style.map(|reference| {
        (
            reference.identifier,
            CHART_NON_STYLE_MESSAGE_TYPE,
            "chart non-style",
        )
    }));
    local_styles.extend(payload.legend_style.map(|reference| {
        (
            reference.identifier,
            LEGEND_STYLE_MESSAGE_TYPE,
            "legend style",
        )
    }));
    local_styles.extend(payload.legend_non_style.map(|reference| {
        (
            reference.identifier,
            LEGEND_NON_STYLE_MESSAGE_TYPE,
            "legend non-style",
        )
    }));
    local_styles.extend(payload.value_axis_styles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_STYLE_MESSAGE_TYPE,
            "value-axis style",
        )
    }));
    local_styles.extend(payload.value_axis_nonstyles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_NON_STYLE_MESSAGE_TYPE,
            "value-axis non-style",
        )
    }));
    local_styles.extend(payload.category_axis_styles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_STYLE_MESSAGE_TYPE,
            "category-axis style",
        )
    }));
    local_styles.extend(payload.category_axis_nonstyles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_NON_STYLE_MESSAGE_TYPE,
            "category-axis non-style",
        )
    }));
    local_styles.extend(payload.series_theme_styles.iter().map(|reference| {
        (
            reference.identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            "series style",
        )
    }));
    for (identifier, message_type, label) in local_styles {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            continue;
        }
        let style = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {label} {identifier} is missing"))
        })?;
        if style
            .messages
            .iter()
            .filter(|message| message.type_ == message_type)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {label} {identifier} must have exactly one expected payload"
            )));
        }
        object_ids.push(identifier);
    }
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} aliases private objects"
        )));
    }
    for (identifier, message_type, label) in [
        (caption_id, STANDIN_MESSAGE_TYPE, "caption stand-in"),
        (title_id, STANDIN_MESSAGE_TYPE, "title stand-in"),
        (mediator_id, CHART_MEDIATOR_MESSAGE_TYPE, "mediator"),
    ] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {label} {identifier} is outside {archive_name}"
            )));
        }
        let private = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {label} {identifier} is missing"))
        })?;
        if private
            .messages
            .iter()
            .filter(|message| message.type_ == message_type)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {label} {identifier} must have exactly one expected payload"
            )));
        }
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    if !registered.is_empty() && uuid_object_ids.len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers component {component_id} UUID map does not cover chart {drawable_object_id}"
        )));
    }
    let private_preset_id = preset_id.filter(|identifier| {
        locations.get(identifier).map(String::as_str) == Some(archive_name.as_str())
    });
    Ok(SheetChartGraph {
        archive_name,
        component_id,
        info: NumbersSheetChartInfo {
            sheet_id,
            drawable_object_id,
            kind: ChartKind::from_raw(
                payload
                    .chart_type
                    .unwrap_or(tsch::ChartType::UndefinedChartType as i32),
            ),
            direction: ChartSeriesDirection::from_raw(
                payload
                    .series_direction
                    .unwrap_or(tsch::SeriesDirection::Unknown as i32),
            ),
            data: chart_data(drawable_object_id, payload)?,
            geometry: drawable_geometry(drawable_object_id, drawable)?,
        },
        object_ids,
        uuid_object_ids,
        private_preset_id,
    })
}

fn chart_data(drawable_object_id: u64, chart: &tsch::ChartArchive) -> Result<ChartData> {
    let grid = chart.grid.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no inline grid"
        ))
    })?;
    ChartData::new(
        grid.row_name.clone(),
        grid.column_name.clone(),
        grid.grid_row
            .iter()
            .map(|row| row.value.iter().map(|value| value.numeric_value).collect())
            .collect(),
    )
}

fn drawable_geometry(
    drawable_object_id: u64,
    drawable: &tsd::DrawableArchive,
) -> Result<DrawableGeometry> {
    let geometry = drawable.geometry.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no geometry"
        ))
    })?;
    DrawableGeometry {
        position: geometry.position.map(|point| DrawablePoint {
            x: point.x,
            y: point.y,
        }),
        size: geometry.size.map(|size| DrawableSize {
            width: size.width,
            height: size.height,
        }),
        flags: geometry.flags,
        angle: geometry.angle,
    }
    .validate()
}

fn geometry_archive(geometry: DrawableGeometry) -> Result<tsd::GeometryArchive> {
    geometry.validate()?;
    Ok(tsd::GeometryArchive {
        position: geometry.position.map(|point| tsp::Point {
            x: point.x,
            y: point.y,
        }),
        size: geometry.size.map(|size| tsp::Size {
            width: size.width,
            height: size.height,
        }),
        flags: geometry.flags,
        angle: geometry.angle,
    })
}

fn chart_geometry(position: DrawablePoint, size: DrawableSize) -> Result<DrawableGeometry> {
    if size.width <= 0.0 || size.height <= 0.0 {
        return Err(Error::ParseError(
            "Numbers chart dimensions must be positive".to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(position),
        size: Some(size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_ROTATION_DEGREES),
    }
    .validate()
}

fn required_chart_reference(
    drawable_object_id: u64,
    reference: Option<&tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {drawable_object_id} has no {label}"))
        })
}

fn require_creatable_kind(kind: ChartKind) -> Result<()> {
    if matches!(kind, ChartKind::Undefined | ChartKind::Unsupported(_)) {
        return Err(Error::ParseError(
            "chart kind must be a supported concrete iWork kind".to_owned(),
        ));
    }
    Ok(())
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
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
        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 72.0, y: 360.0 }),
            size: Some(DrawableSize {
                width: 500.0,
                height: 300.0,
            }),
            flags: Some(DEFAULT_DRAWABLE_FLAGS),
            angle: Some(0.0),
        };
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
