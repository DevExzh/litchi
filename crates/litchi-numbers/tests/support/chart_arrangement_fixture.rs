//! Source-built Numbers chart Arrange fixtures.
//!
//! The fixture deliberately keeps the sheet ownership component separate from
//! the CalculationEngine component, as a chart created by the Numbers host
//! does.  It contains two complete, independently rooted chart graphs.  The
//! test oracle uses generated protobuf values only while constructing hostile
//! input; the production owner still has to discover and rewrite the graph
//! through its bounded wire path.

#![allow(
    dead_code,
    reason = "the arrangement integration suite consumes different fixture views"
)]

use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{tn, tsa, tsce, tsch, tsd, tsk, tsp, tss};
use prost::Message as _;

pub(crate) const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
pub(crate) const CALCULATION_MEMBER: &str = "Index/CalculationEngine.iwa";
pub(crate) const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";
pub(crate) const METADATA_MEMBER: &str = "Index/Metadata.iwa";
pub(crate) const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
pub(crate) const SENTINEL_MEMBER: &str = "Data/chart-arrangement-sentinel.bin";

pub(crate) const DOCUMENT_ID: u64 = 1;
pub(crate) const THEME_ID: u64 = 4;
pub(crate) const SIDEBAR_ID: u64 = 5;
pub(crate) const SHEETS: [u64; 2] = [8, 9];
pub(crate) const STYLESHEET_ID: u64 = 40;
pub(crate) const CALCULATION_ENGINE_ID: u64 = 31;
pub(crate) const FORMULA_OWNER_ID: u64 = 39;
pub(crate) const METADATA_ID: u64 = 2;
pub(crate) const CHARTS: [u64; 2] = [100, 200];

pub(crate) const DOCUMENT_MESSAGE_TYPE: u32 = 1;
pub(crate) const SHEET_MESSAGE_TYPE: u32 = 2;
pub(crate) const THEME_MESSAGE_TYPE: u32 = 12_009;
pub(crate) const SIDEBAR_MESSAGE_TYPE: u32 = 205;
pub(crate) const STYLESHEET_MESSAGE_TYPE: u32 = 401;
pub(crate) const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
pub(crate) const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
pub(crate) const METADATA_MESSAGE_TYPE: u32 = 11_006;

pub(crate) const CHART_MESSAGE_TYPE: u32 = 5_021;
pub(crate) const CHART_MEDIATOR_MESSAGE_TYPE: u32 = 12_006;
pub(crate) const STANDIN_MESSAGE_TYPE: u32 = 3_097;
pub(crate) const CHART_PRESET_MESSAGE_TYPE: u32 = 5_020;
pub(crate) const CHART_STYLE_MESSAGE_TYPE: u32 = 5_022;
pub(crate) const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
pub(crate) const LEGEND_STYLE_MESSAGE_TYPE: u32 = 5_024;
pub(crate) const LEGEND_NON_STYLE_MESSAGE_TYPE: u32 = 5_025;
pub(crate) const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
pub(crate) const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
pub(crate) const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
pub(crate) const SERIES_NON_STYLE_MESSAGE_TYPE: u32 = 5_029;

pub(crate) const DRAWABLE_SUPER_FIELD: u32 = 1;
pub(crate) const DRAWABLE_LOCKED_FIELD: u32 = 5;
pub(crate) const DRAWABLE_ASPECT_RATIO_LOCKED_FIELD: u32 = 7;
pub(crate) const CHART_EXTENSION_FIELD: u32 = 10_000;
pub(crate) const UNKNOWN_DRAWABLE_FIELD: u32 = 4_090;
pub(crate) const UNKNOWN_DRAWABLE_BYTES_FIELD: u32 = 4_091;

pub(crate) type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, Copy)]
pub(crate) struct ChartIds {
    pub(crate) drawable: u64,
    pub(crate) caption: u64,
    pub(crate) title: u64,
    pub(crate) mediator: u64,
    pub(crate) preset: u64,
    pub(crate) chart_style: u64,
    pub(crate) chart_non_style: u64,
    pub(crate) legend_style: u64,
    pub(crate) legend_non_style: u64,
    pub(crate) value_style: [u64; 2],
    pub(crate) value_non_style: [u64; 2],
    pub(crate) category_style: u64,
    pub(crate) category_non_style: u64,
    pub(crate) series_styles: [u64; 6],
}

impl ChartIds {
    pub(crate) const fn for_chart(index: usize) -> Self {
        let base = if index == 0 { 100 } else { 200 };
        Self {
            drawable: base,
            caption: base + 1,
            title: base + 2,
            mediator: base + 3,
            preset: base + 4,
            chart_style: base + 5,
            chart_non_style: base + 6,
            legend_style: base + 7,
            legend_non_style: base + 8,
            value_style: [base + 9, base + 10],
            value_non_style: [base + 11, base + 12],
            category_style: base + 13,
            category_non_style: base + 14,
            series_styles: [
                base + 15,
                base + 16,
                base + 17,
                base + 18,
                base + 19,
                base + 20,
            ],
        }
    }

    pub(crate) fn all(self) -> Vec<u64> {
        let mut ids = vec![self.drawable, self.caption, self.title, self.mediator];
        ids.extend([
            self.preset,
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ]);
        ids.extend(self.value_style);
        ids.extend(self.value_non_style);
        ids.extend([self.category_style, self.category_non_style]);
        ids.extend(self.series_styles);
        ids
    }

    pub(crate) fn style_ids(self) -> Vec<u64> {
        let mut ids = vec![
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ];
        ids.extend(self.value_style);
        ids.extend(self.value_non_style);
        ids.extend([self.category_style, self.category_non_style]);
        ids.extend(self.series_styles);
        ids
    }
}

pub(crate) fn ids(chart: usize) -> ChartIds {
    ChartIds::for_chart(chart)
}

pub(crate) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn uuid(seed: u64) -> tsp::Uuid {
    tsp::Uuid {
        lower: seed.saturating_add(10_000),
        upper: seed.saturating_add(20_000),
    }
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("source fixture object has no message info"))?;
    info.versions = vec![1, 0, 5];
    info.object_references = references.to_vec();
    Ok(object)
}

fn object_with_field_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
    field: u32,
) -> TestResult<ArchiveObject> {
    let mut object = object(identifier, type_, data, references)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("source fixture object has no message info"))?;
    let mut field_info = FieldInfo::new(FieldPath::new(vec![field]));
    field_info.r#type = Some(FieldType::ObjectReference);
    field_info.object_references = references.to_vec();
    info.field_infos = vec![field_info];
    Ok(object)
}

fn style_base() -> tss::StyleArchive {
    tss::StyleArchive {
        stylesheet: Some(reference(STYLESHEET_ID)),
        ..tss::StyleArchive::default()
    }
}

fn style_payload(type_: u32) -> Vec<u8> {
    match type_ {
        CHART_STYLE_MESSAGE_TYPE => tsch::ChartStyleArchive {
            super_: Some(style_base()),
            ..tsch::ChartStyleArchive::default()
        }
        .encode_to_vec(),
        CHART_NON_STYLE_MESSAGE_TYPE => tsch::ChartNonStyleArchive {
            super_: Some(style_base()),
            ..tsch::ChartNonStyleArchive::default()
        }
        .encode_to_vec(),
        LEGEND_STYLE_MESSAGE_TYPE => tsch::LegendStyleArchive {
            super_: Some(style_base()),
            ..tsch::LegendStyleArchive::default()
        }
        .encode_to_vec(),
        LEGEND_NON_STYLE_MESSAGE_TYPE => tsch::LegendNonStyleArchive {
            super_: Some(style_base()),
            ..tsch::LegendNonStyleArchive::default()
        }
        .encode_to_vec(),
        AXIS_STYLE_MESSAGE_TYPE => tsch::ChartAxisStyleArchive {
            super_: Some(style_base()),
            ..tsch::ChartAxisStyleArchive::default()
        }
        .encode_to_vec(),
        AXIS_NON_STYLE_MESSAGE_TYPE => tsch::ChartAxisNonStyleArchive {
            super_: Some(style_base()),
            ..tsch::ChartAxisNonStyleArchive::default()
        }
        .encode_to_vec(),
        SERIES_STYLE_MESSAGE_TYPE => tsch::ChartSeriesStyleArchive {
            super_: Some(style_base()),
            ..tsch::ChartSeriesStyleArchive::default()
        }
        .encode_to_vec(),
        SERIES_NON_STYLE_MESSAGE_TYPE => tsch::ChartSeriesNonStyleArchive {
            super_: Some(style_base()),
            ..tsch::ChartSeriesNonStyleArchive::default()
        }
        .encode_to_vec(),
        _ => Vec::new(),
    }
}

fn chart_grid(seed: u64) -> tsch::ChartGridArchive {
    tsch::ChartGridArchive {
        row_name: vec!["North".to_owned(), "South".to_owned()],
        column_name: vec!["Q1".to_owned(), "Q2".to_owned()],
        grid_row: vec![
            tsch::GridRow {
                value: vec![
                    tsch::GridValue {
                        numeric_value: Some(seed as f64),
                        ..tsch::GridValue::default()
                    },
                    tsch::GridValue {
                        numeric_value: Some(seed as f64 + 1.0),
                        ..tsch::GridValue::default()
                    },
                ],
            },
            tsch::GridRow {
                value: vec![
                    tsch::GridValue {
                        numeric_value: Some(seed as f64 + 2.0),
                        ..tsch::GridValue::default()
                    },
                    tsch::GridValue {
                        numeric_value: Some(seed as f64 + 3.0),
                        ..tsch::GridValue::default()
                    },
                ],
            },
        ],
        id_map: Some(tsch::chart_grid_archive::ChartGridRowColumnIdMap {
            row_id_map: vec![
                tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                    unique_id: format!("row-{seed}-0"),
                    index: 0,
                },
                tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                    unique_id: format!("row-{seed}-1"),
                    index: 1,
                },
            ],
            column_id_map: vec![
                tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                    unique_id: format!("column-{seed}-0"),
                    index: 0,
                },
                tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                    unique_id: format!("column-{seed}-1"),
                    index: 1,
                },
            ],
        }),
    }
}

fn chart_payload(
    chart: usize,
    sheet_identifier: u64,
    state: (Option<bool>, Option<bool>),
    include_unknown: bool,
) -> TestResult<Vec<u8>> {
    let ids = ids(chart);
    let drawable = tsd::DrawableArchive {
        geometry: Some(tsd::GeometryArchive {
            position: Some(tsp::Point { x: 24.0, y: 48.0 }),
            size: Some(tsp::Size {
                width: 320.0,
                height: 220.0,
            }),
            flags: Some(3),
            angle: Some(0.0),
        }),
        parent: Some(reference(sheet_identifier)),
        locked: state.0,
        aspect_ratio_locked: state.1,
        title: Some(reference(ids.title)),
        caption: Some(reference(ids.caption)),
        title_hidden: Some(false),
        caption_hidden: Some(false),
        ..tsd::DrawableArchive::default()
    };
    let chart = tsch::ChartArchive {
        chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
        scatter_format: Some(tsch::ScatterFormat::SharedX as i32),
        preset: Some(reference(ids.preset)),
        series_direction: Some(tsch::SeriesDirection::ByRow as i32),
        grid: Some(chart_grid(ids.drawable)),
        mediator: Some(reference(ids.mediator)),
        chart_style: Some(reference(ids.chart_style)),
        chart_non_style: Some(reference(ids.chart_non_style)),
        legend_style: Some(reference(ids.legend_style)),
        legend_non_style: Some(reference(ids.legend_non_style)),
        value_axis_styles: ids.value_style.into_iter().map(reference).collect(),
        value_axis_nonstyles: ids.value_non_style.into_iter().map(reference).collect(),
        category_axis_styles: vec![reference(ids.category_style)],
        category_axis_nonstyles: vec![reference(ids.category_non_style)],
        series_theme_styles: ids.series_styles.into_iter().map(reference).collect(),
        series_private_styles: Some(tsp::SparseReferenceArray {
            count: 0,
            entries: Vec::new(),
        }),
        series_non_styles: Some(tsp::SparseReferenceArray {
            count: 0,
            entries: Vec::new(),
        }),
        multidataset_index: Some(0),
        needs_calc_engine_deferred_import_action: Some(false),
        is_dirty: Some(false),
        ..tsch::ChartArchive::default()
    };
    let mut payload = Vec::new();
    let mut drawable_payload = drawable.encode_to_vec();
    if include_unknown {
        append_varint_field(&mut drawable_payload, UNKNOWN_DRAWABLE_FIELD, 0x5a5a)?;
        append_length_delimited_field(
            &mut drawable_payload,
            UNKNOWN_DRAWABLE_BYTES_FIELD,
            b"opaque chart drawable metadata",
        )?;
    }
    append_length_delimited_field(&mut payload, DRAWABLE_SUPER_FIELD, &drawable_payload)?;
    append_length_delimited_field(&mut payload, CHART_EXTENSION_FIELD, &chart.encode_to_vec())?;
    Ok(payload)
}

fn chart_objects(
    chart: usize,
    sheet_identifier: u64,
    state: (Option<bool>, Option<bool>),
    include_unknown: bool,
) -> TestResult<Vec<ArchiveObject>> {
    let ids = ids(chart);
    let mut objects = Vec::new();
    let mut chart_object = object(
        ids.drawable,
        CHART_MESSAGE_TYPE,
        chart_payload(chart, sheet_identifier, state, include_unknown)?,
        ids.all()
            .into_iter()
            .filter(|id| *id != ids.drawable)
            .collect::<Vec<_>>()
            .as_slice(),
    )?;
    chart_object.archive_info.message_infos[0].field_infos = vec![FieldInfo::new(vec![1])];
    objects.push(chart_object);
    objects.push(object(
        ids.caption,
        STANDIN_MESSAGE_TYPE,
        tsd::StandinCaptionArchive::default().encode_to_vec(),
        &[],
    )?);
    objects.push(object(
        ids.title,
        STANDIN_MESSAGE_TYPE,
        tsd::StandinCaptionArchive::default().encode_to_vec(),
        &[],
    )?);
    let mediator = tn::ChartMediatorArchive {
        super_: tsch::ChartMediatorArchive {
            local_series_indexes: vec![u32::MAX],
            remote_series_indexes: vec![0],
            ..tsch::ChartMediatorArchive::default()
        },
        entity_id: format!("00000000-0000-4000-8000-{:012X}", ids.mediator),
        formulas: Some(tn::ChartMediatorFormulaStorage {
            direction: Some(0),
            scheme: Some(0),
            ..tn::ChartMediatorFormulaStorage::default()
        }),
        ..tn::ChartMediatorArchive::default()
    };
    objects.push(object(
        ids.mediator,
        CHART_MEDIATOR_MESSAGE_TYPE,
        mediator.encode_to_vec(),
        &[],
    )?);
    let preset = tsch::ChartStylePreset {
        chart_style: Some(reference(ids.chart_style)),
        legend_style: Some(reference(ids.legend_style)),
        value_axis_styles: ids.value_style.into_iter().map(reference).collect(),
        category_axis_styles: vec![reference(ids.category_style)],
        series_styles: ids.series_styles.into_iter().map(reference).collect(),
        ..tsch::ChartStylePreset::default()
    };
    let mut preset_refs = vec![ids.chart_style, ids.legend_style, ids.category_style];
    preset_refs.extend(ids.value_style);
    preset_refs.extend(ids.series_styles);
    objects.push(object(
        ids.preset,
        CHART_PRESET_MESSAGE_TYPE,
        preset.encode_to_vec(),
        &preset_refs,
    )?);
    for (identifier, type_) in [
        (ids.chart_style, CHART_STYLE_MESSAGE_TYPE),
        (ids.chart_non_style, CHART_NON_STYLE_MESSAGE_TYPE),
        (ids.legend_style, LEGEND_STYLE_MESSAGE_TYPE),
        (ids.legend_non_style, LEGEND_NON_STYLE_MESSAGE_TYPE),
        (ids.value_style[0], AXIS_STYLE_MESSAGE_TYPE),
        (ids.value_style[1], AXIS_STYLE_MESSAGE_TYPE),
        (ids.value_non_style[0], AXIS_NON_STYLE_MESSAGE_TYPE),
        (ids.value_non_style[1], AXIS_NON_STYLE_MESSAGE_TYPE),
        (ids.category_style, AXIS_STYLE_MESSAGE_TYPE),
        (ids.category_non_style, AXIS_NON_STYLE_MESSAGE_TYPE),
    ] {
        objects.push(object(
            identifier,
            type_,
            style_payload(type_),
            &[STYLESHEET_ID],
        )?);
    }
    for identifier in ids.series_styles {
        objects.push(object(
            identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            style_payload(SERIES_STYLE_MESSAGE_TYPE),
            &[STYLESHEET_ID],
        )?);
    }
    Ok(objects)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

#[allow(deprecated)]
fn document_objects(include_unknown: bool) -> TestResult<Vec<ArchiveObject>> {
    let root = tn::DocumentArchive {
        sheets: SHEETS.into_iter().map(reference).collect(),
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive {
                locale_identifier: Some("en_US".to_owned()),
                ..tsk::DocumentArchive::default()
            },
            ..tsa::DocumentArchive::default()
        },
        calculation_engine: Some(reference(CALCULATION_ENGINE_ID)),
        stylesheet: reference(STYLESHEET_ID),
        sidebar_order: reference(SIDEBAR_ID),
        theme: reference(THEME_ID),
        ..tn::DocumentArchive::default()
    };
    let mut objects = vec![object_with_field_references(
        DOCUMENT_ID,
        DOCUMENT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[
            SHEETS[0],
            SHEETS[1],
            STYLESHEET_ID,
            SIDEBAR_ID,
            THEME_ID,
            CALCULATION_ENGINE_ID,
        ],
        1,
    )?];
    objects.push(object(
        THEME_ID,
        THEME_MESSAGE_TYPE,
        tn::ThemeArchive {
            super_: tss::ThemeArchive {
                document_stylesheet: Some(reference(STYLESHEET_ID)),
                ..tss::ThemeArchive::default()
            },
            ..tn::ThemeArchive::default()
        }
        .encode_to_vec(),
        &[STYLESHEET_ID],
    )?);
    objects.push(object(
        SIDEBAR_ID,
        SIDEBAR_MESSAGE_TYPE,
        tsk::TreeNode {
            children: SHEETS.into_iter().map(reference).collect(),
            ..tsk::TreeNode::default()
        }
        .encode_to_vec(),
        &SHEETS,
    )?);
    for (index, sheet_identifier) in SHEETS.into_iter().enumerate() {
        let mut payload = tn::SheetArchive {
            name: format!("Sheet {}", index + 1),
            drawable_infos: vec![reference(CHARTS[index])],
            ..tn::SheetArchive::default()
        }
        .encode_to_vec();
        if include_unknown && index == 0 {
            append_varint_field(&mut payload, 4_080, 37)?;
        }
        objects.push(object_with_field_references(
            sheet_identifier,
            SHEET_MESSAGE_TYPE,
            payload,
            &[CHARTS[index]],
            2,
        )?);
    }
    Ok(objects)
}

fn component_info(
    identifier: u64,
    locator: &str,
    object_ids: impl IntoIterator<Item = u64>,
    external_references: Vec<tsp::ComponentExternalReference>,
) -> tsp::ComponentInfo {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        document_read_version: vec![3, 2, 10],
        document_write_version: vec![3, 2, 10],
        save_token: Some(1),
        object_uuid_map_entries: object_ids
            .into_iter()
            .map(|identifier| tsp::ObjectUuidMapEntry {
                identifier,
                uuid: uuid(identifier),
            })
            .collect(),
        external_references,
        ..tsp::ComponentInfo::default()
    }
}

fn external(
    component_identifier: u64,
    object_identifier: Option<u64>,
) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier,
        is_weak: None,
    }
}

fn metadata_objects() -> TestResult<Vec<ArchiveObject>> {
    let chart_object_ids = CHARTS
        .into_iter()
        .flat_map(|chart| ids(if chart == CHARTS[0] { 0 } else { 1 }).all())
        .collect::<Vec<_>>();
    let mut calc_external = Vec::new();
    calc_external.extend([external(STYLESHEET_ID, Some(STYLESHEET_ID))]);
    let mut stylesheet_external = Vec::new();
    stylesheet_external.extend(
        chart_object_ids
            .iter()
            .copied()
            .filter(|identifier| *identifier != ids(0).drawable && *identifier != ids(1).drawable)
            .map(|identifier| external(CALCULATION_ENGINE_ID, Some(identifier))),
    );
    let mut document_external = vec![
        external(STYLESHEET_ID, None),
        external(CALCULATION_ENGINE_ID, None),
    ];
    document_external.extend(
        CHARTS
            .into_iter()
            .map(|identifier| external(CALCULATION_ENGINE_ID, Some(identifier))),
    );
    let document_ids = [DOCUMENT_ID, THEME_ID, SIDEBAR_ID, SHEETS[0], SHEETS[1]];
    let mut calc_ids = vec![CALCULATION_ENGINE_ID, FORMULA_OWNER_ID];
    calc_ids.extend(chart_object_ids);
    let metadata = tsp::PackageMetadata {
        last_object_identifier: 400,
        components: vec![
            component_info(DOCUMENT_ID, "Document", document_ids, document_external),
            component_info(
                CALCULATION_ENGINE_ID,
                "CalculationEngine",
                calc_ids,
                calc_external,
            ),
            component_info(
                STYLESHEET_ID,
                "DocumentStylesheet",
                [STYLESHEET_ID],
                stylesheet_external,
            ),
        ],
        read_version: vec![3, 2, 10],
        write_version: vec![3, 2, 10],
        file_format_version: vec![14, 4, 1],
        save_token: Some(1),
        ..tsp::PackageMetadata::default()
    };
    Ok(vec![object(
        METADATA_ID,
        METADATA_MESSAGE_TYPE,
        metadata.encode_to_vec(),
        &[],
    )?])
}

pub(crate) fn fixture_with_states(
    states: [(Option<bool>, Option<bool>); 2],
    include_unknown: bool,
) -> TestResult<Vec<u8>> {
    let mut document_objects = document_objects(include_unknown)?;
    let mut calculation_objects = vec![object(
        CALCULATION_ENGINE_ID,
        CALCULATION_ENGINE_MESSAGE_TYPE,
        tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive {
                owner_id_map: Some(tsce::OwnerIdMapArchive::default()),
                number_of_formulas: Some(0),
                formula_owner_dependencies: vec![reference(FORMULA_OWNER_ID)],
                ..tsce::DependencyTrackerArchive::default()
            },
            saved_locale_identifier: Some("en_US".to_owned()),
            ..tsce::CalculationEngineArchive::default()
        }
        .encode_to_vec(),
        &[FORMULA_OWNER_ID],
    )?];
    calculation_objects.push(object(
        FORMULA_OWNER_ID,
        FORMULA_OWNER_MESSAGE_TYPE,
        tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: uuid(FORMULA_OWNER_ID),
            internal_formula_owner_id: 6,
            ..tsce::FormulaOwnerDependenciesArchive::default()
        }
        .encode_to_vec(),
        &[],
    )?);
    for (index, sheet_identifier) in SHEETS.into_iter().enumerate() {
        calculation_objects.extend(chart_objects(
            index,
            sheet_identifier,
            states[index],
            include_unknown && index == 0,
        )?);
    }
    let style_ids = CHARTS
        .into_iter()
        .flat_map(|chart| ids(if chart == CHARTS[0] { 0 } else { 1 }).style_ids())
        .collect::<Vec<_>>();
    let stylesheet = object(
        STYLESHEET_ID,
        STYLESHEET_MESSAGE_TYPE,
        tss::StylesheetArchive {
            styles: style_ids.iter().copied().map(reference).collect(),
            ..tss::StylesheetArchive::default()
        }
        .encode_to_vec(),
        &style_ids,
    )?;
    let unrelated = object(
        900,
        99_999,
        b"unrelated Numbers chart component".to_vec(),
        &[],
    )?;
    let metadata_objects = metadata_objects()?;
    let document_component = component(std::mem::take(&mut document_objects))?;
    let calculation_component = component(calculation_objects)?;
    let stylesheet_component = component(vec![stylesheet])?;
    let metadata_component = component(metadata_objects)?;
    let unrelated_component = component(vec![unrelated])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (SENTINEL_MEMBER, b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (CALCULATION_MEMBER, calculation_component.as_slice()),
            (STYLESHEET_MEMBER, stylesheet_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
            (UNRELATED_MEMBER, unrelated_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

pub(crate) fn fixture() -> TestResult<Vec<u8>> {
    fixture_with_states([(None, None), (Some(false), Some(false))], true)
}

pub(crate) fn component_stream(source: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("missing fixture member {member}")))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

pub(crate) fn rewrite_component(
    source: &[u8],
    member: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&component_stream(source, member)?)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?
        .reassemble_to_bytes(&[EntryEdit::new(member, &compressed)], Limits::default())?)
}

pub(crate) fn object_message(
    source: &[u8],
    member: &str,
    identifier: u64,
    type_: u32,
) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&component_stream(source, member)?)?;
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == type_)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("fixture message is missing").into())
}

pub(crate) fn with_chart_payload(
    source: &[u8],
    chart: usize,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    rewrite_component(source, CALCULATION_MEMBER, |archive| {
        let object = archive
            .object_mut(ids(chart).drawable)
            .ok_or_else(|| io::Error::other("chart drawable is missing"))?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("chart payload is missing"))?;
        let length = u32::try_from(payload.len())
            .map_err(|_| io::Error::other("chart payload exceeds u32"))?;
        message.data = payload;
        object.archive_info.message_infos[0].length = length;
        Ok(())
    })
}

pub(crate) fn chart_payload_from_source(source: &[u8], chart: usize) -> TestResult<Vec<u8>> {
    object_message(
        source,
        CALCULATION_MEMBER,
        ids(chart).drawable,
        CHART_MESSAGE_TYPE,
    )
}

pub(crate) fn drawable_payload_from_source(source: &[u8], chart: usize) -> TestResult<Vec<u8>> {
    let payload = chart_payload_from_source(source, chart)?;
    WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .map(|field| field.payload().to_vec())
        .ok_or_else(|| io::Error::other("chart drawable envelope is missing").into())
}

pub(crate) fn raw_drawable_fields(
    source: &[u8],
    chart: usize,
    number: u32,
) -> TestResult<Vec<Vec<u8>>> {
    Ok(
        WireView::parse(&drawable_payload_from_source(source, chart)?)?
            .fields()
            .filter(|field| field.number() == number)
            .map(|field| field.raw().to_vec())
            .collect(),
    )
}

pub(crate) fn with_drawable_wire(
    source: &[u8],
    chart: usize,
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult,
) -> TestResult<Vec<u8>> {
    let original = chart_payload_from_source(source, chart)?;
    let drawable = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("chart drawable envelope is missing"))?;
    let mut payload = drawable.payload().to_vec();
    mutate(&mut payload)?;
    let mut replacement = Vec::new();
    append_length_delimited_field(&mut replacement, DRAWABLE_SUPER_FIELD, &payload)?;
    let mut rewritten = Vec::new();
    let mut replaced = false;
    for field in WireView::parse(&original)?.fields() {
        if !replaced && field.number() == DRAWABLE_SUPER_FIELD {
            rewritten.extend_from_slice(&replacement);
            replaced = true;
        } else {
            rewritten.extend_from_slice(field.raw());
        }
    }
    with_chart_payload(source, chart, rewritten)
}

pub(crate) fn with_duplicate_sheet_owner(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SHEETS[1])
            .ok_or_else(|| io::Error::other("second sheet is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("second sheet payload is missing"))?;
        let mut sheet = tn::SheetArchive::decode(message.data.as_slice())?;
        sheet.drawable_infos.push(reference(CHARTS[0]));
        message.data = sheet.encode_to_vec();
        object.archive_info.message_infos[0].length = u32::try_from(message.data.len())
            .map_err(|_| io::Error::other("sheet payload exceeds u32"))?;
        object.archive_info.message_infos[0]
            .object_references
            .push(CHARTS[0]);
        Ok(())
    })
}

/// Repeat the same chart reference in one sheet's drawable list.  The
/// semantic selector must reject this before exposing two positions for the
/// same native owner.
pub(crate) fn with_duplicate_chart_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SHEETS[0])
            .ok_or_else(|| io::Error::other("first sheet is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("first sheet payload is missing"))?;
        let mut sheet = tn::SheetArchive::decode(message.data.as_slice())?;
        sheet.drawable_infos.push(reference(CHARTS[0]));
        message.data = sheet.encode_to_vec();
        object.archive_info.message_infos[0].length = u32::try_from(message.data.len())
            .map_err(|_| io::Error::other("sheet payload exceeds u32"))?;
        object.archive_info.message_infos[0]
            .object_references
            .push(CHARTS[0]);
        Ok(())
    })
}

/// Keep only the first sheet rooted so the package admission reference budget
/// can be set below the two-pass edit transaction's source-plus-candidate
/// selector work.  The detached second sheet remains physical fixture data,
/// while the semantic document root is still a valid one-sheet Numbers file.
pub(crate) fn with_single_rooted_sheet(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_ID)
            .ok_or_else(|| io::Error::other("Numbers document root is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("Numbers document payload is missing"))?;
        let mut document = tn::DocumentArchive::decode(message.data.as_slice())?;
        document.sheets.truncate(1);
        message.data = document.encode_to_vec();
        object.archive_info.message_infos[0].length = u32::try_from(message.data.len())
            .map_err(|_| io::Error::other("Numbers document payload exceeds u32"))?;
        Ok(())
    })
}

pub(crate) fn with_parent_mismatch(source: &[u8]) -> TestResult<Vec<u8>> {
    with_drawable_wire(source, 0, |payload| {
        let field = WireView::parse(payload)?
            .fields()
            .find(|field| field.number() == 2)
            .ok_or_else(|| io::Error::other("chart parent is missing"))?;
        let replacement = tsp::Reference {
            identifier: SHEETS[1],
            ..tsp::Reference::default()
        }
        .encode_to_vec();
        let mut out = Vec::new();
        for candidate in WireView::parse(payload)?.fields() {
            if candidate.number() == 2 {
                append_length_delimited_field(&mut out, 2, &replacement)?;
            } else {
                out.extend_from_slice(candidate.raw());
            }
        }
        let _ = field;
        *payload = out;
        Ok(())
    })
}

pub(crate) fn with_wrong_chart_message_type(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, CALCULATION_MEMBER, |archive| {
        let object = archive
            .object_mut(ids(0).drawable)
            .ok_or_else(|| io::Error::other("chart drawable is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("chart message is missing"))?;
        message.type_ = 5_022;
        object.archive_info.message_infos[0].type_ = 5_022;
        Ok(())
    })
}

pub(crate) fn with_duplicate_chart_message(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, CALCULATION_MEMBER, |archive| {
        let object = archive
            .object_mut(ids(0).drawable)
            .ok_or_else(|| io::Error::other("chart drawable is missing"))?;
        let message = object
            .messages
            .first()
            .cloned()
            .ok_or_else(|| io::Error::other("chart message is missing"))?;
        object.push_message(message)?;
        Ok(())
    })
}
