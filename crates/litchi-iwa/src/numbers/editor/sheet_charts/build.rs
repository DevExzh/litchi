//! Source-built Numbers chart object and style graphs.

use super::*;

fn repeated_references(count: usize, identifier: u64) -> Vec<tsp::Reference> {
    std::iter::repeat_with(|| reference(identifier))
        .take(count)
        .collect()
}

#[allow(deprecated)]
pub(super) fn chart_objects(
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

pub(super) fn chart_grid(seed: u64, data: ChartData) -> Result<tsch::ChartGridArchive> {
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
