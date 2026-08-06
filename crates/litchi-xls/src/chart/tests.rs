use std::collections::HashSet;
use std::sync::Arc;

use super::codec::{parse_chart, patch_chart_data, serialize_chart};
use super::package::{GENERATED_CHART_OBJECT_ID, chart_bof};
use super::package::{build_workbook_fixture, ranges, remap_extern_sheet};
use super::wire::{
    AREA_FORMAT, BEGIN, BOF, BRAI, CHART, CRT_LINE, DROP_BAR, END, EOF, LABEL, LINE_FORMAT, NUMBER,
    SI_INDEX, record,
};
use super::*;
use crate::Error;

#[cfg(test)]
mod tests {
    use super::*;

    fn area_link(role: Role, first_row: u16, last_row: u16) -> DataLink {
        let mut formula_tokens = vec![0x3b];
        formula_tokens.extend(0u16.to_le_bytes());
        formula_tokens.extend(first_row.to_le_bytes());
        formula_tokens.extend(last_row.to_le_bytes());
        formula_tokens.extend(0u16.to_le_bytes());
        formula_tokens.extend(0u16.to_le_bytes());
        DataLink {
            role,
            source: Source::Cells,
            unlinked_number_format: false,
            number_format: 0,
            formula_tokens,
            references: vec![CellRef {
                extern_sheet_index: 0,
                first_row,
                last_row,
                first_column: 0,
                last_column: 0,
            }],
        }
    }

    fn complete_series() -> Series {
        Series {
            category_count: 2,
            value_count: 2,
            links: vec![
                DataLink {
                    role: Role::Name,
                    source: Source::Automatic,
                    unlinked_number_format: false,
                    number_format: 0,
                    formula_tokens: Vec::new(),
                    references: Vec::new(),
                },
                area_link(Role::Values, 0, 1),
                area_link(Role::Categories, 0, 1),
                DataLink {
                    role: Role::Bubbles,
                    source: Source::Automatic,
                    unlinked_number_format: false,
                    number_format: 0,
                    formula_tokens: Vec::new(),
                    references: Vec::new(),
                },
            ],
            ..Default::default()
        }
    }

    fn data_chart() -> Chart {
        let mut chart = Chart::default();
        chart.series.push(complete_series());
        chart.cached_values.extend([
            Cache::new(CacheKind::Values, 0, 0, 0, Value::Number(1.0)).expect("finite chart value"),
            Cache::new(CacheKind::Categories, 0, 0, 0, Value::Text("Jan".into()))
                .expect("bounded chart label"),
        ]);
        chart
    }

    #[test]
    fn typed_inventory_reports_data_sources_without_rendering() {
        let mut chart = Chart::default();
        chart.series.push(complete_series());
        chart.cached_values.push(Cache {
            kind: CacheKind::Values,
            point: 0,
            series: 0,
            format: 0,
            value: Value::Number(42.0),
        });
        chart.unknown_records.extend([
            Raw {
                record_type: 0x7777,
                data: vec![1],
            },
            Raw {
                record_type: 0x7777,
                data: vec![2],
            },
            Raw {
                record_type: 0x7778,
                data: vec![3],
            },
        ]);

        let inventory = chart.inventory(Limits::default()).unwrap();
        assert_eq!(inventory.kind, Kind::Line);
        assert_eq!(inventory.series_count, 1);
        assert_eq!(inventory.group_count, 1);
        assert_eq!(inventory.axis_count, 0);
        assert_eq!(inventory.data_link_count, 4);
        assert_eq!(inventory.cell_reference_count, 2);
        assert_eq!(inventory.opaque_formula_count, 0);
        assert_eq!(inventory.cached_value_count, 1);
        assert_eq!(inventory.unknown_record_count, 3);
        assert_eq!(inventory.unknown_record_types, [0x7777, 0x7778]);
        assert_eq!(
            inventory.semantic_completeness,
            SemanticCompleteness::Partial
        );
    }

    #[test]
    fn semantic_validation_checks_ordered_ai_links_and_cache_identity() {
        let mut chart = Chart::default();
        chart.series.push(complete_series());
        chart
            .validate_semantics(Limits::default())
            .expect("complete MS-XLS series links are valid");

        chart.series[0].links.swap(1, 2);
        assert!(chart.validate_semantics(Limits::default()).is_err());

        let mut chart = Chart::default();
        chart.series.push(complete_series());
        chart.cached_values.extend([
            Cache {
                kind: CacheKind::Values,
                point: 0,
                series: 0,
                format: 0,
                value: Value::Blank,
            },
            Cache {
                kind: CacheKind::Values,
                point: 0,
                series: 0,
                format: 0,
                value: Value::Blank,
            },
        ]);
        assert!(chart.validate_semantics(Limits::default()).is_err());

        let mut chart = Chart::default();
        chart.series.push(complete_series());
        chart.cached_values.push(Cache {
            kind: CacheKind::Values,
            point: 2,
            series: 0,
            format: 0,
            value: Value::Blank,
        });
        assert!(chart.validate_semantics(Limits::default()).is_err());
    }

    #[test]
    fn semantic_validation_checks_reference_cardinality_and_scatter_types() {
        let mut chart = Chart::default();
        chart.series.push(complete_series());
        chart.series[0].value_count = 3;
        assert!(chart.validate_semantics(Limits::default()).is_err());

        let mut chart = Chart::default();
        chart.groups[0].kind = GroupKind::Scatter {
            bubble_size_percent: 100,
            bubble_size_type: 1,
            flags: 0,
        };
        chart.series.push(complete_series());
        assert!(chart.validate_semantics(Limits::default()).is_err());
    }

    #[test]
    fn abbreviated_test_fixture_exercises_chart_parser() {
        let mut chart = Chart {
            title: Some("Sales".into()),
            ..Default::default()
        };
        chart.series.push(Series {
            category_count: 2,
            value_count: 2,
            links: vec![DataLink {
                role: Role::Values,
                source: Source::Cells,
                unlinked_number_format: false,
                number_format: 0,
                formula_tokens: vec![0x3b, 0, 0, 1, 0, 2, 0, 1, 0, 1, 0],
                references: vec![CellRef {
                    extern_sheet_index: 0,
                    first_row: 1,
                    last_row: 2,
                    first_column: 1,
                    last_column: 1,
                }],
            }],
            ..Default::default()
        });
        chart.cached_values.push(Cache {
            kind: CacheKind::Values,
            point: 0,
            series: 0,
            format: 0,
            value: Value::Number(42.0),
        });
        let bytes = build_workbook_fixture(chart, Limits::default()).unwrap();
        let editor = Editor::open(bytes, Limits::default()).unwrap();
        let mut charts = editor.charts();
        assert_eq!(charts.len(), 1);
        let chart = charts.next().unwrap();
        assert_eq!(
            chart.location,
            Location::Embedded {
                sheet_index: 0,
                object_id: GENERATED_CHART_OBJECT_ID
            }
        );
        assert_eq!(chart.chart.title.as_deref(), Some("Sales"));
        assert_eq!(chart.chart.series.len(), 1);
        assert!(
            chart
                .chart
                .cached_values
                .iter()
                .any(|v| v.value == Value::Number(42.0))
        );
    }

    #[test]
    fn public_authoring_refuses_atomically_with_typed_error() {
        fn assert_unsupported(error: Error) {
            assert!(matches!(
                error,
                Error::Graph(litchi_ograph::Error::UnsupportedAuthoring { .. })
            ));
        }

        let original = build_workbook_fixture(Chart::default(), Limits::default()).unwrap();
        let location = Location::Embedded {
            sheet_index: 0,
            object_id: GENERATED_CHART_OBJECT_ID,
        };

        let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
        assert_unsupported(
            editor
                .replace_at(&location, Chart::default())
                .expect_err("replacement must be refused"),
        );
        assert_eq!(editor.finish().unwrap(), original);

        assert_unsupported(
            build_workbook(Chart::default(), Limits::default())
                .expect_err("fresh workbook authoring must be refused"),
        );
    }

    #[test]
    fn embedded_identity_reorder_is_exact_and_removal_is_atomic_refusal() {
        let original = build_workbook_fixture(Chart::default(), Limits::default()).unwrap();

        let mut reordered = Editor::open(original.clone(), Limits::default()).unwrap();
        reordered.reorder("Sheet1", &[0]).unwrap();
        assert_eq!(reordered.finish().unwrap(), original);

        let mut removed = Editor::open(original.clone(), Limits::default()).unwrap();
        assert!(matches!(
            removed
                .remove(Selector::Embedded {
                    sheet: "Sheet1",
                    index: 0,
                })
                .expect_err("embedded removal must be refused"),
            Error::Graph(litchi_ograph::Error::UnsupportedMutation { .. })
        ));
        assert_eq!(removed.finish().unwrap(), original);
    }

    #[test]
    fn editor_and_package_share_the_workbook_stream_allocation() {
        let bytes = build_workbook_fixture(Chart::default(), Limits::default()).unwrap();
        let editor = Editor::open(bytes, Limits::default()).unwrap();
        let captured = editor.package.stream_shared(&editor.workbook_path).unwrap();
        assert!(Arc::ptr_eq(&captured, &editor.workbook));
    }

    #[test]
    fn generated_combo_round_trips_and_refuses_unplaced_opaque_records() {
        let mut chart = Chart::default();
        chart.groups.push(Group {
            order: 1,
            vary_colors: true,
            kind: GroupKind::Pie {
                rotation: 45,
                hole_size: 50,
                flags: 0,
            },
            lines: Vec::new(),
            drop_bars: Vec::new(),
        });
        chart.series.push(Series {
            chart_group: 1,
            ..Default::default()
        });
        let bytes = serialize_chart(&chart, Limits::default()).unwrap();
        let parsed = parse_chart(&bytes, Limits::default()).unwrap();
        assert_eq!(parsed.kind(), Kind::Combo);

        chart.unknown_records.push(Raw {
            record_type: 0x7777,
            data: b"opaque".to_vec(),
        });
        assert!(matches!(
            serialize_chart(&chart, Limits::default()),
            Err(Error::UnsafeEdit(_))
        ));
    }

    #[test]
    fn axis_line_format_and_blank_cache_round_trip_once() {
        let format = LineFormat {
            color: [1, 2, 3, 4],
            pattern: 1,
            weight: 2,
            flags: 0,
            color_index: 8,
        };
        let mut chart = Chart::default();
        chart.axes.push(Axis {
            kind: AxisKind::CategoryOrHorizontal,
            scale: None,
            tick: None,
            lines: vec![AxisLine {
                kind: AxisLineKind::Axis,
                format: format.clone(),
            }],
        });
        chart.series.push(complete_series());
        chart.cached_values.push(Cache {
            kind: CacheKind::Values,
            point: 1,
            series: 0,
            format: 9,
            value: Value::Blank,
        });

        let bytes = serialize_chart(&chart, Limits::default()).unwrap();
        let parsed = parse_chart(&bytes, Limits::default()).unwrap();
        assert_eq!(parsed.axes[0].lines[0].format, format);
        assert!(
            parsed
                .formatting
                .iter()
                .all(|value| !matches!(value, Format::Line(_)))
        );
        assert!(parsed.cached_values.iter().any(|value| {
            value.kind == CacheKind::Values
                && value.point == 1
                && value.series == 0
                && value.format == 9
                && value.value == Value::Blank
        }));
    }

    #[test]
    fn typed_formula_and_cache_edits_round_trip() {
        let limits = Limits::default();
        let mut chart = data_chart();
        let formula = Formula::range(CellRef {
            extern_sheet_index: 0,
            first_row: 4,
            last_row: 5,
            first_column: 1,
            last_column: 1,
        })
        .expect("absolute chart range");
        chart
            .set_formula(0, Role::Values, formula.clone())
            .expect("same-cardinality formula edit");
        chart
            .set_cache(CacheKind::Values, 0, 0, 9, Value::Number(7.5))
            .expect("existing cache edit");

        let bytes = serialize_chart(&chart, limits).expect("bounded chart serialization");
        let parsed = parse_chart(&bytes, limits).expect("chart round trip");
        assert_eq!(
            parsed.series[0]
                .link(Role::Values)
                .expect("values link")
                .formula(limits)
                .expect("typed values formula"),
            formula
        );
        assert!(parsed.cached_values.iter().any(|value| {
            value.kind == CacheKind::Values
                && value.series == 0
                && value.point == 0
                && value.format == 9
                && value.value == Value::Number(7.5)
        }));
        assert!(parsed.cached_values.iter().any(|value| {
            value.kind == CacheKind::Categories && value.value == Value::Text("Jan".into())
        }));
    }

    #[test]
    fn formula_and_cache_edits_are_bounded_and_transactional() {
        let mut chart = data_chart();
        let original = chart.clone();
        let transposed = Formula::range(CellRef {
            extern_sheet_index: 0,
            first_row: 9,
            last_row: 9,
            first_column: 1,
            last_column: 2,
        })
        .expect("absolute chart range");
        assert!(matches!(
            chart.set_formula(0, Role::Values, transposed),
            Err(Error::UnsafeEdit(_))
        ));
        assert_eq!(chart, original);

        let one_cell = Formula::cell(CellRef {
            extern_sheet_index: 0,
            first_row: 9,
            last_row: 9,
            first_column: 1,
            last_column: 1,
        })
        .expect("absolute chart cell");
        assert!(matches!(
            chart.set_formula(0, Role::Values, one_cell),
            Err(Error::InvalidRecord {
                record_type: BRAI,
                ..
            })
        ));
        assert_eq!(chart, original);

        assert!(matches!(
            chart.set_cache(CacheKind::Values, 0, 0, 0, Value::Number(f64::NAN)),
            Err(Error::InvalidRecord {
                record_type: NUMBER,
                ..
            })
        ));
        assert_eq!(chart, original);
        assert!(matches!(
            chart.set_cache(CacheKind::Values, 0, 99, 0, Value::Blank),
            Err(Error::UnsafeEdit(_))
        ));
        assert_eq!(chart, original);
    }

    #[test]
    fn malformed_formula_and_cache_metadata_are_rejected() {
        let limits = Limits::default();
        let absolute = Formula::cell(CellRef {
            extern_sheet_index: 0,
            first_row: 0,
            last_row: 0,
            first_column: 0,
            last_column: 0,
        })
        .expect("absolute chart cell");
        let mut relative = absolute.tokens().to_vec();
        relative[5..7].copy_from_slice(&0x4000u16.to_le_bytes());
        assert!(Formula::parse(relative, limits).is_err());
        assert!(Formula::parse(vec![0x15], limits).is_err());

        let source = serialize_chart(&data_chart(), limits).expect("chart serialization");
        let mut invalid_section = source.clone();
        let section = ranges(&invalid_section)
            .expect("framed chart")
            .into_iter()
            .find(|value| value.kind == SI_INDEX)
            .expect("SIIndex record");
        invalid_section[section.body_start..section.body_end].copy_from_slice(&0u16.to_le_bytes());
        assert!(matches!(
            parse_chart(&invalid_section, limits),
            Err(Error::InvalidRecord {
                record_type: SI_INDEX,
                ..
            })
        ));

        let mut invalid_number = source.clone();
        let number = ranges(&invalid_number)
            .expect("framed chart")
            .into_iter()
            .find(|value| value.kind == NUMBER)
            .expect("Number cache record");
        invalid_number[number.body_start + 6..number.body_end]
            .copy_from_slice(&f64::NAN.to_le_bytes());
        assert!(matches!(
            parse_chart(&invalid_number, limits),
            Err(Error::InvalidRecord {
                record_type: NUMBER,
                ..
            })
        ));

        let mut invalid_label = source;
        let label = ranges(&invalid_label)
            .expect("framed chart")
            .into_iter()
            .find(|value| value.kind == LABEL)
            .expect("Label cache record");
        invalid_label[label.body_start + 6..label.body_start + 8]
            .copy_from_slice(&4u16.to_le_bytes());
        assert!(matches!(
            parse_chart(&invalid_label, limits),
            Err(Error::InvalidRecord {
                record_type: LABEL,
                ..
            })
        ));
    }

    #[test]
    fn chart_data_patch_preserves_unknown_records_and_refuses_graph_changes() {
        let limits = Limits::default();
        let base = serialize_chart(&data_chart(), limits).expect("chart serialization");
        let eof = ranges(&base)
            .expect("framed chart")
            .last()
            .expect("chart EOF")
            .start;
        let opaque = record(0x7777, b"opaque chart extension").expect("opaque record framing");
        let mut source = base;
        source.splice(eof..eof, opaque.iter().copied());
        let old = parse_chart(&source, limits).expect("chart with opaque record");
        let mut updated = old.clone();
        let formula = Formula::range(CellRef {
            extern_sheet_index: 0,
            first_row: 6,
            last_row: 7,
            first_column: 1,
            last_column: 1,
        })
        .expect("absolute chart range");
        updated
            .set_formula(0, Role::Values, formula.clone())
            .expect("same-cardinality formula edit");
        updated
            .set_cache(CacheKind::Values, 0, 0, 11, Value::Number(8.25))
            .expect("existing cache edit");

        let patched = patch_chart_data(&source, &old, &updated, limits).expect("safe patch");
        let opaque_range = ranges(&patched)
            .expect("patched chart framing")
            .into_iter()
            .find(|value| value.kind == 0x7777)
            .expect("opaque record survives");
        assert_eq!(
            &patched[opaque_range.start..opaque_range.end],
            opaque.as_slice()
        );
        let parsed = parse_chart(&patched, limits).expect("patched chart parses");
        assert_eq!(parsed.unknown_records, old.unknown_records);
        assert_eq!(
            parsed.series[0]
                .link(Role::Values)
                .expect("values link")
                .formula(limits)
                .expect("patched formula"),
            formula
        );
        assert!(parsed.cached_values.iter().any(|value| {
            value.kind == CacheKind::Values
                && value.format == 11
                && value.value == Value::Number(8.25)
        }));

        let mut graph_change = updated;
        graph_change.title = Some("unsafe graph edit".into());
        assert!(matches!(
            patch_chart_data(&source, &old, &graph_change, limits),
            Err(Error::UnsafeEdit(_))
        ));
    }

    #[test]
    fn embedded_formula_edit_is_an_atomic_refusal() {
        let limits = Limits::default();
        let original = build_workbook_fixture(Chart::default(), limits).unwrap();
        let mut editor = Editor::open(original.clone(), limits).unwrap();
        let formula = Formula::cell(CellRef {
            extern_sheet_index: 0,
            first_row: 0,
            last_row: 0,
            first_column: 0,
            last_column: 0,
        })
        .expect("absolute chart cell");
        assert!(matches!(
            editor.set_formula(
                Selector::Embedded {
                    sheet: "Sheet1",
                    index: 0,
                },
                0,
                Role::Values,
                formula,
            ),
            Err(Error::Graph(
                litchi_ograph::Error::UnsupportedMutation { .. }
            ))
        ));
        assert_eq!(editor.finish().unwrap(), original);
    }

    #[test]
    fn group_lines_and_drop_bars_round_trip_without_collapsing_kinds() {
        let line_format = format::Line {
            color: [1, 2, 3, 4],
            pattern: 1,
            weight: 2,
            flags: 0,
            color_index: 8,
        };
        let area_format = format::Area {
            foreground: [5, 6, 7, 8],
            background: [9, 10, 11, 12],
            pattern: 1,
            flags: 0,
            foreground_index: 9,
            background_index: 10,
        };
        let mut chart = Chart::default();
        chart.groups[0].lines.extend([
            group::Line {
                kind: line::Kind::HighLow,
                format: line_format,
            },
            group::Line {
                kind: line::Kind::Series,
                format: line_format,
            },
        ]);
        chart.groups[0].drop_bars.push(group::DropBar {
            gap: group::Gap::new(257).expect("valid DropBar gap"),
            line: line_format,
            area: area_format,
        });

        let bytes = serialize_chart(&chart, Limits::default()).unwrap();
        let kinds = ranges(&bytes)
            .expect("framed chart")
            .into_iter()
            .map(|record| record.kind)
            .collect::<Vec<_>>();
        for (index, kind) in kinds.iter().enumerate() {
            if *kind == CRT_LINE {
                assert_eq!(kinds.get(index + 1), Some(&LINE_FORMAT));
            }
        }
        assert!(
            kinds
                .windows(5)
                .any(|window| { window == [DROP_BAR, BEGIN, LINE_FORMAT, AREA_FORMAT, END] })
        );
        let parsed = parse_chart(&bytes, Limits::default()).unwrap();
        assert_eq!(parsed.kind(), Kind::Stock);
        assert_eq!(parsed.groups[0].lines, chart.groups[0].lines);
        assert_eq!(parsed.groups[0].drop_bars, chart.groups[0].drop_bars);
    }

    #[test]
    fn sheet_properties_and_group_numeric_domains_are_strict() {
        let limits = Limits::default();
        let mut chart = Chart::default();
        for blank in 0..=2 {
            chart.sheet_properties = blank << 16;
            chart.validate(limits).expect("valid blank mode");
        }
        chart.sheet_properties = 3 << 16;
        assert!(chart.validate(limits).is_err());
        chart.sheet_properties = 1 << 2;
        chart
            .validate(limits)
            .expect("fNotSizeWith is a defined ShtProps bit");
        chart.sheet_properties = 1 << 4;
        assert!(chart.validate(limits).is_err());
        chart.sheet_properties = (1 << 4) | (1 << 3);
        chart.validate(limits).expect("paired plot-area flags");

        chart.groups[0].kind = GroupKind::Bar {
            overlap: 101,
            gap: 150,
            flags: 0,
        };
        assert!(chart.validate(limits).is_err());
        chart.groups[0].kind = GroupKind::Scatter {
            bubble_size_percent: 301,
            bubble_size_type: 1,
            flags: 0,
        };
        assert!(chart.validate(limits).is_err());
        chart.groups[0].kind = GroupKind::Scatter {
            bubble_size_percent: 100,
            bubble_size_type: 3,
            flags: 0,
        };
        assert!(chart.validate(limits).is_err());
    }

    #[test]
    fn malformed_nesting_formula_and_axis_are_rejected() {
        let mut bytes = record(BOF, &chart_bof()).unwrap();
        bytes.extend(record(CHART, &[0; 16]).unwrap());
        bytes.extend(record(END, &[]).unwrap());
        bytes.extend(record(EOF, &[]).unwrap());
        assert!(parse_chart(&bytes, Limits::default()).is_err());
        let mut chart = Chart::default();
        chart.series.push(Series {
            links: vec![DataLink {
                role: Role::Values,
                source: Source::Cells,
                unlinked_number_format: false,
                number_format: 0,
                formula_tokens: vec![0x3b, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                references: vec![CellRef {
                    extern_sheet_index: 2,
                    first_row: 0,
                    last_row: 0,
                    first_column: 300,
                    last_column: 300,
                }],
            }],
            ..Default::default()
        });
        assert!(chart.validate(Limits::default()).is_err());
    }

    #[test]
    fn referenced_sheet_removal_and_noncontiguous_reorder_are_rejected_atomically() {
        let mut extern_sheet = 1u16.to_le_bytes().to_vec();
        extern_sheet.extend(0u16.to_le_bytes());
        extern_sheet.extend(0u16.to_le_bytes());
        extern_sheet.extend(2u16.to_le_bytes());
        let internal = HashSet::from([0u16]);
        assert!(
            remap_extern_sheet(&extern_sheet, &internal, &[Some(0), None, Some(1)], None).is_err()
        );
        assert!(
            remap_extern_sheet(&extern_sheet, &internal, &[Some(0), Some(2), Some(1)], None)
                .is_ok()
        );
    }
}
