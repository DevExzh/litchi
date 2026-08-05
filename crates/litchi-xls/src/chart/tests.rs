use std::collections::HashSet;
use std::sync::Arc;

use super::codec::{parse_chart, serialize_chart};
use super::package::{GENERATED_CHART_OBJECT_ID, chart_bof};
use super::package::{build_workbook_fixture, ranges, remap_extern_sheet};
use super::wire::{
    AREA_FORMAT, BEGIN, BOF, CHART, CRT_LINE, DROP_BAR, END, EOF, LINE_FORMAT, record,
};
use super::*;
use crate::Error;

#[cfg(test)]
mod tests {
    use super::*;

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
            cache_index: 0,
            row: 0,
            column: 0,
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
        chart.cached_values.push(Cache {
            cache_index: 3,
            row: 4,
            column: 5,
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
            value.cache_index == 3
                && value.row == 4
                && value.column == 5
                && value.format == 9
                && value.value == Value::Blank
        }));
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
