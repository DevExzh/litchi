//! Record-order characterization of the chart-substream parser.
//!
//! Each base stream below is a chart the parser accepts. The matrix edits
//! the stream at every record index (delete, duplicate, swap with the next
//! record, insert a Begin or an End, cut the stream short, delete or repeat
//! a whole collection) and compares what the parser reports with the frozen
//! table in `tests/data/parse_mutation_matrix.txt`. For a stream with more
//! than one defect the table also pins which defect is reported first.
//!
//! A table row is `<stream> <record index> <record kind> <edit> => <result>`.
//! `ok same` means the parsed chart has the same summary as the unedited
//! stream; `framing …;` precedes the parser result when the substream
//! framing check also refuses the edited bytes.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    reason = "test assertions panic by design"
)]

use super::super::super::model::{
    Ai, Binding, Cache, Context, Count, Label, Legend, Link, Model, Owner, Role, RowCol, Series,
    Source, Value,
};
use super::super::super::{Kind as ChartKind, Ref, axis, cache, format, group};
use super::wire::{AXES_USED, AXIS_PARENT, BEGIN, CHART_FORMAT, DIMENSIONS, END, TEXT, UNITS};
use super::{encode, parse};
use crate::{Error, Limits};
use litchi_biff::{Encoder, Kind as RecordKind, Records};
use std::collections::BTreeSet;
use std::fmt::Write as _;

const EXPECTED: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/data/parse_mutation_matrix.txt"
));

type Wire = Vec<(RecordKind, Vec<u8>)>;

fn limits() -> Limits {
    Limits::default().validate().expect("default limits")
}

fn line_format() -> format::Line {
    format::Line {
        color: [1, 2, 3, 0],
        pattern: 0,
        weight: 0,
        flags: 0,
        color_index: 8,
    }
}

fn area_format() -> format::Area {
    format::Area {
        foreground: [4, 5, 6, 0],
        background: [7, 8, 9, 0],
        pattern: 1,
        flags: 0,
        foreground_index: 9,
        background_index: 10,
    }
}

/// Adds the members shared by every base chart: a category and a value
/// axis with lines, a line group with a high-low line and one drop-bar
/// collection, a legend, and loose formats.
fn decorate(chart: &mut Model) {
    let mut category = axis::Axis::new(axis::Kind::Category);
    category.tick = Some(axis::Tick {
        major: 2,
        minor: 0,
        label: 3,
        background: 1,
        color: [0; 4],
        flags: 0x23,
    });
    category.lines.push(axis::Line {
        kind: axis::LineKind::Axis,
        format: line_format(),
    });
    chart.add_axis(category).expect("category axis");
    let mut value = axis::Axis::new(axis::Kind::Value);
    value.scale = Some(axis::Scale {
        min: 0.0,
        max: 10.0,
        major: 2.0,
        minor: 1.0,
        crossing: 0.0,
        flags: 0x1F,
    });
    value.lines.push(axis::Line {
        kind: axis::LineKind::MajorGrid,
        format: line_format(),
    });
    chart.add_axis(value).expect("value axis");
    {
        let mut groups = chart.groups_mut();
        let group = groups.first_mut().expect("default line group");
        group.family = super::super::super::model::Family::Line { flags: 1 };
        group.lines.push(group::Line {
            kind: crate::record::line::Kind::HighLow,
            format: line_format(),
        });
        group.drop_bars.push(group::DropBar {
            gap: group::Gap::new(20).expect("bounded gap"),
            line: line_format(),
            area: area_format(),
        });
    }
    chart.set_legend(Some(Legend {
        x: 1,
        y: 2,
        width: 3,
        height: 4,
        position: 3,
        spacing: 1,
        flags: 0x1F,
    }));
    chart
        .add_format(format::Format::Data {
            point: 0xFFFF,
            series: 0,
            flags: 0,
        })
        .expect("data format");
    chart
        .add_format(format::Format::Line(line_format()))
        .expect("line format");
    chart
        .add_format(format::Format::Area(area_format()))
        .expect("area format");
}

fn auxiliary(context: Context, error_bar: bool) -> Series {
    let parent = crate::record::series::Parent::try_new(1).expect("parent series");
    let mut series = Series::new(context);
    series.owner = if error_bar {
        Owner::ErrorBar {
            parent,
            data: [0; 14],
        }
    } else {
        Owner::Trend {
            parent,
            data: [0; 28],
        }
    };
    series
}

fn excel_context() -> Context {
    Context::excel().with_external_sheets(1)
}

fn excel_chart() -> Model {
    let context = excel_context();
    let mut chart = Model::new(context).expect("new chart");
    let mut series = Series::new(context);
    series.category_count = Count::new(2).expect("count");
    series.value_count = Count::new(2).expect("count");
    series.ai = Ai::new(
        Binding::new(
            Link::excel(Role::Name, Source::Automatic, Vec::new()),
            Some("FY26".into()),
        ),
        Binding::new(
            Link::excel(
                Role::Values,
                Source::Cells,
                vec![0x1B, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
            ),
            None,
        ),
        Binding::new(
            Link::excel(Role::Categories, Source::Automatic, Vec::new()),
            None,
        ),
        Binding::new(
            Link::excel(Role::Bubbles, Source::Automatic, Vec::new()),
            None,
        ),
    )
    .expect("canonical AI roles");
    chart.add_series(series).expect("series");
    chart
        .add_series(auxiliary(context, false))
        .expect("trend series");
    chart
        .add_series(auxiliary(context, true))
        .expect("error-bar series");
    chart
        .add_cache(Cache::excel(
            cache::Index::Values,
            0,
            0,
            cache::Xf::new(0),
            Value::Number(42.5),
        ))
        .expect("numeric cache");
    chart
        .add_cache(Cache::excel(
            cache::Index::Values,
            1,
            0,
            cache::Xf::new(2),
            Value::Text("safe".into()),
        ))
        .expect("text cache");
    chart
        .add_cache(Cache::excel(
            cache::Index::Categories,
            2,
            0,
            cache::Xf::new(3),
            Value::Blank,
        ))
        .expect("blank cache");
    decorate(&mut chart);
    chart
        .add_label(Label {
            kind: TEXT,
            data: vec![0; 32],
        })
        .expect("label");
    chart
}

/// A Graph chart; `datasheet` selects the layout that starts with a
/// complete datasheet collection holding the cached cells.
fn graph_chart(datasheet: bool) -> Model {
    let context = Context::graph();
    let mut chart = Model::new(context).expect("new chart");
    let row = RowCol::new(1).expect("series row");
    let mut series = Series::new(context);
    series.category_count = Count::new(2).expect("count");
    series.value_count = Count::new(2).expect("count");
    series.ai = Ai::new(
        Binding::new(
            Link::graph(Role::Name, Source::Literal, row),
            Some("Sales".into()),
        ),
        Binding::new(Link::graph(Role::Values, Source::Literal, row), None),
        Binding::new(
            Link::graph(Role::Categories, Source::Literal, RowCol::ZERO),
            None,
        ),
        Binding::new(
            Link::graph(Role::Bubbles, Source::Automatic, RowCol::ZERO),
            None,
        ),
    )
    .expect("Graph AI roles");
    chart.add_series(series).expect("series");
    chart
        .add_series(auxiliary(context, false))
        .expect("trend series");
    chart
        .add_cache(Cache::graph(
            row,
            RowCol::ZERO,
            cache::Ifmt::new(0),
            Value::Text("Sales".into()),
        ))
        .expect("name cache");
    chart
        .add_cache(Cache::graph(
            row,
            RowCol::new(1).expect("column"),
            cache::Ifmt::new(0),
            Value::Number(10.0),
        ))
        .expect("value cache");
    chart
        .add_cache(Cache::graph(
            row,
            RowCol::new(2).expect("column"),
            cache::Ifmt::new(0),
            Value::Blank,
        ))
        .expect("blank cache");
    decorate(&mut chart);
    // The encoder writes a title as a Graph attached label only.
    chart.set_title(Some("Title".into()));
    chart.graph_authoring_profile = datasheet;
    chart
}

fn wire(mut chart: Model) -> Wire {
    chart.authoring_proven = true;
    let bytes = encode(&chart, limits()).expect("base chart encoding");
    Records::new(&bytes, limits().biff)
        .expect("valid BIFF limits")
        .map(|item| {
            let record = item.expect("encoder framing");
            (record.kind(), record.payload().to_vec())
        })
        .collect()
}

/// Repeats the axis-parent collection as a secondary axis group, which the
/// encoder itself never emits. The copy gets axis-group index one and the
/// next chart-group order, and `AxesUsed` is raised to two.
fn with_secondary_parent(mut base: Wire) -> Wire {
    let start = position(&base, AXIS_PARENT);
    let end = collection_end(&base, start).expect("AxisParent collection");
    let mut copy = base[start..=end].to_vec();
    copy[0].1[0] = 1;
    for (kind, data) in &mut copy {
        if *kind == CHART_FORMAT {
            data[18] = 1;
        }
    }
    for (kind, data) in &mut base {
        if *kind == AXES_USED {
            data[0] = 2;
        }
    }
    base.splice(end + 1..end + 1, copy);
    base
}

fn position(wire: &Wire, target: RecordKind) -> usize {
    wire.iter()
        .position(|(kind, _)| *kind == target)
        .expect("record kind is present")
}

/// Index of the End that closes the collection owned by record `index`,
/// when that record is followed by a Begin.
fn collection_end(wire: &Wire, index: usize) -> Option<usize> {
    if wire.get(index + 1)?.0 != BEGIN || wire[index].0 == BEGIN {
        return None;
    }
    let mut depth = 0usize;
    for (offset, (kind, _)) in wire.iter().enumerate().skip(index + 1) {
        if *kind == BEGIN {
            depth += 1;
        } else if *kind == END {
            depth -= 1;
            if depth == 0 {
                return Some(offset);
            }
        }
    }
    None
}

fn bytes(wire: &Wire) -> Vec<u8> {
    let mut out = Encoder::new(limits().biff).expect("valid BIFF limits");
    for (kind, data) in wire {
        out.push(*kind, data).expect("record replay");
    }
    out.finish()
}

fn summary(chart: &Model) -> String {
    let mut out = String::new();
    write!(out, "series=[").unwrap();
    for series in chart.series() {
        let texts = series
            .ai
            .ordered()
            .into_iter()
            .filter(|binding| binding.text().is_some())
            .count();
        match &series.owner {
            Owner::Group(group) => write!(out, "g{}t{texts} ", group.get()),
            Owner::Trend { parent, .. } => write!(out, "trend{}t{texts} ", parent.series().get()),
            Owner::ErrorBar { parent, .. } => write!(out, "err{}t{texts} ", parent.series().get()),
        }
        .unwrap();
    }
    write!(out, "] groups=[").unwrap();
    for group in chart.groups() {
        write!(
            out,
            "p{}o{}l{}d{} ",
            group.parent.get(),
            group.order.get(),
            group.lines.len(),
            group.drop_bars.len()
        )
        .unwrap();
    }
    write!(out, "] axes=[").unwrap();
    for axis in chart.axes() {
        write!(
            out,
            "p{}{:?}s{}t{}l{} ",
            axis.parent.get(),
            axis.kind,
            u8::from(axis.scale.is_some()),
            u8::from(axis.tick.is_some()),
            axis.lines.len()
        )
        .unwrap();
    }
    write!(
        out,
        "] parents={} title={:?} legend={} plot={} formats={} labels={} unknown={} caches={}",
        chart.parents().len(),
        chart.title(),
        u8::from(chart.legend.is_some()),
        u8::from(chart.props.plot_area),
        chart.formats().len(),
        chart.labels().len(),
        chart.unknown().len(),
        chart.caches().len()
    )
    .unwrap();
    out
}

/// One table cell: what the framing check and the parser say about `wire`.
/// The parser runs even when the framing check refuses the stream, so the
/// table also pins the parser's own handling of a displaced BOF or EOF.
fn outcome(wire: &Wire, kind: ChartKind, context: Context, base: Option<&str>) -> String {
    let bytes = bytes(wire);
    let mut out = String::new();
    if let Err(error) = Ref::open(&bytes, limits()) {
        match error {
            Error::InvalidChart { offset, reason } => {
                write!(out, "framing {offset} {reason:?}; ").unwrap();
            },
            other => write!(out, "framing {other:?}; ").unwrap(),
        }
    }
    match parse(
        Ref::from_validated(&bytes, kind, 0, limits()),
        context,
        limits(),
    ) {
        Ok(chart) => {
            let summary = summary(&chart);
            if base == Some(summary.as_str()) {
                out.push_str("ok same");
            } else {
                write!(out, "ok {summary}").unwrap();
            }
        },
        Err(Error::InvalidChart { offset, reason }) => {
            write!(out, "invalid {offset} {reason:?}").unwrap();
        },
        Err(other) => write!(out, "error {other:?}").unwrap(),
    }
    out
}

fn matrix() -> String {
    let streams: [(&str, ChartKind, Context, Wire); 5] = [
        (
            "excel",
            ChartKind::Excel,
            excel_context(),
            wire(excel_chart()),
        ),
        (
            "excel2",
            ChartKind::Excel,
            excel_context(),
            with_secondary_parent(wire(excel_chart())),
        ),
        (
            "graph",
            ChartKind::Graph,
            Context::graph(),
            wire(graph_chart(false)),
        ),
        (
            "sheet",
            ChartKind::Graph,
            Context::graph(),
            wire(graph_chart(true)),
        ),
        (
            "sheet2",
            ChartKind::Graph,
            Context::graph(),
            with_secondary_parent(wire(graph_chart(true))),
        ),
    ];
    let mut out = String::new();
    let sheet = &streams[3].3;
    let units = position(sheet, UNITS);
    let mut early = sheet.clone();
    let dimensions = early.remove(position(sheet, DIMENSIONS));
    early.insert(units + 1, dimensions);
    writeln!(
        out,
        "sheet Dimensions-before-Chart => {}",
        outcome(&early, ChartKind::Graph, Context::graph(), None)
    )
    .unwrap();
    early.splice(
        units + 1..units + 1,
        [(BEGIN, Vec::new()), (END, Vec::new())],
    );
    writeln!(
        out,
        "sheet Dimensions-before-Chart-after-empty-collection => {}",
        outcome(&early, ChartKind::Graph, Context::graph(), None)
    )
    .unwrap();
    for (name, kind, context, base) in &streams {
        let base_outcome = outcome(base, *kind, *context, None);
        writeln!(out, "{name} base => {base_outcome}").unwrap();
        let base_summary = base_outcome.strip_prefix("ok ");
        for index in 0..base.len() {
            let record = base[index].0.get();
            let mut cases: Vec<(&str, Wire)> = Vec::new();
            let mut deleted = base.clone();
            deleted.remove(index);
            cases.push(("delete", deleted));
            let mut duplicated = base.clone();
            duplicated.insert(index, base[index].clone());
            cases.push(("duplicate", duplicated));
            if index + 1 < base.len() {
                let mut swapped = base.clone();
                swapped.swap(index, index + 1);
                cases.push(("swap", swapped));
            }
            let mut begin = base.clone();
            begin.insert(index + 1, (BEGIN, Vec::new()));
            cases.push(("begin", begin));
            let mut end = base.clone();
            end.insert(index + 1, (END, Vec::new()));
            cases.push(("end", end));
            if index + 1 < base.len() {
                let mut truncated = base[..=index].to_vec();
                truncated.push(base[base.len() - 1].clone());
                cases.push(("truncate", truncated));
            }
            if let Some(end) = collection_end(base, index) {
                let mut removed = base.clone();
                removed.drain(index..=end);
                cases.push(("delete-collection", removed));
                let mut repeated = base.clone();
                repeated.splice(end + 1..end + 1, base[index..=end].iter().cloned());
                cases.push(("duplicate-collection", repeated));
            }
            for (mutation, wire) in &cases {
                writeln!(
                    out,
                    "{name} {index:03} {record:04X} {mutation} => {}",
                    outcome(wire, *kind, *context, base_summary)
                )
                .unwrap();
            }
        }
    }
    out
}

#[test]
fn single_record_edits_report_the_frozen_outcome() {
    let actual = matrix();
    let mut differences = Vec::new();
    let mut expected = EXPECTED.lines();
    let mut observed = actual.lines();
    loop {
        match (expected.next(), observed.next()) {
            (None, None) => break,
            (want, got) if want == got => {},
            (want, got) => differences.push(format!("expected {want:?}\n  actual {got:?}")),
        }
    }
    assert!(
        differences.is_empty(),
        "{} rows differ; first rows:\n{}",
        differences.len(),
        differences
            .iter()
            .take(20)
            .cloned()
            .collect::<Vec<_>>()
            .join("\n")
    );
}

/// The matrix is only useful while it reaches the ordered checks that run
/// before record dispatch and the checks that run when a collection closes.
#[test]
fn frozen_table_reaches_the_ordered_checks() {
    let reasons: BTreeSet<&str> = EXPECTED
        .lines()
        .filter_map(|line| {
            line.split_once("=> ")?
                .1
                .split_once("invalid ")?
                .1
                .split_once(' ')
        })
        .map(|(_, reason)| reason.trim_matches('"'))
        .collect();
    for reason in REACHED {
        assert!(
            reasons.contains(reason),
            "table no longer reaches {reason:?}"
        );
    }
    for (name, _) in EXPECTED
        .lines()
        .filter_map(|line| line.split_once(" base => "))
    {
        let accepted = EXPECTED
            .lines()
            .filter(|line| line.starts_with(&format!("{name} ")) && line.contains("=> ok"))
            .count();
        let rejected = EXPECTED
            .lines()
            .filter(|line| line.starts_with(&format!("{name} ")) && line.contains("invalid "))
            .count();
        assert!(accepted > 1 && rejected > 1, "{name} has a one-sided table");
    }
}

const REACHED: &[&str] = &[
    // Checks that run before record dispatch, in source order.
    "chart substream does not start with BOF",
    "line owner is not followed immediately by LineFormat",
    "DropBar Begin is not followed by LineFormat",
    "DropBar LineFormat is not followed by AreaFormat",
    "SerParent is not followed immediately by SerAuxTrend or SerAuxErrBar",
    "Series must begin with exactly four ordered AI bindings",
    "AxisParent Begin is not followed immediately by Pos",
    "ChartFormat Begin is not followed by a chart family",
    "chart family is not followed immediately by CrtLink",
    "BOFDatasheet is not followed immediately by Begin",
    "attached Text is not followed immediately by Begin",
    "collection-owning record is not followed immediately by Begin",
    // Checks that run when a Begin or an End is read.
    "datasheet collection cannot contain nested Begin",
    "Begin record has no chart-level collection owner",
    "End record has no matching Begin",
    "datasheet is missing Orient or mandatory window/selection records",
    "chart contains more than one title label",
    "Series requires exactly four AI bindings and one SerToCrt",
    "ChartFormat collection is missing its family or required Excel CrtLink",
    "AxisParent requires Pos and one through four ChartFormat collections",
    "CHARTFOMATS is missing or misorders a mandatory collection",
    // Checks that run after the last record.
    "chart Begin/End collections are unbalanced",
    "Datasheet or attached-label collection is incomplete",
    "Graph chart sheet is missing a mandatory Datasheet or chart record",
    "chart collection owner is missing its complete Begin/End collection",
    "chart is missing Chart, Scl, PlotGrowth, or Dimensions",
    "AxesUsed does not match the AxisParent collection count",
    "Excel SERIESDATA does not contain exactly three SIIndex sections",
];
