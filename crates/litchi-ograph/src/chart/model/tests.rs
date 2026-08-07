#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic by design"
)]

use super::super::{Ref, Stream, axis, cache, format, group};
use super::{
    Ai, Binding, Cache, Chart, Context, Count, DataKind, Group, GroupId, Link, Order, Owner, Props,
    Role, Series, Source, Value, ValueRef, XlValue,
};
use crate::chart::RowCol;
use crate::{Error, Limits};
use litchi_biff::{Encoder, Kind as RecordKind, Records};

const UNKNOWN: RecordKind = RecordKind::from_wire(0x7777);

fn count(value: u16) -> Count {
    Count::new(value).expect("bounded fixture count")
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

fn fixture(mut chart: Chart) -> Stream {
    chart.authoring_proven = true;
    chart.encode().expect("internal parser fixture")
}

fn omit(stream: &Stream, target: RecordKind) -> Vec<u8> {
    let mut out = Encoder::new();
    for item in Records::new(stream.as_bytes()) {
        let record = item.expect("valid fixture record");
        if record.kind() != target {
            out.push_ref(record).expect("record replay");
        }
    }
    out.finish()
}

fn excel_input(bytes: &[u8]) -> Ref<'_> {
    Ref::open(bytes).expect("well-framed chart rewrite")
}

fn excel_chart() -> Chart {
    let context = Context::excel().with_external_sheets(1);
    let mut chart = Chart::new(context).expect("new chart");
    let mut series = Series::new(context);
    series.category_kind = DataKind::Text;
    series.category_count = count(2);
    series.value_count = count(2);
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
            cache::Index::Values,
            2,
            0,
            cache::Xf::new(3),
            Value::Blank,
        ))
        .expect("blank cache");
    chart
}

#[test]
fn fresh_encode_refuses_while_internal_fixture_replay_moves_original_stream() {
    assert!(matches!(
        excel_chart().encode(),
        Err(Error::UnsupportedAuthoring { .. })
    ));
    let stream = fixture(excel_chart());
    let pointer = stream.as_bytes().as_ptr();
    let parsed =
        Chart::open(stream, Context::excel().with_external_sheets(1)).expect("semantic parse");
    assert!(parsed.is_pristine());
    assert_eq!(parsed.title(), None);
    assert_eq!(parsed.series().len(), 1);
    assert_eq!(parsed.caches().len(), 3);
    assert!(matches!(parsed.caches()[2].value(), ValueRef::Blank));
    let replay = parsed.encode().expect("exact replay");
    assert_eq!(replay.as_bytes().as_ptr(), pointer);
}

#[test]
fn graph_and_excel_links_have_distinct_checked_wire_grammars() {
    let row_col = RowCol::new(7).expect("Graph coordinate");
    let mut graph = Chart::new(Context::graph()).expect("Graph chart");
    let mut series = Series::new(Context::graph());
    series.ai.replace(Binding::new(
        Link::graph(Role::Values, Source::Literal, row_col),
        None,
    ));
    graph.add_series(series).expect("Graph series");
    graph
        .add_cache(Cache::graph(
            RowCol::new(1).expect("row"),
            RowCol::new(2).expect("column"),
            cache::Ifmt::new(4),
            Value::Blank,
        ))
        .expect("Graph blank");
    let stream = fixture(graph);
    let parsed = Chart::open(stream, Context::graph()).expect("Graph parse");
    assert!(matches!(
        parsed.series()[0].ai.get(Role::Values).link(),
        Link::Graph { .. }
    ));
    assert!(matches!(parsed.caches()[0], Cache::Graph { .. }));

    let mut wrong = Chart::new(Context::graph()).expect("Graph chart");
    let mut wrong_series = Series::new(Context::graph());
    wrong_series.ai.replace(Binding::new(
        Link::excel(Role::Values, Source::Automatic, Vec::new()),
        None,
    ));
    assert!(matches!(
        wrong.add_series(wrong_series),
        Err(Error::InvalidModel { field: "link", .. })
    ));
}

#[test]
fn parsed_mutation_is_refused_and_unknown_order_replays_exactly() {
    let original = fixture(excel_chart());
    let mut out = Encoder::new();
    for item in Records::new(original.as_bytes()) {
        let record = item.expect("valid record");
        if record.kind() == super::super::EOF {
            out.push(UNKNOWN, &[9, 8, 7]).expect("unknown record");
        }
        out.push_ref(record).expect("record replay");
    }
    let bytes = out.finish();
    let pointer = bytes.as_ptr();
    let stream = Stream::open(bytes).expect("raw chart");
    let parsed =
        Chart::open(stream, Context::excel().with_external_sheets(1)).expect("semantic chart");
    assert_eq!(parsed.unknown().len(), 1);
    assert_eq!(parsed.unknown()[0].kind(), UNKNOWN);
    let replay = parsed.encode().expect("exact replay");
    assert_eq!(replay.as_bytes().as_ptr(), pointer);
    let mut parsed =
        Chart::open(replay, Context::excel().with_external_sheets(1)).expect("semantic chart");
    parsed.set_title(Some("Changed".into()));
    assert!(matches!(parsed.encode(), Err(Error::UnsafeEdit { .. })));
}

#[test]
fn rejects_context_mismatch_tighter_replay_limit_and_invalid_properties() {
    let stream = fixture(excel_chart());
    let bytes = stream.as_bytes().len();
    assert!(matches!(
        Chart::open(stream, Context::graph()),
        Err(Error::InvalidModel {
            field: "context",
            ..
        })
    ));

    let stream = fixture(excel_chart());
    let chart =
        Chart::open(stream, Context::excel().with_external_sheets(1)).expect("parsed chart");
    let mut limits = Limits::default();
    limits.biff.max_output_bytes = bytes.saturating_sub(1);
    assert!(matches!(
        chart.encode_with(limits),
        Err(Error::LimitExceeded {
            resource: "output bytes",
            ..
        })
    ));

    let mut valid_blank_mode = Chart::new(Context::excel()).expect("chart");
    valid_blank_mode.set_props(Props {
        flags: 2 | (2 << 16),
        plot_area: true,
    });
    valid_blank_mode.authoring_proven = true;
    assert!(valid_blank_mode.encode().is_ok());

    let mut reserved_bit = Chart::new(Context::excel()).expect("chart");
    reserved_bit.set_props(Props {
        flags: 1 << 2,
        plot_area: true,
    });
    reserved_bit.authoring_proven = true;
    assert!(matches!(
        reserved_bit.encode(),
        Err(Error::InvalidModel { .. })
    ));

    let mut reserved = Chart::new(Context::excel()).expect("chart");
    reserved.set_props(Props {
        flags: 1 << 5,
        plot_area: true,
    });
    reserved.authoring_proven = true;
    assert!(matches!(reserved.encode(), Err(Error::InvalidModel { .. })));

    let mut dependency = Chart::new(Context::excel()).expect("chart");
    dependency.set_props(Props {
        flags: 1 << 4,
        plot_area: true,
    });
    dependency.authoring_proven = true;
    assert!(matches!(
        dependency.encode(),
        Err(Error::InvalidModel { .. })
    ));
}

#[test]
fn add_methods_enforce_authoring_limits_before_growth() {
    let limits = Limits {
        max_series: 1,
        max_groups: 1,
        max_axes: 1,
        max_cached_values: 1,
        ..Limits::default()
    };
    let mut chart = Chart::new_with(Context::excel(), limits).expect("bounded chart");
    chart
        .add_series(Series::new(Context::excel()))
        .expect("first series");
    assert!(matches!(
        chart.add_series(Series::new(Context::excel())),
        Err(Error::LimitExceeded {
            resource: "series count",
            ..
        })
    ));
    assert!(matches!(
        chart.add_group(Group::line()),
        Err(Error::LimitExceeded {
            resource: "group count",
            ..
        })
    ));
    chart
        .add_axis(axis::Axis::new(axis::Kind::Category))
        .expect("first axis");
    assert!(chart.add_axis(axis::Axis::new(axis::Kind::Value)).is_err());
}

#[test]
fn group_lines_and_drop_bars_emit_mandatory_owned_formats() {
    let mut chart = Chart::new(Context::excel()).expect("chart");
    let mut groups = chart.groups_mut();
    let group = groups.first_mut().expect("default line group");
    group
        .lines
        .try_reserve_exact(1)
        .expect("line fixture allocation");
    group.lines.push(group::Line {
        kind: crate::record::line::Kind::HighLow,
        format: line_format(),
    });
    group
        .drop_bars
        .try_reserve_exact(1)
        .expect("DropBar fixture allocation");
    group.drop_bars.push(group::DropBar {
        gap: group::Gap::new(20).expect("bounded gap"),
        line: line_format(),
        area: area_format(),
    });

    let stream = fixture(chart);
    let kinds = stream
        .records()
        .map(|record| record.expect("valid record").kind())
        .collect::<Vec<_>>();
    let crt = kinds
        .iter()
        .position(|kind| *kind == RecordKind::from_wire(0x101C))
        .expect("CrtLine");
    assert_eq!(kinds.get(crt + 1), Some(&RecordKind::from_wire(0x1007)));
    let drop = kinds
        .iter()
        .position(|kind| *kind == RecordKind::from_wire(0x103D))
        .expect("DropBar");
    assert_eq!(
        kinds.get(drop..drop + 5),
        Some(
            [0x103D, 0x1033, 0x1007, 0x100A, 0x1034]
                .map(RecordKind::from_wire)
                .as_slice()
        )
    );

    let parsed = Chart::open(stream, Context::excel()).expect("parse");
    let group = parsed.groups().first().expect("group");
    assert_eq!(group.lines.len(), 1);
    assert_eq!(group.drop_bars.len(), 1);
    assert_eq!(group.drop_bars[0].gap.get(), 20);
}

#[test]
fn rejects_missing_collection_begin_and_nesting_over_limit() {
    let stream = fixture(excel_chart());
    let mut out = Encoder::new();
    let mut after_series = false;
    let mut removed = false;
    for item in Records::new(stream.as_bytes()) {
        let record = item.expect("valid record");
        if after_series && record.kind() == RecordKind::from_wire(0x1033) {
            removed = true;
            after_series = false;
            continue;
        }
        after_series = record.kind() == RecordKind::from_wire(0x1003);
        out.push_ref(record).expect("record replay");
    }
    assert!(removed);
    let malformed = out.finish();
    let input = Ref::open(&malformed).expect("raw boundaries remain valid");
    assert!(matches!(
        Chart::parse(input, Context::excel().with_external_sheets(1)),
        Err(Error::InvalidChart {
            reason: "collection-owning record is not followed immediately by Begin",
            ..
        })
    ));

    assert!(matches!(
        Chart::parse_with(
            stream.as_ref(),
            Context::excel().with_external_sheets(1),
            Limits {
                max_nesting: 2,
                ..Limits::default()
            }
        ),
        Err(Error::LimitExceeded {
            resource: "chart nesting",
            ..
        })
    ));
}

#[test]
fn regular_series_requires_one_owner_and_series_text_remains_ai_local() {
    let stream = fixture(excel_chart());
    let missing = omit(&stream, RecordKind::from_wire(0x1045));
    assert!(matches!(
        Chart::parse(
            excel_input(&missing),
            Context::excel().with_external_sheets(1)
        ),
        Err(Error::InvalidChart {
            reason: "Series requires exactly four AI bindings and one SerToCrt",
            ..
        })
    ));

    let mut out = Encoder::new();
    let mut ai = 0usize;
    for item in Records::new(stream.as_bytes()) {
        let record = item.expect("valid fixture record");
        out.push_ref(record).expect("record replay");
        if record.kind() == RecordKind::from_wire(0x1051) {
            ai += 1;
            if ai == Role::ALL.len() {
                let text = [0, 0, 1, 0, b'x'];
                out.push(RecordKind::from_wire(0x100D), &text)
                    .expect("first optional SeriesText");
                out.push(RecordKind::from_wire(0x100D), &text)
                    .expect("misplaced second SeriesText");
            }
        }
    }
    let duplicate = out.finish();
    assert!(matches!(
        Chart::parse(
            excel_input(&duplicate),
            Context::excel().with_external_sheets(1)
        ),
        Err(Error::InvalidChart {
            reason: "SeriesText in Series must immediately follow one BRAI",
            ..
        })
    ));
}

#[test]
fn auxiliary_owner_round_trips_and_series_removal_preserves_dependencies() {
    let context = Context::excel().with_external_sheets(1);
    let mut chart = excel_chart();
    let mut auxiliary = Series::new(context);
    auxiliary.owner = Owner::Trend {
        parent: crate::record::series::Parent::try_new(1).expect("parent series"),
        data: [0; 28],
    };
    chart.add_series(auxiliary).expect("auxiliary series");
    let parsed = Chart::open(fixture(chart), context).expect("auxiliary parse");
    assert!(matches!(parsed.series()[1].owner, Owner::Trend { .. }));

    let mut blocked = excel_chart();
    let mut auxiliary = Series::new(context);
    auxiliary.owner = Owner::ErrorBar {
        parent: crate::record::series::Parent::try_new(1).expect("parent series"),
        data: [0; 14],
    };
    blocked.add_series(auxiliary).expect("auxiliary series");
    assert!(matches!(
        blocked.remove_series(0),
        Err(Error::InvalidModel {
            field: "series",
            ..
        })
    ));
    assert_eq!(blocked.series().len(), 2);

    let mut shifted = Chart::new(context).expect("chart");
    shifted
        .add_series(Series::new(context))
        .expect("first regular series");
    shifted
        .add_series(Series::new(context))
        .expect("second regular series");
    let mut auxiliary = Series::new(context);
    auxiliary.owner = Owner::Trend {
        parent: crate::record::series::Parent::try_new(2).expect("second parent"),
        data: [0; 28],
    };
    shifted.add_series(auxiliary).expect("auxiliary series");
    assert!(shifted.remove_series(0).expect("safe removal").is_some());
    let Owner::Trend { parent, .. } = &shifted.series()[1].owner else {
        panic!("shifted auxiliary owner");
    };
    assert_eq!(parent.series().get(), 1);
}

#[test]
fn cache_dimensions_and_bool_err_follow_the_typed_crud_model() {
    let mut chart = excel_chart();
    chart
        .add_cache(Cache::excel(
            cache::Index::Values,
            3,
            0,
            cache::Xf::new(4),
            XlValue::Bool(true),
        ))
        .expect("Boolean cache");
    chart
        .add_cache(Cache::excel(
            cache::Index::Values,
            4,
            0,
            cache::Xf::new(5),
            XlValue::Error(cache::Fault::DivZero),
        ))
        .expect("error cache");
    assert!(matches!(
        chart.dimensions(),
        cache::Dims::Excel(value) if value.row_after() == 5 && value.col_after() == 1
    ));
    assert!(
        chart
            .set_dimensions(cache::Dims::Excel(
                cache::ExcelDims::new(0, 4, 0, 1).expect("smaller range")
            ))
            .is_err()
    );

    let parsed = Chart::open(fixture(chart), Context::excel().with_external_sheets(1))
        .expect("BoolErr round trip");
    assert_eq!(parsed.caches()[3].value(), ValueRef::Bool(true));
    assert_eq!(
        parsed.caches()[4].value(),
        ValueRef::Error(cache::Fault::DivZero)
    );
}

#[test]
fn excel_cache_label_uses_xl_unicode_string_wire_format() {
    let text = "界".repeat(300);
    let mut chart = excel_chart();
    {
        let mut caches = chart.caches_mut();
        let Cache::Excel { value, .. } = caches.get_mut(1).expect("text cache") else {
            panic!("expected Excel text cache");
        };
        *value = XlValue::Text(text.clone());
    }

    let stream = fixture(chart);
    let label = stream
        .records()
        .map(|record| record.expect("valid fixture record"))
        .find(|record| record.kind() == RecordKind::from_wire(0x0204))
        .expect("Label record");
    assert_eq!(
        label.payload().get(6..9),
        Some([0x2C, 0x01, 0x01].as_slice())
    );
    assert_eq!(label.payload().len(), 6 + 3 + 300 * 2);

    let parsed = Chart::open(stream, Context::excel().with_external_sheets(1))
        .expect("XLUnicodeString Label round trip");
    assert_eq!(parsed.caches()[1].value(), ValueRef::Text(&text));
}

#[test]
fn group_and_parent_crud_preserves_semantic_ownership() {
    let context = Context::excel();
    let mut chart = Chart::new(context).expect("chart");
    let mut second = Group::line();
    second.order = Order::new(1).expect("drawing order");
    chart.add_group(second).expect("second group");
    let mut series = Series::new(context);
    series.owner = Owner::Group(GroupId::new(1).expect("second group index"));
    chart.add_series(series).expect("series");

    assert!(matches!(
        chart.remove_group(1),
        Err(Error::InvalidModel { field: "group", .. })
    ));
    assert!(chart.remove_group(0).expect("safe removal").is_some());
    assert_eq!(chart.groups().len(), 1);
    assert_eq!(chart.series()[0].owner.group(), Some(GroupId::ZERO));

    chart
        .add_axis(axis::Axis::new(axis::Kind::Category))
        .expect("primary axis");
    let parsed = Chart::open(fixture(chart), context).expect("parent ownership parse");
    assert_eq!(parsed.groups()[0].parent, axis::ParentId::PRIMARY);
    assert_eq!(parsed.axes()[0].parent, axis::ParentId::PRIMARY);
}

#[test]
fn mutable_borrow_marks_parsed_input_dirty_only_after_write() {
    let context = Context::excel().with_external_sheets(1);
    let mut chart = Chart::open(fixture(excel_chart()), context).expect("parsed chart");
    {
        let groups = chart.groups_mut();
        assert_eq!(groups.len(), 1);
    }
    assert!(chart.is_pristine());
    {
        let mut groups = chart.groups_mut();
        groups[0].vary_colors = true;
    }
    assert!(!chart.is_pristine());
}

#[test]
fn excel_rejects_proven_topology_violations_and_bad_siindex_order() {
    let context = Context::excel().with_external_sheets(1);
    let stream = fixture(excel_chart());
    for kind in [0x00A0, 0x1022, 0x104F].map(RecordKind::from_wire) {
        let malformed = omit(&stream, kind);
        assert!(Chart::parse(excel_input(&malformed), context).is_err());
    }

    let mut out = Encoder::new();
    let mut section = 0usize;
    for item in Records::new(stream.as_bytes()) {
        let record = item.expect("valid fixture record");
        if record.kind() == RecordKind::from_wire(0x1065) {
            section += 1;
            if section == 2 {
                out.push(record.kind(), &3u16.to_le_bytes())
                    .expect("out-of-order SIIndex");
                continue;
            }
        }
        out.push_ref(record).expect("record replay");
    }
    let malformed = out.finish();
    assert!(matches!(
        Chart::parse(excel_input(&malformed), context),
        Err(Error::InvalidChart {
            reason: "SIIndex sections are missing, duplicated, or out of order",
            ..
        })
    ));

    let mut out = Encoder::new();
    for item in Records::new(stream.as_bytes()) {
        let record = item.expect("valid fixture record");
        if record.kind() == RecordKind::from_wire(0x0200) {
            out.push(RecordKind::from_wire(0x1033), &[])
                .expect("orphan Begin");
            out.push(RecordKind::from_wire(0x1034), &[])
                .expect("balanced orphan End");
        }
        out.push_ref(record).expect("record replay");
    }
    let malformed = out.finish();
    assert!(matches!(
        Chart::parse(excel_input(&malformed), context),
        Err(Error::InvalidChart {
            reason: "Begin record has no chart-level collection owner",
            ..
        })
    ));
}

#[test]
fn graph_does_not_inherit_excel_outer_order_but_rejects_siindex() {
    let stream = fixture(Chart::new(Context::graph()).expect("Graph chart"));
    let without_excel_scl = omit(&stream, RecordKind::from_wire(0x00A0));
    let parsed = Chart::parse(
        Ref::open(&without_excel_scl).expect("Graph rewrite"),
        Context::graph(),
    )
    .expect("Graph does not require Excel CHARTFOMATS order");
    assert!(parsed.is_pristine());
    assert_eq!(
        parsed.encode().expect("exact Graph replay").as_bytes(),
        without_excel_scl
    );

    let without_excel_crt_link = omit(&stream, RecordKind::from_wire(0x1022));
    let parsed = Chart::parse(
        Ref::open(&without_excel_crt_link).expect("Graph rewrite"),
        Context::graph(),
    )
    .expect("Graph does not require the Excel-mandatory CrtLink");
    assert_eq!(
        parsed.encode().expect("exact Graph replay").as_bytes(),
        without_excel_crt_link
    );

    let mut out = Encoder::new();
    for item in Records::new(stream.as_bytes()) {
        let record = item.expect("valid fixture record");
        if record.kind() == super::super::EOF {
            out.push(RecordKind::from_wire(0x1065), &1u16.to_le_bytes())
                .expect("Graph SIIndex");
        }
        out.push_ref(record).expect("record replay");
    }
    let malformed = out.finish();
    let input = Ref::open(&malformed).expect("Graph SIIndex framing");
    assert!(matches!(
        Chart::parse(input, Context::graph()),
        Err(Error::InvalidChart {
            reason: "SIIndex is not part of the standalone Graph grammar",
            ..
        })
    ));
}
