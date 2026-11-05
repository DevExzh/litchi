#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic by design"
)]

use std::io::Cursor;

use litchi_biff::{Encoder, Kind, Records};
use litchi_cfb::OleWriter;

use crate::chart;
use crate::{Error, Limits};

use super::codec::{COMP_OBJ, OLE, WORKBOOK};
use super::semantic::{Package, PackageRef, Workbook, WorkbookRef};
use super::validation::{
    BOF, BOF_BYTES, CHART_SHEET, EOF, GLOBALS, OGRAPH_VERSION, REQUIRED_PLATFORM_FLAGS,
};

fn bof(doc_type: u16) -> [u8; BOF_BYTES] {
    let mut payload = [0; BOF_BYTES];
    payload[0..2].copy_from_slice(&OGRAPH_VERSION.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    payload[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14)).to_le_bytes());
    payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    payload
}

fn workbook() -> Vec<u8> {
    workbook_with_bofs(bof(GLOBALS), bof(CHART_SHEET))
}

fn workbook_with_bofs(globals: [u8; BOF_BYTES], chart: [u8; BOF_BYTES]) -> Vec<u8> {
    let mut out = Encoder::new(Limits::default().biff).expect("BIFF limits");
    out.push(BOF, &globals).expect("globals BOF");
    out.push(Kind::from_wire(0x7777), &[1, 2, 3])
        .expect("unknown record");
    out.push(EOF, &[]).expect("globals EOF");
    out.push(BOF, &chart).expect("chart BOF");
    out.push(Kind::from_wire(0x7778), &[4, 5])
        .expect("unknown record");
    out.push(EOF, &[]).expect("chart EOF");
    out.finish()
}

fn editable_chart(with_cache: bool) -> Vec<u8> {
    let mut out = Encoder::new(Limits::default().biff).expect("BIFF limits");
    out.push(BOF, &bof(CHART_SHEET)).expect("chart BOF");

    let mut rect = [0u8; 16];
    rect[8..12].copy_from_slice(&(4_000_i32 << 16).to_le_bytes());
    rect[12..16].copy_from_slice(&(3_000_i32 << 16).to_le_bytes());
    out.push(Kind::from_wire(0x1002), &rect)
        .expect("Chart record");
    out.push(Kind::from_wire(0x1033), &[]).expect("chart Begin");
    out.push(Kind::from_wire(0x00A0), &[1, 0, 1, 0])
        .expect("Scl");
    out.push(Kind::from_wire(0x1064), &[0, 0, 1, 0, 0, 0, 1, 0])
        .expect("PlotGrowth");
    out.push(Kind::from_wire(0x1044), &[2, 0, 0, 0])
        .expect("ShtProps");
    out.push(Kind::from_wire(0x1046), &[1, 0])
        .expect("AxesUsed");
    out.push(Kind::from_wire(0x1041), &[0; 18])
        .expect("AxisParent");
    out.push(Kind::from_wire(0x1033), &[])
        .expect("parent Begin");
    let mut pos = [0u8; 20];
    pos[0..2].copy_from_slice(&2_u16.to_le_bytes());
    pos[2..4].copy_from_slice(&2_u16.to_le_bytes());
    pos[12..14].copy_from_slice(&4_000_i16.to_le_bytes());
    pos[16..18].copy_from_slice(&4_000_i16.to_le_bytes());
    out.push(Kind::from_wire(0x104F), &pos).expect("Pos");
    out.push(Kind::from_wire(0x1034), &[]).expect("parent End");
    out.push(Kind::from_wire(0x1035), &[]).expect("PlotArea");
    out.push(Kind::from_wire(0x1034), &[]).expect("chart End");

    let mut dimensions = [0u8; 14];
    if with_cache {
        dimensions[4..8].copy_from_slice(&1_u32.to_le_bytes());
        dimensions[10..12].copy_from_slice(&1_u16.to_le_bytes());
    }
    out.push(Kind::from_wire(0x0200), &dimensions)
        .expect("Dimensions");
    if with_cache {
        let mut number = [0u8; 15];
        number[5..7].copy_from_slice(&4_u16.to_le_bytes());
        number[7..15].copy_from_slice(&1.25_f64.to_le_bytes());
        out.push(Kind::from_wire(0x0003), &number)
            .expect("Graph Number");
    }
    out.push(Kind::from_wire(0x7777), &[0xA1, 0xB2, 0xC3])
        .expect("opaque chart record");
    out.push(EOF, &[]).expect("chart EOF");
    out.finish()
}

fn editable_workbook(with_cache: bool) -> Vec<u8> {
    let chart = editable_chart(with_cache);
    let mut out = Encoder::new(Limits::default().biff).expect("BIFF limits");
    out.push(BOF, &bof(GLOBALS)).expect("globals BOF");
    out.push(Kind::from_wire(0x7776), &[0x10, 0x20])
        .expect("opaque globals record");
    out.push(EOF, &[]).expect("globals EOF");
    for item in Records::new(&chart, litchi_biff::Limits::default()).expect("valid BIFF limits") {
        out.push_ref(item.expect("chart record"))
            .expect("chart replay");
    }
    out.finish()
}

fn package(workbook: Option<&[u8]>, extras: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    if let Some(workbook) = workbook {
        writer
            .create_stream(&[WORKBOOK], workbook)
            .expect("Workbook stream");
    }
    for (name, bytes) in extras {
        writer.create_stream(&[*name], bytes).expect("extra stream");
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write CFB");
    output.into_inner()
}

fn authored_graph_chart() -> chart::Model {
    let mut chart = chart::Model::new_graph_authoring(chart::GraphFamily::Pie, Limits::default())
        .expect("Graph scaffold");
    let count = chart::Count::new(2).expect("count");
    let row = chart::RowCol::new(1).expect("row");
    let series = chart::Series {
        category_kind: chart::DataKind::Text,
        category_count: count,
        value_count: count,
        bubble_count: chart::Count::ZERO,
        owner: chart::Owner::PRIMARY,
        ai: chart::Ai::new(
            chart::Binding::new(
                chart::Link::graph(chart::Role::Name, chart::Source::Literal, row),
                Some("Slices".into()),
            ),
            chart::Binding::new(
                chart::Link::graph(chart::Role::Values, chart::Source::Literal, row),
                None,
            ),
            chart::Binding::new(
                chart::Link::graph(
                    chart::Role::Categories,
                    chart::Source::Literal,
                    chart::RowCol::ZERO,
                ),
                None,
            ),
            chart::Binding::new(
                chart::Link::graph(
                    chart::Role::Bubbles,
                    chart::Source::Automatic,
                    chart::RowCol::ZERO,
                ),
                None,
            ),
        )
        .expect("AI"),
    };
    chart.add_series(series).expect("series");
    chart
        .add_cache(chart::Cache::graph(
            row,
            chart::RowCol::ZERO,
            chart::cache::Ifmt::new(0),
            chart::Value::Text("Slices".into()),
        ))
        .expect("series name cache");
    for (col, value) in [(1, "A"), (2, "B")] {
        chart
            .add_cache(chart::Cache::graph(
                chart::RowCol::ZERO,
                chart::RowCol::new(col).expect("column"),
                chart::cache::Ifmt::new(0),
                chart::Value::Text(value.into()),
            ))
            .expect("category cache");
    }
    for (col, value) in [(1, 1.0), (2, 2.0)] {
        chart
            .add_cache(chart::Cache::graph(
                row,
                chart::RowCol::new(col).expect("column"),
                chart::cache::Ifmt::new(0),
                chart::Value::Number(value),
            ))
            .expect("value cache");
    }
    chart
}

#[test]
fn authored_graph_package_reopens_as_owned_graph_chart() {
    let package = Package::from_chart(authored_graph_chart(), crate::OpenOptions::default())
        .expect("Graph package");
    assert_eq!(package.topology().stream_count(), 1);
    let bytes = package.finish().into_bytes();
    let reopened = Package::open(bytes).expect("reopen Graph package");
    let workbook = reopened.workbook().expect("Workbook");
    let parsed =
        chart::Model::parse(workbook.chart(), chart::Context::graph()).expect("typed Graph chart");
    assert_eq!(parsed.groups().len(), 1);
    assert!(parsed.axes().is_empty());
    assert_eq!(parsed.series().len(), 1);
    assert_eq!(parsed.caches().len(), 5);
}

#[test]
fn accepts_exact_standalone_topology_and_reads_workbook() {
    let workbook = workbook();
    let bytes = package(Some(&workbook), &[(COMP_OBJ, &[1, 2]), (OLE, &[3, 4, 5])]);
    let parsed = PackageRef::open(&bytes).expect("valid package");
    assert_eq!(parsed.topology().stream_count(), 3);
    assert_eq!(parsed.topology().workbook_bytes(), workbook.len() as u64);
    let opened = parsed.workbook().expect("read Workbook");
    assert_eq!(opened.as_bytes(), workbook);
    assert_eq!(opened.chart().kind(), chart::Kind::Graph);
    assert_eq!(
        opened.chart().records().expect("validated records").count(),
        3
    );
}

#[test]
fn owned_finish_reuses_the_input_allocation() {
    let bytes = package(Some(&workbook()), &[]);
    let pointer = bytes.as_ptr();
    let capacity = bytes.capacity();
    let payload = Package::open(bytes).expect("valid").finish();
    let bytes = payload.into_bytes();
    assert_eq!(bytes.as_ptr(), pointer);
    assert_eq!(bytes.capacity(), capacity);
}

#[test]
fn rejects_missing_unknown_and_nested_root_entries() {
    let missing = package(None, &[(COMP_OBJ, &[])]);
    assert!(matches!(
        PackageRef::open(&missing),
        Err(Error::MissingStream { name: WORKBOOK })
    ));

    let unknown = package(Some(&workbook()), &[("Other", &[])]);
    assert!(matches!(
        PackageRef::open(&unknown),
        Err(Error::UnexpectedEntry { name, .. }) if name == "Other"
    ));

    let mut writer = OleWriter::new();
    writer.create_storage(&["Nested"]).expect("storage");
    writer
        .create_stream(&[WORKBOOK], &workbook())
        .expect("Workbook");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write CFB");
    assert!(matches!(
        PackageRef::open(output.get_ref()),
        Err(Error::UnexpectedEntry { name, .. }) if name == "Nested"
    ));
}

#[test]
fn rejects_wrong_substream_shape_and_trailing_records() {
    let mut wrong = Encoder::new(Limits::default().biff).expect("BIFF limits");
    wrong.push(BOF, &bof(GLOBALS)).expect("BOF");
    wrong.push(EOF, &[]).expect("EOF");
    wrong.push(BOF, &bof(GLOBALS)).expect("wrong BOF");
    wrong.push(EOF, &[]).expect("EOF");
    let bytes = package(Some(&wrong.finish()), &[]);
    assert!(matches!(
        PackageRef::open(&bytes),
        Err(Error::InvalidWorkbook { .. })
    ));

    let mut trailing = workbook();
    trailing.extend_from_slice(&[0x0A, 0, 0, 0]);
    let bytes = package(Some(&trailing), &[]);
    assert!(matches!(
        PackageRef::open(&bytes),
        Err(Error::InvalidWorkbook { .. })
    ));
}

#[test]
fn checks_package_and_workbook_limits_before_exposing_data() {
    let bytes = package(Some(&workbook()), &[]);
    let package_limits = Limits {
        max_package_bytes: bytes.len() - 1,
        ..Limits::default()
    };
    assert!(matches!(
        PackageRef::open_with_options(&bytes, crate::OpenOptions::new().limits(package_limits)),
        Err(Error::Limit {
            kind: crate::LimitKind::PackageBytes,
            ..
        })
    ));

    let workbook_limits = Limits {
        max_workbook_bytes: 1,
        ..Limits::default()
    };
    assert!(matches!(
        PackageRef::open_with_options(&bytes, crate::OpenOptions::new().limits(workbook_limits)),
        Err(Error::Limit {
            kind: crate::LimitKind::WorkbookBytes,
            ..
        })
    ));
}

#[test]
fn strict_bof_checks_must_fields_but_preserves_undefined_bits() {
    let mut globals = bof(GLOBALS);
    let ignored = (0b11 << 6) | (1 << 9) | (1 << 10) | (0b11 << 11) | (1 << 13) | (1 << 18);
    globals[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | ignored | (6 << 14)).to_le_bytes());
    let bytes = workbook_with_bofs(globals, bof(CHART_SHEET));
    let opened = Workbook::open(bytes).expect("undefined BOF bits are preserved");
    assert_eq!(opened.as_bytes(), opened.as_ref().as_bytes());

    let mut invalid = Vec::new();

    let mut year = bof(GLOBALS);
    year[6..8].copy_from_slice(&0x07CB_u16.to_le_bytes());
    invalid.push(year);

    let mut platform = bof(GLOBALS);
    platform[8..12].copy_from_slice(&(6_u32 << 14).to_le_bytes());
    invalid.push(platform);

    let mut forbidden = bof(GLOBALS);
    forbidden[8..12]
        .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (1 << 1) | (6 << 14)).to_le_bytes());
    invalid.push(forbidden);

    let mut reserved1 = bof(GLOBALS);
    reserved1[8..12]
        .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14) | (1 << 19)).to_le_bytes());
    invalid.push(reserved1);

    let mut highest = bof(GLOBALS);
    highest[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (5 << 14)).to_le_bytes());
    invalid.push(highest);

    let mut lowest = bof(GLOBALS);
    lowest[12..16].copy_from_slice(&(0x05_u32 | (4 << 8)).to_le_bytes());
    invalid.push(lowest);

    let mut last = bof(GLOBALS);
    last[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (4 << 14)).to_le_bytes());
    last[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    invalid.push(last);

    let mut reserved2 = bof(GLOBALS);
    reserved2[12..16].copy_from_slice(&(0x06_u32 | (6 << 8) | (1 << 12)).to_le_bytes());
    invalid.push(reserved2);

    for globals in invalid {
        assert!(matches!(
            WorkbookRef::open(&workbook_with_bofs(globals, bof(CHART_SHEET))),
            Err(Error::InvalidWorkbook { .. })
        ));
    }
}

#[test]
fn package_edit_replays_typed_graph_metadata_and_opaque_streams() {
    let source = package(
        Some(&editable_workbook(true)),
        &[(COMP_OBJ, &[0xCA, 0xFE]), (OLE, &[0xBA, 0xBE, 0x01])],
    );
    let snapshot = Package::open(source.clone()).expect("snapshot");
    let before = chart::Model::parse(
        snapshot.workbook().unwrap().chart(),
        chart::Context::graph(),
    )
    .expect("typed chart");
    assert_eq!(before.rect(), chart::Rect::default());
    assert_eq!(before.caches()[0].value(), chart::ValueRef::Number(1.25));
    assert_eq!(before.unknown().len(), 1);

    let target = chart::Rect {
        x: 0,
        y: 0,
        width: 4_500 << 16,
        height: 2_500 << 16,
    };
    let mut transaction = snapshot.edit();
    transaction.set_rect(target).expect("chart area");
    transaction
        .set_props(chart::Props {
            flags: 2 | (1 << 16),
            plot_area: true,
        })
        .expect("chart properties");
    transaction
        .set_cache_value(0, chart::Value::Number(2.5))
        .expect("cache value");
    let commit = transaction.commit().expect("package commit");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().chart().change_count(), 3);

    let output = commit.package();
    let after = chart::Model::parse(output.workbook().unwrap().chart(), chart::Context::graph())
        .expect("reparsed chart");
    assert_eq!(after.rect(), target);
    assert_eq!(after.props().flags, 2 | (1 << 16));
    assert_eq!(after.caches()[0].value(), chart::ValueRef::Number(2.5));
    assert_eq!(after.unknown().len(), 1);

    let mut cfb = litchi_cfb::OleFile::open(
        Cursor::new(output.as_bytes()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("CFB");
    assert_eq!(cfb.open_stream(&[COMP_OBJ]).expect("CompObj"), [0xCA, 0xFE]);
    assert_eq!(cfb.open_stream(&[OLE]).expect("Ole"), [0xBA, 0xBE, 0x01]);
    assert_ne!(output.as_bytes(), source.as_slice());
}

#[test]
fn package_no_op_is_exact_and_inverse_patch_restores_source() {
    let source = package(Some(&editable_workbook(false)), &[(COMP_OBJ, &[1, 2, 3])]);
    let snapshot = Package::open(source.clone()).expect("snapshot");
    let no_op = snapshot.edit().commit().expect("commit");
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().as_bytes(), source.as_slice());

    let mut edit = snapshot.edit();
    edit.set_rect(chart::Rect {
        x: 0,
        y: 0,
        width: 4_200 << 16,
        height: 3_100 << 16,
    })
    .expect("stage");
    let commit = edit.commit().expect("commit");
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.package().as_bytes(), source.as_slice());
}

#[test]
fn package_patch_rejects_stale_sources_and_invalid_edits_atomically() {
    let source = package(Some(&editable_workbook(false)), &[]);
    let snapshot = Package::open(source.clone()).expect("snapshot");

    let mut invalid = snapshot.edit();
    assert!(matches!(
        invalid.set_rect(chart::Rect {
            x: 1,
            y: 0,
            width: 1,
            height: 1,
        }),
        Err(Error::InvalidModel {
            field: "chart area",
            ..
        })
    ));
    let invalid_commit = invalid.commit().expect("atomic no-op");
    assert!(!invalid_commit.diagnostics().changed());
    assert_eq!(invalid_commit.package().as_bytes(), source.as_slice());

    let mut edit = snapshot.edit();
    edit.set_props(chart::Props {
        flags: 2 | (1 << 16),
        plot_area: true,
    })
    .expect("stage");
    let commit = edit.commit().expect("commit");
    let patch = commit.patch().clone();
    let stale_bytes = package(Some(&editable_workbook(false)), &[(COMP_OBJ, &[9])]);
    let stale = Package::open(stale_bytes).expect("stale package");
    assert!(matches!(
        stale.apply(&patch),
        Err(Error::PatchConflict {
            operation: "package-patch",
            ..
        })
    ));
}

#[test]
fn shared_owner_edits_are_lazy_and_drop_preserves_the_exact_artifact() {
    // Valid frame lengths but an invalid typed Chart payload: edit creation
    // must not parse it; either staging or committing must still reject it.
    let original = workbook();
    let mut malformed = Encoder::new(Limits::default().biff).unwrap();
    for record in Records::new(&original, Limits::default().biff).unwrap() {
        let record = record.unwrap();
        let kind = if record.kind() == Kind::from_wire(0x7778) {
            Kind::from_wire(0x1002)
        } else {
            record.kind()
        };
        malformed.push(kind, record.payload()).unwrap();
    }
    let bytes = package(Some(&malformed.finish()), &[(COMP_OBJ, &[7; 4096])]);
    let pointer = bytes.as_ptr();
    let source = Package::open(bytes).unwrap();
    let mut edit = source.edit();
    assert!(edit.is_empty());
    assert_eq!(edit.len(), 0);
    assert_eq!(edit.source().as_bytes().as_ptr(), pointer);
    assert!(edit.set_rect(chart::Rect::default()).is_err());
    drop(edit);
    assert!(source.edit().commit().is_err());
    let recovered = source.finish().into_bytes();
    assert_eq!(recovered.as_ptr(), pointer);
}

#[test]
fn package_replay_is_content_based_and_returns_an_editable_owner() {
    let bytes = package(Some(&editable_workbook(true)), &[(COMP_OBJ, &[9, 8, 7])]);
    let source = Package::open(bytes.clone()).unwrap();
    let mut edit = source.edit();
    edit.set_cache_value(0, chart::Value::Number(4.5)).unwrap();
    let commit = edit.commit().unwrap();
    let recipient = Package::open_with_options(
        bytes.clone(),
        crate::OpenOptions::new().limits(Limits {
            max_cached_values: 1,
            ..Limits::default()
        }),
    )
    .unwrap();
    assert_ne!(source.as_bytes().as_ptr(), recipient.as_bytes().as_ptr());
    let replayed: crate::Commit = recipient.apply(commit.patch()).unwrap();
    assert!(replayed.diagnostics().changed());
    assert_eq!(replayed.package().as_bytes(), commit.package().as_bytes());
    assert_eq!(
        replayed.package().as_ref().limits(),
        recipient.as_ref().limits()
    );
    let mut next = replayed.package().edit();
    next.set_cache_value(0, chart::Value::Number(6.5)).unwrap();
    assert!(next.commit().unwrap().diagnostics().changed());
    let reopened = Package::open(replayed.package().as_bytes().to_vec()).unwrap();
    let inverse = reopened.apply(&commit.patch().inverse()).unwrap();
    assert_eq!(inverse.package().as_bytes(), bytes);
    let noop = source.edit().commit().unwrap();
    let noop_replayed = recipient.apply(noop.patch()).unwrap();
    assert!(!noop_replayed.diagnostics().changed());
    assert!(noop_replayed.patch().is_noop());
    assert_eq!(noop_replayed.package().as_bytes(), bytes);
    assert_eq!(source.as_bytes(), bytes);
}

#[test]
fn concurrent_package_borrows_promote_one_immutable_allocation() {
    let bytes = package(Some(&editable_workbook(true)), &[]);
    let pointer = bytes.as_ptr() as usize;
    let source = Package::open(bytes).unwrap();
    std::thread::scope(|scope| {
        let handles: Vec<_> = (0..4)
            .map(|_| {
                scope.spawn(|| {
                    let copy = source.clone();
                    assert_eq!(copy.as_bytes(), source.as_bytes());
                    assert_eq!(copy.as_bytes().as_ptr() as usize, pointer);
                    assert_eq!(
                        copy.workbook().unwrap().as_bytes(),
                        source.workbook().unwrap().as_bytes()
                    );
                })
            })
            .collect();
        for handle in handles {
            handle.join().unwrap();
        }
    });
    let recovered = source.finish().into_bytes();
    assert_eq!(recovered.as_ptr() as usize, pointer);
}
