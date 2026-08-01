use litchi_cfb::OleWriter;
use litchi_xls::XlsError;
use litchi_xls::chart::{Chart, Editor, Limits, Location, Selector, build_workbook};
use std::io::Cursor;

fn record(kind: u16, data: &[u8]) -> Vec<u8> {
    let mut out = kind.to_le_bytes().to_vec();
    out.extend((data.len() as u16).to_le_bytes());
    out.extend(data);
    out
}
fn bof(kind: u16) -> Vec<u8> {
    let mut d = Vec::new();
    d.extend(0x0600u16.to_le_bytes());
    d.extend(kind.to_le_bytes());
    d.extend([0; 12]);
    record(0x0809, &d)
}
fn workbook_with(records: &[u8]) -> Vec<u8> {
    let mut globals = bof(5);
    let bound_at = globals.len();
    let mut bound = vec![0; 8];
    bound[6] = 1;
    bound.extend(b"S");
    globals.extend(record(0x0085, &bound));
    globals.extend(record(0x000a, &[]));
    let offset = globals.len() as u32;
    globals[bound_at + 4..bound_at + 8].copy_from_slice(&offset.to_le_bytes());
    globals.extend(bof(0x0010));
    globals.extend(records);
    globals.extend(record(0x000a, &[]));
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &globals).unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    out.into_inner()
}
fn workbook() -> Vec<u8> {
    workbook_with(&[])
}

fn assert_unsupported<T>(result: Result<T, XlsError>) {
    match result {
        Err(XlsError::Graph(litchi_ograph::Error::UnsupportedAuthoring { reason })) => {
            assert!(!reason.is_empty());
        },
        Err(error) => panic!("expected typed unsupported-authoring error, found {error}"),
        Ok(_) => panic!("incomplete chart authoring unexpectedly succeeded"),
    }
}

fn assert_unsupported_mutation<T>(result: Result<T, XlsError>) {
    match result {
        Err(XlsError::Graph(litchi_ograph::Error::UnsupportedMutation { operation, reason })) => {
            assert!(!operation.is_empty());
            assert!(!reason.is_empty());
        },
        Err(error) => panic!("expected typed unsupported-mutation error, found {error}"),
        Ok(_) => panic!("unsafe embedded chart mutation unexpectedly succeeded"),
    }
}

#[test]
fn public_fresh_and_replacement_authoring_is_typed_and_atomic() {
    let original = workbook();

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported(editor.add("S", Chart::default()));
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported(editor.insert_at(0, 1, 0, Chart::default()));
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported(editor.add_sheet("Chart", Chart::default()));
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported(editor.insert_sheet_at(1, "Chart", Chart::default()));
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported(editor.replace(
        Selector::Embedded {
            sheet: "S",
            index: 0,
        },
        Chart::default(),
    ));
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported(editor.replace_at(
        &Location::Embedded {
            sheet_index: 0,
            object_id: 1,
        },
        Chart::default(),
    ));
    assert_eq!(editor.finish().unwrap(), original);

    assert_unsupported(build_workbook(Chart::default(), Limits::default()));
}

#[test]
fn clean_inventory_replay_and_existing_sheet_reorder_remain_available() {
    let original = workbook();
    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_eq!(editor.charts().len(), 0);
    editor.reorder_sheets(&["S"]).unwrap();
    let finished = editor.finish().unwrap();
    let reopened = Editor::open(finished, Limits::default()).unwrap();
    assert_eq!(reopened.charts().len(), 0);
}

#[test]
fn real_embedded_chart_mutations_are_typed_atomic_refusals() {
    let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls/WithThreeCharts.xls");
    let original = std::fs::read(path).unwrap();
    let inventory = Editor::open(original.clone(), Limits::default())
        .expect("bundled POI three-chart fixture must remain readable")
        .charts()
        .map(|entry| entry.location.clone())
        .collect::<Vec<_>>();
    let (sheet_index, object_ids) = inventory
        .iter()
        .find_map(|location| {
            let Location::Embedded {
                sheet_index,
                object_id: _,
            } = location
            else {
                return None;
            };
            let ids = inventory
                .iter()
                .filter_map(|candidate| match candidate {
                    Location::Embedded {
                        sheet_index: candidate_sheet,
                        object_id,
                    } if candidate_sheet == sheet_index => Some(*object_id),
                    _ => None,
                })
                .collect::<Vec<_>>();
            (ids.len() >= 2).then_some((*sheet_index, ids))
        })
        .expect("fixture must contain two embedded charts on one worksheet");

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    let first_object_id = object_ids
        .first()
        .copied()
        .expect("fixture chart inventory is non-empty");
    assert_unsupported_mutation(editor.remove_at(&Location::Embedded {
        sheet_index,
        object_id: first_object_id,
    }));
    assert_eq!(editor.finish().unwrap(), original);

    let mut reversed = object_ids.clone();
    reversed.reverse();
    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    assert_unsupported_mutation(editor.reorder_at(sheet_index, &reversed));
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original.clone(), Limits::default()).unwrap();
    editor.reorder_at(sheet_index, &object_ids).unwrap();
    assert_eq!(editor.finish().unwrap(), original);
}

#[test]
fn bundled_poi_and_libreoffice_chart_fixtures_are_strictly_gated() {
    let root = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for path in [
        root.join("test-data/poi/test-data/spreadsheet/44010-TwoCharts.xls"),
        root.join("test-data/poi/test-data/spreadsheet/SimpleScatterChart.xls"),
        root.join("test-data/libreoffice-core/sc/qa/unit/data/xls/embedded-chart.xls"),
        root.join("test-data/libreoffice-core/chart2/qa/extras/data/xls/chart.xls"),
    ] {
        let original = std::fs::read(&path).unwrap();
        match Editor::open(original.clone(), Limits::default()) {
            Ok(editor) => {
                let count = editor.charts().len();
                let finished = editor.finish().expect("finish clean fixture editor");
                assert_eq!(finished, original, "clean finish must preserve exact bytes");
                let reopened =
                    Editor::open(finished, Limits::default()).expect("reopen clean fixture editor");
                assert_eq!(reopened.charts().len(), count);
            },
            Err(_) => assert_eq!(std::fs::read(&path).unwrap(), original),
        }
    }
}
