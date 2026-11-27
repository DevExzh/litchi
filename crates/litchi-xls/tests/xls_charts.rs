use litchi_xls::Error;
use litchi_xls::builder::Chart;
use litchi_xls::chart::{Limits, Location, Selector, build_workbook};

fn workbook() -> Vec<u8> {
    // A complete workbook with the one worksheet `S`, which the package-level
    // chart family opens.
    let mut builder = litchi_xls::Builder::new();
    builder.add_worksheet("S").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

fn assert_unsupported<T>(result: Result<T, Error>) {
    match result {
        Err(Error::Refused(refusal)) => {
            assert_eq!(refusal.operation(), "author");
            assert!(!refusal.reason().is_empty());
        },
        Err(error) => panic!("expected typed unsupported-authoring error, found {error}"),
        Ok(_) => panic!("incomplete chart authoring unexpectedly succeeded"),
    }
}

fn assert_unsupported_mutation<T>(result: Result<T, Error>) {
    match result {
        Err(Error::Refused(refusal)) => {
            assert!(!refusal.operation().is_empty());
            assert!(!refusal.reason().is_empty());
        },
        Err(error) => panic!("expected typed unsupported-mutation error, found {error}"),
        Ok(_) => panic!("unsafe embedded chart mutation unexpectedly succeeded"),
    }
}

#[test]
fn public_fresh_and_replacement_authoring_is_typed_and_atomic() {
    let original = workbook();

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported(editor.add("S", Chart::default()));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported(editor.insert_at(0, 1, 0, Chart::default()));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported(editor.add_sheet("Chart", Chart::default()));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported(editor.insert_sheet_at(1, "Chart", Chart::default()));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported(editor.replace(
        Selector::Embedded {
            sheet: "S",
            index: 0,
        },
        Chart::default(),
    ));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported(editor.replace_at(
        &Location::Embedded {
            sheet_index: 0,
            object_id: 1,
        },
        Chart::default(),
    ));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    assert_unsupported(build_workbook(Chart::default(), Limits::default()));
}

#[test]
fn clean_inventory_replay_and_existing_sheet_reorder_remain_available() {
    let original = workbook();
    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_eq!(editor.charts().unwrap().len(), 0);
    editor.reorder_sheets(&["S"]).unwrap();
    let finished = editor.commit().unwrap().patch().after().to_vec();
    let mut reopened = litchi_xls::Package::from_bytes(&finished)
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_eq!(reopened.charts().unwrap().len(), 0);
}

#[test]
fn real_embedded_chart_mutations_are_typed_atomic_refusals() {
    let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls/WithThreeCharts.xls");
    let original = std::fs::read(path).unwrap();
    let inventory = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .expect("bundled POI three-chart fixture must remain readable")
        .charts()
        .unwrap()
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

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    let first_object_id = object_ids
        .first()
        .copied()
        .expect("fixture chart inventory is non-empty");
    assert_unsupported_mutation(editor.remove_at(&Location::Embedded {
        sheet_index,
        object_id: first_object_id,
    }));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut reversed = object_ids.clone();
    reversed.reverse();
    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    assert_unsupported_mutation(editor.reorder_at(sheet_index, &reversed));
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);

    let mut editor = litchi_xls::Package::from_bytes(&original.clone())
        .unwrap()
        .edit_charts()
        .unwrap();
    editor.reorder_at(sheet_index, &object_ids).unwrap();
    assert_eq!(editor.commit().unwrap().patch().after().to_vec(), original);
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
        match litchi_xls::Package::from_bytes(&original.clone())
            .unwrap()
            .edit_charts()
        {
            Ok(mut editor) => {
                let count = editor.charts().unwrap().len();
                let finished = editor
                    .commit()
                    .expect("finish clean fixture editor")
                    .patch()
                    .after()
                    .to_vec();
                assert_eq!(finished, original, "clean finish must preserve exact bytes");
                let mut reopened = litchi_xls::Package::from_bytes(&finished)
                    .unwrap()
                    .edit_charts()
                    .expect("reopen clean fixture editor");
                assert_eq!(reopened.charts().unwrap().len(), count);
            },
            Err(_) => assert_eq!(std::fs::read(&path).unwrap(), original),
        }
    }
}
