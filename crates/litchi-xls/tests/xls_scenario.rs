use std::fs::File;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_xls::{Scenario, ScenarioCell, ScenarioManager, ScenarioRange, Workbook, Writer};

fn poi_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn scenario_manager_round_trip_keeps_values_inert() {
    let mut base = Scenario::new(
        "Base",
        vec![
            ScenarioCell::new(1, 2, "=EXEC(\"not evaluated\")"),
            ScenarioCell::new(4, 5, "42"),
        ],
    );
    base.set_creator(Some("作者".to_string()));
    base.set_comment(Some("Baseline scenario".to_string()));
    base.set_locked(true);
    let mut alternate = Scenario::new(
        "Alternative",
        vec![ScenarioCell::deleted(7, 3, "plain text")],
    );
    alternate.set_hidden(true);
    let mut manager = ScenarioManager::new(vec![base, alternate]);
    manager.set_current_scenario(Some(1));
    manager.set_shown_scenario(Some(0));
    manager.set_result_ranges(vec![ScenarioRange::new(0, 2, 0, 1).unwrap()]);

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Scenarios").unwrap();
    writer.set_scenario_manager(sheet, manager).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let manager = workbook
        .xls_worksheet(0)
        .unwrap()
        .scenario_manager()
        .unwrap();
    assert_eq!(manager.current_scenario(), Some(1));
    assert_eq!(manager.shown_scenario(), Some(0));
    assert_eq!(manager.result_ranges().len(), 1);
    assert_eq!(manager.scenarios()[0].creator(), Some("作者"));
    assert!(manager.scenarios()[0].is_locked());
    assert_eq!(
        manager.scenarios()[0].cells()[0].value(),
        "=EXEC(\"not evaluated\")"
    );
    assert!(manager.scenarios()[1].cells()[0].is_deleted());
}

#[test]
fn reads_poi_empty_scenario_manager_fixture() {
    let workbook = Workbook::new(File::open(poi_fixture("15228.xls")).unwrap()).unwrap();
    let manager = (0..workbook.sheets().len())
        .filter_map(|index| workbook.xls_worksheet(index).ok())
        .find_map(|sheet| sheet.scenario_manager())
        .expect("fixture contains ScenMan");
    assert!(manager.scenarios().is_empty());
    assert!(manager.result_ranges().is_empty());
    assert_eq!(manager.current_scenario(), None);
}

#[test]
fn writer_rejects_scenario_resource_bounds() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Scenarios").unwrap();
    let too_many = Scenario::new(
        "Too many",
        (0..33).map(|row| ScenarioCell::new(row, 0, "x")).collect(),
    );
    assert!(
        writer
            .set_scenario_manager(sheet, ScenarioManager::new(vec![too_many]))
            .is_err()
    );
}
