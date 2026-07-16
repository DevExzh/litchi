use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::xls::writer::{XlsCalculationSettings, XlsWriter};
use litchi_ole::xls::{XlsCalculationMode, XlsReferenceMode, XlsWorkbook};

#[test]
fn calculation_settings_round_trip() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Calculation").unwrap();
    writer.set_calculation_settings(XlsCalculationSettings {
        mode: XlsCalculationMode::Manual,
        maximum_iterations: 321,
        iteration_enabled: true,
        iteration_delta: 0.000_25,
        full_precision: false,
        reference_mode: XlsReferenceMode::R1C1,
        recalculate_before_save: false,
        recalculation_engine_id: 0x1234_5678,
        force_full_calculation: true,
    }).unwrap();
    writer.set_recalculation_pending(sheet, true).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let globals = workbook.calculation();
    assert!(!globals.full_precision());
    assert!(globals.force_full_calculation());
    assert_eq!(globals.recalculation_engine_id(), Some(0x1234_5678));
    let calculation = workbook.xls_worksheet(0).unwrap().calculation();
    assert_eq!(calculation.mode(), XlsCalculationMode::Manual);
    assert_eq!(calculation.maximum_iterations(), 321);
    assert!(calculation.iteration_enabled());
    assert_eq!(calculation.iteration_delta(), 0.000_25);
    assert_eq!(calculation.reference_mode(), XlsReferenceMode::R1C1);
    assert!(!calculation.recalculate_before_save());
    assert!(calculation.formulas_pending_recalculation());
}

#[test]
fn reads_poi_calculation_fixture() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet/UncalcedRecord.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    let calculation = workbook.xls_worksheet(0).unwrap().calculation();
    assert_eq!(calculation.mode(), XlsCalculationMode::Automatic);
    assert_eq!(calculation.maximum_iterations(), 100);
    assert_eq!(calculation.reference_mode(), XlsReferenceMode::A1);
    assert!(!calculation.iteration_enabled());
    assert!(calculation.recalculate_before_save());
}

#[test]
fn writer_rejects_invalid_calculation_bounds() {
    let mut writer = XlsWriter::new();
    let invalid_count = XlsCalculationSettings {
        maximum_iterations: 0,
        ..XlsCalculationSettings::default()
    };
    assert!(writer.set_calculation_settings(invalid_count).is_err());
    let invalid_delta = XlsCalculationSettings {
        iteration_delta: f64::NAN,
        ..XlsCalculationSettings::default()
    };
    assert!(writer.set_calculation_settings(invalid_delta).is_err());
}
