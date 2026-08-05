use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_xls::writer::{CalculationSettings, Writer};
use litchi_xls::{CalculationMode, MultithreadedCalculation, ReferenceMode, Workbook};

fn workbook_global_record_types(bytes: &[u8]) -> Vec<u16> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut record_types = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= stream.len() {
        let record_type = u16::from_le_bytes(stream[offset..offset + 2].try_into().unwrap());
        let length = usize::from(u16::from_le_bytes(
            stream[offset + 2..offset + 4].try_into().unwrap(),
        ));
        let end = offset.checked_add(4 + length).unwrap();
        assert!(
            end <= stream.len(),
            "truncated BIFF record in writer output"
        );
        record_types.push(record_type);
        offset = end;
        if record_type == 0x000A {
            break;
        }
    }
    record_types
}

#[test]
fn calculation_settings_round_trip() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Calculation").unwrap();
    writer
        .set_calculation_settings(CalculationSettings {
            mode: CalculationMode::Manual,
            maximum_iterations: 321,
            iteration_enabled: true,
            iteration_delta: 0.000_25,
            full_precision: false,
            reference_mode: ReferenceMode::R1C1,
            recalculate_before_save: false,
            recalculation_engine_id: 0x1234_5678,
            multithreaded_calculation: Some(
                MultithreadedCalculation::try_with_thread_count(true, 8).unwrap(),
            ),
            force_full_calculation: true,
        })
        .unwrap();
    writer.set_recalculation_pending(sheet, true).unwrap();
    writer.write_string(sheet, 0, 0, "shared string").unwrap();
    writer.add_comment(sheet, 0, 0, "Author", "note").unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let bytes = bytes.into_inner();
    let relevant = workbook_global_record_types(&bytes)
        .into_iter()
        .filter(|record_type| {
            matches!(
                record_type,
                0x089A | 0x08A3 | 0x008C | 0x01C1 | 0x00EB | 0x00FC
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        relevant,
        vec![0x089A, 0x08A3, 0x008C, 0x01C1, 0x00EB, 0x00FC]
    );
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    let globals = workbook.calculation();
    assert!(!globals.full_precision());
    assert!(globals.force_full_calculation());
    assert_eq!(globals.recalculation_engine_id(), Some(0x1234_5678));
    let threading = globals.multithreaded_calculation().unwrap();
    assert!(threading.enabled());
    assert_eq!(threading.user_thread_count(), Some(8));
    let calculation = workbook.xls_worksheet(0).unwrap().calculation();
    assert_eq!(calculation.mode(), CalculationMode::Manual);
    assert_eq!(calculation.maximum_iterations(), 321);
    assert!(calculation.iteration_enabled());
    assert_eq!(calculation.iteration_delta(), 0.000_25);
    assert_eq!(calculation.reference_mode(), ReferenceMode::R1C1);
    assert!(!calculation.recalculate_before_save());
    assert!(calculation.formulas_pending_recalculation());
}

#[test]
fn reads_poi_calculation_fixture() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/UncalcedRecord.xls");
    let workbook = Workbook::new(File::open(fixture).unwrap()).unwrap();
    let calculation = workbook.xls_worksheet(0).unwrap().calculation();
    assert_eq!(calculation.mode(), CalculationMode::Automatic);
    assert_eq!(calculation.maximum_iterations(), 100);
    assert_eq!(calculation.reference_mode(), ReferenceMode::A1);
    assert!(!calculation.iteration_enabled());
    assert!(calculation.recalculate_before_save());
}

#[test]
fn writer_rejects_invalid_calculation_bounds() {
    let mut writer = Writer::new();
    let invalid_count = CalculationSettings {
        maximum_iterations: 0,
        ..CalculationSettings::default()
    };
    assert!(writer.set_calculation_settings(invalid_count).is_err());
    let invalid_delta = CalculationSettings {
        iteration_delta: f64::NAN,
        ..CalculationSettings::default()
    };
    assert!(writer.set_calculation_settings(invalid_delta).is_err());
    assert!(MultithreadedCalculation::try_with_thread_count(true, 0).is_err());
    assert!(MultithreadedCalculation::try_with_thread_count(true, 1025).is_err());
}

#[test]
fn automatic_multithreaded_calculation_round_trip() {
    let mut writer = Writer::new();
    writer.add_worksheet("Automatic threading").unwrap();
    writer
        .set_calculation_settings(CalculationSettings {
            multithreaded_calculation: Some(MultithreadedCalculation::automatic(true)),
            ..CalculationSettings::default()
        })
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let threading = workbook.calculation().multithreaded_calculation().unwrap();
    assert!(threading.enabled());
    assert_eq!(threading.user_thread_count(), None);
}
