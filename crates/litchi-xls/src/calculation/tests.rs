//! Regression coverage for the BIFF8 calculation owner.

use super::model::{Mode, Multithreaded};
use super::package::{WorkbookCalculationCollector, WorksheetCalculationCollector};
use super::{
    CALC_COUNT_RECORD_TYPE, CALC_DELTA_RECORD_TYPE, CALC_ITER_RECORD_TYPE, CALC_MODE_RECORD_TYPE,
    CALC_PRECISION_RECORD_TYPE, CALC_REF_MODE_RECORD_TYPE, CALC_SAVE_RECALC_RECORD_TYPE,
    FORCE_FULL_CALCULATION_RECORD_TYPE, MTR_SETTINGS_RECORD_TYPE, RECALC_ID_RECORD_TYPE,
    UNCALCED_RECORD_TYPE,
};

fn complete_sheet(collector: &mut WorksheetCalculationCollector) {
    collector
        .feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
    collector
        .feed_record(CALC_COUNT_RECORD_TYPE, &100u16.to_le_bytes())
        .unwrap();
    collector
        .feed_record(CALC_REF_MODE_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
    collector
        .feed_record(CALC_ITER_RECORD_TYPE, &0u16.to_le_bytes())
        .unwrap();
    collector
        .feed_record(CALC_DELTA_RECORD_TYPE, &0.001f64.to_le_bytes())
        .unwrap();
    collector
        .feed_record(CALC_SAVE_RECALC_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
}

#[test]
fn rejects_malformed_lengths_values_and_reserved_fields() {
    let mut sheet = WorksheetCalculationCollector::new();
    assert!(sheet.feed_record(CALC_MODE_RECORD_TYPE, &[1]).is_err());
    let mut sheet = WorksheetCalculationCollector::new();
    assert!(
        sheet
            .feed_record(CALC_ITER_RECORD_TYPE, &2u16.to_le_bytes())
            .is_err()
    );
    let mut sheet = WorksheetCalculationCollector::new();
    assert!(
        sheet
            .feed_record(UNCALCED_RECORD_TYPE, &1u16.to_le_bytes())
            .is_err()
    );
    let mut globals = WorkbookCalculationCollector::new();
    let mut force = [0u8; 16];
    force[0..2].copy_from_slice(&FORCE_FULL_CALCULATION_RECORD_TYPE.to_le_bytes());
    force[2] = 1;
    assert!(
        globals
            .feed_record(FORCE_FULL_CALCULATION_RECORD_TYPE, &force)
            .is_err()
    );

    let mut globals = WorkbookCalculationCollector::new();
    let mut mtr = [0u8; 24];
    mtr[0..2].copy_from_slice(&MTR_SETTINGS_RECORD_TYPE.to_le_bytes());
    mtr[12..16].copy_from_slice(&1u32.to_le_bytes());
    mtr[16..20].copy_from_slice(&2u32.to_le_bytes());
    mtr[20..24].copy_from_slice(&4u32.to_le_bytes());
    assert!(globals.feed_record(MTR_SETTINGS_RECORD_TYPE, &mtr).is_err());

    let mut globals = WorkbookCalculationCollector::new();
    mtr[16..20].copy_from_slice(&1u32.to_le_bytes());
    mtr[20..24].copy_from_slice(&1025u32.to_le_bytes());
    assert!(globals.feed_record(MTR_SETTINGS_RECORD_TYPE, &mtr).is_err());
}

#[test]
fn rejects_duplicate_out_of_order_and_incomplete_blocks() {
    let mut duplicate = WorksheetCalculationCollector::new();
    duplicate
        .feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
    assert!(
        duplicate
            .feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes())
            .is_err()
    );
    let mut order = WorksheetCalculationCollector::new();
    order
        .feed_record(CALC_COUNT_RECORD_TYPE, &100u16.to_le_bytes())
        .unwrap();
    assert!(
        order
            .feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes())
            .is_err()
    );
    let mut incomplete = WorksheetCalculationCollector::new();
    incomplete
        .feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
    assert!(incomplete.finish().is_err());
}

#[test]
fn parses_complete_blocks_and_global_future_records() {
    let mut sheet = WorksheetCalculationCollector::new();
    sheet
        .feed_record(UNCALCED_RECORD_TYPE, &0u16.to_le_bytes())
        .unwrap();
    complete_sheet(&mut sheet);
    assert!(sheet.finish().unwrap().formulas_pending_recalculation());
    let mut globals = WorkbookCalculationCollector::new();
    globals
        .feed_record(CALC_PRECISION_RECORD_TYPE, &0u16.to_le_bytes())
        .unwrap();
    let mut mtr = [0u8; 24];
    mtr[0..2].copy_from_slice(&MTR_SETTINGS_RECORD_TYPE.to_le_bytes());
    mtr[12..16].copy_from_slice(&1u32.to_le_bytes());
    mtr[16..20].copy_from_slice(&1u32.to_le_bytes());
    mtr[20..24].copy_from_slice(&8u32.to_le_bytes());
    globals.feed_record(MTR_SETTINGS_RECORD_TYPE, &mtr).unwrap();
    let mut force = [0u8; 16];
    force[0..2].copy_from_slice(&FORCE_FULL_CALCULATION_RECORD_TYPE.to_le_bytes());
    force[12..16].copy_from_slice(&1u32.to_le_bytes());
    globals
        .feed_record(FORCE_FULL_CALCULATION_RECORD_TYPE, &force)
        .unwrap();
    let mut recalc = [0u8; 8];
    recalc[0..2].copy_from_slice(&RECALC_ID_RECORD_TYPE.to_le_bytes());
    recalc[4..8].copy_from_slice(&0x1234_5678u32.to_le_bytes());
    globals.feed_record(RECALC_ID_RECORD_TYPE, &recalc).unwrap();
    let calculation = globals.finish();
    assert!(!calculation.full_precision());
    assert_eq!(
        calculation.multithreaded_calculation(),
        Some(Multithreaded::try_with_thread_count(true, 8).unwrap())
    );
    assert!(calculation.force_full_calculation());
    assert_eq!(calculation.recalculation_engine_id(), Some(0x1234_5678));
}

#[test]
fn accepts_libreoffice_block_without_calc_mode() {
    let mut sheet = WorksheetCalculationCollector::new();
    sheet
        .feed_record(CALC_COUNT_RECORD_TYPE, &100u16.to_le_bytes())
        .unwrap();
    sheet
        .feed_record(CALC_REF_MODE_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
    sheet
        .feed_record(CALC_ITER_RECORD_TYPE, &0u16.to_le_bytes())
        .unwrap();
    sheet
        .feed_record(CALC_DELTA_RECORD_TYPE, &0.001f64.to_le_bytes())
        .unwrap();
    sheet
        .feed_record(CALC_SAVE_RECALC_RECORD_TYPE, &1u16.to_le_bytes())
        .unwrap();
    let calculation = sheet.finish().unwrap();
    assert_eq!(calculation.mode(), Mode::Automatic);
    assert_eq!(calculation.maximum_iterations(), 100);
}

#[test]
fn validates_user_specified_thread_counts() {
    assert!(Multithreaded::try_with_thread_count(true, 0).is_err());
    assert!(Multithreaded::try_with_thread_count(true, 1025).is_err());
    let automatic = Multithreaded::automatic(false);
    assert!(!automatic.enabled());
    assert_eq!(automatic.user_thread_count(), None);
}
