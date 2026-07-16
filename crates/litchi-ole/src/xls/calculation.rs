//! Typed BIFF8 calculation and recalculation records.

use super::{XlsError, XlsResult};

pub(crate) const CALC_COUNT_RECORD_TYPE: u16 = 0x000C;
pub(crate) const CALC_MODE_RECORD_TYPE: u16 = 0x000D;
pub(crate) const CALC_PRECISION_RECORD_TYPE: u16 = 0x000E;
pub(crate) const CALC_REF_MODE_RECORD_TYPE: u16 = 0x000F;
pub(crate) const CALC_DELTA_RECORD_TYPE: u16 = 0x0010;
pub(crate) const CALC_ITER_RECORD_TYPE: u16 = 0x0011;
pub(crate) const UNCALCED_RECORD_TYPE: u16 = 0x005E;
pub(crate) const CALC_SAVE_RECALC_RECORD_TYPE: u16 = 0x005F;
pub(crate) const RECALC_ID_RECORD_TYPE: u16 = 0x01C1;
pub(crate) const FORCE_FULL_CALCULATION_RECORD_TYPE: u16 = 0x08A3;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsCalculationMode {
    Manual,
    Automatic,
    AutomaticExceptTables,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsReferenceMode {
    R1C1,
    A1,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsWorkbookCalculation {
    full_precision: bool,
    force_full_calculation: bool,
    recalculation_engine_id: Option<u32>,
}

impl Default for XlsWorkbookCalculation {
    fn default() -> Self {
        Self { full_precision: true, force_full_calculation: false, recalculation_engine_id: None }
    }
}

impl XlsWorkbookCalculation {
    pub fn full_precision(&self) -> bool { self.full_precision }
    pub fn force_full_calculation(&self) -> bool { self.force_full_calculation }
    pub fn recalculation_engine_id(&self) -> Option<u32> { self.recalculation_engine_id }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsWorksheetCalculation {
    mode: XlsCalculationMode,
    maximum_iterations: u16,
    iteration_enabled: bool,
    iteration_delta: f64,
    reference_mode: XlsReferenceMode,
    recalculate_before_save: bool,
    formulas_pending_recalculation: bool,
}

impl Default for XlsWorksheetCalculation {
    fn default() -> Self {
        Self {
            mode: XlsCalculationMode::Automatic,
            maximum_iterations: 100,
            iteration_enabled: false,
            iteration_delta: 0.001,
            reference_mode: XlsReferenceMode::A1,
            recalculate_before_save: true,
            formulas_pending_recalculation: false,
        }
    }
}

impl XlsWorksheetCalculation {
    pub fn mode(&self) -> XlsCalculationMode { self.mode }
    pub fn maximum_iterations(&self) -> u16 { self.maximum_iterations }
    pub fn iteration_enabled(&self) -> bool { self.iteration_enabled }
    pub fn iteration_delta(&self) -> f64 { self.iteration_delta }
    pub fn reference_mode(&self) -> XlsReferenceMode { self.reference_mode }
    pub fn recalculate_before_save(&self) -> bool { self.recalculate_before_save }
    pub fn formulas_pending_recalculation(&self) -> bool { self.formulas_pending_recalculation }
}

pub(crate) struct WorkbookCalculationCollector {
    calculation: XlsWorkbookCalculation,
    precision_seen: bool,
    force_seen: bool,
    recalc_id_seen: bool,
    last_rank: Option<u8>,
}

impl WorkbookCalculationCollector {
    pub(crate) fn new() -> Self {
        Self {
            calculation: XlsWorkbookCalculation::default(),
            precision_seen: false,
            force_seen: false,
            recalc_id_seen: false,
            last_rank: None,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        let rank = match record_type {
            CALC_PRECISION_RECORD_TYPE => 0,
            FORCE_FULL_CALCULATION_RECORD_TYPE => 1,
            RECALC_ID_RECORD_TYPE => 2,
            _ => return Ok(()),
        };
        if self.last_rank.is_some_and(|previous| rank < previous) {
            return invalid(record_type, "calculation record is out of BIFF8 order");
        }
        self.last_rank = Some(rank);
        match record_type {
            CALC_PRECISION_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.precision_seen)?;
                self.calculation.full_precision = parse_bool16(record_type, data)?;
            },
            FORCE_FULL_CALCULATION_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.force_seen)?;
                require_length(record_type, data, 16)?;
                if read_u16(data, 0) != FORCE_FULL_CALCULATION_RECORD_TYPE {
                    return invalid(record_type, "future-record header type does not match containing record");
                }
                if read_u16(data, 2) != 0 || data[4..12].iter().any(|byte| *byte != 0) {
                    return invalid(record_type, "future-record flags and reserved bytes must be zero");
                }
                self.calculation.force_full_calculation = parse_bool32(record_type, &data[12..16])?;
            },
            RECALC_ID_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.recalc_id_seen)?;
                require_length(record_type, data, 8)?;
                if read_u16(data, 0) != RECALC_ID_RECORD_TYPE || read_u16(data, 2) != 0 {
                    return invalid(record_type, "RecalcId type must match and reserved field must be zero");
                }
                self.calculation.recalculation_engine_id = Some(read_u32(data, 4));
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> XlsWorkbookCalculation { self.calculation }
}

pub(crate) struct WorksheetCalculationCollector {
    calculation: XlsWorksheetCalculation,
    mode: Option<XlsCalculationMode>,
    maximum_iterations: Option<u16>,
    iteration_enabled: Option<bool>,
    iteration_delta: Option<f64>,
    reference_mode: Option<XlsReferenceMode>,
    recalculate_before_save: Option<bool>,
    uncalced_seen: bool,
    last_rank: Option<u8>,
}

impl WorksheetCalculationCollector {
    pub(crate) fn new() -> Self {
        Self {
            calculation: XlsWorksheetCalculation::default(),
            mode: None,
            maximum_iterations: None,
            iteration_enabled: None,
            iteration_delta: None,
            reference_mode: None,
            recalculate_before_save: None,
            uncalced_seen: false,
            last_rank: None,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        let rank = match record_type {
            UNCALCED_RECORD_TYPE => 0,
            CALC_MODE_RECORD_TYPE => 1,
            CALC_COUNT_RECORD_TYPE => 2,
            CALC_REF_MODE_RECORD_TYPE => 3,
            CALC_ITER_RECORD_TYPE => 4,
            CALC_DELTA_RECORD_TYPE => 5,
            CALC_SAVE_RECALC_RECORD_TYPE => 6,
            _ => return Ok(()),
        };
        if self.last_rank.is_some_and(|previous| rank < previous) {
            return invalid(record_type, "worksheet calculation record is out of BIFF8 order");
        }
        self.last_rank = Some(rank);
        match record_type {
            UNCALCED_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.uncalced_seen)?;
                require_length(record_type, data, 2)?;
                if read_u16(data, 0) != 0 {
                    return invalid(record_type, "Uncalced reserved field must be zero");
                }
                self.calculation.formulas_pending_recalculation = true;
            },
            CALC_MODE_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.mode)?;
                require_length(record_type, data, 2)?;
                self.mode = Some(match read_u16(data, 0) {
                    0 => XlsCalculationMode::Manual,
                    1 => XlsCalculationMode::Automatic,
                    2 => XlsCalculationMode::AutomaticExceptTables,
                    value => return invalid(record_type, format!("invalid calculation mode {value}")),
                });
            },
            CALC_COUNT_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.maximum_iterations)?;
                require_length(record_type, data, 2)?;
                let value = read_u16(data, 0);
                if !(1..=32_767).contains(&value) {
                    return invalid(record_type, format!("iteration count must be 1..=32767, got {value}"));
                }
                self.maximum_iterations = Some(value);
            },
            CALC_REF_MODE_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.reference_mode)?;
                self.reference_mode = Some(if parse_bool16(record_type, data)? {
                    XlsReferenceMode::A1
                } else {
                    XlsReferenceMode::R1C1
                });
            },
            CALC_ITER_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.iteration_enabled)?;
                self.iteration_enabled = Some(parse_bool16(record_type, data)?);
            },
            CALC_DELTA_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.iteration_delta)?;
                require_length(record_type, data, 8)?;
                let value = f64::from_le_bytes(data.try_into().unwrap());
                if !value.is_finite() || value < 0.0 {
                    return invalid(record_type, "iteration delta must be finite and non-negative");
                }
                self.iteration_delta = Some(value);
            },
            CALC_SAVE_RECALC_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.recalculate_before_save)?;
                self.recalculate_before_save = Some(parse_bool16(record_type, data)?);
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    pub(crate) fn finish(mut self) -> XlsResult<XlsWorksheetCalculation> {
        // LibreOffice BIFF8 files commonly omit CalcMode while emitting the
        // complete CalcCount..CalcSaveRecalc tail. Treat that producer variant
        // as automatic mode, but continue to reject every partial tail.
        let tail_present = [
            self.maximum_iterations.is_some(), self.reference_mode.is_some(),
            self.iteration_enabled.is_some(), self.iteration_delta.is_some(),
            self.recalculate_before_save.is_some(),
        ];
        if tail_present.iter().any(|value| *value)
            && !tail_present.iter().all(|value| *value)
        {
            return invalid(CALC_MODE_RECORD_TYPE, "worksheet calculation block is incomplete");
        }
        if self.mode.is_some() && !tail_present.iter().all(|value| *value) {
            return invalid(CALC_MODE_RECORD_TYPE, "worksheet calculation block is incomplete");
        }
        if tail_present.iter().all(|value| *value) {
            self.calculation.mode = self.mode.unwrap_or(XlsCalculationMode::Automatic);
            self.calculation.maximum_iterations = self.maximum_iterations.unwrap();
            self.calculation.reference_mode = self.reference_mode.unwrap();
            self.calculation.iteration_enabled = self.iteration_enabled.unwrap();
            self.calculation.iteration_delta = self.iteration_delta.unwrap();
            self.calculation.recalculate_before_save = self.recalculate_before_save.unwrap();
        }
        Ok(self.calculation)
    }
}

fn reject_duplicate(record_type: u16, seen: &mut bool) -> XlsResult<()> {
    if *seen { return invalid(record_type, "duplicate calculation record"); }
    *seen = true;
    Ok(())
}

fn reject_option_duplicate<T>(record_type: u16, value: &Option<T>) -> XlsResult<()> {
    if value.is_some() { return invalid(record_type, "duplicate calculation record"); }
    Ok(())
}

fn parse_bool16(record_type: u16, data: &[u8]) -> XlsResult<bool> {
    require_length(record_type, data, 2)?;
    match read_u16(data, 0) {
        0 => Ok(false), 1 => Ok(true),
        value => invalid(record_type, format!("Boolean must be 0 or 1, got {value}")),
    }
}

fn parse_bool32(record_type: u16, data: &[u8]) -> XlsResult<bool> {
    require_length(record_type, data, 4)?;
    match read_u32(data, 0) {
        0 => Ok(false), 1 => Ok(true),
        value => invalid(record_type, format!("Boolean must be 0 or 1, got {value}")),
    }
}

fn require_length(record_type: u16, data: &[u8], expected: usize) -> XlsResult<()> {
    if data.len() != expected {
        return invalid(record_type, format!("payload must be exactly {expected} bytes, got {}", data.len()));
    }
    Ok(())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
}
fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}
fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord { record_type, message: message.into() })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn complete_sheet(collector: &mut WorksheetCalculationCollector) {
        collector.feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
        collector.feed_record(CALC_COUNT_RECORD_TYPE, &100u16.to_le_bytes()).unwrap();
        collector.feed_record(CALC_REF_MODE_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
        collector.feed_record(CALC_ITER_RECORD_TYPE, &0u16.to_le_bytes()).unwrap();
        collector.feed_record(CALC_DELTA_RECORD_TYPE, &0.001f64.to_le_bytes()).unwrap();
        collector.feed_record(CALC_SAVE_RECALC_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
    }

    #[test]
    fn rejects_malformed_lengths_values_and_reserved_fields() {
        let mut sheet = WorksheetCalculationCollector::new();
        assert!(sheet.feed_record(CALC_MODE_RECORD_TYPE, &[1]).is_err());
        let mut sheet = WorksheetCalculationCollector::new();
        assert!(sheet.feed_record(CALC_ITER_RECORD_TYPE, &2u16.to_le_bytes()).is_err());
        let mut sheet = WorksheetCalculationCollector::new();
        assert!(sheet.feed_record(UNCALCED_RECORD_TYPE, &1u16.to_le_bytes()).is_err());
        let mut globals = WorkbookCalculationCollector::new();
        let mut force = [0u8; 16];
        force[0..2].copy_from_slice(&FORCE_FULL_CALCULATION_RECORD_TYPE.to_le_bytes());
        force[2] = 1;
        assert!(globals.feed_record(FORCE_FULL_CALCULATION_RECORD_TYPE, &force).is_err());
    }

    #[test]
    fn rejects_duplicate_out_of_order_and_incomplete_blocks() {
        let mut duplicate = WorksheetCalculationCollector::new();
        duplicate.feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
        assert!(duplicate.feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes()).is_err());
        let mut order = WorksheetCalculationCollector::new();
        order.feed_record(CALC_COUNT_RECORD_TYPE, &100u16.to_le_bytes()).unwrap();
        assert!(order.feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes()).is_err());
        let mut incomplete = WorksheetCalculationCollector::new();
        incomplete.feed_record(CALC_MODE_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
        assert!(incomplete.finish().is_err());
    }

    #[test]
    fn parses_complete_blocks_and_global_future_records() {
        let mut sheet = WorksheetCalculationCollector::new();
        sheet.feed_record(UNCALCED_RECORD_TYPE, &0u16.to_le_bytes()).unwrap();
        complete_sheet(&mut sheet);
        assert!(sheet.finish().unwrap().formulas_pending_recalculation());
        let mut globals = WorkbookCalculationCollector::new();
        globals.feed_record(CALC_PRECISION_RECORD_TYPE, &0u16.to_le_bytes()).unwrap();
        let mut force = [0u8; 16];
        force[0..2].copy_from_slice(&FORCE_FULL_CALCULATION_RECORD_TYPE.to_le_bytes());
        force[12..16].copy_from_slice(&1u32.to_le_bytes());
        globals.feed_record(FORCE_FULL_CALCULATION_RECORD_TYPE, &force).unwrap();
        let mut recalc = [0u8; 8];
        recalc[0..2].copy_from_slice(&RECALC_ID_RECORD_TYPE.to_le_bytes());
        recalc[4..8].copy_from_slice(&0x1234_5678u32.to_le_bytes());
        globals.feed_record(RECALC_ID_RECORD_TYPE, &recalc).unwrap();
        let calculation = globals.finish();
        assert!(!calculation.full_precision());
        assert!(calculation.force_full_calculation());
        assert_eq!(calculation.recalculation_engine_id(), Some(0x1234_5678));
    }

    #[test]
    fn accepts_libreoffice_block_without_calc_mode() {
        let mut sheet = WorksheetCalculationCollector::new();
        sheet.feed_record(CALC_COUNT_RECORD_TYPE, &100u16.to_le_bytes()).unwrap();
        sheet.feed_record(CALC_REF_MODE_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
        sheet.feed_record(CALC_ITER_RECORD_TYPE, &0u16.to_le_bytes()).unwrap();
        sheet.feed_record(CALC_DELTA_RECORD_TYPE, &0.001f64.to_le_bytes()).unwrap();
        sheet.feed_record(CALC_SAVE_RECALC_RECORD_TYPE, &1u16.to_le_bytes()).unwrap();
        let calculation = sheet.finish().unwrap();
        assert_eq!(calculation.mode(), XlsCalculationMode::Automatic);
        assert_eq!(calculation.maximum_iterations(), 100);
    }
}
