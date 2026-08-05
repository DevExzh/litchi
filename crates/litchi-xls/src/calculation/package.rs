//! Ordered workbook and worksheet calculation-block assembly.

use crate::error::Result;

use super::codec::{
    parse_bool16, parse_bool32, read_u16, read_u32, require_future_record_header, require_length,
};
use super::model::{Mode, Multithreaded, ReferenceMode, Workbook, Worksheet};
use super::{
    CALC_COUNT_RECORD_TYPE, CALC_DELTA_RECORD_TYPE, CALC_ITER_RECORD_TYPE, CALC_MODE_RECORD_TYPE,
    CALC_PRECISION_RECORD_TYPE, CALC_REF_MODE_RECORD_TYPE, CALC_SAVE_RECALC_RECORD_TYPE,
    FORCE_FULL_CALCULATION_RECORD_TYPE, MAX_CALCULATION_THREADS, MTR_SETTINGS_RECORD_TYPE,
    RECALC_ID_RECORD_TYPE, UNCALCED_RECORD_TYPE, invalid,
};

pub(crate) struct WorkbookCalculationCollector {
    calculation: Workbook,
    precision_seen: bool,
    mtr_settings_seen: bool,
    force_seen: bool,
    recalc_id_seen: bool,
    last_rank: Option<u8>,
}

impl WorkbookCalculationCollector {
    pub(crate) fn new() -> Self {
        Self {
            calculation: Workbook::default(),
            precision_seen: false,
            mtr_settings_seen: false,
            force_seen: false,
            recalc_id_seen: false,
            last_rank: None,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        // BIFF8 overloads these future-record IDs. After the workbook
        // calculation block has advanced past their rank, 0x089A is
        // CompressPictures rather than MTRSettings and 0x08A3 is Compat12
        // rather than ForceFullCalculation. Their payloads are otherwise
        // indistinguishable, so workbook-global record position is the
        // required discriminator.
        if (record_type == MTR_SETTINGS_RECORD_TYPE
            && self.last_rank.is_some_and(|previous| previous > 1))
            || (record_type == FORCE_FULL_CALCULATION_RECORD_TYPE
                && self.last_rank.is_some_and(|previous| previous > 2))
        {
            return Ok(());
        }
        let rank = match record_type {
            CALC_PRECISION_RECORD_TYPE => 0,
            MTR_SETTINGS_RECORD_TYPE => 1,
            FORCE_FULL_CALCULATION_RECORD_TYPE => 2,
            RECALC_ID_RECORD_TYPE => 3,
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
            MTR_SETTINGS_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.mtr_settings_seen)?;
                require_future_record_header(record_type, data, 24)?;
                let enabled = parse_bool32(record_type, &data[12..16])?;
                let user_set_thread_count = parse_bool32(record_type, &data[16..20])?;
                let thread_count = read_u32(data, 20);
                if !(1..=u32::from(MAX_CALCULATION_THREADS)).contains(&thread_count) {
                    return invalid(
                        record_type,
                        format!(
                            "calculation thread count must be 1..={MAX_CALCULATION_THREADS}, got {thread_count}"
                        ),
                    );
                }
                self.calculation.multithreaded_calculation = Some(if user_set_thread_count {
                    Multithreaded::try_with_thread_count(enabled, thread_count as u16)?
                } else {
                    Multithreaded::automatic(enabled)
                });
            },
            FORCE_FULL_CALCULATION_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.force_seen)?;
                require_future_record_header(record_type, data, 16)?;
                self.calculation.force_full_calculation = parse_bool32(record_type, &data[12..16])?;
            },
            RECALC_ID_RECORD_TYPE => {
                reject_duplicate(record_type, &mut self.recalc_id_seen)?;
                require_length(record_type, data, 8)?;
                if read_u16(data, 0) != RECALC_ID_RECORD_TYPE || read_u16(data, 2) != 0 {
                    return invalid(
                        record_type,
                        "RecalcId type must match and reserved field must be zero",
                    );
                }
                self.calculation.recalculation_engine_id = Some(read_u32(data, 4));
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> Workbook {
        self.calculation
    }
}

pub(crate) struct WorksheetCalculationCollector {
    calculation: Worksheet,
    mode: Option<Mode>,
    maximum_iterations: Option<u16>,
    iteration_enabled: Option<bool>,
    iteration_delta: Option<f64>,
    reference_mode: Option<ReferenceMode>,
    recalculate_before_save: Option<bool>,
    uncalced_seen: bool,
    last_rank: Option<u8>,
}

impl WorksheetCalculationCollector {
    pub(crate) fn new() -> Self {
        Self {
            calculation: Worksheet::default(),
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

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
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
            return invalid(
                record_type,
                "worksheet calculation record is out of BIFF8 order",
            );
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
                    0 => Mode::Manual,
                    1 => Mode::Automatic,
                    2 => Mode::AutomaticExceptTables,
                    value => {
                        return invalid(record_type, format!("invalid calculation mode {value}"));
                    },
                });
            },
            CALC_COUNT_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.maximum_iterations)?;
                require_length(record_type, data, 2)?;
                let value = read_u16(data, 0);
                if !(1..=32_767).contains(&value) {
                    return invalid(
                        record_type,
                        format!("iteration count must be 1..=32767, got {value}"),
                    );
                }
                self.maximum_iterations = Some(value);
            },
            CALC_REF_MODE_RECORD_TYPE => {
                reject_option_duplicate(record_type, &self.reference_mode)?;
                self.reference_mode = Some(if parse_bool16(record_type, data)? {
                    ReferenceMode::A1
                } else {
                    ReferenceMode::R1C1
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
                    return invalid(
                        record_type,
                        "iteration delta must be finite and non-negative",
                    );
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

    pub(crate) fn finish(mut self) -> Result<Worksheet> {
        // LibreOffice BIFF8 files commonly omit CalcMode while emitting the
        // complete CalcCount..CalcSaveRecalc tail. Treat that producer variant
        // as automatic mode, but continue to reject every partial tail.
        let tail_present = [
            self.maximum_iterations.is_some(),
            self.reference_mode.is_some(),
            self.iteration_enabled.is_some(),
            self.iteration_delta.is_some(),
            self.recalculate_before_save.is_some(),
        ];
        if tail_present.iter().any(|value| *value) && !tail_present.iter().all(|value| *value) {
            return invalid(
                CALC_MODE_RECORD_TYPE,
                "worksheet calculation block is incomplete",
            );
        }
        if self.mode.is_some() && !tail_present.iter().all(|value| *value) {
            return invalid(
                CALC_MODE_RECORD_TYPE,
                "worksheet calculation block is incomplete",
            );
        }
        if tail_present.iter().all(|value| *value) {
            self.calculation.mode = self.mode.unwrap_or(Mode::Automatic);
            self.calculation.maximum_iterations = self.maximum_iterations.unwrap();
            self.calculation.reference_mode = self.reference_mode.unwrap();
            self.calculation.iteration_enabled = self.iteration_enabled.unwrap();
            self.calculation.iteration_delta = self.iteration_delta.unwrap();
            self.calculation.recalculate_before_save = self.recalculate_before_save.unwrap();
        }
        Ok(self.calculation)
    }
}

fn reject_duplicate(record_type: u16, seen: &mut bool) -> Result<()> {
    if *seen {
        return invalid(record_type, "duplicate calculation record");
    }
    *seen = true;
    Ok(())
}

fn reject_option_duplicate<T>(record_type: u16, value: &Option<T>) -> Result<()> {
    if value.is_some() {
        return invalid(record_type, "duplicate calculation record");
    }
    Ok(())
}
