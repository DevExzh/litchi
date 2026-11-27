//! Typed BIFF8 calculation and recalculation records.
//!
//! The owner keeps the public calculation facade at this module boundary;
//! semantic values, record codecs, and regression coverage live in focused
//! layers below it.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

use crate::error::{Error, Result};

pub(crate) const CALC_COUNT_RECORD_TYPE: u16 = 0x000C;
pub(crate) const CALC_MODE_RECORD_TYPE: u16 = 0x000D;
pub(crate) const CALC_PRECISION_RECORD_TYPE: u16 = 0x000E;
pub(crate) const CALC_REF_MODE_RECORD_TYPE: u16 = 0x000F;
pub(crate) const CALC_DELTA_RECORD_TYPE: u16 = 0x0010;
pub(crate) const CALC_ITER_RECORD_TYPE: u16 = 0x0011;
pub(crate) const UNCALCED_RECORD_TYPE: u16 = 0x005E;
pub(crate) const CALC_SAVE_RECALC_RECORD_TYPE: u16 = 0x005F;
pub(crate) const RECALC_ID_RECORD_TYPE: u16 = 0x01C1;
pub(crate) const MTR_SETTINGS_RECORD_TYPE: u16 = 0x089A;
pub(crate) const FORCE_FULL_CALCULATION_RECORD_TYPE: u16 = 0x08A3;
pub(crate) const MAX_CALCULATION_THREADS: u16 = 1024;

fn invalid<T>(record_type: u16, message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidRecord {
        record_type,
        message: message.into(),
    })
}

pub use model::{
    Mode as CalculationMode, Multithreaded as MultithreadedCalculation, ReferenceMode,
    Workbook as WorkbookCalculation, Worksheet as WorksheetCalculation,
};
pub(crate) use package::{WorkbookCalculationCollector, WorksheetCalculationCollector};
