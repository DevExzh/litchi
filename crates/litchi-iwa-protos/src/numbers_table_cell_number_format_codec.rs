//! Neutral strict plain-number format seam for Numbers table-cell adapters.
//!
//! The wire implementation remains in the audited popup/control codec so
//! there is one `FormatStructArchive` policy and one Buffa lazy parity view.
//! This module publishes only the scalar Number API and shared finite
//! execution types; generated values never cross the crate boundary.

#![allow(clippy::module_name_repetitions)]

pub use crate::numbers_table_cell_control_codec::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, MAX_NUMBER_DECIMAL_PLACES,
    NATIVE_AUTOMATIC_DECIMAL_PLACES, NATIVE_NUMBER_FORMAT_TYPE, NumberFormatSnapshot,
    NumberFormatWrite, PreparedNumberFormatRewrite, PreparedNumberFormatWrite,
    RewriteExecutionLimits, RewriteExecutionRequirements, RewriteOutput, canonical_number_format,
    decode_number_format, decode_number_format_with_report, prepare_number_format_append,
    prepare_number_format_rewrite, prepare_number_format_write, rewrite_number_format,
    rewrite_table_cell_number_format,
};
