//! Glyph mapping and OpenType subsetting.

use std::mem::size_of;

mod allsorts;
pub mod mapping;

use ::allsorts::{
    binary::budget::{self, BudgetDomain},
    cff::CFFError,
    error::{ParseError, ReadWriteError, WriteError},
    subset::SubsetError,
};

use crate::model::{FontData, FontError};

pub use allsorts::Allsorts;
pub use mapping::glyph_ids;

/// Resource bounds for one raw glyph-mapping or subsetting operation.
///
/// Both dimensions are required at the raw boundary. The working cap covers
/// source and glyph-ID staging as well as checked dependency allocations; the
/// output cap covers retained subset bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_working_bytes: usize,
    pub max_output_bytes: usize,
}

impl Limits {
    pub const fn new(max_working_bytes: usize, max_output_bytes: usize) -> Self {
        Self {
            max_working_bytes,
            max_output_bytes,
        }
    }
}

/// One crate-private bounded subsetting session shared by mapping and writing.
pub(crate) struct SubsetSession {
    scope: budget::Scope,
    budget: budget::Budget,
}

impl SubsetSession {
    pub(crate) fn new(limits: Limits) -> Self {
        let scope = budget::enter(limits.max_working_bytes, limits.max_output_bytes);
        let budget = budget::current().expect("subset budget scope is active");
        Self { scope, budget }
    }

    fn budget(&self) -> budget::Budget {
        let _scope = &self.scope;
        self.budget.clone()
    }

    pub(crate) fn charge_mapping_staging(
        &self,
        source_bytes: usize,
        requested_codepoints: u64,
    ) -> Result<(), FontError> {
        let requested_codepoints =
            usize::try_from(requested_codepoints).map_err(|_| self.staging_overflow())?;
        let glyph_slots = requested_codepoints
            .checked_add(1)
            .ok_or_else(|| self.staging_overflow())?;
        self.charge_staging_slots(source_bytes, glyph_slots)
    }

    pub(crate) fn charge_subset_staging(
        &self,
        source_bytes: usize,
        glyph_ids: usize,
    ) -> Result<(), FontError> {
        self.charge_staging_slots(source_bytes, glyph_ids)
    }

    fn charge_staging_slots(
        &self,
        source_bytes: usize,
        glyph_slots: usize,
    ) -> Result<(), FontError> {
        let glyph_id_bytes = glyph_slots
            .checked_mul(size_of::<u16>())
            .ok_or_else(|| self.staging_overflow())?;
        let bytes = source_bytes
            .checked_add(glyph_id_bytes)
            .ok_or_else(|| self.staging_overflow())?;
        self.budget()
            .charge(bytes)
            .map_err(|error| FontError::LimitExceeded {
                resource: "font subsetting working bytes",
                limit: error.limit,
                actual: error.requested,
            })
    }

    fn staging_overflow(&self) -> FontError {
        FontError::LimitExceeded {
            resource: "font subsetting working bytes",
            limit: self.budget.limit(),
            actual: usize::MAX,
        }
    }
}

/// A pluggable font-program reducer.
pub trait Subsetter {
    /// Reduce one font inside the caller-selected bounds.
    fn subset(
        &self,
        font: &FontData,
        glyph_ids: &[u16],
        limits: Limits,
    ) -> Result<Vec<u8>, FontError>;
}

fn budget_limit_error() -> Option<FontError> {
    budget::failure().map(|failure| FontError::LimitExceeded {
        resource: match failure.domain {
            BudgetDomain::WorkingBytes => "font subsetting working bytes",
            BudgetDomain::OutputBytes => "font subsetting output bytes",
            BudgetDomain::WorkUnits => "font subsetting work units",
        },
        limit: failure.limit,
        actual: failure.requested,
    })
}

pub(crate) fn parse_error(context: &str, error: ParseError) -> FontError {
    if matches!(error, ParseError::LimitExceeded) {
        if let Some(error) = budget_limit_error() {
            return error;
        }
    }
    FontError::SubsettingFailed(format!("{context}: {error}"))
}

pub(crate) fn write_error(context: &str, error: WriteError) -> FontError {
    if matches!(error, WriteError::LimitExceeded) {
        if let Some(error) = budget_limit_error() {
            return error;
        }
    }
    FontError::SubsettingFailed(format!("{context}: {error}"))
}

pub(crate) fn read_write_error(context: &str, error: ReadWriteError) -> FontError {
    match error {
        ReadWriteError::Read(error) => parse_error(context, error),
        ReadWriteError::Write(error) => write_error(context, error),
    }
}

pub(crate) fn subset_error(error: SubsetError) -> FontError {
    match error {
        SubsetError::Parse(error) => parse_error("font subsetting parse failed", error),
        SubsetError::Write(error) => write_error("font subsetting write failed", error),
        SubsetError::CFF(CFFError::ParseError(error)) => {
            parse_error("font subsetting CFF parse failed", error)
        },
        error => FontError::SubsettingFailed(format!("font subsetting failed: {error}")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_only_an_actual_working_budget_failure() {
        let _scope = budget::enter(4, usize::MAX);
        budget::current().unwrap().charge(5).unwrap_err();
        assert!(matches!(
            parse_error("parse", ParseError::LimitExceeded),
            FontError::LimitExceeded {
                resource: "font subsetting working bytes",
                limit: 4,
                actual: 5,
            }
        ));
    }

    #[test]
    fn maps_output_budget_failures_without_relabeling_structural_limits() {
        let _scope = budget::enter(usize::MAX, 4);
        budget::current().unwrap().charge_output(5).unwrap_err();
        assert!(matches!(
            write_error("write", WriteError::LimitExceeded),
            FontError::LimitExceeded {
                resource: "font subsetting output bytes",
                limit: 4,
                actual: 5,
            }
        ));

        let _scope = budget::enter(usize::MAX, usize::MAX);
        assert!(matches!(
            parse_error("parse", ParseError::LimitExceeded),
            FontError::SubsettingFailed(message) if message == "parse: limit exceeded"
        ));
    }

    #[test]
    fn staging_counts_mapping_slots_and_raw_ids_separately() {
        let source_bytes = 7;
        let mapping_bytes = source_bytes + (3 + 1) * size_of::<u16>();
        let mapping = SubsetSession::new(Limits::new(mapping_bytes, usize::MAX));
        assert!(mapping.charge_mapping_staging(source_bytes, 3).is_ok());

        let raw_bytes = source_bytes + 3 * size_of::<u16>();
        let raw = SubsetSession::new(Limits::new(raw_bytes, usize::MAX));
        assert!(raw.charge_subset_staging(source_bytes, 3).is_ok());
        let raw_tight = SubsetSession::new(Limits::new(raw_bytes - 1, usize::MAX));
        assert!(matches!(
            raw_tight.charge_subset_staging(source_bytes, 3),
            Err(FontError::LimitExceeded {
                resource: "font subsetting working bytes",
                actual,
                ..
            }) if actual == raw_bytes
        ));
    }

    #[test]
    fn oversized_mapping_count_is_a_typed_staging_refusal() {
        let session = SubsetSession::new(Limits::new(16, usize::MAX));
        assert!(matches!(
            session.charge_mapping_staging(0, u64::MAX),
            Err(FontError::LimitExceeded {
                resource: "font subsetting working bytes",
                actual: usize::MAX,
                ..
            })
        ));
    }
}
