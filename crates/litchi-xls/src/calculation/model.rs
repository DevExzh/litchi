//! Semantic calculation settings exposed by the XLS owner.

use crate::Result;

use super::{MAX_CALCULATION_THREADS, MTR_SETTINGS_RECORD_TYPE, invalid};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    Manual,
    Automatic,
    AutomaticExceptTables,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceMode {
    R1C1,
    A1,
}

/// Multithreaded calculation settings stored in an `MTRSettings` record.
///
/// This is inert workbook metadata. The reader does not evaluate formulas or
/// create calculation threads.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Multithreaded {
    enabled: bool,
    user_thread_count: Option<u16>,
}

impl Multithreaded {
    /// Use the application's automatic thread count, with multithreaded
    /// calculation either enabled or disabled.
    pub const fn automatic(enabled: bool) -> Self {
        Self {
            enabled,
            user_thread_count: None,
        }
    }

    /// Use a user-specified calculation thread count.
    pub fn try_with_thread_count(enabled: bool, thread_count: u16) -> Result<Self> {
        if !(1..=MAX_CALCULATION_THREADS).contains(&thread_count) {
            return invalid(
                MTR_SETTINGS_RECORD_TYPE,
                format!(
                    "calculation thread count must be 1..={MAX_CALCULATION_THREADS}, got {thread_count}"
                ),
            );
        }
        Ok(Self {
            enabled,
            user_thread_count: Some(thread_count),
        })
    }

    pub const fn enabled(&self) -> bool {
        self.enabled
    }

    /// A user-specified thread count, or `None` when the producer selected the
    /// count automatically. The count is metadata when calculation is disabled.
    pub const fn user_thread_count(&self) -> Option<u16> {
        self.user_thread_count
    }

    pub(crate) const fn serialized_thread_count(&self) -> u16 {
        match self.user_thread_count {
            Some(value) => value,
            None => 1,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Workbook {
    pub(super) full_precision: bool,
    pub(super) multithreaded_calculation: Option<Multithreaded>,
    pub(super) force_full_calculation: bool,
    pub(super) recalculation_engine_id: Option<u32>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self {
            full_precision: true,
            multithreaded_calculation: None,
            force_full_calculation: false,
            recalculation_engine_id: None,
        }
    }
}

impl Workbook {
    pub fn full_precision(&self) -> bool {
        self.full_precision
    }
    pub fn multithreaded_calculation(&self) -> Option<Multithreaded> {
        self.multithreaded_calculation
    }
    pub fn force_full_calculation(&self) -> bool {
        self.force_full_calculation
    }
    pub fn recalculation_engine_id(&self) -> Option<u32> {
        self.recalculation_engine_id
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Worksheet {
    pub(super) mode: Mode,
    pub(super) maximum_iterations: u16,
    pub(super) iteration_enabled: bool,
    pub(super) iteration_delta: f64,
    pub(super) reference_mode: ReferenceMode,
    pub(super) recalculate_before_save: bool,
    pub(super) formulas_pending_recalculation: bool,
}

impl Default for Worksheet {
    fn default() -> Self {
        Self {
            mode: Mode::Automatic,
            maximum_iterations: 100,
            iteration_enabled: false,
            iteration_delta: 0.001,
            reference_mode: ReferenceMode::A1,
            recalculate_before_save: true,
            formulas_pending_recalculation: false,
        }
    }
}

impl Worksheet {
    pub fn mode(&self) -> Mode {
        self.mode
    }
    pub fn maximum_iterations(&self) -> u16 {
        self.maximum_iterations
    }
    pub fn iteration_enabled(&self) -> bool {
        self.iteration_enabled
    }
    pub fn iteration_delta(&self) -> f64 {
        self.iteration_delta
    }
    pub fn reference_mode(&self) -> ReferenceMode {
        self.reference_mode
    }
    pub fn recalculate_before_save(&self) -> bool {
        self.recalculate_before_save
    }
    pub fn formulas_pending_recalculation(&self) -> bool {
        self.formulas_pending_recalculation
    }
}
