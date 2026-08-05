//! Typed workbook calculation policy.

use crate::error::{Result, invalid};

/// Workbook formula calculation mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Mode {
    Manual,
    #[default]
    Automatic,
    AutomaticExceptTables,
}

impl Mode {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "manual" => Ok(Self::Manual),
            "auto" => Ok(Self::Automatic),
            "autoNoTable" => Ok(Self::AutomaticExceptTables),
            _ => Err(invalid(format!("invalid calcPr calcMode '{value}'"))),
        }
    }
}

/// Cell-reference style used by formulas in the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ReferenceMode {
    #[default]
    A1,
    R1C1,
}

impl ReferenceMode {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "A1" => Ok(Self::A1),
            "R1C1" => Ok(Self::R1C1),
            _ => Err(invalid(format!("invalid calcPr refMode '{value}'"))),
        }
    }
}

/// Effective workbook calculation policy from `calcPr`.
#[derive(Debug, Clone, PartialEq)]
pub struct Properties {
    calculation_id: u32,
    calculation_mode: Mode,
    full_calculation_on_load: bool,
    reference_mode: ReferenceMode,
    iterative_calculation: bool,
    iteration_count: u32,
    iteration_delta: f64,
    full_precision: bool,
    calculation_completed: bool,
    calculate_on_save: bool,
    concurrent_calculation: bool,
    concurrent_manual_count: Option<u32>,
    force_full_calculation: bool,
}

impl Properties {
    pub(super) fn new(
        calculation_id: u32,
        calculation_mode: Mode,
        full_calculation_on_load: bool,
        reference_mode: ReferenceMode,
        iterative_calculation: bool,
        iteration_count: u32,
        iteration_delta: f64,
        full_precision: bool,
        calculation_completed: bool,
        calculate_on_save: bool,
        concurrent_calculation: bool,
        concurrent_manual_count: Option<u32>,
        force_full_calculation: bool,
    ) -> Self {
        Self {
            calculation_id,
            calculation_mode,
            full_calculation_on_load,
            reference_mode,
            iterative_calculation,
            iteration_count,
            iteration_delta,
            full_precision,
            calculation_completed,
            calculate_on_save,
            concurrent_calculation,
            concurrent_manual_count,
            force_full_calculation,
        }
    }

    /// Calculation-engine identifier. Excel's effective default is zero.
    pub fn calculation_id(&self) -> u32 {
        self.calculation_id
    }

    pub fn calculation_mode(&self) -> Mode {
        self.calculation_mode
    }

    pub fn full_calculation_on_load(&self) -> bool {
        self.full_calculation_on_load
    }

    pub fn reference_mode(&self) -> ReferenceMode {
        self.reference_mode
    }

    pub fn iterative_calculation(&self) -> bool {
        self.iterative_calculation
    }

    pub fn iteration_count(&self) -> u32 {
        self.iteration_count
    }

    pub fn iteration_delta(&self) -> f64 {
        self.iteration_delta
    }

    pub fn full_precision(&self) -> bool {
        self.full_precision
    }

    pub fn calculation_completed(&self) -> bool {
        self.calculation_completed
    }

    pub fn calculate_on_save(&self) -> bool {
        self.calculate_on_save
    }

    pub fn concurrent_calculation(&self) -> bool {
        self.concurrent_calculation
    }

    pub fn concurrent_manual_count(&self) -> Option<u32> {
        self.concurrent_manual_count
    }

    /// Whether Excel should perform a full calculation on the next calculation cycle.
    pub fn force_full_calculation(&self) -> bool {
        self.force_full_calculation
    }
}
