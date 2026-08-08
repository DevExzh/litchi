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

/// Workbook calculation properties, retaining whether each value was authored.
#[derive(Debug, Clone, Default)]
pub struct Properties {
    calculation_id: Option<u32>,
    calculation_mode: Option<Mode>,
    full_calculation_on_load: Option<bool>,
    reference_mode: Option<ReferenceMode>,
    iterative_calculation: Option<bool>,
    iteration_count: Option<u32>,
    iteration_delta: Option<f64>,
    full_precision: Option<bool>,
    calculation_completed: Option<bool>,
    calculate_on_save: Option<bool>,
    concurrent_calculation: Option<bool>,
    concurrent_manual_count: Option<u32>,
    force_full_calculation: Option<bool>,
}

impl Properties {
    /// Creates calculation properties with no explicitly authored values.
    pub const fn new() -> Self {
        Self {
            calculation_id: None,
            calculation_mode: None,
            full_calculation_on_load: None,
            reference_mode: None,
            iterative_calculation: None,
            iteration_count: None,
            iteration_delta: None,
            full_precision: None,
            calculation_completed: None,
            calculate_on_save: None,
            concurrent_calculation: None,
            concurrent_manual_count: None,
            force_full_calculation: None,
        }
    }

    /// Starts a builder which records explicitly authored values.
    pub const fn builder() -> Builder {
        Builder::new()
    }

    /// Returns the exact authored-value view, before effective defaults are applied.
    pub const fn specified(&self) -> Specified<'_> {
        Specified { properties: self }
    }

    /// Returns whether the same values were authored, including attribute presence.
    pub fn same_specification(&self, other: &Self) -> bool {
        self.calculation_id == other.calculation_id
            && self.calculation_mode == other.calculation_mode
            && self.full_calculation_on_load == other.full_calculation_on_load
            && self.reference_mode == other.reference_mode
            && self.iterative_calculation == other.iterative_calculation
            && self.iteration_count == other.iteration_count
            && same_specified_double(self.iteration_delta, other.iteration_delta)
            && self.full_precision == other.full_precision
            && self.calculation_completed == other.calculation_completed
            && self.calculate_on_save == other.calculate_on_save
            && self.concurrent_calculation == other.concurrent_calculation
            && self.concurrent_manual_count == other.concurrent_manual_count
            && self.force_full_calculation == other.force_full_calculation
    }

    /// Calculation-engine identifier. Excel's effective default is zero.
    pub fn calculation_id(&self) -> u32 {
        self.calculation_id.unwrap_or(0)
    }

    pub fn calculation_mode(&self) -> Mode {
        self.calculation_mode.unwrap_or_default()
    }

    pub fn full_calculation_on_load(&self) -> bool {
        self.full_calculation_on_load.unwrap_or(false)
    }

    pub fn reference_mode(&self) -> ReferenceMode {
        self.reference_mode.unwrap_or_default()
    }

    pub fn iterative_calculation(&self) -> bool {
        self.iterative_calculation.unwrap_or(false)
    }

    pub fn iteration_count(&self) -> u32 {
        self.iteration_count.unwrap_or(100)
    }

    pub fn iteration_delta(&self) -> f64 {
        self.iteration_delta.unwrap_or(0.001)
    }

    pub fn full_precision(&self) -> bool {
        self.full_precision.unwrap_or(true)
    }

    pub fn calculation_completed(&self) -> bool {
        self.calculation_completed.unwrap_or(true)
    }

    pub fn calculate_on_save(&self) -> bool {
        self.calculate_on_save.unwrap_or(true)
    }

    pub fn concurrent_calculation(&self) -> bool {
        self.concurrent_calculation.unwrap_or(true)
    }

    pub fn concurrent_manual_count(&self) -> Option<u32> {
        self.concurrent_manual_count
    }

    /// Whether Excel should perform a full calculation on the next calculation cycle.
    pub fn force_full_calculation(&self) -> bool {
        self.force_full_calculation.unwrap_or(false)
    }

    pub fn set_calculation_id(&mut self, value: Option<u32>) {
        self.calculation_id = value;
    }

    pub fn set_calculation_mode(&mut self, value: Option<Mode>) {
        self.calculation_mode = value;
    }

    pub fn set_full_calculation_on_load(&mut self, value: Option<bool>) {
        self.full_calculation_on_load = value;
    }

    pub fn set_reference_mode(&mut self, value: Option<ReferenceMode>) {
        self.reference_mode = value;
    }

    pub fn set_iterative_calculation(&mut self, value: Option<bool>) {
        self.iterative_calculation = value;
    }

    pub fn set_iteration_count(&mut self, value: Option<u32>) {
        self.iteration_count = value;
    }

    /// Sets the authored iteration delta.
    pub fn set_iteration_delta(&mut self, value: Option<f64>) -> Result<()> {
        self.iteration_delta = value;
        Ok(())
    }

    pub fn set_full_precision(&mut self, value: Option<bool>) {
        self.full_precision = value;
    }

    pub fn set_calculation_completed(&mut self, value: Option<bool>) {
        self.calculation_completed = value;
    }

    pub fn set_calculate_on_save(&mut self, value: Option<bool>) {
        self.calculate_on_save = value;
    }

    pub fn set_concurrent_calculation(&mut self, value: Option<bool>) {
        self.concurrent_calculation = value;
    }

    pub fn set_concurrent_manual_count(&mut self, value: Option<u32>) {
        self.concurrent_manual_count = value;
    }

    pub fn set_force_full_calculation(&mut self, value: Option<bool>) {
        self.force_full_calculation = value;
    }

    pub fn with_calculation_id(mut self, value: Option<u32>) -> Self {
        self.set_calculation_id(value);
        self
    }

    pub fn with_calculation_mode(mut self, value: Option<Mode>) -> Self {
        self.set_calculation_mode(value);
        self
    }

    pub fn with_full_calculation_on_load(mut self, value: Option<bool>) -> Self {
        self.set_full_calculation_on_load(value);
        self
    }

    pub fn with_reference_mode(mut self, value: Option<ReferenceMode>) -> Self {
        self.set_reference_mode(value);
        self
    }

    pub fn with_iterative_calculation(mut self, value: Option<bool>) -> Self {
        self.set_iterative_calculation(value);
        self
    }

    pub fn with_iteration_count(mut self, value: Option<u32>) -> Self {
        self.set_iteration_count(value);
        self
    }

    pub fn with_iteration_delta(mut self, value: Option<f64>) -> Result<Self> {
        self.set_iteration_delta(value)?;
        Ok(self)
    }

    pub fn with_full_precision(mut self, value: Option<bool>) -> Self {
        self.set_full_precision(value);
        self
    }

    pub fn with_calculation_completed(mut self, value: Option<bool>) -> Self {
        self.set_calculation_completed(value);
        self
    }

    pub fn with_calculate_on_save(mut self, value: Option<bool>) -> Self {
        self.set_calculate_on_save(value);
        self
    }

    pub fn with_concurrent_calculation(mut self, value: Option<bool>) -> Self {
        self.set_concurrent_calculation(value);
        self
    }

    pub fn with_concurrent_manual_count(mut self, value: Option<u32>) -> Self {
        self.set_concurrent_manual_count(value);
        self
    }

    pub fn with_force_full_calculation(mut self, value: Option<bool>) -> Self {
        self.set_force_full_calculation(value);
        self
    }
}

impl PartialEq for Properties {
    fn eq(&self, other: &Self) -> bool {
        self.calculation_id() == other.calculation_id()
            && self.calculation_mode() == other.calculation_mode()
            && self.full_calculation_on_load() == other.full_calculation_on_load()
            && self.reference_mode() == other.reference_mode()
            && self.iterative_calculation() == other.iterative_calculation()
            && self.iteration_count() == other.iteration_count()
            && self.iteration_delta() == other.iteration_delta()
            && self.full_precision() == other.full_precision()
            && self.calculation_completed() == other.calculation_completed()
            && self.calculate_on_save() == other.calculate_on_save()
            && self.concurrent_calculation() == other.concurrent_calculation()
            && self.concurrent_manual_count() == other.concurrent_manual_count()
            && self.force_full_calculation() == other.force_full_calculation()
    }
}

/// Borrowed view of the values explicitly authored on `calcPr`.
#[derive(Debug, Clone, Copy)]
pub struct Specified<'a> {
    properties: &'a Properties,
}

macro_rules! specified_getters {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {
        impl Specified<'_> {
            $(pub fn $name(self) -> Option<$ty> { self.properties.$name })+
        }
    };
}

specified_getters!(
    (calculation_id, u32),
    (calculation_mode, Mode),
    (full_calculation_on_load, bool),
    (reference_mode, ReferenceMode),
    (iterative_calculation, bool),
    (iteration_count, u32),
    (iteration_delta, f64),
    (full_precision, bool),
    (calculation_completed, bool),
    (calculate_on_save, bool),
    (concurrent_calculation, bool),
    (concurrent_manual_count, u32),
    (force_full_calculation, bool),
);

/// Builder for exact authored calculation-property values.
#[derive(Debug, Clone, Default)]
pub struct Builder {
    properties: Properties,
}

macro_rules! builder_methods {
    ($(($name:ident, $ty:ty, $field:ident)),+ $(,)?) => {
        impl Builder {
            $(pub fn $name(mut self, value: Option<$ty>) -> Self {
                self.properties.$field = value;
                self
            })+
        }
    };
}

impl Builder {
    pub const fn new() -> Self {
        Self {
            properties: Properties::new(),
        }
    }

    pub fn iteration_delta(mut self, value: Option<f64>) -> Result<Self> {
        self.properties.set_iteration_delta(value)?;
        Ok(self)
    }

    pub fn build(self) -> Properties {
        self.properties
    }
}

builder_methods!(
    (calculation_id, u32, calculation_id),
    (calculation_mode, Mode, calculation_mode),
    (full_calculation_on_load, bool, full_calculation_on_load),
    (reference_mode, ReferenceMode, reference_mode),
    (iterative_calculation, bool, iterative_calculation),
    (iteration_count, u32, iteration_count),
    (full_precision, bool, full_precision),
    (calculation_completed, bool, calculation_completed),
    (calculate_on_save, bool, calculate_on_save),
    (concurrent_calculation, bool, concurrent_calculation),
    (concurrent_manual_count, u32, concurrent_manual_count),
    (force_full_calculation, bool, force_full_calculation),
);

fn same_specified_double(left: Option<f64>, right: Option<f64>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.to_bits() == right.to_bits(),
        (None, None) => true,
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn equality_is_effective_but_specification_is_exact() {
        let omitted = Properties::new();
        let explicit = Properties::new()
            .with_calculation_mode(Some(Mode::Automatic))
            .with_iteration_count(Some(100));
        assert_eq!(omitted, explicit);
        assert!(!omitted.same_specification(&explicit));
        assert_eq!(omitted.specified().iteration_count(), None);
        assert_eq!(explicit.specified().iteration_count(), Some(100));
    }

    #[test]
    fn iteration_delta_accepts_the_xsd_double_value_space() {
        for value in [-1.0, 0.0, -0.0, f64::INFINITY, f64::NEG_INFINITY, f64::NAN] {
            let properties = Properties::new().with_iteration_delta(Some(value)).unwrap();
            assert_eq!(
                properties.specified().iteration_delta().unwrap().to_bits(),
                value.to_bits()
            );
        }
    }

    #[test]
    fn exact_delta_comparison_is_bit_deterministic() {
        let nan = Properties::new()
            .with_iteration_delta(Some(f64::NAN))
            .unwrap();
        assert!(nan.same_specification(&nan.clone()));

        let positive_zero = Properties::new().with_iteration_delta(Some(0.0)).unwrap();
        let negative_zero = Properties::new().with_iteration_delta(Some(-0.0)).unwrap();
        assert_eq!(positive_zero, negative_zero);
        assert!(!positive_zero.same_specification(&negative_zero));
    }
}
