//! Semantic values for the BIFF8 `Formula` record metadata.

/// Calculation and preservation metadata surrounding one BIFF8 cell formula.
///
/// The flags are inert: this type never evaluates a formula, starts a
/// recalculation engine, or interprets the application-specific `chn` cache.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Metadata {
    always_calculate: bool,
    fill_alignment: bool,
    shared_formula: bool,
    clear_errors: bool,
    calculation_cache: u32,
}

impl Metadata {
    /// Construct metadata with all flags cleared and an empty application
    /// cache.
    pub const fn new() -> Self {
        Self {
            always_calculate: false,
            fill_alignment: false,
            shared_formula: false,
            clear_errors: false,
            calculation_cache: 0,
        }
    }

    /// Set whether the formula requests calculation during the next
    /// recalculation.
    pub const fn with_always_calculate(mut self, value: bool) -> Self {
        self.always_calculate = value;
        self
    }

    /// Set whether the cell has fill or center-across-selection alignment.
    pub const fn with_fill_alignment(mut self, value: bool) -> Self {
        self.fill_alignment = value;
        self
    }

    /// Set whether the formula is represented by a following `ShrFmla`
    /// record.
    ///
    /// The reader preserves this bit. The high-level writer currently refuses
    /// to author it because it does not yet own the corresponding `ShrFmla`
    /// sequence.
    pub const fn with_shared_formula(mut self, value: bool) -> Self {
        self.shared_formula = value;
        self
    }

    /// Set whether formula error checking is disabled for this cell.
    pub const fn with_clear_errors(mut self, value: bool) -> Self {
        self.clear_errors = value;
        self
    }

    /// Set the opaque application-specific calculation cache.
    pub const fn with_calculation_cache(mut self, value: u32) -> Self {
        self.calculation_cache = value;
        self
    }

    pub const fn always_calculate(self) -> bool {
        self.always_calculate
    }

    pub const fn fill_alignment(self) -> bool {
        self.fill_alignment
    }

    pub const fn shared_formula(self) -> bool {
        self.shared_formula
    }

    pub const fn clear_errors(self) -> bool {
        self.clear_errors
    }

    pub const fn calculation_cache(self) -> u32 {
        self.calculation_cache
    }

    pub(crate) const fn from_wire(flags: u16, calculation_cache: u32) -> Self {
        Self {
            always_calculate: flags & 0x0001 != 0,
            fill_alignment: flags & 0x0004 != 0,
            shared_formula: flags & 0x0008 != 0,
            clear_errors: flags & 0x0020 != 0,
            calculation_cache,
        }
    }

    pub(crate) const fn wire_flags(self) -> u16 {
        (if self.always_calculate { 1 } else { 0 })
            | ((if self.fill_alignment { 1 } else { 0 }) << 2)
            | ((if self.shared_formula { 1 } else { 0 }) << 3)
            | ((if self.clear_errors { 1 } else { 0 }) << 5)
    }
}
