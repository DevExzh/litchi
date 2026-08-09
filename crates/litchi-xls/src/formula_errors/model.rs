//! Typed semantic values for BIFF8 formula error-checking features.

/// Header starting a worksheet formula error-checking feature collection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Header;

impl Header {
    #[must_use]
    pub const fn new() -> Self {
        Self
    }
}

/// Inclusive BIFF8 cell range targeted by a formula error feature.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_column: u8,
    pub(super) last_column: u8,
}

impl Range {
    #[must_use]
    pub const fn first_row(self) -> u16 {
        self.first_row
    }

    #[must_use]
    pub const fn last_row(self) -> u16 {
        self.last_row
    }

    #[must_use]
    pub const fn first_column(self) -> u8 {
        self.first_column
    }

    #[must_use]
    pub const fn last_column(self) -> u8 {
        self.last_column
    }
}

/// Formula conditions selected by an `FFErrorCheck` bit field.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Checks {
    pub(super) bits: u8,
}

impl Checks {
    #[must_use]
    pub const fn from_bits(bits: u8) -> Self {
        Self { bits }
    }

    #[must_use]
    pub const fn bits(self) -> u8 {
        self.bits
    }

    #[must_use]
    pub const fn calculation_errors(self) -> bool {
        self.bits & 0x01 != 0
    }

    #[must_use]
    pub const fn empty_cell_references(self) -> bool {
        self.bits & 0x02 != 0
    }

    #[must_use]
    pub const fn numbers_stored_as_text(self) -> bool {
        self.bits & 0x04 != 0
    }

    #[must_use]
    pub const fn inconsistent_ranges(self) -> bool {
        self.bits & 0x08 != 0
    }

    #[must_use]
    pub const fn inconsistent_formulas(self) -> bool {
        self.bits & 0x10 != 0
    }

    #[must_use]
    pub const fn insufficient_date_time_formats(self) -> bool {
        self.bits & 0x20 != 0
    }

    #[must_use]
    pub const fn unprotected_formulas(self) -> bool {
        self.bits & 0x40 != 0
    }

    #[must_use]
    pub const fn data_validation(self) -> bool {
        self.bits & 0x80 != 0
    }
}

/// One worksheet formula error-checking shared feature.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Feature {
    pub(super) ranges: Vec<Range>,
    pub(super) checks: Checks,
}

impl Feature {
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }

    #[must_use]
    pub const fn checks(&self) -> Checks {
        self.checks
    }
}
