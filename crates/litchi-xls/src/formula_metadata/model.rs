//! Semantic values for the BIFF8 `Formula` record metadata.

use std::sync::Arc;

use super::array::Owner as ArrayOwner;
use super::shared::Owner;

/// Calculation and preservation metadata surrounding one BIFF8 cell formula.
///
/// The flags are inert: this type never evaluates a formula, starts a
/// recalculation engine, or interprets the application-specific `chn` cache.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Metadata {
    always_calculate: bool,
    fill_alignment: bool,
    shared_formula: bool,
    clear_errors: bool,
    calculation_cache: u32,
    shared_owner: Option<Arc<Owner>>,
    array_owner: Option<Arc<ArrayOwner>>,
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
            shared_owner: None,
            array_owner: None,
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
    /// This low-level flag is preserved by the reader. For authoring, prefer
    /// [`Self::with_shared`], which binds the bit to a checked `ShrFmla`
    /// owner.
    pub const fn with_shared_formula(mut self, value: bool) -> Self {
        self.shared_formula = value;
        self
    }

    /// Bind the Formula metadata to a checked shared-formula owner.
    pub fn with_shared(mut self, owner: Owner) -> Self {
        self.shared_formula = true;
        self.shared_owner = Some(Arc::new(owner));
        self.array_owner = None;
        self
    }

    /// Bind this Formula record to one checked, inert BIFF8 array-formula
    /// owner. This clears shared-formula ownership because the wire formats
    /// are mutually exclusive.
    pub(crate) fn with_array(mut self, owner: ArrayOwner) -> Self {
        self.shared_formula = false;
        self.shared_owner = None;
        self.array_owner = Some(Arc::new(owner));
        self
    }

    pub(crate) fn attach_array(&mut self, owner: Arc<ArrayOwner>) {
        self.shared_formula = false;
        self.shared_owner = None;
        self.array_owner = Some(owner);
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

    pub const fn always_calculate(&self) -> bool {
        self.always_calculate
    }

    pub const fn fill_alignment(&self) -> bool {
        self.fill_alignment
    }

    pub const fn shared_formula(&self) -> bool {
        self.shared_formula
    }

    pub const fn clear_errors(&self) -> bool {
        self.clear_errors
    }

    pub const fn calculation_cache(&self) -> u32 {
        self.calculation_cache
    }

    /// The checked owner used to emit the following `ShrFmla`, if any.
    pub fn shared_owner(&self) -> Option<&Owner> {
        self.shared_owner.as_deref()
    }

    /// The checked owner of the following `Array` record, if this Formula is
    /// an array-formula participant.
    pub fn array_owner(&self) -> Option<&ArrayOwner> {
        self.array_owner.as_deref()
    }

    pub(crate) const fn from_wire(flags: u16, calculation_cache: u32) -> Self {
        Self {
            always_calculate: flags & 0x0001 != 0,
            fill_alignment: flags & 0x0004 != 0,
            shared_formula: flags & 0x0008 != 0,
            clear_errors: flags & 0x0020 != 0,
            calculation_cache,
            shared_owner: None,
            array_owner: None,
        }
    }

    pub(crate) const fn wire_flags(&self) -> u16 {
        (if self.always_calculate { 1 } else { 0 })
            | ((if self.fill_alignment { 1 } else { 0 }) << 2)
            | ((if self.shared_formula { 1 } else { 0 }) << 3)
            | ((if self.clear_errors { 1 } else { 0 }) << 5)
    }
}
