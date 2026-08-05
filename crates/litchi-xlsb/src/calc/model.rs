//! Typed calculation-property values for `BrtCalcProp`.

use std::num::FpCategory;

use bitflags::bitflags;

use super::{Error, Result};

const DEFAULT_ID: u32 = 0x0001_EB1D;
const DEFAULT_ITERS: u32 = 100;
const DEFAULT_DELTA: f64 = 0.001;

/// Workbook calculation mode from `[MS-XLSB]` section 2.4.318.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u32)]
pub enum Mode {
    /// Recalculate only when explicitly requested.
    Manual = 0,
    /// Recalculate formulas automatically.
    #[default]
    Auto = 1,
    /// Recalculate formulas automatically, excluding tables.
    AutoNoTables = 2,
}

impl Mode {
    pub(super) fn read(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Manual),
            1 => Ok(Self::Auto),
            2 => Ok(Self::AutoNoTables),
            value => Err(Error::Mode { value }),
        }
    }

    pub(super) const fn wire(self) -> u32 {
        self as u32
    }
}

bitflags! {
    /// Compact calculation switches from `[MS-XLSB]` section 2.4.318.
    ///
    /// Unknown/reserved bits are rejected by [`Props::set_opts`] and [`crate::calc::read`].
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Opts: u16 {
        /// Perform a full calculation when the workbook opens.
        const FULL_ON_LOAD = 0x0001;
        /// Use A1 rather than R1C1 references.
        const A1 = 0x0002;
        /// Enable iterative calculation.
        const ITERATE = 0x0004;
        /// Do not use precision-as-displayed mode.
        const FULL_PRECISION = 0x0008;
        /// Calculation was incomplete when the workbook was saved.
        const UNCALCULATED = 0x0010;
        /// Recalculate before saving in manual mode.
        const RECALC_ON_SAVE = 0x0020;
        /// Enable concurrent calculation processes.
        const MTR = 0x0040;
        /// Use the workbook's explicit thread count.
        const USER_THREADS = 0x0080;
        /// Ignore dependencies and fully calculate every formula.
        const IGNORE_DEPS = 0x0100;
    }
}

/// A checked `Xnum` iterative-calculation delta.
///
/// `[MS-XLSB]` section 2.5.172 excludes infinity, subnormal values, NaN, and
/// negative zero. The exact accepted IEEE-754 bits are retained.
#[derive(Debug, Clone, Copy, PartialEq)]
#[repr(transparent)]
pub struct Delta(f64);

impl Delta {
    /// Validate one `xnumDelta` value.
    pub fn new(value: f64) -> Result<Self> {
        if matches!(
            value.classify(),
            FpCategory::Nan | FpCategory::Infinite | FpCategory::Subnormal
        ) || (value == 0.0 && value.is_sign_negative())
        {
            return Err(Error::Delta {
                bits: value.to_bits(),
            });
        }
        Ok(Self(value))
    }

    /// Return the exact validated value.
    #[must_use]
    pub const fn get(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for Delta {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::new(value)
    }
}

impl From<Delta> for f64 {
    fn from(value: Delta) -> Self {
        value.get()
    }
}

/// A checked concurrent-calculation process count.
///
/// `[MS-XLSB]` section 2.4.318 requires `cUserThreadCount` to be in
/// `1..=1024`, regardless of whether its option bits currently use it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Threads(u16);

impl Threads {
    /// Validate a concurrent process count.
    pub fn new(value: u16) -> Result<Self> {
        if !(1..=1024).contains(&value) {
            return Err(Error::Threads {
                value: i32::from(value),
            });
        }
        Ok(Self(value))
    }

    /// Return the validated process count.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }

    pub(super) fn from_wire(value: i32) -> Result<Self> {
        if !(1..=1024).contains(&value) {
            return Err(Error::Threads { value });
        }
        let value = u16::try_from(value).map_err(|_| Error::Threads { value })?;
        Ok(Self(value))
    }
}

impl TryFrom<u16> for Threads {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self> {
        Self::new(value)
    }
}

impl From<Threads> for u16 {
    fn from(value: Threads) -> Self {
        value.get()
    }
}

/// Validated workbook calculation properties.
///
/// Private fields prevent mutation into an unwritable state. Options occupy a
/// single `u16`; the complete value has an explicit structural size bound in
/// the module tests.
#[derive(Debug, Clone, PartialEq)]
pub struct Props {
    delta: Delta,
    id: u32,
    iters: u32,
    mode: Mode,
    threads: Threads,
    opts: Opts,
}

impl Props {
    /// Construct the same valid defaults used by [`Default`].
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Recalculation engine identifier.
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// Set the recalculation engine identifier.
    pub fn set_id(&mut self, id: u32) -> &mut Self {
        self.id = id;
        self
    }

    /// Set the recalculation engine identifier while consuming `self`.
    #[must_use]
    pub fn with_id(mut self, id: u32) -> Self {
        self.set_id(id);
        self
    }

    /// Workbook calculation mode.
    #[must_use]
    pub const fn mode(&self) -> Mode {
        self.mode
    }

    /// Set the workbook calculation mode.
    pub fn set_mode(&mut self, mode: Mode) -> &mut Self {
        self.mode = mode;
        self
    }

    /// Set the workbook calculation mode while consuming `self`.
    #[must_use]
    pub fn with_mode(mut self, mode: Mode) -> Self {
        self.set_mode(mode);
        self
    }

    /// Maximum iterative-calculation pass count.
    #[must_use]
    pub const fn iters(&self) -> u32 {
        self.iters
    }

    /// Set the iterative-calculation pass count.
    pub fn set_iters(&mut self, iters: u32) -> &mut Self {
        self.iters = iters;
        self
    }

    /// Set the iterative-calculation pass count while consuming `self`.
    #[must_use]
    pub fn with_iters(mut self, iters: u32) -> Self {
        self.set_iters(iters);
        self
    }

    /// Minimum iterative-calculation change.
    #[must_use]
    pub const fn delta(&self) -> Delta {
        self.delta
    }

    /// Set a previously checked iterative-calculation delta.
    pub fn set_delta(&mut self, delta: Delta) -> &mut Self {
        self.delta = delta;
        self
    }

    /// Set a checked delta while consuming `self`.
    #[must_use]
    pub fn with_delta(mut self, delta: Delta) -> Self {
        self.set_delta(delta);
        self
    }

    /// Configured concurrent-calculation process count.
    #[must_use]
    pub const fn threads(&self) -> Threads {
        self.threads
    }

    /// Set a previously checked process count.
    pub fn set_threads(&mut self, threads: Threads) -> &mut Self {
        self.threads = threads;
        self
    }

    /// Set a checked process count while consuming `self`.
    #[must_use]
    pub fn with_threads(mut self, threads: Threads) -> Self {
        self.set_threads(threads);
        self
    }

    /// All compact calculation switches.
    #[must_use]
    pub const fn opts(&self) -> Opts {
        self.opts
    }

    /// Whether every switch in `opts` is enabled.
    #[must_use]
    pub const fn has(&self, opts: Opts) -> bool {
        self.opts.contains(opts)
    }

    /// Replace the compact switches after rejecting retained reserved bits.
    pub fn set_opts(&mut self, opts: Opts) -> Result<&mut Self> {
        validate_opts(opts)?;
        self.opts = opts;
        Ok(self)
    }

    /// Replace the compact switches while consuming `self`.
    pub fn with_opts(mut self, opts: Opts) -> Result<Self> {
        self.set_opts(opts)?;
        Ok(self)
    }

    /// Enable or disable one or more switches.
    pub fn set_opt(&mut self, opts: Opts, enabled: bool) -> Result<&mut Self> {
        validate_opts(opts)?;
        self.opts.set(opts, enabled);
        Ok(self)
    }

    /// Enable or disable switches while consuming `self`.
    pub fn with_opt(mut self, opts: Opts, enabled: bool) -> Result<Self> {
        self.set_opt(opts, enabled)?;
        Ok(self)
    }

    pub(super) fn from_wire(
        id: u32,
        mode: Mode,
        iters: u32,
        delta: Delta,
        threads: Threads,
        opts: Opts,
    ) -> Self {
        Self {
            delta,
            id,
            iters,
            mode,
            threads,
            opts,
        }
    }
}

impl Default for Props {
    fn default() -> Self {
        Self {
            delta: Delta(DEFAULT_DELTA),
            id: DEFAULT_ID,
            iters: DEFAULT_ITERS,
            mode: Mode::Auto,
            threads: Threads(1),
            opts: Opts::A1 | Opts::FULL_PRECISION | Opts::RECALC_ON_SAVE | Opts::MTR,
        }
    }
}

pub(super) fn validate_opts(opts: Opts) -> Result<()> {
    let bits = opts.bits();
    if bits & !Opts::all().bits() != 0 {
        return Err(Error::ReservedOpts {
            bits: bits & !Opts::all().bits(),
        });
    }
    Ok(())
}
