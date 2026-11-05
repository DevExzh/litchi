//! Caller-selected source and semantic limits.
//!
//! Every reader in this crate takes untrusted input, so each bounds the
//! bytes it reads ([`SourceLimits`]) and the cells it materializes
//! ([`SemanticLimits`]). A sparse SYLK or DIF file can address a far cell
//! with a few bytes; the cell bound counts every cell up to it.

use std::io::Read;

use crate::error::{Error, Result};

/// Default maximum input: 512 MiB, the shared source-byte default.
pub const DEFAULT_INPUT_BYTES: usize = 512 * 1024 * 1024;
/// Default maximum materialized cells: 16,777,216 (2^24).
pub const DEFAULT_CELLS: usize = 1 << 24;

/// A physical ingress dimension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Bytes read from the source.
    InputBytes,
}

impl SourceLimitKind {
    /// The dimension noun: `input_bytes`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::InputBytes => "input_bytes",
        }
    }
}

/// A semantic dimension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Cells materialized in the row grid, empty padding included.
    Cells,
}

impl SemanticLimitKind {
    /// The dimension noun: `cells`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Cells => "cells",
        }
    }
}

/// The dimension of an [`Error::Limit`] or [`Error::InvalidLimit`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// A [`SourceLimits`] dimension.
    Source(SourceLimitKind),
    /// A [`SemanticLimits`] dimension.
    Semantic(SemanticLimitKind),
}

impl LimitKind {
    /// The dimension noun.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Source(kind) => kind.as_str(),
            Self::Semantic(kind) => kind.as_str(),
        }
    }
}

/// Physical ingress limits.
///
/// ```
/// use litchi_sheet_text::SourceLimits;
///
/// let limits = SourceLimits::builder().max_input_bytes(4096)?.build()?;
/// assert_eq!(limits.max_input_bytes(), 4096);
/// assert!(SourceLimits::builder().max_input_bytes(0).is_err());
/// # Ok::<(), litchi_sheet_text::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: usize,
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self {
            input_bytes: DEFAULT_INPUT_BYTES,
        }
    }
}

impl SourceLimits {
    /// Start from the defaults.
    #[must_use]
    pub fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder(Self::default())
    }

    /// Maximum bytes read from a source.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Read `reader` to its end, refusing more than `max_input_bytes`.
    pub(crate) fn read(self, reader: impl Read) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        let ceiling = u64::try_from(self.input_bytes)
            .unwrap_or(u64::MAX)
            .saturating_add(1);
        reader.take(ceiling).read_to_end(&mut bytes)?;
        self.check(bytes.len())?;
        Ok(bytes)
    }

    /// Refuse an input longer than `max_input_bytes`.
    pub(crate) fn check(self, observed: usize) -> Result<()> {
        if observed > self.input_bytes {
            return Err(Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                observed,
                maximum: self.input_bytes,
            });
        }
        Ok(())
    }
}

/// Consuming builder of checked [`SourceLimits`].
#[derive(Debug, Clone, Copy)]
pub struct SourceLimitsBuilder(SourceLimits);

impl SourceLimitsBuilder {
    /// Set the maximum bytes read from a source.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for zero or a value above `isize::MAX`.
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self> {
        check_value(
            LimitKind::Source(SourceLimitKind::InputBytes),
            value,
            isize::MAX.unsigned_abs(),
        )?;
        self.0.input_bytes = value;
        Ok(self)
    }

    /// Finish the checked limits.
    ///
    /// # Errors
    ///
    /// None today: every setter validates its value. The signature keeps
    /// room for cross-dimension checks.
    pub fn build(self) -> Result<SourceLimits> {
        Ok(self.0)
    }
}

/// Semantic limits on what a reader materializes.
///
/// ```
/// use litchi_sheet_text::SemanticLimits;
///
/// let limits = SemanticLimits::builder().max_cells(100)?.build()?;
/// assert_eq!(limits.max_cells(), 100);
/// # Ok::<(), litchi_sheet_text::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    cells: usize,
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            cells: DEFAULT_CELLS,
        }
    }
}

impl SemanticLimits {
    /// Start from the defaults.
    #[must_use]
    pub fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder(Self::default())
    }

    /// Maximum cells materialized in the row grid, empty padding included.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.cells
    }

    /// Refuse a grid of more than `max_cells` cells.
    pub(crate) fn check_cells(self, observed: usize) -> Result<()> {
        if observed > self.cells {
            return Err(Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::Cells),
                observed,
                maximum: self.cells,
            });
        }
        Ok(())
    }
}

/// Consuming builder of checked [`SemanticLimits`].
#[derive(Debug, Clone, Copy)]
pub struct SemanticLimitsBuilder(SemanticLimits);

impl SemanticLimitsBuilder {
    /// Set the maximum materialized cells.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for zero or a value above `isize::MAX`.
    pub fn max_cells(mut self, value: usize) -> Result<Self> {
        check_value(
            LimitKind::Semantic(SemanticLimitKind::Cells),
            value,
            isize::MAX.unsigned_abs(),
        )?;
        self.0.cells = value;
        Ok(self)
    }

    /// Finish the checked limits.
    ///
    /// # Errors
    ///
    /// None today: every setter validates its value.
    pub fn build(self) -> Result<SemanticLimits> {
        Ok(self.0)
    }
}

fn check_value(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

/// A running count of materialized cells against [`SemanticLimits`].
#[derive(Debug)]
pub(crate) struct CellBudget {
    used: usize,
    limits: SemanticLimits,
}

impl CellBudget {
    pub(crate) const fn new(limits: SemanticLimits) -> Self {
        Self { used: 0, limits }
    }

    /// Charge `cells` more cells.
    pub(crate) fn charge(&mut self, cells: usize) -> Result<()> {
        self.used = self.used.saturating_add(cells);
        self.limits.check_cells(self.used)
    }
}
