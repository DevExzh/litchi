//! Archive-free Pages section pagination values.

use std::num::NonZeroU32;

use thiserror::Error;

const RAW_START_NEXT_PAGE: u32 = 0;
const RAW_START_RIGHT_PAGE: u32 = 1;
const RAW_START_LEFT_PAGE: u32 = 2;
const RAW_NUMBERING_CONTINUE: u32 = 0;
const RAW_NUMBERING_RESTART: u32 = 1;

/// Validation failures for section pagination values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A known native section-start value was represented as `Unknown`.
    #[error("Pages section start must use its canonical variant for a known value")]
    NonCanonicalStart,
    /// A known native page-numbering value was represented as `Unknown`.
    #[error("Pages page numbering must use its canonical variant for a known value")]
    NonCanonicalNumbering,
    /// A section page number was zero.
    #[error("Pages section page numbers must be greater than zero")]
    ZeroPageNumber,
}

/// Result type for section pagination construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Page on which a Pages section begins when facing pages are enabled.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Start {
    /// Begin on the next available page.
    NextPage,
    /// Begin on a right-hand page.
    RightPage,
    /// Begin on a left-hand page.
    LeftPage,
    /// A value written by a newer Pages version.
    Unknown(u32),
}

impl Start {
    /// Decode a lossless native section-start value.
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            RAW_START_NEXT_PAGE => Self::NextPage,
            RAW_START_RIGHT_PAGE => Self::RightPage,
            RAW_START_LEFT_PAGE => Self::LeftPage,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown start value without shadowing a named variant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalStart`] when `raw` is already assigned to
    /// a named native value.
    pub const fn unknown(raw: u32) -> Result<Self> {
        match raw {
            RAW_START_NEXT_PAGE | RAW_START_RIGHT_PAGE | RAW_START_LEFT_PAGE => {
                Err(Error::NonCanonicalStart)
            },
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the lossless native value.
    #[must_use]
    pub const fn as_raw(self) -> u32 {
        match self {
            Self::NextPage => RAW_START_NEXT_PAGE,
            Self::RightPage => RAW_START_RIGHT_PAGE,
            Self::LeftPage => RAW_START_LEFT_PAGE,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether an `Unknown` value is canonical for its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(RAW_START_NEXT_PAGE | RAW_START_RIGHT_PAGE | RAW_START_LEFT_PAGE)
        )
    }
}

/// Whether a Pages section continues or restarts page numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PageNumbering {
    /// Continue from the previous section.
    ContinueFromPrevious,
    /// Restart numbering at the section's configured starting page.
    Restart,
    /// A value written by a newer Pages version.
    Unknown(u32),
}

impl PageNumbering {
    /// Decode a lossless native page-numbering value.
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            RAW_NUMBERING_CONTINUE => Self::ContinueFromPrevious,
            RAW_NUMBERING_RESTART => Self::Restart,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown numbering value without shadowing a named variant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalNumbering`] when `raw` is already assigned
    /// to a named native value.
    pub const fn unknown(raw: u32) -> Result<Self> {
        match raw {
            RAW_NUMBERING_CONTINUE | RAW_NUMBERING_RESTART => Err(Error::NonCanonicalNumbering),
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the lossless native value.
    #[must_use]
    pub const fn as_raw(self) -> u32 {
        match self {
            Self::ContinueFromPrevious => RAW_NUMBERING_CONTINUE,
            Self::Restart => RAW_NUMBERING_RESTART,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether an `Unknown` value is canonical for its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(RAW_NUMBERING_CONTINUE | RAW_NUMBERING_RESTART)
        )
    }
}

/// A validated, non-zero Pages section page number.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PageNumber(NonZeroU32);

impl PageNumber {
    /// Validate and construct a section page number.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroPageNumber`] for zero.
    pub fn new(value: u32) -> Result<Self> {
        NonZeroU32::new(value)
            .map(Self)
            .ok_or(Error::ZeroPageNumber)
    }

    /// Return the numeric page number.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

impl TryFrom<u32> for PageNumber {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<PageNumber> for u32 {
    fn from(value: PageNumber) -> Self {
        value.get()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Error, PageNumber, PageNumbering, Start};

    #[test]
    fn native_values_round_trip_without_aliasing_unknowns() {
        assert_eq!(Start::from_raw(0), Start::NextPage);
        assert_eq!(Start::from_raw(1), Start::RightPage);
        assert_eq!(Start::from_raw(2), Start::LeftPage);
        assert_eq!(Start::from_raw(7), Start::Unknown(7));
        assert_eq!(
            PageNumbering::from_raw(0),
            PageNumbering::ContinueFromPrevious
        );
        assert_eq!(PageNumbering::from_raw(1), PageNumbering::Restart);
        assert_eq!(PageNumbering::from_raw(3), PageNumbering::Unknown(3));
        assert_eq!(Start::unknown(0), Err(Error::NonCanonicalStart));
        assert_eq!(PageNumbering::unknown(1), Err(Error::NonCanonicalNumbering));
        assert!(!Start::Unknown(0).is_canonical());
        assert!(!PageNumbering::Unknown(1).is_canonical());
    }

    #[test]
    fn page_number_is_non_zero_and_compact() {
        assert_eq!(size_of::<PageNumber>(), size_of::<u32>());
        assert_eq!(PageNumber::new(0), Err(Error::ZeroPageNumber));
        assert_eq!(PageNumber::new(42).map(PageNumber::get), Ok(42));
    }
}
