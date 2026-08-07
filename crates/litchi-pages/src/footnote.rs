//! Archive-free Pages footnote and endnote formatter values.

pub mod body;

use thiserror::Error;

const RAW_KIND_FOOTNOTES: i32 = 0;
const RAW_KIND_DOCUMENT_ENDNOTES: i32 = 1;
const RAW_KIND_SECTION_ENDNOTES: i32 = 2;
const RAW_FORMAT_NUMERIC: i32 = 0;
const RAW_FORMAT_ROMAN: i32 = 1;
const RAW_FORMAT_SYMBOLIC: i32 = 2;
const RAW_FORMAT_JAPANESE_NUMERIC: i32 = 3;
const RAW_FORMAT_JAPANESE_IDEOGRAPHIC: i32 = 4;
const RAW_FORMAT_ARABIC_NUMERIC: i32 = 5;
const RAW_NUMBERING_CONTINUOUS: i32 = 0;
const RAW_NUMBERING_RESTART_EACH_PAGE: i32 = 1;
const RAW_NUMBERING_RESTART_EACH_SECTION: i32 = 2;

/// Validation failures for note formatter values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A known native note-kind value was represented as `Unknown`.
    #[error("Pages footnote kind must use its canonical variant for a known value")]
    NonCanonicalKind,
    /// A known native marker-format value was represented as `Unknown`.
    #[error("Pages footnote format must use its canonical variant for a known value")]
    NonCanonicalFormat,
    /// A known native numbering value was represented as `Unknown`.
    #[error("Pages footnote numbering must use its canonical variant for a known value")]
    NonCanonicalNumbering,
    /// A note gap exceeded the signed native range.
    #[error("Pages footnote gap exceeds the native signed integer range")]
    GapOutOfRange,
}

/// Result type for note formatter construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Where Pages places notes belonging to the document body.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// Notes at the bottom of each page.
    Footnotes,
    /// Notes collected at the end of the document.
    DocumentEndnotes,
    /// Notes collected at the end of each section.
    SectionEndnotes,
    /// A value written by a newer Pages version.
    Unknown(i32),
}

impl Kind {
    /// Decode a lossless native note-kind value.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            RAW_KIND_FOOTNOTES => Self::Footnotes,
            RAW_KIND_DOCUMENT_ENDNOTES => Self::DocumentEndnotes,
            RAW_KIND_SECTION_ENDNOTES => Self::SectionEndnotes,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown note-kind value without shadowing a named variant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalKind`] when `raw` is already assigned to a
    /// named native value.
    pub const fn unknown(raw: i32) -> Result<Self> {
        match raw {
            RAW_KIND_FOOTNOTES | RAW_KIND_DOCUMENT_ENDNOTES | RAW_KIND_SECTION_ENDNOTES => {
                Err(Error::NonCanonicalKind)
            },
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the lossless native value.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Footnotes => RAW_KIND_FOOTNOTES,
            Self::DocumentEndnotes => RAW_KIND_DOCUMENT_ENDNOTES,
            Self::SectionEndnotes => RAW_KIND_SECTION_ENDNOTES,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether an `Unknown` value is canonical for its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(
                RAW_KIND_FOOTNOTES | RAW_KIND_DOCUMENT_ENDNOTES | RAW_KIND_SECTION_ENDNOTES
            )
        )
    }
}

/// Marker sequence used for Pages footnotes and endnotes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Format {
    /// Arabic decimal digits.
    Numeric,
    /// Roman numerals.
    Roman,
    /// Symbolic markers.
    Symbolic,
    /// Japanese numeric markers.
    JapaneseNumeric,
    /// Japanese ideographic markers.
    JapaneseIdeographic,
    /// Arabic numeric markers with the native Arabic style.
    ArabicNumeric,
    /// A value written by a newer Pages version.
    Unknown(i32),
}

impl Format {
    /// Decode a lossless native marker-format value.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            RAW_FORMAT_NUMERIC => Self::Numeric,
            RAW_FORMAT_ROMAN => Self::Roman,
            RAW_FORMAT_SYMBOLIC => Self::Symbolic,
            RAW_FORMAT_JAPANESE_NUMERIC => Self::JapaneseNumeric,
            RAW_FORMAT_JAPANESE_IDEOGRAPHIC => Self::JapaneseIdeographic,
            RAW_FORMAT_ARABIC_NUMERIC => Self::ArabicNumeric,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown marker-format value without shadowing a named variant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalFormat`] when `raw` is already assigned to
    /// a named native value.
    pub const fn unknown(raw: i32) -> Result<Self> {
        match raw {
            RAW_FORMAT_NUMERIC
            | RAW_FORMAT_ROMAN
            | RAW_FORMAT_SYMBOLIC
            | RAW_FORMAT_JAPANESE_NUMERIC
            | RAW_FORMAT_JAPANESE_IDEOGRAPHIC
            | RAW_FORMAT_ARABIC_NUMERIC => Err(Error::NonCanonicalFormat),
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the lossless native value.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Numeric => RAW_FORMAT_NUMERIC,
            Self::Roman => RAW_FORMAT_ROMAN,
            Self::Symbolic => RAW_FORMAT_SYMBOLIC,
            Self::JapaneseNumeric => RAW_FORMAT_JAPANESE_NUMERIC,
            Self::JapaneseIdeographic => RAW_FORMAT_JAPANESE_IDEOGRAPHIC,
            Self::ArabicNumeric => RAW_FORMAT_ARABIC_NUMERIC,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether an `Unknown` value is canonical for its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(
                RAW_FORMAT_NUMERIC
                    | RAW_FORMAT_ROMAN
                    | RAW_FORMAT_SYMBOLIC
                    | RAW_FORMAT_JAPANESE_NUMERIC
                    | RAW_FORMAT_JAPANESE_IDEOGRAPHIC
                    | RAW_FORMAT_ARABIC_NUMERIC
            )
        )
    }
}

/// How Pages restarts footnote or endnote numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Numbering {
    /// Continue numbering through the whole note story.
    Continuous,
    /// Restart numbering on every page.
    RestartEachPage,
    /// Restart numbering for every section.
    RestartEachSection,
    /// A value written by a newer Pages version.
    Unknown(i32),
}

impl Numbering {
    /// Decode a lossless native note-numbering value.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            RAW_NUMBERING_CONTINUOUS => Self::Continuous,
            RAW_NUMBERING_RESTART_EACH_PAGE => Self::RestartEachPage,
            RAW_NUMBERING_RESTART_EACH_SECTION => Self::RestartEachSection,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown numbering value without shadowing a named variant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalNumbering`] when `raw` is already assigned
    /// to a named native value.
    pub const fn unknown(raw: i32) -> Result<Self> {
        match raw {
            RAW_NUMBERING_CONTINUOUS
            | RAW_NUMBERING_RESTART_EACH_PAGE
            | RAW_NUMBERING_RESTART_EACH_SECTION => Err(Error::NonCanonicalNumbering),
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the lossless native value.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Continuous => RAW_NUMBERING_CONTINUOUS,
            Self::RestartEachPage => RAW_NUMBERING_RESTART_EACH_PAGE,
            Self::RestartEachSection => RAW_NUMBERING_RESTART_EACH_SECTION,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether an `Unknown` value is canonical for its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(
                RAW_NUMBERING_CONTINUOUS
                    | RAW_NUMBERING_RESTART_EACH_PAGE
                    | RAW_NUMBERING_RESTART_EACH_SECTION
            )
        )
    }
}

/// Validated spacing between Pages footnotes or endnotes, in whole points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Gap(u32);

impl Gap {
    /// Construct a note gap in whole points.
    ///
    /// # Errors
    ///
    /// Returns [`Error::GapOutOfRange`] when the value cannot be represented by
    /// Pages' signed native field.
    pub fn new(points: u32) -> Result<Self> {
        i32::try_from(points)
            .map(|_| Self(points))
            .map_err(|_conversion_error| Error::GapOutOfRange)
    }

    /// Return the gap in whole points.
    #[must_use]
    pub const fn points(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for Gap {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<Gap> for u32 {
    fn from(value: Gap) -> Self {
        value.points()
    }
}

/// Lossless settings shown by Pages' Footnotes formatter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Settings {
    /// Where Pages places notes.
    pub kind: Option<Kind>,
    /// Marker sequence used for notes.
    pub format: Option<Format>,
    /// Numbering restart behavior.
    pub numbering: Option<Numbering>,
    /// Gap between the note marker and note text.
    pub gap: Option<Gap>,
}

impl Settings {
    /// Validate unknown discriminants before an archive adapter publishes them.
    ///
    /// # Errors
    ///
    /// Returns an error when a known native discriminant is represented by an
    /// `Unknown` variant.
    pub fn validate(self) -> Result<()> {
        if self.kind.is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalKind);
        }
        if self.format.is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalFormat);
        }
        if self.numbering.is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalNumbering);
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Format, Gap, Kind, Numbering, Settings};

    #[test]
    fn note_values_are_lossless_and_strict() {
        assert_eq!(Kind::from_raw(2), Kind::SectionEndnotes);
        assert_eq!(Format::from_raw(5), Format::ArabicNumeric);
        assert_eq!(Numbering::from_raw(9), Numbering::Unknown(9));
        assert_eq!(Kind::unknown(0), Err(Error::NonCanonicalKind));
        assert_eq!(Format::unknown(1), Err(Error::NonCanonicalFormat));
        assert_eq!(Numbering::unknown(2), Err(Error::NonCanonicalNumbering));
        assert_eq!(Gap::new(u32::MAX), Err(Error::GapOutOfRange));
        assert!(
            Settings {
                kind: Some(Kind::Unknown(0)),
                ..Settings::default()
            }
            .validate()
            .is_err()
        );
    }
}
