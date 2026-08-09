//! Document-level footnote and endnote configuration.

use crate::{RtfError, RtfResult};

/// Kinds of notes declared present by `\fet`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresentNoteKinds {
    FootnotesOnly,
    EndnotesOnly,
    FootnotesAndEndnotes,
}

/// Placement of footnotes or endnotes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NotePlacement {
    EndOfSection,
    EndOfDocument,
    BeneathText,
    BottomOfPage,
}

/// Footnote numbering restart policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FootnoteRestart {
    Continuous,
    EachSection,
    EachPage,
}

/// Endnote numbering restart policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EndnoteRestart {
    Continuous,
    EachSection,
}

/// RTF footnote/endnote numbering style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteNumberingStyle {
    Arabic,
    LowercaseLetter,
    UppercaseLetter,
    LowercaseRoman,
    UppercaseRoman,
    Chicago,
    KoreanChosung,
    Circle,
    KanjiDigitless,
    KanjiWithDigit,
    KanjiThree,
    KanjiFour,
    DoubleByte,
    KoreanGanada,
    ChineseOne,
    ChineseTwo,
    ChineseThree,
    ChineseFour,
    ZodiacOne,
    ZodiacTwo,
    ZodiacThree,
}

/// Explicit document-level note settings. `None` preserves omission from the source.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct NoteOptions {
    pub present_kinds: Option<PresentNoteKinds>,
    pub footnote_placement: Option<NotePlacement>,
    pub endnote_placement: Option<NotePlacement>,
    pub footnote_start: Option<i32>,
    pub endnote_start: Option<i32>,
    pub footnote_restart: Option<FootnoteRestart>,
    pub endnote_restart: Option<EndnoteRestart>,
    pub footnote_numbering: Option<NoteNumberingStyle>,
    pub endnote_numbering: Option<NoteNumberingStyle>,
}

impl NoteOptions {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.footnote_start.is_some_and(|value| value <= 0)
            || self.endnote_start.is_some_and(|value| value <= 0)
        {
            return Err(RtfError::MalformedDocument(
                "RTF note starting numbers must be positive".to_string(),
            ));
        }
        Ok(())
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}
