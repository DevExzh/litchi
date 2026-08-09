//! Semantic `PowerPoint` header/footer metadata.

use crate::package::{Error, Result};

/// A validated `PowerPoint` datetime format identifier.
///
/// Values 0 through 12 are the ordinary locale-dependent formats. Value 13
/// is permitted by `HeadersFootersAtom`, although producers are advised not to
/// emit it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DateTimeFormatId(u8);

impl DateTimeFormatId {
    /// Lowest valid format identifier.
    pub const MIN: u8 = 0;
    /// Highest valid format identifier.
    pub const MAX: u8 = 13;

    /// Construct a validated format identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(value: u8) -> Result<Self> {
        if value > Self::MAX {
            return Err(Error::Corrupted(
                "header/footer datetime format ID is outside 0..=13".to_string(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the on-disk identifier.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for DateTimeFormatId {
    fn default() -> Self {
        Self(Self::MIN)
    }
}

impl TryFrom<u8> for DateTimeFormatId {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

impl From<DateTimeFormatId> for u8 {
    fn from(value: DateTimeFormatId) -> Self {
        value.get()
    }
}

/// The direct parent of a local header/footer container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HeaderFooterParent {
    /// A presentation slide.
    Slide,
    /// A main-master slide.
    MainMaster,
}

/// A zero-based ordinal among parents of the same kind in record order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct HeaderFooterParentOrdinal(pub(super) usize);

impl HeaderFooterParentOrdinal {
    /// Construct an ordinal from a zero-based parent index.
    #[inline]
    #[must_use]
    pub const fn new(value: usize) -> Self {
        Self(value)
    }

    /// Return the zero-based ordinal.
    #[inline]
    #[must_use]
    pub const fn get(self) -> usize {
        self.0
    }
}

/// The specification-defined scope of a header/footer container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HeaderFooterScope {
    /// Presentation-wide defaults for ordinary slides.
    PresentationSlides,
    /// Presentation-wide defaults for notes pages and handouts.
    NotesAndHandouts,
    /// Overrides or defaults attached directly to one slide or main master.
    Local {
        /// Kind of direct parent.
        parent: HeaderFooterParent,
        /// Parent ordinal in `PowerPoint` record order.
        parent_ordinal: HeaderFooterParentOrdinal,
    },
}

/// Display options stored by `HeadersFootersAtom`.
#[allow(
    clippy::struct_excessive_bools,
    reason = "each bool maps one-to-one to an independent flag bit of the MS-PPT \
              `HeadersFootersAtom` bitfield; grouping them into enums would misrepresent \
              the on-disk layout and churn the public API"
)]
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct HeaderFooterOptions {
    /// Locale-dependent datetime format identifier.
    pub datetime_format: DateTimeFormatId,
    /// Display a date placeholder.
    pub show_date: bool,
    /// Use the current date and time.
    pub use_current_datetime: bool,
    /// Use the custom user-date string.
    pub use_user_date: bool,
    /// Display the slide number.
    pub show_slide_number: bool,
    /// Display a header. This bit is retained even where the specification says
    /// it has no effect.
    pub show_header: bool,
    /// Display a footer.
    pub show_footer: bool,
}

/// Text derived from inert header/footer placeholder shapes.
///
/// Office 2007 can save binary presentations with visible header/footer text
/// in placeholders while leaving the corresponding `CString` atoms absent. This
/// view is kept separate so record-local serialization remains lossless.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooterDisplayText {
    /// Custom date text visible through a datetime placeholder.
    pub user_date: Option<String>,
    /// Header text visible through a header placeholder.
    pub header: Option<String>,
    /// Footer text visible through a footer placeholder.
    pub footer: Option<String>,
}

/// Placeholder-derived display text associated with a specification scope.
///
/// A scoped display can exist without a corresponding local record because
/// Office 2007 binary presentations can inherit document-level options while
/// storing slide-specific text only in placeholder shapes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScopedHeaderFooterDisplayText {
    /// Slide, master, or document-level scope of the placeholder text.
    pub scope: HeaderFooterScope,
    /// Inert text extracted from placeholder shapes.
    pub text: HeaderFooterDisplayText,
}

/// Typed, inert metadata from one `PowerPoint` header/footer container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderFooter {
    /// Container scope and parent association.
    pub scope: HeaderFooterScope,
    /// Display and datetime-format options.
    pub options: HeaderFooterOptions,
    /// Optional custom date text.
    pub user_date: Option<String>,
    /// Optional notes/handout header text.
    pub header: Option<String>,
    /// Optional footer text.
    pub footer: Option<String>,
    /// Optional text derived from inert placeholders. This is never serialized
    /// into the record-local `CString` fields.
    pub placeholder_display: Option<HeaderFooterDisplayText>,
}

impl HeaderFooter {
    /// Return visible custom-date text, preferring an attached Office 2007
    /// placeholder and otherwise using the stored `UserDateAtom`.
    #[must_use]
    pub fn display_user_date(&self) -> Option<&str> {
        self.placeholder_display
            .as_ref()
            .and_then(|display| display.user_date.as_deref())
            .or(self.user_date.as_deref())
    }

    /// Return visible header text, preferring an attached Office 2007
    /// placeholder and otherwise using the stored `HeaderAtom`.
    #[must_use]
    pub fn display_header(&self) -> Option<&str> {
        self.placeholder_display
            .as_ref()
            .and_then(|display| display.header.as_deref())
            .or(self.header.as_deref())
    }

    /// Return visible footer text, preferring an attached Office 2007
    /// placeholder and otherwise using the stored `FooterAtom`.
    #[must_use]
    pub fn display_footer(&self) -> Option<&str> {
        self.placeholder_display
            .as_ref()
            .and_then(|display| display.footer.as_deref())
            .or(self.footer.as_deref())
    }
}

/// All strictly located header/footer containers in a presentation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooters {
    pub(super) entries: Vec<HeaderFooter>,
    pub(super) placeholder_displays: Vec<ScopedHeaderFooterDisplayText>,
    pub(super) placeholder_display_bytes: usize,
}

impl HeaderFooters {
    /// Return entries in `PowerPoint` record order.
    #[inline]
    #[must_use]
    pub fn entries(&self) -> &[HeaderFooter] {
        &self.entries
    }

    /// Return placeholder-derived displays in physical `PowerPoint` record order.
    ///
    /// Unlike [`Self::entries`], these values are not necessarily backed by a
    /// local `RT_HeadersFooters` record and cannot be serialized as one.
    #[inline]
    #[must_use]
    pub fn placeholder_displays(&self) -> &[ScopedHeaderFooterDisplayText] {
        &self.placeholder_displays
    }

    /// Return placeholder-derived display text for an exact scope.
    #[must_use]
    pub fn placeholder_display(
        &self,
        scope: HeaderFooterScope,
    ) -> Option<&HeaderFooterDisplayText> {
        self.placeholder_displays
            .iter()
            .find(|display| display.scope == scope)
            .map(|display| &display.text)
    }

    /// Return the presentation-wide ordinary-slide defaults, if present.
    #[must_use]
    pub fn presentation_slides(&self) -> Option<&HeaderFooter> {
        self.entries
            .iter()
            .find(|entry| entry.scope == HeaderFooterScope::PresentationSlides)
    }

    /// Return the presentation-wide notes/handout defaults, if present.
    #[must_use]
    pub fn notes_and_handouts(&self) -> Option<&HeaderFooter> {
        self.entries
            .iter()
            .find(|entry| entry.scope == HeaderFooterScope::NotesAndHandouts)
    }
}
