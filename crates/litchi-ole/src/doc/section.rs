//! Public section-layout model for Word 97+ documents.

/// A section and the character-position range to which its properties apply.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocSection {
    /// Inclusive start character position in the main document story.
    pub start_cp: u32,
    /// Exclusive end character position in the main document story.
    pub end_cp: u32,
    /// The break that terminates this section.
    pub break_kind: SectionBreakKind,
    /// Page geometry for this section.
    pub page: SectionPageLayout,
    /// Column geometry for this section.
    pub columns: SectionColumnLayout,
}

/// The kind of break that terminates a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionBreakKind {
    Continuous,
    NewColumn,
    NewPage,
    EvenPage,
    OddPage,
}

/// Page orientation explicitly stored in section properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// A vertical page margin as defined by `sprmSDyaTop` and `sprmSDyaBottom`.
///
/// Positive values are minimum margins that can grow to accommodate headers,
/// footers, or footnotes. Negative values are fixed distances whose absolute
/// value must be used.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalMargin {
    Minimum(u16),
    Fixed(u16),
}

impl VerticalMargin {
    /// Return the lossless signed-twips representation used by the file format.
    pub fn signed_twips(self) -> i16 {
        match self {
            Self::Minimum(value) => value as i16,
            Self::Fixed(value) => -(value as i16),
        }
    }

    /// Return the physical distance in twips, independent of margin behavior.
    pub fn distance_twips(self) -> u16 {
        match self {
            Self::Minimum(value) | Self::Fixed(value) => value,
        }
    }

    pub(crate) fn from_signed_twips(value: i16) -> Self {
        if value < 0 {
            Self::Fixed(value.unsigned_abs())
        } else {
            Self::Minimum(value as u16)
        }
    }
}

/// Page margins for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionMargins {
    pub left_twips: u16,
    pub right_twips: u16,
    pub top: VerticalMargin,
    pub bottom: VerticalMargin,
    pub gutter_twips: u16,
}

/// Page geometry and header/footer distances for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageLayout {
    pub width_twips: u16,
    pub height_twips: u16,
    pub orientation: PageOrientation,
    pub margins: SectionMargins,
    pub header_distance_twips: u16,
    pub footer_distance_twips: u16,
}

/// Column geometry for a section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SectionColumnLayout {
    /// Equal-width columns separated by a common spacing value.
    Even {
        count: u8,
        spacing_twips: u16,
        line_between: bool,
    },
    /// Individually sized columns and their following spacing.
    Unequal {
        columns: Vec<SectionColumn>,
        line_between: bool,
    },
}

impl SectionColumnLayout {
    /// Number of columns in this section.
    pub fn count(&self) -> usize {
        match self {
            Self::Even { count, .. } => usize::from(*count),
            Self::Unequal { columns, .. } => columns.len(),
        }
    }

    /// Whether a vertical line is drawn between columns.
    pub fn line_between(&self) -> bool {
        match self {
            Self::Even { line_between, .. } | Self::Unequal { line_between, .. } => *line_between,
        }
    }
}

/// Width and following spacing for one unequal-width column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionColumn {
    pub width_twips: u16,
    /// Space after this column. The final column has no following spacing.
    pub spacing_after_twips: Option<u16>,
}
