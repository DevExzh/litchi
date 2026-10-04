//! Story and part axes, and A1 cell-reference segments.

/// A story or part axis named by a path step.
///
/// Story axes reach content outside the main body: `/header[1]/p[1]`,
/// `/footnote[2]`, `/comments/comment[3]`, `/slide[2]/notes`,
/// `/slideMaster[1]/layout[2]`, `/Sheet1/chart[1]/series[2]`. Part axes name
/// document-level parts (`/styles`, `/numbering`, `/theme`). Classification is
/// syntactic and case-insensitive; whether the axis exists in a given format
/// is decided at resolution.
///
/// ```
/// use litchi_core::selector::{Axis, Path};
///
/// let path = Path::parse("/slideMaster[1]/layout[2]/shape[1]")?;
/// let axes: Vec<_> = path.axes().map(|(_, axis)| axis).collect();
/// assert_eq!(axes, [Axis::SlideMaster, Axis::SlideLayout]);
/// # Ok::<(), litchi_core::selector::ParseError>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Axis {
    /// `body`: the main word-processing story.
    Body,
    /// `header`
    Header,
    /// `footer`
    Footer,
    /// `footnote`, `footnotes`
    Footnote,
    /// `endnote`, `endnotes`
    Endnote,
    /// `comment`, `comments`, `modernComment`
    Comment,
    /// `notes` (slide speaker notes), `noteSlide`
    Notes,
    /// `notesMaster`
    NotesMaster,
    /// `slide`
    Slide,
    /// `slideMaster`
    SlideMaster,
    /// `slideLayout`, `layout`
    SlideLayout,
    /// `chart`
    Chart,
    /// `sheet` (positional worksheet axis, `/sheet[2]/A1`)
    Sheet,
    /// `styles`
    Styles,
    /// `numbering`
    Numbering,
    /// `theme`
    Theme,
    /// `settings`
    Settings,
    /// `watermark`
    Watermark,
    /// `customXml`
    CustomXml,
}

impl Axis {
    const TABLE: &'static [(&'static str, Axis)] = &[
        ("body", Axis::Body),
        ("chart", Axis::Chart),
        ("comment", Axis::Comment),
        ("comments", Axis::Comment),
        ("customxml", Axis::CustomXml),
        ("endnote", Axis::Endnote),
        ("endnotes", Axis::Endnote),
        ("footer", Axis::Footer),
        ("footnote", Axis::Footnote),
        ("footnotes", Axis::Footnote),
        ("header", Axis::Header),
        ("layout", Axis::SlideLayout),
        ("moderncomment", Axis::Comment),
        ("notes", Axis::Notes),
        ("noteslide", Axis::Notes),
        ("notesmaster", Axis::NotesMaster),
        ("numbering", Axis::Numbering),
        ("settings", Axis::Settings),
        ("sheet", Axis::Sheet),
        ("slide", Axis::Slide),
        ("slidelayout", Axis::SlideLayout),
        ("slidemaster", Axis::SlideMaster),
        ("styles", Axis::Styles),
        ("theme", Axis::Theme),
        ("watermark", Axis::Watermark),
    ];

    #[cfg(test)]
    pub(super) fn table_is_sorted() -> bool {
        Self::TABLE.windows(2).all(|pair| pair[0].0 < pair[1].0)
    }

    /// Classifies a step name (ASCII case-insensitive). Allocation-free.
    #[must_use]
    pub fn classify(name: &str) -> Option<Self> {
        if name.len() > 16 {
            return None;
        }
        let mut buf = [0u8; 16];
        let lower = &mut buf[..name.len()];
        lower.copy_from_slice(name.as_bytes());
        lower.make_ascii_lowercase();
        Self::TABLE
            .binary_search_by(|(key, _)| key.as_bytes().cmp(lower))
            .ok()
            .map(|index| Self::TABLE[index].1)
    }

    /// Returns `true` for axes holding text stories (body, header, footer,
    /// notes, comments, ...) rather than document-level parts.
    #[must_use]
    pub const fn is_story(self) -> bool {
        matches!(
            self,
            Self::Body
                | Self::Header
                | Self::Footer
                | Self::Footnote
                | Self::Endnote
                | Self::Comment
                | Self::Notes
                | Self::NotesMaster
                | Self::Slide
                | Self::SlideMaster
                | Self::SlideLayout
                | Self::Chart
                | Self::Sheet
        )
    }
}

/// Excel's last column (`XFD`), 1-based.
const MAX_COLUMN: u32 = 16_384;
/// Excel's last row, 1-based.
const MAX_ROW: u32 = 1_048_576;

/// A syntactic A1 cell reference with zero-based column and row.
///
/// ```
/// use litchi_core::selector::CellRef;
///
/// let cell = CellRef::parse("B3").unwrap();
/// assert_eq!((cell.column(), cell.row()), (1, 2));
/// assert_eq!(CellRef::parse("XFE1"), None); // beyond Excel's last column
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellRef {
    column: u32,
    row: u32,
}

impl CellRef {
    /// Parses `A1`-style text (column letters case-insensitive, no `$`).
    #[must_use]
    pub fn parse(text: &str) -> Option<Self> {
        let bytes = text.as_bytes();
        let letters = bytes.iter().take_while(|b| b.is_ascii_alphabetic()).count();
        if !(1..=3).contains(&letters) {
            return None;
        }
        let column = column_number(&text[..letters])?;
        let digits = &bytes[letters..];
        if digits.is_empty() || digits.len() > 7 || digits[0] == b'0' {
            return None;
        }
        let mut row: u32 = 0;
        for &b in digits {
            if !b.is_ascii_digit() {
                return None;
            }
            row = row * 10 + u32::from(b - b'0');
        }
        (row <= MAX_ROW).then_some(Self {
            column: column - 1,
            row: row - 1,
        })
    }

    /// Returns the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Returns the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }
}

/// A syntactic `A1:C5` range (or a single cell) with ordered corners as
/// written.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    start: CellRef,
    end: CellRef,
}

impl CellRange {
    /// Parses `A1` or `A1:C5`.
    #[must_use]
    pub fn parse(text: &str) -> Option<Self> {
        match text.split_once(':') {
            Some((start, end)) => Some(Self {
                start: CellRef::parse(start)?,
                end: CellRef::parse(end)?,
            }),
            None => CellRef::parse(text).map(|cell| Self {
                start: cell,
                end: cell,
            }),
        }
    }

    /// Returns the first corner.
    #[must_use]
    pub const fn start(self) -> CellRef {
        self.start
    }

    /// Returns the second corner.
    #[must_use]
    pub const fn end(self) -> CellRef {
        self.end
    }
}

/// Converts column letters (`A`, `xfd`) to a 1-based column number.
///
/// ```
/// use litchi_core::selector::column_number;
///
/// assert_eq!(column_number("B"), Some(2));
/// assert_eq!(column_number("XFD"), Some(16_384));
/// assert_eq!(column_number("XFE"), None);
/// ```
#[must_use]
pub fn column_number(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut column: u32 = 0;
    for b in letters.bytes() {
        if !b.is_ascii_alphabetic() {
            return None;
        }
        column = column * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1);
    }
    (column <= MAX_COLUMN).then_some(column)
}
