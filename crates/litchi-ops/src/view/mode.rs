//! View modes and their options (record 0010 §6.1).
//!
//! File rendering methods take a [`Mode`]. Internal [`Options`] retain the
//! adapter controls pending typed public option integration.

// Internal option parsing is used by adapter tests until public options land.
#![cfg_attr(
    not(test),
    allow(dead_code, reason = "typed public view options remain pending")
)]

use litchi_core::selector::{CellRange, column_number};

use crate::error::{Code, Error, Result};

/// A view mode, named as OfficeCLI names it (`CommandBuilder.View.cs`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Mode {
    /// The document's text (`text`, `t`).
    Text,
    /// Text with per-element annotations (`annotated`, `a`).
    Annotated,
    /// The structure (`outline`, `o`).
    Outline,
    /// Counts (`stats`, `s`).
    Stats,
    /// Validation findings (`issues`, `i`): `unsupported_capability` in M1.
    Issues,
    /// Form fields (`forms`, `f`): `unsupported_capability` in M1.
    Forms,
    /// HTML (`html`, `h`): `unsupported_capability`.
    Html,
    /// SVG (`svg`, `g`): `unsupported_capability`.
    Svg,
    /// A screenshot (`screenshot`, `p`): `unsupported_capability`.
    Screenshot,
    /// PDF (`pdf`): `unsupported_capability`.
    Pdf,
}

impl Mode {
    /// Every rendering mode, including modes with an explicit capability gap.
    pub const ALL: &'static [Self] = &[
        Self::Text,
        Self::Annotated,
        Self::Outline,
        Self::Stats,
        Self::Issues,
        Self::Forms,
        Self::Html,
        Self::Svg,
        Self::Screenshot,
        Self::Pdf,
    ];

    /// The mode named `text`, by OfficeCLI's name or one-letter alias,
    /// ASCII case-insensitively.
    pub fn parse(text: &str) -> Option<Self> {
        const NAMES: [(&str, &str, Mode); 10] = [
            ("text", "t", Mode::Text),
            ("annotated", "a", Mode::Annotated),
            ("outline", "o", Mode::Outline),
            ("stats", "s", Mode::Stats),
            ("issues", "i", Mode::Issues),
            ("forms", "f", Mode::Forms),
            ("html", "h", Mode::Html),
            ("svg", "g", Mode::Svg),
            ("screenshot", "p", Mode::Screenshot),
            ("pdf", "pdf", Mode::Pdf),
        ];
        NAMES
            .iter()
            .find(|(name, alias, _)| {
                text.eq_ignore_ascii_case(name) || text.eq_ignore_ascii_case(alias)
            })
            .map(|(_, _, mode)| *mode)
    }

    /// OfficeCLI's name of the mode.
    pub const fn name(self) -> &'static str {
        match self {
            Self::Text => "text",
            Self::Annotated => "annotated",
            Self::Outline => "outline",
            Self::Stats => "stats",
            Self::Issues => "issues",
            Self::Forms => "forms",
            Self::Html => "html",
            Self::Svg => "svg",
            Self::Screenshot => "screenshot",
            Self::Pdf => "pdf",
        }
    }

    /// The `unsupported_capability` error of a mode the format does not
    /// render (record 0010 §6.1, M1-Δ20).
    pub(crate) fn unsupported(self) -> Error {
        let reason = match self {
            Self::Issues | Self::Forms => "needs validators that M2's `validate` adds",
            Self::Html | Self::Svg | Self::Screenshot | Self::Pdf => {
                "rendering is outside the unified layer (design §10)"
            },
            _ => "is not supported for this format",
        };
        Error::new(
            Code::UnsupportedCapability,
            format!("the `{}` view {reason}", self.name()),
        )
        .with_suggestion("use the text, annotated, outline or stats view")
    }
}

/// Options of a view (OfficeCLI's `--start`, `--end`, `--max-lines`,
/// `--cols`, `--range`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub(crate) struct Options {
    /// The first line (1-based) a text or annotated view shows; for a
    /// workbook, counted in row records per sheet.
    pub(crate) start: Option<u32>,
    /// The last line (1-based, inclusive).
    pub(crate) end: Option<u32>,
    /// The most lines (workbook rows) the view shows in all.
    pub(crate) max_lines: Option<u32>,
    /// Workbooks: only these zero-based columns, ascending.
    pub(crate) columns: Option<Vec<u32>>,
    /// Workbooks: only this rectangle of one sheet (text view).
    pub(crate) range: Option<SheetRange>,
}

impl Options {
    /// Whether a workbook view shows the zero-based `column`.
    pub(crate) fn shows_column(&self, column: u32) -> bool {
        self.columns
            .as_ref()
            .is_none_or(|columns| columns.binary_search(&column).is_ok())
    }

    /// Workbooks: the columns of OfficeCLI's `--cols A,C`, as letters.
    ///
    /// # Errors
    ///
    /// `invalid_argument` for anything but column letters.
    pub(crate) fn with_columns(mut self, letters: &str) -> Result<Self> {
        let mut columns = Vec::new();
        for token in letters
            .split(',')
            .map(str::trim)
            .filter(|token| !token.is_empty())
        {
            let number = column_number(token).ok_or_else(|| {
                Error::new(Code::InvalidArgument, "invalid worksheet column")
                    .with_suggestion("list column letters, e.g. A,C,AB")
            })?;
            columns.push(number - 1);
        }
        columns.sort_unstable();
        columns.dedup();
        self.columns = Some(columns);
        Ok(self)
    }
}

/// A rectangle of one sheet (`Sheet1!A1:C10`, `/Sheet1/A1:C10`,
/// `Sheet1!B5`), with zero-based inclusive corners in normal order.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct SheetRange {
    /// The sheet, matched without regard to case.
    pub(crate) sheet: String,
    /// The first `(column, row)`.
    pub(crate) first: (u32, u32),
    /// The last `(column, row)`.
    pub(crate) last: (u32, u32),
}

impl std::fmt::Debug for SheetRange {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("SheetRange { .. }")
    }
}

impl SheetRange {
    /// Parses OfficeCLI's `--range` forms.
    ///
    /// # Errors
    ///
    /// `invalid_argument` for any other text.
    pub(crate) fn parse(text: &str) -> Result<Self> {
        let text = text.trim();
        let parts = match text.strip_prefix('/') {
            Some(rest) => rest.split_once('/'),
            None => text.split_once('!'),
        };
        let parsed = parts.and_then(|(sheet, cells)| {
            let sheet = sheet.trim_matches('\'');
            let range = CellRange::parse(cells)?;
            (!sheet.is_empty()).then(|| {
                let (start, end) = (range.start(), range.end());
                Self {
                    sheet: sheet.to_owned(),
                    first: (start.column().min(end.column()), start.row().min(end.row())),
                    last: (start.column().max(end.column()), start.row().max(end.row())),
                }
            })
        });
        parsed.ok_or_else(|| {
            Error::new(Code::InvalidArgument, "invalid worksheet range")
                .with_suggestion("write Sheet1!A1:C10, /Sheet1/A1:C10 or one cell, Sheet1!B5")
        })
    }
}

#[cfg(test)]
mod tests {
    use super::{Mode, Options, SheetRange};
    use crate::Code;

    #[test]
    fn modes_parse_by_name_and_alias() {
        assert_eq!(Mode::parse("TEXT"), Some(Mode::Text));
        assert_eq!(Mode::parse("a"), Some(Mode::Annotated));
        assert_eq!(Mode::parse("p"), Some(Mode::Screenshot));
        assert_eq!(Mode::parse("tree"), None);
        assert_eq!(Mode::Html.unsupported().code(), Code::UnsupportedCapability);
    }

    #[test]
    fn ranges_and_columns_parse() {
        let range = SheetRange::parse("'My Data'!C10:A1").expect("a range");
        assert_eq!(
            (range.sheet.as_str(), range.first, range.last),
            ("My Data", (0, 0), (2, 9))
        );
        assert_eq!(
            SheetRange::parse("/Sheet1/B5").expect("a cell").first,
            (1, 4)
        );
        assert_eq!(
            SheetRange::parse("Sheet1").unwrap_err().code(),
            Code::InvalidArgument
        );
        let options = Options::default().with_columns("c, A,C").expect("letters");
        assert_eq!(options.columns, Some(vec![0, 2]));
        assert!(options.shows_column(2) && !options.shows_column(1));
        assert!(Options::default().with_columns("A,1").is_err());
    }
}
