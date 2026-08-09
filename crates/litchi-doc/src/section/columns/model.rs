//! Checked semantic column values for Word section properties.

use std::fmt;

/// A validated unequal-width section column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Column {
    width_twips: u16,
    /// Space after this column. The final column has no following spacing.
    spacing_after_twips: Option<u16>,
}

impl Column {
    /// Construct a column from its width and optional following spacing.
    ///
    /// The final-column rule is checked when the column is installed in a
    /// [`Layout`], because it depends on the column's position in the list.
    pub fn new(width_twips: u16, spacing_after_twips: Option<u16>) -> Result<Self, Error> {
        Self::from_parts(0, width_twips, spacing_after_twips)
    }

    /// Return the column width in twips.
    #[must_use]
    pub fn width_twips(self) -> u16 {
        self.width_twips
    }

    /// Return the spacing after this column, if it is not the final column.
    #[must_use]
    pub fn spacing_after_twips(self) -> Option<u16> {
        self.spacing_after_twips
    }

    /// Construct a column while retaining its position for parser diagnostics.
    pub(crate) fn from_parts(
        index: usize,
        width_twips: u16,
        spacing_after_twips: Option<u16>,
    ) -> Result<Self, Error> {
        if !(Layout::MIN_UNEQUAL_WIDTH_TWIPS..=Layout::MAX_TWIPS).contains(&width_twips) {
            return Err(Error::InvalidWidth { index, width_twips });
        }
        if let Some(spacing_twips) = spacing_after_twips
            && spacing_twips > Layout::MAX_TWIPS
        {
            return Err(Error::InvalidSpacing {
                index,
                spacing_twips,
            });
        }
        Ok(Self {
            width_twips,
            spacing_after_twips,
        })
    }
}

/// Checked section column layout.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Layout {
    /// Equal-width columns separated by a common spacing value.
    Even {
        count: u8,
        spacing_twips: u16,
        line_between: bool,
    },
    /// Individually sized columns and their following spacing.
    Unequal {
        columns: Vec<Column>,
        line_between: bool,
    },
}

/// Validation failure for a section column layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    InvalidCount(usize),
    InvalidWidth { index: usize, width_twips: u16 },
    InvalidSpacing { index: usize, spacing_twips: u16 },
    MissingSpacing { index: usize },
    FinalColumnHasSpacing,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidCount(count) => {
                write!(formatter, "section column count {count} is outside 1..=44")
            },
            Self::InvalidWidth { index, width_twips } => write!(
                formatter,
                "section column {index} width {width_twips} is outside 718..=31680 twips"
            ),
            Self::InvalidSpacing {
                index,
                spacing_twips,
            } => write!(
                formatter,
                "section column {index} spacing {spacing_twips} exceeds 31680 twips"
            ),
            Self::MissingSpacing { index } => {
                write!(
                    formatter,
                    "section column {index} is missing following spacing"
                )
            },
            Self::FinalColumnHasSpacing => {
                formatter.write_str("the final section column cannot have following spacing")
            },
        }
    }
}

impl std::error::Error for Error {}

impl Layout {
    /// Maximum number of columns supported by the Word section grammar.
    pub const MAX_COLUMNS: usize = 44;
    /// Maximum representable non-negative column spacing or width.
    pub const MAX_TWIPS: u16 = 31_680;
    /// Minimum width of an unequal-width column.
    pub const MIN_UNEQUAL_WIDTH_TWIPS: u16 = 718;

    /// Construct and validate an equal-width layout.
    pub fn even(count: u8, spacing_twips: u16, line_between: bool) -> Result<Self, Error> {
        let layout = Self::Even {
            count,
            spacing_twips,
            line_between,
        };
        layout.validate()?;
        Ok(layout)
    }

    /// Construct and validate an unequal-width layout.
    pub fn unequal(columns: Vec<Column>, line_between: bool) -> Result<Self, Error> {
        let layout = Self::Unequal {
            columns,
            line_between,
        };
        layout.validate()?;
        Ok(layout)
    }

    /// Validate all cross-field constraints without depending on SPRM order.
    pub fn validate(&self) -> Result<(), Error> {
        let count = self.count();
        if !(1..=Self::MAX_COLUMNS).contains(&count) {
            return Err(Error::InvalidCount(count));
        }
        match self {
            Self::Even { spacing_twips, .. } => {
                if *spacing_twips > Self::MAX_TWIPS {
                    return Err(Error::InvalidSpacing {
                        index: 0,
                        spacing_twips: *spacing_twips,
                    });
                }
            },
            Self::Unequal { columns, .. } => {
                for (index, column) in columns.iter().enumerate() {
                    // Column::new checks the scalar domains. Recheck here so
                    // future in-crate construction paths cannot bypass the
                    // aggregate validator.
                    if !(Self::MIN_UNEQUAL_WIDTH_TWIPS..=Self::MAX_TWIPS)
                        .contains(&column.width_twips)
                    {
                        return Err(Error::InvalidWidth {
                            index,
                            width_twips: column.width_twips,
                        });
                    }
                    if index + 1 == columns.len() {
                        if column.spacing_after_twips.is_some() {
                            return Err(Error::FinalColumnHasSpacing);
                        }
                    } else {
                        let spacing_twips = column
                            .spacing_after_twips
                            .ok_or(Error::MissingSpacing { index })?;
                        if spacing_twips > Self::MAX_TWIPS {
                            return Err(Error::InvalidSpacing {
                                index,
                                spacing_twips,
                            });
                        }
                    }
                }
            },
        }
        Ok(())
    }

    /// Replace this layout with a validated equal-width layout.
    pub fn set_even(
        &mut self,
        count: u8,
        spacing_twips: u16,
        line_between: bool,
    ) -> Result<(), Error> {
        *self = Self::even(count, spacing_twips, line_between)?;
        Ok(())
    }

    /// Replace this layout with a validated unequal-width layout.
    pub fn set_unequal(&mut self, columns: Vec<Column>, line_between: bool) -> Result<(), Error> {
        *self = Self::unequal(columns, line_between)?;
        Ok(())
    }

    /// Change only the line-between flag without affecting column geometry.
    pub fn set_line_between(&mut self, value: bool) {
        match self {
            Self::Even { line_between, .. } | Self::Unequal { line_between, .. } => {
                *line_between = value;
            },
        }
    }

    /// Number of columns in this section.
    #[must_use]
    pub fn count(&self) -> usize {
        match self {
            Self::Even { count, .. } => usize::from(*count),
            Self::Unequal { columns, .. } => columns.len(),
        }
    }

    /// Whether a vertical line is drawn between columns.
    #[must_use]
    pub fn line_between(&self) -> bool {
        match self {
            Self::Even { line_between, .. } | Self::Unequal { line_between, .. } => *line_between,
        }
    }

    /// Return equal-column parameters when this is an equal-width layout.
    #[must_use]
    pub fn equal(&self) -> Option<(u8, u16, bool)> {
        match self {
            Self::Even {
                count,
                spacing_twips,
                line_between,
            } => Some((*count, *spacing_twips, *line_between)),
            Self::Unequal { .. } => None,
        }
    }

    /// Return unequal columns when this is an individually sized layout.
    #[must_use]
    pub fn unequal_columns(&self) -> Option<&[Column]> {
        match self {
            Self::Even { .. } => None,
            Self::Unequal { columns, .. } => Some(columns),
        }
    }

    /// Return the checked data needed by the section wire codec.
    pub(crate) fn wire_view(&self) -> WireView<'_> {
        match self {
            Self::Even {
                spacing_twips,
                line_between,
                ..
            } => WireView::Even {
                spacing_twips: *spacing_twips,
                line_between: *line_between,
            },
            Self::Unequal {
                columns,
                line_between,
            } => WireView::Unequal {
                columns,
                line_between: *line_between,
            },
        }
    }
}

/// Borrowed representation for the section `SEPX` codec.
#[derive(Debug, Clone, Copy)]
pub(crate) enum WireView<'a> {
    Even {
        spacing_twips: u16,
        line_between: bool,
    },
    Unequal {
        columns: &'a [Column],
        line_between: bool,
    },
}
