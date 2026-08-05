//! Archive-free Numbers document semantics.

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use crate::Sheet;
use crate::cell::Value;

/// Maximum number of ordered sheets retained by one semantic document.
pub const MAX_SHEETS: usize = 4096;
/// Maximum number of tables retained by one semantic document.
pub const MAX_TABLES: usize = 65_536;
/// Maximum number of materialized cells retained by one semantic document.
pub const MAX_MATERIALIZED_CELLS: usize = 16_000_000;
/// Maximum UTF-8 bytes retained by semantic names, headers, and textual cells.
pub const DEFAULT_MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Finite construction limits for an immutable Numbers document.
#[allow(
    clippy::struct_field_names,
    reason = "The public budget accessors intentionally share one max_* vocabulary"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_sheets: usize,
    max_tables: usize,
    max_materialized_cells: usize,
    max_text_bytes: usize,
}

impl Limits {
    /// Creates a bounded profile. Values above the hard semantic ceilings are
    /// clamped during document construction.
    #[must_use]
    pub const fn new(
        max_sheets: usize,
        max_tables: usize,
        max_materialized_cells: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_sheets,
            max_tables,
            max_materialized_cells,
            max_text_bytes,
        }
    }

    /// Returns the configured sheet ceiling.
    #[must_use]
    pub const fn max_sheets(self) -> usize {
        self.max_sheets
    }

    /// Returns the configured table ceiling.
    #[must_use]
    pub const fn max_tables(self) -> usize {
        self.max_tables
    }

    /// Returns the configured materialized-cell ceiling.
    #[must_use]
    pub const fn max_materialized_cells(self) -> usize {
        self.max_materialized_cells
    }

    /// Returns the configured text-byte ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new(
            MAX_SHEETS,
            MAX_TABLES,
            MAX_MATERIALIZED_CELLS,
            DEFAULT_MAX_TEXT_BYTES,
        )
    }
}

/// Errors returned while constructing a bounded semantic document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The supplied sheet sequence exceeds the selected bound.
    TooManySheets {
        /// Number of supplied sheets.
        actual: usize,
        /// Maximum accepted sheets.
        limit: usize,
    },
    /// A sheet does not carry its canonical position in the ordered sequence.
    InvalidSheetIndex {
        /// Position occupied by the sheet in the supplied sequence.
        expected: usize,
        /// Index stored by the sheet.
        actual: usize,
    },
    /// Two sheets use the same semantic name.
    DuplicateSheetName {
        /// Earlier sheet position using the name.
        first: usize,
        /// Later sheet position using the name.
        duplicate: usize,
    },
    /// The table aggregate exceeds the selected bound.
    TooManyTables {
        /// Number of supplied tables.
        actual: usize,
        /// Maximum accepted tables.
        limit: usize,
    },
    /// The materialized-cell aggregate exceeds the selected bound.
    TooManyMaterializedCells {
        /// Number of supplied materialized cells.
        actual: usize,
        /// Maximum accepted materialized cells.
        limit: usize,
    },
    /// Semantic names, headers, and textual cell values exceed the budget.
    TextTooLarge {
        /// Maximum accepted UTF-8 bytes.
        limit: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::TooManySheets { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} sheets; maximum is {limit}"
            ),
            Self::InvalidSheetIndex { expected, actual } => write!(
                formatter,
                "Numbers sheet index {actual} is not the expected index {expected}"
            ),
            Self::DuplicateSheetName { first, duplicate } => write!(
                formatter,
                "Numbers sheets {first} and {duplicate} have the same semantic name"
            ),
            Self::TooManyTables { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} tables; maximum is {limit}"
            ),
            Self::TooManyMaterializedCells { actual, limit } => write!(
                formatter,
                "Numbers document contains {actual} materialized cells; maximum is {limit}"
            ),
            Self::TextTooLarge { limit } => write!(
                formatter,
                "Numbers document semantic text exceeds {limit} bytes"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for bounded Numbers semantic construction.
pub type Result<T> = std::result::Result<T, Error>;

/// An immutable, archive-free Numbers document snapshot.
///
/// The document owns only semantic [`Sheet`] values. Its hidden state is
/// reference counted so cloning or taking a snapshot never copies the sheet
/// or table storage. Native archives, protobuf values, package entries, and
/// physical object identifiers are intentionally outside this API.
#[derive(Debug, Clone)]
pub struct Document {
    sheets: Arc<[Sheet]>,
}

impl Document {
    /// Build a document from sheets in source order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManySheets`] when the hard semantic bound is
    /// exceeded, or [`Error::InvalidSheetIndex`] when a sheet is not numbered
    /// by its zero-based position in the supplied sequence.
    pub fn from_sheets(sheets: Vec<Sheet>) -> Result<Self> {
        Self::from_sheets_with_limits(sheets, Limits::default())
    }

    /// Build a document under a caller-selected sheet-count budget.
    ///
    /// The package-independent hard cap [`MAX_SHEETS`] cannot be relaxed by a
    /// caller. The input vector is consumed without rebuilding its sheet
    /// values when construction succeeds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManySheets`] when the supplied count exceeds either
    /// the caller budget or the hard semantic cap, or
    /// [`Error::InvalidSheetIndex`] when a sheet is not numbered by its
    /// zero-based position in the supplied sequence.
    pub fn from_sheets_with_max_sheets(sheets: Vec<Sheet>, max_sheets: usize) -> Result<Self> {
        Self::from_sheets_with_limits(
            sheets,
            Limits::new(
                max_sheets,
                MAX_TABLES,
                MAX_MATERIALIZED_CELLS,
                DEFAULT_MAX_TEXT_BYTES,
            ),
        )
    }

    /// Build a document under explicit finite semantic budgets.
    ///
    /// The source vector is consumed without cloning its sheet or table
    /// values. Validation runs before the vector is moved into one shared
    /// immutable allocation, so a rejected archive cannot publish a partial
    /// semantic snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed error when any count, index, name, or textual-data
    /// budget is exceeded.
    pub fn from_sheets_with_limits(sheets: Vec<Sheet>, limits: Limits) -> Result<Self> {
        let max_sheets = limits.max_sheets.min(MAX_SHEETS);
        let max_tables = limits.max_tables.min(MAX_TABLES);
        let max_materialized_cells = limits.max_materialized_cells.min(MAX_MATERIALIZED_CELLS);
        let max_text_bytes = limits.max_text_bytes.min(DEFAULT_MAX_TEXT_BYTES);

        if sheets.len() > max_sheets {
            return Err(Error::TooManySheets {
                actual: sheets.len(),
                limit: max_sheets,
            });
        }

        let mut names = HashSet::new();
        let mut table_count = 0usize;
        let mut materialized_cell_count = 0usize;
        let mut text_bytes = 0usize;
        for (expected, sheet) in sheets.iter().enumerate() {
            if sheet.index() != expected {
                return Err(Error::InvalidSheetIndex {
                    expected,
                    actual: sheet.index(),
                });
            }
            if !names.insert(sheet.name()) {
                let first = sheets[..expected]
                    .iter()
                    .position(|previous| previous.name() == sheet.name())
                    .unwrap_or(expected);
                return Err(Error::DuplicateSheetName {
                    first,
                    duplicate: expected,
                });
            }

            text_bytes = checked_text_add(text_bytes, sheet.name().len(), max_text_bytes)?;
            for table in sheet.tables() {
                table_count = table_count.checked_add(1).ok_or(Error::TooManyTables {
                    actual: usize::MAX,
                    limit: max_tables,
                })?;
                if table_count > max_tables {
                    return Err(Error::TooManyTables {
                        actual: table_count,
                        limit: max_tables,
                    });
                }

                materialized_cell_count = materialized_cell_count
                    .checked_add(table.cell_count())
                    .ok_or(Error::TooManyMaterializedCells {
                    actual: usize::MAX,
                    limit: max_materialized_cells,
                })?;
                if materialized_cell_count > max_materialized_cells {
                    return Err(Error::TooManyMaterializedCells {
                        actual: materialized_cell_count,
                        limit: max_materialized_cells,
                    });
                }

                text_bytes = checked_text_add(text_bytes, table.name().len(), max_text_bytes)?;
                for header in table.column_headers().chain(table.row_headers()) {
                    text_bytes = checked_text_add(text_bytes, header.len(), max_text_bytes)?;
                }
                for cell in table.iter_cells() {
                    let value_bytes = match cell.value() {
                        Value::Text(value) | Value::Formula(value) | Value::Error(value) => {
                            value.len()
                        },
                        Value::Empty
                        | Value::Number(_)
                        | Value::Boolean(_)
                        | Value::Date(_)
                        | Value::Duration(_) => 0,
                    };
                    text_bytes = checked_text_add(text_bytes, value_bytes, max_text_bytes)?;
                }
            }
        }

        Ok(Self {
            sheets: Arc::from(sheets.into_boxed_slice()),
        })
    }

    /// Capture another cheap handle to the same immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow all sheets in stable source order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Select a sheet by checked zero-based position.
    #[must_use]
    pub fn sheet(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    /// Select a sheet by its unique semantic name.
    #[must_use]
    pub fn sheet_named(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|sheet| sheet.name() == name)
    }

    /// Return the number of semantic sheets.
    #[must_use]
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Return whether the document contains no semantic sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }
}

fn checked_text_add(current: usize, added: usize, limit: usize) -> Result<usize> {
    let total = current
        .checked_add(added)
        .ok_or(Error::TextTooLarge { limit })?;
    if total > limit {
        return Err(Error::TextTooLarge { limit });
    }
    Ok(total)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn empty_document_is_a_valid_bounded_snapshot() {
        let document = Document::from_sheets(Vec::new())
            .unwrap_or_else(|error| panic!("empty document should be valid: {error}"));

        assert_send_sync::<Document>();
        assert!(document.is_empty());
        assert_eq!(document.sheet_count(), 0);
        assert!(document.sheets().is_empty());
        assert!(document.sheet(0).is_none());
    }

    #[test]
    fn construction_checks_budget_and_canonical_order() {
        let too_many = Document::from_sheets_with_max_sheets(vec![Sheet::new("Sheet 1", 0)], 0);
        assert!(matches!(
            too_many,
            Err(Error::TooManySheets {
                actual: 1,
                limit: 0,
            })
        ));

        let invalid = Document::from_sheets(vec![Sheet::new("Sheet 2", 1)]);
        assert!(matches!(
            invalid,
            Err(Error::InvalidSheetIndex {
                expected: 0,
                actual: 1,
            })
        ));
    }

    #[test]
    fn clones_share_ordered_semantic_storage() {
        let document =
            Document::from_sheets(vec![Sheet::new("Sheet 1", 0), Sheet::new("Sheet 2", 1)])
                .unwrap_or_else(|error| panic!("document should be valid: {error}"));
        let snapshot = document.snapshot();

        assert!(Arc::ptr_eq(&document.sheets, &snapshot.sheets));
        assert_eq!(snapshot.sheet_count(), 2);
        assert_eq!(snapshot.sheet(0).map(Sheet::name), Some("Sheet 1"));
        assert_eq!(snapshot.sheet(1).map(Sheet::name), Some("Sheet 2"));
        assert!(snapshot.sheet(2).is_none());
        assert_eq!(snapshot.sheet_named("Sheet 2").map(Sheet::index), Some(1));
    }

    #[test]
    fn construction_rejects_duplicate_names_and_aggregate_budgets() {
        let duplicate =
            Document::from_sheets(vec![Sheet::new("Summary", 0), Sheet::new("Summary", 1)]);
        assert!(matches!(
            duplicate,
            Err(Error::DuplicateSheetName {
                first: 0,
                duplicate: 1,
            })
        ));

        let mut table = crate::table::Builder::new("Data", crate::Dimensions::new(1, 1));
        assert!(
            table
                .set(crate::Position::new(0, 0), Value::Text("value".to_owned()))
                .is_ok()
        );
        let table = table
            .finish()
            .unwrap_or_else(|error| panic!("table should be valid: {error}"));
        let mut sheet = crate::sheet::Builder::new("Summary", 0);
        assert!(sheet.push_table(table).is_ok());
        let sheet = sheet.finish();

        let table_limit = Limits::new(1, 0, MAX_MATERIALIZED_CELLS, DEFAULT_MAX_TEXT_BYTES);
        assert!(matches!(
            Document::from_sheets_with_limits(vec![sheet.clone()], table_limit),
            Err(Error::TooManyTables {
                actual: 1,
                limit: 0,
            })
        ));

        let text_limit = Limits::new(1, MAX_TABLES, MAX_MATERIALIZED_CELLS, 4);
        assert!(matches!(
            Document::from_sheets_with_limits(vec![sheet], text_limit),
            Err(Error::TextTooLarge { limit: 4 })
        ));
    }
}
