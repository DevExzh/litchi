//! Source-bound existing-cell edits for positional ODS packages.
//!
//! The transaction owns no complete package bytes. It retains the already
//! validated worksheet projection, clones only selected worksheet models, and
//! prepares one checked `content.xml` replacement. Publication delegates to
//! the common ODF positional writer so every other ZIP member remains a raw
//! source copy.

use std::{collections::BTreeMap, fmt, ops::Range, sync::Arc};

use litchi_core::{Error, Result, SourceVersion};

use super::SourceBackedSpreadsheet;
use crate::worksheet::Row;
use crate::worksheet::{Cell, CellChange, CellValue, Link, Merge, Selector, Sheet, validation};

/// One existing formula-cell replacement in a source-backed ODS transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormulaChange {
    row: usize,
    column: usize,
    formula: String,
}

impl FormulaChange {
    /// Construct one formula replacement at a zero-based logical coordinate.
    #[must_use]
    pub fn new(row: usize, column: usize, formula: impl Into<String>) -> Self {
        Self {
            row,
            column,
            formula: formula.into(),
        }
    }

    /// Zero-based logical row coordinate.
    #[must_use]
    pub const fn row(&self) -> usize {
        self.row
    }

    /// Zero-based logical column coordinate.
    #[must_use]
    pub const fn column(&self) -> usize {
        self.column
    }

    /// Replacement formula text.
    #[must_use]
    pub fn formula(&self) -> &str {
        &self.formula
    }
}

/// One ordered mutation of the direct hyperlinks in an existing ODS cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HyperlinkOperation {
    /// Replace the complete displayed cell value with one hyperlink.
    Set(Link),
    /// Add one hyperlink around a checked UTF-8 byte range of the cell text.
    Add {
        /// UTF-8 byte range occupied by the new link.
        range: Range<usize>,
        /// Link metadata and visible text to add.
        hyperlink: Link,
    },
    /// Replace one hyperlink by its document-order index.
    Replace {
        /// Document-order hyperlink index.
        index: usize,
        /// Replacement link metadata and visible text.
        hyperlink: Link,
    },
    /// Replace the hyperlink occupying an exact UTF-8 byte range.
    ReplaceAt {
        /// Exact existing hyperlink range.
        range: Range<usize>,
        /// Replacement link metadata and visible text.
        hyperlink: Link,
    },
    /// Remove one hyperlink by its document-order index.
    Remove {
        /// Document-order hyperlink index.
        index: usize,
    },
    /// Remove all hyperlinks while retaining the cell text.
    Clear,
}

impl HyperlinkOperation {
    /// Construct [`Self::Set`].
    #[must_use]
    pub fn set(hyperlink: Link) -> Self {
        Self::Set(hyperlink)
    }

    /// Construct [`Self::Add`].
    #[must_use]
    pub fn add(range: Range<usize>, hyperlink: Link) -> Self {
        Self::Add { range, hyperlink }
    }

    /// Construct [`Self::Replace`].
    #[must_use]
    pub fn replace(index: usize, hyperlink: Link) -> Self {
        Self::Replace { index, hyperlink }
    }

    /// Construct [`Self::ReplaceAt`].
    #[must_use]
    pub fn replace_at(range: Range<usize>, hyperlink: Link) -> Self {
        Self::ReplaceAt { range, hyperlink }
    }

    /// Construct [`Self::Remove`].
    #[must_use]
    pub const fn remove(index: usize) -> Self {
        Self::Remove { index }
    }

    /// Construct [`Self::Clear`].
    #[must_use]
    pub const fn clear() -> Self {
        Self::Clear
    }
}

/// One existing cell and its ordered direct-hyperlink operations.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HyperlinkChange {
    row: usize,
    column: usize,
    operations: Vec<HyperlinkOperation>,
}

impl HyperlinkChange {
    /// Construct a hyperlink change at a zero-based logical coordinate.
    #[must_use]
    pub fn new(row: usize, column: usize, operations: impl Into<Vec<HyperlinkOperation>>) -> Self {
        Self {
            row,
            column,
            operations: operations.into(),
        }
    }

    /// Construct a change containing one operation.
    #[must_use]
    pub fn one(row: usize, column: usize, operation: HyperlinkOperation) -> Self {
        Self::new(row, column, vec![operation])
    }

    /// Construct a complete-cell [`HyperlinkOperation::Set`] change.
    #[must_use]
    pub fn set(row: usize, column: usize, hyperlink: Link) -> Self {
        Self::one(row, column, HyperlinkOperation::set(hyperlink))
    }

    /// Construct an [`HyperlinkOperation::Add`] change.
    #[must_use]
    pub fn add(row: usize, column: usize, range: Range<usize>, hyperlink: Link) -> Self {
        Self::one(row, column, HyperlinkOperation::add(range, hyperlink))
    }

    /// Construct a [`HyperlinkOperation::Replace`] change.
    #[must_use]
    pub fn replace(row: usize, column: usize, index: usize, hyperlink: Link) -> Self {
        Self::one(row, column, HyperlinkOperation::replace(index, hyperlink))
    }

    /// Construct a [`HyperlinkOperation::ReplaceAt`] change.
    #[must_use]
    pub fn replace_at(row: usize, column: usize, range: Range<usize>, hyperlink: Link) -> Self {
        Self::one(
            row,
            column,
            HyperlinkOperation::replace_at(range, hyperlink),
        )
    }

    /// Construct a [`HyperlinkOperation::Remove`] change.
    #[must_use]
    pub fn remove(row: usize, column: usize, index: usize) -> Self {
        Self::one(row, column, HyperlinkOperation::remove(index))
    }

    /// Construct a [`HyperlinkOperation::Clear`] change.
    #[must_use]
    pub fn clear(row: usize, column: usize) -> Self {
        Self::one(row, column, HyperlinkOperation::clear())
    }

    /// Zero-based logical row coordinate.
    #[must_use]
    pub const fn row(&self) -> usize {
        self.row
    }

    /// Zero-based logical column coordinate.
    #[must_use]
    pub const fn column(&self) -> usize {
        self.column
    }

    /// Ordered operations applied to the cell.
    #[must_use]
    pub fn operations(&self) -> &[HyperlinkOperation] {
        &self.operations
    }
}

/// An immutable semantic ODS cell snapshot bound to one positional source.
///
/// The snapshot shares the retained `content.xml` and worksheet projection.
/// It is not a complete-package snapshot: source publication remains the
/// responsibility of the originating [`SourceBackedSpreadsheet`].
#[derive(Clone)]
pub(crate) struct SourceCellSnapshot<'source> {
    owner: &'source SourceBackedSpreadsheet,
    content_xml: Arc<str>,
    sheets: Arc<[Sheet]>,
    source_version: SourceVersion,
}

impl fmt::Debug for SourceCellSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceCellSnapshot")
            .field("source_version", &self.source_version)
            .field("content_xml_bytes", &self.content_xml.len())
            .field("sheets", &self.sheets.len())
            .finish_non_exhaustive()
    }
}

/// A sparse, failure-atomic existing-cell transaction.
pub(crate) struct SourceCellEdit<'source> {
    before: SourceCellSnapshot<'source>,
    touched: BTreeMap<usize, Sheet>,
    coordinates: Vec<(usize, usize, usize)>,
    hyperlink_edit: bool,
}

impl fmt::Debug for SourceCellEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceCellEdit")
            .field("source_version", &self.before.source_version)
            .field("touched_sheets", &self.touched.len())
            .field("changed_cells", &self.coordinates.len())
            .finish_non_exhaustive()
    }
}

/// A checked source-backed cell edit: the candidate `content.xml`.
pub(crate) struct SourceCellCommit<'source> {
    snapshot: SourceCellSnapshot<'source>,
    changed_cells: usize,
}

impl fmt::Debug for SourceCellCommit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceCellCommit")
            .field("changed_cells", &self.changed_cells)
            .field("content_xml_bytes", &self.snapshot.content_xml.len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedSpreadsheet {
    /// Capture a cheap immutable cell-edit snapshot of this exact source.
    pub(crate) fn cell_snapshot(&self) -> Result<SourceCellSnapshot<'_>> {
        self.check_source()?;
        let snapshot = SourceCellSnapshot {
            owner: self,
            content_xml: Arc::clone(&self.content_xml),
            sheets: Arc::clone(&self.sheets),
            source_version: self.source_version,
        };
        self.check_source()?;
        Ok(snapshot)
    }
}

impl<'source> SourceCellSnapshot<'source> {
    /// Start an isolated sparse transaction from this snapshot.
    pub(crate) fn edit(&self) -> Result<SourceCellEdit<'source>> {
        self.check_source()?;
        Ok(SourceCellEdit {
            before: self.clone(),
            touched: BTreeMap::new(),
            coordinates: Vec::new(),
            hyperlink_edit: false,
        })
    }

    fn check_source(&self) -> Result<()> {
        let observed = self.owner.package.source_version()?;
        if observed == self.source_version {
            Ok(())
        } else {
            Err(Error::SourceChanged {
                expected: self.source_version,
                observed,
            })
        }
    }
}

impl<'source> SourceCellEdit<'source> {
    /// Number of effective unique cell changes staged so far.
    #[must_use]
    pub(crate) fn changed_cells(&self) -> usize {
        self.coordinates.len()
    }

    /// Whether this transaction currently has no effective changes.
    #[must_use]
    pub(crate) fn is_no_op(&self) -> bool {
        self.coordinates.is_empty()
    }

    fn check_protection(&self, sheet_name: &str) -> Result<()> {
        let (structure_protected, protected_sheets) = self.before.owner.edit_protection()?;
        if *structure_protected || protected_sheets.iter().any(|name| name == sheet_name) {
            return Err(Error::InvalidFormat(
                "ODS source cell edits refuse protected spreadsheets and worksheets".to_string(),
            ));
        }
        Ok(())
    }

    /// Replace one existing ordinary scalar cell.
    pub(crate) fn set_cell<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> Result<Option<bool>> {
        self.set_cells(selector, vec![CellChange::new(row, column, cell)])
            .map(|selected| selected.map(|changed| changed != 0))
    }

    /// Replace one existing formula cell while retaining its cached value,
    /// displayed text, style, and other cell metadata.
    pub(crate) fn set_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<Option<bool>> {
        self.set_formulas(selector, vec![FormulaChange::new(row, column, formula)])
            .map(|selected| selected.map(|changed| changed != 0))
    }

    /// Atomically apply ordered direct-hyperlink operations to existing cells.
    ///
    /// Missing cells, repeated physical rows, merged covered cells, duplicate
    /// coordinates, and unsupported row shapes are refused.  New or changed
    /// links are authoring-validated; unchanged parsed links are only required
    /// to pass storage validation.  Scalar and formula operations retain their
    /// separate refusal for rows containing hyperlinks.
    pub(crate) fn set_hyperlinks<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        mut changes: Vec<HyperlinkChange>,
    ) -> Result<Option<usize>> {
        self.before.check_source()?;
        validate_hyperlink_changes(&mut changes)?;
        let Some(sheet_index) =
            crate::worksheet::snapshot::select(&self.before.sheets, selector.into())?
        else {
            self.before.check_source()?;
            return Ok(None);
        };
        let source_sheet = self
            .touched
            .get(&sheet_index)
            .unwrap_or(&self.before.sheets[sheet_index]);
        for change in &changes {
            if self
                .coordinates
                .iter()
                .any(|coordinate| *coordinate == (sheet_index, change.row(), change.column()))
            {
                return Err(Error::InvalidFormat(format!(
                    "ODS source cell edit repeats coordinate ({}, {}) on sheet '{}'",
                    change.row(),
                    change.column(),
                    source_sheet.name
                )));
            }
        }
        let replacements = hyperlink_replacements(source_sheet, &changes)?;
        let effective =
            crate::worksheet::snapshot::effective_cell_changes(source_sheet, replacements);
        if effective.is_empty() {
            self.before.check_source()?;
            return Ok(Some(0));
        }
        self.check_protection(&source_sheet.name)?;
        let attempted = self
            .coordinates
            .len()
            .checked_add(effective.len())
            .unwrap_or(usize::MAX);
        if attempted > crate::worksheet::MAX_CELL_CHANGES {
            return Err(Error::InvalidFormat(format!(
                "ODS source cell edit exceeds the {} transaction safety limit",
                crate::worksheet::MAX_CELL_CHANGES
            )));
        }

        let mut candidate = source_sheet.clone();
        candidate.set_cells_prevalidated(
            effective
                .iter()
                .map(|change| (change.row(), change.column(), change.cell().clone()))
                .collect(),
        )?;
        validate_candidate_rows(source_sheet, &candidate, &effective)?;
        crate::worksheet::package::validate_owned_link_delta(
            std::slice::from_ref(source_sheet),
            std::slice::from_ref(&candidate),
        )?;
        self.before.check_source()?;

        self.coordinates
            .try_reserve(effective.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS source cell coordinates",
                source,
            })?;
        let previous = self.touched.insert(sheet_index, candidate);
        let old_len = self.coordinates.len();
        self.coordinates.extend(
            effective
                .iter()
                .map(|change| (sheet_index, change.row(), change.column())),
        );
        let previous_hyperlink_edit = self.hyperlink_edit;
        self.hyperlink_edit = true;
        if let Err(error) = self.before.check_source() {
            self.coordinates.truncate(old_len);
            if let Some(previous) = previous {
                self.touched.insert(sheet_index, previous);
            } else {
                self.touched.remove(&sheet_index);
            }
            self.hyperlink_edit = previous_hyperlink_edit;
            return Err(error);
        }
        Ok(Some(effective.len()))
    }

    /// Atomically replace a bounded batch of existing formula cells.
    ///
    /// Missing cells, scalar cells, repeated physical rows, merged cells,
    /// unknown values, and duplicate coordinates are refused. Repeated
    /// physical columns may be split because the containing row remains the
    /// same owner. Formula changes never evaluate or replace cached values.
    pub(crate) fn set_formulas<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        mut changes: Vec<FormulaChange>,
    ) -> Result<Option<usize>> {
        self.before.check_source()?;
        validate_formula_changes(&mut changes)?;
        let Some(sheet_index) =
            crate::worksheet::snapshot::select(&self.before.sheets, selector.into())?
        else {
            self.before.check_source()?;
            return Ok(None);
        };
        let source_sheet = self
            .touched
            .get(&sheet_index)
            .unwrap_or(&self.before.sheets[sheet_index]);
        for change in &changes {
            if self
                .coordinates
                .iter()
                .any(|coordinate| *coordinate == (sheet_index, change.row(), change.column()))
            {
                return Err(Error::InvalidFormat(format!(
                    "ODS source cell edit repeats coordinate ({}, {}) on sheet '{}'",
                    change.row(),
                    change.column(),
                    source_sheet.name
                )));
            }
        }
        let replacements = formula_replacements(source_sheet, &changes)?;
        let effective =
            crate::worksheet::snapshot::effective_cell_changes(source_sheet, replacements);
        if effective.is_empty() {
            self.before.check_source()?;
            return Ok(Some(0));
        }
        self.check_protection(&source_sheet.name)?;
        let attempted = self
            .coordinates
            .len()
            .checked_add(effective.len())
            .unwrap_or(usize::MAX);
        if attempted > crate::worksheet::MAX_CELL_CHANGES {
            return Err(Error::InvalidFormat(format!(
                "ODS source cell edit exceeds the {} transaction safety limit",
                crate::worksheet::MAX_CELL_CHANGES
            )));
        }

        let mut candidate = source_sheet.clone();
        candidate.set_cells_prevalidated(
            effective
                .iter()
                .map(|change| (change.row(), change.column(), change.cell().clone()))
                .collect(),
        )?;
        validate_candidate_rows(source_sheet, &candidate, &effective)?;
        self.before.check_source()?;

        self.coordinates
            .try_reserve(effective.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS source cell coordinates",
                source,
            })?;
        let previous = self.touched.insert(sheet_index, candidate);
        let old_len = self.coordinates.len();
        self.coordinates.extend(
            effective
                .iter()
                .map(|change| (sheet_index, change.row(), change.column())),
        );
        if let Err(error) = self.before.check_source() {
            self.coordinates.truncate(old_len);
            if let Some(previous) = previous {
                self.touched.insert(sheet_index, previous);
            } else {
                self.touched.remove(&sheet_index);
            }
            return Err(error);
        }
        Ok(Some(effective.len()))
    }

    /// Atomically replace one bounded batch of existing ordinary scalar cells.
    ///
    /// Missing cells, repeated physical rows, formulas, merged cells, unknown
    /// values, and style retargeting are refused. Repeated cell runs may be
    /// split because the containing physical row remains the same owner.
    pub(crate) fn set_cells<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        mut changes: Vec<CellChange>,
    ) -> Result<Option<usize>> {
        self.before.check_source()?;
        crate::worksheet::snapshot::validate_cell_changes(&mut changes)?;
        let Some(sheet_index) =
            crate::worksheet::snapshot::select(&self.before.sheets, selector.into())?
        else {
            self.before.check_source()?;
            return Ok(None);
        };
        let source_sheet = self
            .touched
            .get(&sheet_index)
            .unwrap_or(&self.before.sheets[sheet_index]);
        for change in &changes {
            if self
                .coordinates
                .iter()
                .any(|coordinate| *coordinate == (sheet_index, change.row(), change.column()))
            {
                return Err(Error::InvalidFormat(format!(
                    "ODS source cell edit repeats coordinate ({}, {}) on sheet '{}'",
                    change.row(),
                    change.column(),
                    source_sheet.name
                )));
            }
        }
        validate_existing_changes(source_sheet, &changes)?;
        let effective = crate::worksheet::snapshot::effective_cell_changes(source_sheet, changes);
        if effective.is_empty() {
            self.before.check_source()?;
            return Ok(Some(0));
        }
        self.check_protection(&source_sheet.name)?;
        let attempted = self
            .coordinates
            .len()
            .checked_add(effective.len())
            .unwrap_or(usize::MAX);
        if attempted > crate::worksheet::MAX_CELL_CHANGES {
            return Err(Error::InvalidFormat(format!(
                "ODS source cell edit exceeds the {} transaction safety limit",
                crate::worksheet::MAX_CELL_CHANGES
            )));
        }

        let mut candidate = source_sheet.clone();
        candidate.set_cells_prevalidated(
            effective
                .iter()
                .map(|change| (change.row(), change.column(), change.cell().clone()))
                .collect(),
        )?;
        validate_candidate_rows(source_sheet, &candidate, &effective)?;
        self.before.check_source()?;

        self.coordinates
            .try_reserve(effective.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS source cell coordinates",
                source,
            })?;
        let previous = self.touched.insert(sheet_index, candidate);
        let old_len = self.coordinates.len();
        self.coordinates.extend(
            effective
                .iter()
                .map(|change| (sheet_index, change.row(), change.column())),
        );
        if let Err(error) = self.before.check_source() {
            self.coordinates.truncate(old_len);
            if let Some(previous) = previous {
                self.touched.insert(sheet_index, previous);
            } else {
                self.touched.remove(&sheet_index);
            }
            return Err(error);
        }
        Ok(Some(effective.len()))
    }

    /// Validate the row-local replacement and freeze a publication commit.
    pub(crate) fn commit(self) -> Result<SourceCellCommit<'source>> {
        self.before.check_source()?;
        if self.touched.is_empty() {
            return Ok(SourceCellCommit {
                snapshot: self.before,
                changed_cells: 0,
            });
        }

        let mut candidates = Vec::new();
        candidates
            .try_reserve_exact(self.before.sheets.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS source cell candidate worksheets",
                source,
            })?;
        candidates.extend(
            self.before
                .sheets
                .iter()
                .enumerate()
                .map(|(index, _)| self.touched.get(&index)),
        );
        // Reuse the owner's cached content layout when a previous transaction
        // already scanned this immutable projection; otherwise scan now and
        // retain the layout for later transactions. Both paths run the same
        // gates before the scan in the same order.
        // The owner's cached layout is keyed to the exact immutable
        // `content.xml` Arc retained at open.  A committed source-cell
        // snapshot may be edited again before it is published; its XML is a
        // different allocation (and may have different row offsets), so the
        // owner's layout must not be reused for that derived snapshot.  A
        // stale layout would splice the second batch at offsets from the
        // original source and could corrupt otherwise valid chained patches.
        let rewrite_mode = if self.hyperlink_edit {
            crate::worksheet::package::RowRewriteMode::SourceBackedHyperlinks
        } else {
            crate::worksheet::package::RowRewriteMode::SourceBacked
        };
        let content_xml = if Arc::ptr_eq(&self.before.content_xml, &self.before.owner.content_xml)
            && let Some(layout) = self.before.owner.cached_content_layout()
        {
            crate::worksheet::package::replace_changed_rows_from_content_xml_with_layout_mode(
                &self.before.content_xml,
                layout,
                &self.before.sheets,
                &candidates,
                validation::MAX_CONTENT_XML_BYTES,
                rewrite_mode,
            )?
        } else {
            let (content, layout) = crate::worksheet::package::replace_changed_rows_from_content_xml_retaining_layout_mode(
                    &self.before.content_xml,
                    &self.before.sheets,
                    &candidates,
                    validation::MAX_CONTENT_XML_BYTES,
                    rewrite_mode,
                )?;
            if let Some(layout) = layout {
                self.before.owner.cache_content_layout(layout);
            }
            content
        }
        .ok_or_else(|| {
            Error::InvalidFormat(
                "ODS source cell edit is not eligible for exact row-local publication".to_string(),
            )
        })?;
        crate::authoring::validate_content_xml(&content_xml)?;
        let parsed = crate::worksheet::codec::parse(&content_xml)?;
        if parsed.len() != self.before.sheets.len()
            || parsed.iter().enumerate().any(|(index, sheet)| {
                self.touched
                    .get(&index)
                    .unwrap_or(&self.before.sheets[index])
                    != sheet
            })
        {
            return Err(Error::InvalidFormat(
                "ODS source cell candidate readback differs from staged state".to_string(),
            ));
        }
        self.before.check_source()?;
        let snapshot = SourceCellSnapshot {
            owner: self.before.owner,
            content_xml: Arc::from(content_xml),
            sheets: Arc::from(parsed),
            source_version: self.before.source_version,
        };
        Ok(SourceCellCommit {
            snapshot,
            changed_cells: self.coordinates.len(),
        })
    }
}

impl SourceCellCommit<'_> {
    /// The candidate `content.xml`.
    pub(crate) fn content_xml(&self) -> &Arc<str> {
        &self.snapshot.content_xml
    }

    /// Number of effective changed cells.
    pub(crate) const fn changed_cells(&self) -> usize {
        self.changed_cells
    }
}

fn validate_hyperlink_changes(changes: &mut [HyperlinkChange]) -> Result<()> {
    if changes.len() > crate::worksheet::MAX_CELL_CHANGES {
        return Err(Error::InvalidFormat(format!(
            "ODS hyperlink batch exceeds the {} operation safety limit",
            crate::worksheet::MAX_CELL_CHANGES
        )));
    }
    let mut operation_count = 0usize;
    let mut payload_bytes = 0usize;
    for change in changes.iter() {
        if change.row() >= validation::MAX_LOGICAL_ROWS {
            return Err(Error::InvalidFormat(format!(
                "ODS hyperlink batch row {} is outside the {}-row logical grid",
                change.row(),
                validation::MAX_LOGICAL_ROWS
            )));
        }
        if change.column() >= validation::MAX_LOGICAL_COLUMNS {
            return Err(Error::InvalidFormat(format!(
                "ODS hyperlink batch column {} is outside the {}-column logical grid",
                change.column(),
                validation::MAX_LOGICAL_COLUMNS
            )));
        }
        operation_count = operation_count
            .checked_add(change.operations().len())
            .ok_or_else(|| {
                Error::InvalidFormat("ODS hyperlink operation count overflows".to_string())
            })?;
        if operation_count > crate::worksheet::MAX_CELL_CHANGES {
            return Err(Error::InvalidFormat(format!(
                "ODS hyperlink batch exceeds the {} operation safety limit",
                crate::worksheet::MAX_CELL_CHANGES
            )));
        }
        for operation in change.operations() {
            let hyperlink = match operation {
                HyperlinkOperation::Set(hyperlink)
                | HyperlinkOperation::Add { hyperlink, .. }
                | HyperlinkOperation::Replace { hyperlink, .. }
                | HyperlinkOperation::ReplaceAt { hyperlink, .. } => Some(hyperlink),
                HyperlinkOperation::Remove { .. } | HyperlinkOperation::Clear => None,
            };
            if let Some(hyperlink) = hyperlink {
                hyperlink.validate_storage()?;
                for value in [
                    hyperlink.href.len(),
                    hyperlink.text.len(),
                    hyperlink.name.as_ref().map_or(0, String::len),
                    hyperlink.title.as_ref().map_or(0, String::len),
                    hyperlink.target_frame_name.as_ref().map_or(0, String::len),
                    hyperlink.style_name.as_ref().map_or(0, String::len),
                    hyperlink.visited_style_name.as_ref().map_or(0, String::len),
                ] {
                    payload_bytes = payload_bytes.checked_add(value).ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODS hyperlink batch retained string payload overflows".to_string(),
                        )
                    })?;
                    if payload_bytes > validation::MAX_CONTENT_XML_BYTES {
                        return Err(Error::InvalidFormat(format!(
                            "ODS hyperlink batch retained string payload exceeds the {}-byte safety limit",
                            validation::MAX_CONTENT_XML_BYTES
                        )));
                    }
                }
            }
        }
    }
    changes.sort_by_key(|change| (change.row(), change.column()));
    for repeated in changes.windows(2) {
        if repeated[0].row() == repeated[1].row() && repeated[0].column() == repeated[1].column() {
            return Err(Error::InvalidFormat(format!(
                "ODS hyperlink batch repeats logical coordinate ({}, {})",
                repeated[0].row(),
                repeated[0].column()
            )));
        }
    }
    Ok(())
}

fn hyperlink_replacements(sheet: &Sheet, changes: &[HyperlinkChange]) -> Result<Vec<CellChange>> {
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(changes.len())
        .map_err(|source| Error::Allocation {
            resource: "ODS source hyperlink replacements",
            source,
        })?;
    visit_existing_hyperlink_cells(sheet, changes, |_row, existing, change| {
        if existing.merge == Merge::Covered && !change.operations().is_empty() {
            return Err(Error::InvalidFormat(
                "ODS source hyperlink edits refuse covered cells".to_string(),
            ));
        }
        let mut replacement = existing.clone();
        for operation in change.operations() {
            apply_hyperlink_operation(&mut replacement, operation)?;
        }
        validation::validate_cell(&replacement)?;
        if !existing.equivalent_run(&replacement) {
            replacements.push(CellChange::new(change.row(), change.column(), replacement));
        }
        Ok(())
    })?;
    Ok(replacements)
}

fn apply_hyperlink_operation(cell: &mut Cell, operation: &HyperlinkOperation) -> Result<()> {
    match operation {
        HyperlinkOperation::Set(hyperlink) => cell.set_hyperlink(hyperlink.clone()),
        HyperlinkOperation::Add { range, hyperlink } => {
            cell.add_hyperlink(range.clone(), hyperlink.clone())
        },
        HyperlinkOperation::Replace { index, hyperlink } => {
            if cell.hyperlinks().get(*index) == Some(hyperlink) {
                return Ok(());
            }
            cell.replace_hyperlink(*index, hyperlink.clone())?;
            Ok(())
        },
        HyperlinkOperation::ReplaceAt { range, hyperlink } => {
            if cell
                .hyperlinks()
                .iter()
                .any(|existing| existing.range() == *range && existing == hyperlink)
            {
                return Ok(());
            }
            cell.replace_hyperlink_at(range.clone(), hyperlink.clone())?;
            Ok(())
        },
        HyperlinkOperation::Remove { index } => {
            let _removed = cell.remove_hyperlink(*index);
            Ok(())
        },
        HyperlinkOperation::Clear => {
            let _removed = cell.clear_hyperlinks();
            Ok(())
        },
    }
}

fn validate_existing_changes(sheet: &Sheet, changes: &[CellChange]) -> Result<()> {
    let mut validated_row = None;
    visit_existing_cells(sheet, changes, |row, existing, change| {
        if validated_row != Some(change.row()) {
            if row
                .cells
                .iter()
                .any(|cell| matches!(cell.value, CellValue::Unknown { .. }))
            {
                return Err(Error::InvalidFormat(
                    "ODS source cell edits refuse rows containing unknown values".to_string(),
                ));
            }
            if row.cells.iter().any(|cell| !cell.hyperlinks.is_empty()) {
                return Err(Error::InvalidFormat(
                    "ODS source cell edits refuse rows containing hyperlinks".to_string(),
                ));
            }
            validated_row = Some(change.row());
        }
        validate_plain_cell(existing, "source")?;
        validate_plain_cell(change.cell(), "replacement")?;
        if existing.style_name != change.cell().style_name {
            return Err(Error::InvalidFormat(
                "ODS source cell edits cannot retarget direct styles".to_string(),
            ));
        }
        Ok(())
    })
}

fn validate_formula_changes(changes: &mut [FormulaChange]) -> Result<()> {
    if changes.len() > crate::worksheet::MAX_CELL_CHANGES {
        return Err(Error::InvalidFormat(format!(
            "ODS formula batch exceeds the {} operation safety limit",
            crate::worksheet::MAX_CELL_CHANGES
        )));
    }
    let maximum_payload = validation::MAX_TEXT_BYTES.min(validation::MAX_CONTENT_XML_BYTES);
    let mut payload_bytes = 0u64;
    for change in changes.iter() {
        if change.row() >= validation::MAX_LOGICAL_ROWS {
            return Err(Error::InvalidFormat(format!(
                "ODS formula batch row {} is outside the {}-row logical grid",
                change.row(),
                validation::MAX_LOGICAL_ROWS
            )));
        }
        if change.column() >= validation::MAX_LOGICAL_COLUMNS {
            return Err(Error::InvalidFormat(format!(
                "ODS formula batch column {} is outside the {}-column logical grid",
                change.column(),
                validation::MAX_LOGICAL_COLUMNS
            )));
        }
        let formula_bytes = u64::try_from(change.formula().len()).unwrap_or(u64::MAX);
        payload_bytes = payload_bytes
            .checked_add(formula_bytes)
            .ok_or_else(|| formula_payload_limit(u64::MAX, maximum_payload))?;
        if payload_bytes > maximum_payload as u64 {
            return Err(formula_payload_limit(payload_bytes, maximum_payload));
        }
    }
    for change in changes.iter() {
        let mut validated = Cell::empty();
        validated.set_formula(change.formula().to_string())?;
    }
    changes.sort_by_key(|change| (change.row(), change.column()));
    for repeated in changes.windows(2) {
        if repeated[0].row() == repeated[1].row() && repeated[0].column() == repeated[1].column() {
            return Err(Error::InvalidFormat(format!(
                "ODS formula batch repeats logical coordinate ({}, {})",
                repeated[0].row(),
                repeated[0].column()
            )));
        }
    }
    Ok(())
}

fn formula_payload_limit(observed: u64, maximum: usize) -> Error {
    Error::ResourceLimit(litchi_core::ResourceLimit {
        resource: litchi_core::Resource::InputBytes,
        observed,
        limit: maximum as u64,
        scope: Arc::from("ODS source formula batch payload"),
    })
}

fn formula_replacements(sheet: &Sheet, changes: &[FormulaChange]) -> Result<Vec<CellChange>> {
    let mut replacements = Vec::with_capacity(changes.len());
    visit_existing_formula_cells(sheet, changes, |row, existing, change| {
        if row
            .cells
            .iter()
            .any(|cell| matches!(cell.value, CellValue::Unknown { .. }))
        {
            return Err(Error::InvalidFormat(
                "ODS source formula edits refuse rows containing unknown values".to_string(),
            ));
        }
        if row.cells.iter().any(|cell| !cell.hyperlinks.is_empty()) {
            return Err(Error::InvalidFormat(
                "ODS source formula edits refuse rows containing hyperlinks".to_string(),
            ));
        }
        if existing.formula.is_none() {
            return Err(Error::InvalidFormat(
                "ODS source formula edits require existing formula cells".to_string(),
            ));
        }
        if existing.merge != Merge::None {
            return Err(Error::InvalidFormat(
                "ODS source formula edits refuse merged cells".to_string(),
            ));
        }
        let mut replacement = existing.clone();
        replacement.set_formula(change.formula().to_string())?;
        replacements.push(CellChange::new(change.row(), change.column(), replacement));
        Ok(())
    })?;
    Ok(replacements)
}

fn validate_plain_cell(cell: &Cell, role: &str) -> Result<()> {
    if cell.formula.is_some() {
        return Err(Error::InvalidFormat(format!(
            "ODS source cell edits refuse {role} formulas"
        )));
    }
    if cell.merge != Merge::None {
        return Err(Error::InvalidFormat(format!(
            "ODS source cell edits refuse {role} merged cells"
        )));
    }
    if matches!(cell.value, CellValue::Unknown { .. }) {
        return Err(Error::InvalidFormat(format!(
            "ODS source cell edits refuse {role} unknown values"
        )));
    }
    if !cell.hyperlinks.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "ODS source cell edits refuse {role} hyperlinks"
        )));
    }
    Ok(())
}

fn validate_candidate_rows(
    source: &Sheet,
    candidate: &Sheet,
    changes: &[CellChange],
) -> Result<()> {
    if source.name != candidate.name
        || source.style_name != candidate.style_name
        || source.rows.len() != candidate.rows.len()
    {
        return Err(Error::InvalidFormat(
            "ODS source cell edit changed worksheet topology".to_string(),
        ));
    }
    visit_existing_cells(candidate, changes, |_, cell, change| {
        if !cell.equivalent_run(change.cell()) {
            return Err(Error::InvalidFormat(
                "ODS source cell candidate differs from staged replacement".to_string(),
            ));
        }
        Ok(())
    })
}

fn visit_existing_cells(
    sheet: &Sheet,
    changes: &[CellChange],
    mut visit: impl FnMut(&Row, &Cell, &CellChange) -> Result<()>,
) -> Result<()> {
    let mut row_index = 0usize;
    let mut row_start = 0usize;
    let mut logical_row = None;
    let mut cell_index = 0usize;
    let mut cell_start = 0usize;
    for change in changes {
        while let Some(row) = sheet.rows.get(row_index) {
            let row_end = row_start.checked_add(row.repeat()).ok_or_else(|| {
                Error::InvalidFormat("ODS source row range overflows".to_string())
            })?;
            if change.row() < row_end {
                break;
            }
            row_start = row_end;
            row_index += 1;
        }
        let Some(row) = sheet.rows.get(row_index) else {
            return missing_cell(change);
        };
        if row.repeat() != 1 {
            return Err(Error::InvalidFormat(
                "ODS source cell edits refuse repeated physical rows".to_string(),
            ));
        }
        if logical_row != Some(change.row()) {
            logical_row = Some(change.row());
            cell_index = 0;
            cell_start = 0;
        }
        while let Some(cell) = row.cells.get(cell_index) {
            let cell_end = cell_start.checked_add(cell.repeat()).ok_or_else(|| {
                Error::InvalidFormat("ODS source cell range overflows".to_string())
            })?;
            if change.column() < cell_end {
                visit(row, cell, change)?;
                break;
            }
            cell_start = cell_end;
            cell_index += 1;
        }
        if row.cells.get(cell_index).is_none() {
            return missing_cell(change);
        }
    }
    Ok(())
}

fn visit_existing_formula_cells(
    sheet: &Sheet,
    changes: &[FormulaChange],
    mut visit: impl FnMut(&Row, &Cell, &FormulaChange) -> Result<()>,
) -> Result<()> {
    let mut row_index = 0usize;
    let mut row_start = 0usize;
    let mut logical_row = None;
    let mut cell_index = 0usize;
    let mut cell_start = 0usize;
    for change in changes {
        while let Some(row) = sheet.rows.get(row_index) {
            let row_end = row_start.checked_add(row.repeat()).ok_or_else(|| {
                Error::InvalidFormat("ODS source row range overflows".to_string())
            })?;
            if change.row() < row_end {
                break;
            }
            row_start = row_end;
            row_index += 1;
        }
        let Some(row) = sheet.rows.get(row_index) else {
            return missing_formula_cell(change);
        };
        if row.repeat() != 1 {
            return Err(Error::InvalidFormat(
                "ODS source formula edits refuse repeated physical rows".to_string(),
            ));
        }
        if logical_row != Some(change.row()) {
            logical_row = Some(change.row());
            cell_index = 0;
            cell_start = 0;
        }
        while let Some(cell) = row.cells.get(cell_index) {
            let cell_end = cell_start.checked_add(cell.repeat()).ok_or_else(|| {
                Error::InvalidFormat("ODS source cell range overflows".to_string())
            })?;
            if change.column() < cell_end {
                visit(row, cell, change)?;
                break;
            }
            cell_start = cell_end;
            cell_index += 1;
        }
        if row.cells.get(cell_index).is_none() {
            return missing_formula_cell(change);
        }
    }
    Ok(())
}

fn visit_existing_hyperlink_cells(
    sheet: &Sheet,
    changes: &[HyperlinkChange],
    mut visit: impl FnMut(&Row, &Cell, &HyperlinkChange) -> Result<()>,
) -> Result<()> {
    let mut row_index = 0usize;
    let mut row_start = 0usize;
    let mut logical_row = None;
    let mut cell_index = 0usize;
    let mut cell_start = 0usize;
    for change in changes {
        while let Some(row) = sheet.rows.get(row_index) {
            let row_end = row_start.checked_add(row.repeat()).ok_or_else(|| {
                Error::InvalidFormat("ODS source row range overflows".to_string())
            })?;
            if change.row() < row_end {
                break;
            }
            row_start = row_end;
            row_index += 1;
        }
        let Some(row) = sheet.rows.get(row_index) else {
            return missing_hyperlink_cell(change);
        };
        if row.repeat() != 1 {
            return Err(Error::InvalidFormat(
                "ODS source hyperlink edits refuse repeated physical rows".to_string(),
            ));
        }
        if logical_row != Some(change.row()) {
            logical_row = Some(change.row());
            cell_index = 0;
            cell_start = 0;
        }
        while let Some(cell) = row.cells.get(cell_index) {
            let cell_end = cell_start.checked_add(cell.repeat()).ok_or_else(|| {
                Error::InvalidFormat("ODS source cell range overflows".to_string())
            })?;
            if change.column() < cell_end {
                visit(row, cell, change)?;
                break;
            }
            cell_start = cell_end;
            cell_index += 1;
        }
        if row.cells.get(cell_index).is_none() {
            return missing_hyperlink_cell(change);
        }
    }
    Ok(())
}

fn missing_cell(change: &CellChange) -> Result<()> {
    Err(Error::InvalidFormat(format!(
        "ODS source cell ({}, {}) does not exist",
        change.row(),
        change.column()
    )))
}

fn missing_formula_cell(change: &FormulaChange) -> Result<()> {
    Err(Error::InvalidFormat(format!(
        "ODS source formula cell ({}, {}) does not exist",
        change.row(),
        change.column()
    )))
}

fn missing_hyperlink_cell(change: &HyperlinkChange) -> Result<()> {
    Err(Error::InvalidFormat(format!(
        "ODS source hyperlink cell ({}, {}) does not exist",
        change.row(),
        change.column()
    )))
}
