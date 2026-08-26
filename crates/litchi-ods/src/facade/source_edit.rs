//! Source-bound existing-cell edits for positional ODS packages.
//!
//! The transaction owns no complete package bytes. It retains the already
//! validated worksheet projection, clones only selected worksheet models, and
//! prepares one checked `content.xml` replacement. Publication delegates to
//! the common ODF positional writer so every other ZIP member remains a raw
//! source copy.

use std::{collections::BTreeMap, fmt, io::Write, sync::Arc};

use litchi_core::{Error, Result, SourceVersion};
use litchi_odf_common::core::{
    SourceContentPublicationError, SourceContentPublicationOptions, SourceContentPublicationReport,
};

use super::SourceBackedSpreadsheet;
use crate::worksheet::Row;
use crate::worksheet::{Cell, CellChange, CellValue, Merge, Selector, Sheet, validation};

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

/// An immutable semantic ODS cell snapshot bound to one positional source.
///
/// The snapshot shares the retained `content.xml` and worksheet projection.
/// It is not a complete-package snapshot: source publication remains the
/// responsibility of the originating [`SourceBackedSpreadsheet`].
#[derive(Clone)]
pub struct SourceCellSnapshot<'source> {
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
pub struct SourceCellEdit<'source> {
    before: SourceCellSnapshot<'source>,
    touched: BTreeMap<usize, Sheet>,
    coordinates: Vec<(usize, usize, usize)>,
    structure_protected: bool,
    protected_sheets: Vec<String>,
}

impl fmt::Debug for SourceCellEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceCellEdit")
            .field("source_version", &self.before.source_version)
            .field("touched_sheets", &self.touched.len())
            .field("changed_cells", &self.coordinates.len())
            .field("structure_protected", &self.structure_protected)
            .finish_non_exhaustive()
    }
}

/// A semantic `content.xml` patch bound to one retained positional owner.
///
/// This patch is intentionally not an exact ZIP-artifact patch. Applying it
/// validates the complete semantic source snapshot and returns the retained
/// target projection; publication still requires the checked commit that
/// created it. Reopening a streamed artifact establishes a new source lineage.
#[derive(Clone)]
pub struct SourceCellPatch<'source> {
    before: SourceCellSnapshot<'source>,
    after: SourceCellSnapshot<'source>,
}

impl fmt::Debug for SourceCellPatch<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceCellPatch")
            .field("changed", &self.changed())
            .field("source_content_bytes", &self.before.content_xml.len())
            .field("target_content_bytes", &self.after.content_xml.len())
            .finish()
    }
}

/// A checked source-backed cell edit ready for sequential publication.
pub struct SourceCellCommit<'source> {
    snapshot: SourceCellSnapshot<'source>,
    patch: SourceCellPatch<'source>,
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

/// Complete report for one source-backed ODS cell publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceCellPublicationReport {
    changed_cells: usize,
    content: SourceContentPublicationReport,
}

impl SourceCellPublicationReport {
    /// Number of semantic cells changed by the transaction.
    #[must_use]
    pub const fn changed_cells(self) -> usize {
        self.changed_cells
    }

    /// Whether the publication copied the exact source artifact.
    #[must_use]
    pub const fn is_no_op(self) -> bool {
        self.content.is_no_op()
    }

    /// Complete candidate artifact bytes accepted by the sink.
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.content.bytes()
    }

    /// Source version checked throughout publication.
    #[must_use]
    pub const fn source_version(self) -> SourceVersion {
        self.content.source_version()
    }

    /// The common ODF transport report.
    #[must_use]
    pub const fn content(self) -> SourceContentPublicationReport {
        self.content
    }
}

impl SourceBackedSpreadsheet {
    /// Capture a cheap immutable cell-edit snapshot of this exact source.
    pub fn cell_snapshot(&self) -> Result<SourceCellSnapshot<'_>> {
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

    /// Begin one sparse existing-cell edit over this exact source.
    pub fn edit_cells(&self) -> Result<SourceCellEdit<'_>> {
        self.cell_snapshot()?.edit()
    }

    /// Alias for [`Self::edit_cells`].
    pub fn begin_cell_edit(&self) -> Result<SourceCellEdit<'_>> {
        self.edit_cells()
    }
}

impl<'source> SourceCellSnapshot<'source> {
    /// Exact source version captured by this snapshot.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Borrow the validated candidate `content.xml`.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.content_xml
    }

    /// Borrow the complete validated worksheet projection.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Select a worksheet by exact name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<&Sheet>> {
        crate::worksheet::snapshot::select(&self.sheets, selector.into())
            .map(|selected| selected.map(|index| &self.sheets[index]))
    }

    /// Start an isolated sparse transaction from this snapshot.
    pub fn edit(&self) -> Result<SourceCellEdit<'source>> {
        self.check_source()?;
        let (structure_protected, protected_sheets) = self.owner.edit_protection()?.clone();
        self.check_source()?;
        Ok(SourceCellEdit {
            before: self.clone(),
            touched: BTreeMap::new(),
            coordinates: Vec::new(),
            structure_protected,
            protected_sheets,
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

    fn same_snapshot(&self, other: &Self) -> bool {
        std::ptr::eq(self.owner, other.owner)
            && self.source_version == other.source_version
            && self.content_xml.as_ref() == other.content_xml.as_ref()
            && self.sheets.as_ref() == other.sheets.as_ref()
    }
}

impl<'source> SourceCellEdit<'source> {
    /// Exact source snapshot captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &SourceCellSnapshot<'source> {
        &self.before
    }

    /// Number of effective unique cell changes staged so far.
    #[must_use]
    pub fn changed_cells(&self) -> usize {
        self.coordinates.len()
    }

    /// Whether this transaction currently has no effective changes.
    #[must_use]
    pub fn is_no_op(&self) -> bool {
        self.coordinates.is_empty()
    }

    /// Replace one existing ordinary scalar cell.
    pub fn set_cell<'a>(
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
    pub fn set_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<Option<bool>> {
        self.set_formulas(selector, vec![FormulaChange::new(row, column, formula)])
            .map(|selected| selected.map(|changed| changed != 0))
    }

    /// Atomically replace a bounded batch of existing formula cells.
    ///
    /// Missing cells, scalar cells, repeated physical rows, merged cells,
    /// unknown values, and duplicate coordinates are refused. Repeated
    /// physical columns may be split because the containing row remains the
    /// same owner. Formula changes never evaluate or replace cached values.
    pub fn set_formulas<'a>(
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
        if self.structure_protected
            || self
                .protected_sheets
                .iter()
                .any(|name| name == &source_sheet.name)
        {
            return Err(Error::InvalidFormat(
                "ODS source cell edits refuse protected spreadsheets and worksheets".to_string(),
            ));
        }
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
    pub fn set_cells<'a>(
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
        if self.structure_protected
            || self
                .protected_sheets
                .iter()
                .any(|name| name == &source_sheet.name)
        {
            return Err(Error::InvalidFormat(
                "ODS source cell edits refuse protected spreadsheets and worksheets".to_string(),
            ));
        }
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
    pub fn commit(self) -> Result<SourceCellCommit<'source>> {
        self.before.check_source()?;
        if self.touched.is_empty() {
            let snapshot = self.before.clone();
            return Ok(SourceCellCommit {
                patch: SourceCellPatch {
                    before: self.before,
                    after: snapshot.clone(),
                },
                snapshot,
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
        let content_xml = if Arc::ptr_eq(&self.before.content_xml, &self.before.owner.content_xml)
            && let Some(layout) = self.before.owner.cached_content_layout()
        {
            crate::worksheet::package::replace_changed_rows_from_content_xml_with_layout(
                &self.before.content_xml,
                layout,
                &self.before.sheets,
                &candidates,
                validation::MAX_CONTENT_XML_BYTES,
            )?
        } else {
            let (content, layout) =
                crate::worksheet::package::replace_changed_rows_from_content_xml_retaining_layout(
                    &self.before.content_xml,
                    &self.before.sheets,
                    &candidates,
                    validation::MAX_CONTENT_XML_BYTES,
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
        let patch = SourceCellPatch {
            before: self.before,
            after: snapshot.clone(),
        };
        Ok(SourceCellCommit {
            snapshot,
            patch,
            changed_cells: self.coordinates.len(),
        })
    }
}

impl<'source> SourceCellPatch<'source> {
    /// Exact semantic source required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceCellSnapshot<'source> {
        &self.before
    }

    /// Exact semantic target produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceCellSnapshot<'source> {
        &self.after
    }

    /// Whether the candidate `content.xml` changes.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.before.content_xml.as_ref() != self.after.content_xml.as_ref()
    }

    /// Apply this semantic patch to its exact in-memory source snapshot.
    pub fn apply(
        &self,
        snapshot: &SourceCellSnapshot<'source>,
    ) -> Result<SourceCellSnapshot<'source>> {
        snapshot.check_source()?;
        if !snapshot.same_snapshot(&self.before) {
            return Err(Error::InvalidFormat(
                "ODS source cell patch source snapshot does not match".to_string(),
            ));
        }
        Ok(self.after.clone())
    }

    /// Return the exact inverse semantic `content.xml` patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

impl<'source> SourceCellCommit<'source> {
    /// Candidate semantic snapshot after this edit.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceCellSnapshot<'source> {
        &self.snapshot
    }

    /// Source-content semantic patch for this edit.
    #[must_use]
    pub const fn patch(&self) -> &SourceCellPatch<'source> {
        &self.patch
    }

    /// Number of effective changed cells.
    #[must_use]
    pub const fn changed_cells(&self) -> usize {
        self.changed_cells
    }

    /// Whether the edit changed `content.xml`.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.changed()
    }

    /// Publish the candidate to a sequential sink under the default policy.
    pub fn write_to<W: Write>(
        &self,
        writer: W,
    ) -> std::result::Result<SourceCellPublicationReport, SourceContentPublicationError> {
        self.write_to_with_options(writer, SourceContentPublicationOptions::default())
    }

    /// Publish under explicit replacement/output, cancellation, payload, and
    /// hierarchical execution policies.
    pub fn write_to_with_options<W: Write>(
        &self,
        writer: W,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceCellPublicationReport, SourceContentPublicationError> {
        self.patch
            .before
            .check_source()
            .map_err(|error| match error {
                Error::SourceChanged { expected, observed } => {
                    SourceContentPublicationError::SourceChanged {
                        expected,
                        observed,
                        progress:
                            litchi_odf_common::core::SourceContentPublicationProgress::Untouched,
                    }
                },
                other => SourceContentPublicationError::Core(other),
            })?;
        let content = self
            .patch
            .before
            .owner
            .package
            .write_content_xml_to_stream_with_options(
                writer,
                self.snapshot.content_xml.as_bytes(),
                options,
            )?;
        Ok(SourceCellPublicationReport {
            changed_cells: self.changed_cells,
            content,
        })
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
