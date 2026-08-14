//! Package-bound worksheet snapshots and reversible edits.

use std::{fmt, sync::Arc};

use litchi_core::{Error, Position, Result};

use super::{Cell, Sheet, validation};
use crate::package::Package;

/// Maximum number of logical cell replacements accepted by one batch.
pub const MAX_CELL_CHANGES: usize = 4_096;

/// One owned logical-cell replacement in a worksheet batch.
#[derive(Clone, Debug, PartialEq)]
pub struct CellChange {
    row: usize,
    column: usize,
    cell: Cell,
}

impl CellChange {
    /// Create one logical-cell replacement.
    #[must_use]
    pub const fn new(row: usize, column: usize, cell: Cell) -> Self {
        Self { row, column, cell }
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

    /// Replacement cell.
    #[must_use]
    pub const fn cell(&self) -> &Cell {
        &self.cell
    }
}

/// Exact sheet name or checked zero-based source position.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Selector<'a> {
    Name(&'a str),
    Position(Position),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

/// An immutable worksheet graph bound to exact ODS package bytes.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<Vec<u8>>,
    package: Arc<Package>,
    sheets: Arc<[Sheet]>,
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.source.len())
            .field("sheets", &self.sheets.len())
            .finish()
    }
}

impl Snapshot {
    /// Parse an owned ODS package and capture its bounded worksheet graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or worksheet graph is invalid.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        Self::from_arc(Arc::new(source))
    }

    fn from_arc(source: Arc<Vec<u8>>) -> Result<Self> {
        let package = Package::from_shared_bytes(Arc::clone(&source))?;
        Self::from_package(package)
    }

    /// Adopt a package that has already passed the ODS package boundary.
    ///
    /// Worksheet graph parsing and validation remain mandatory; only the
    /// immutable ZIP/archive ownership is reused.
    pub(crate) fn from_package(package: Package) -> Result<Self> {
        Self::from_shared_package(Arc::new(package))
    }

    pub(crate) fn from_shared_package(package: Arc<Package>) -> Result<Self> {
        let source = package.shared_bytes_owner();
        let sheets = Arc::from(package.sheets()?);
        Ok(Self {
            source,
            package,
            sheets,
        })
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.source
    }

    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    #[cfg(test)]
    pub(crate) fn prepared_index_identity(&self) -> usize {
        self.package.prepared_index_identity()
    }

    /// Select one sheet.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact name is ambiguous.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<&Sheet>> {
        select(&self.sheets, selector.into())
            .map(|selected| selected.map(|position| &self.sheets[position]))
    }

    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.sheets.to_vec(),
        }
    }
}

/// A clone-staged worksheet edit.
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Vec<Sheet>,
}

impl Edit {
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.draft
    }

    /// Append one worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete graph is invalid.
    pub fn add(&mut self, sheet: Sheet) -> Result<()> {
        let mut candidate = self.draft.clone();
        candidate.push(sheet);
        self.publish(candidate)
    }

    /// Replace one selected worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error when selection or graph validation fails.
    pub fn replace<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        sheet: Sheet,
    ) -> Result<Option<Sheet>> {
        let Some(index) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        let mut candidate = self.draft.clone();
        let previous = std::mem::replace(&mut candidate[index], sheet);
        self.publish(candidate)?;
        Ok(Some(previous))
    }

    /// Remove one selected worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error when selection or graph validation fails.
    pub fn remove<'a>(&mut self, selector: impl Into<Selector<'a>>) -> Result<Option<Sheet>> {
        let Some(index) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        let mut candidate = self.draft.clone();
        let removed = candidate.remove(index);
        self.publish(candidate)?;
        Ok(Some(removed))
    }

    /// Move one selected worksheet to its final position.
    ///
    /// # Errors
    ///
    /// Returns an error when selection fails or the destination is out of range.
    pub fn move_to<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        destination: Position,
    ) -> Result<Option<()>> {
        let Some(source) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        let destination_index = destination.get();
        if destination_index >= self.draft.len() {
            return Err(bounds(destination_index, self.draft.len()));
        }
        if source != destination_index {
            let mut candidate = self.draft.clone();
            let sheet = candidate.remove(source);
            candidate.insert(destination_index, sheet);
            self.publish(candidate)?;
        }
        Ok(Some(()))
    }

    /// Replace one logical cell.
    ///
    /// # Errors
    ///
    /// Returns an error when selection, mutation, or graph validation fails.
    pub fn set_cell<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> Result<Option<()>> {
        self.update(selector.into(), |sheet| sheet.set_cell(row, column, cell))
    }

    /// Atomically replace a bounded collection of logical cells.
    ///
    /// Semantic no-ops are omitted. The returned count is the number of
    /// replacements staged in the selected worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error when the batch exceeds [`MAX_CELL_CHANGES`], contains
    /// an invalid or repeated cell, repeats a coordinate, or cannot produce a
    /// valid worksheet graph. Every failure leaves this edit unchanged.
    pub fn set_cells<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        mut changes: Vec<CellChange>,
    ) -> Result<Option<usize>> {
        if changes.len() > MAX_CELL_CHANGES {
            return Err(Error::InvalidFormat(format!(
                "ODS cell batch exceeds the {MAX_CELL_CHANGES} operation safety limit"
            )));
        }
        let Some(index) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };

        validate_aggregate_cell_payload(&changes)?;
        for change in &changes {
            if change.row >= validation::MAX_LOGICAL_ROWS {
                return Err(Error::InvalidFormat(format!(
                    "ODS cell batch row {} is outside the {}-row logical grid",
                    change.row,
                    validation::MAX_LOGICAL_ROWS
                )));
            }
            if change.column >= validation::MAX_LOGICAL_COLUMNS {
                return Err(Error::InvalidFormat(format!(
                    "ODS cell batch column {} is outside the {}-column logical grid",
                    change.column,
                    validation::MAX_LOGICAL_COLUMNS
                )));
            }
            change.cell.validate()?;
            if change.cell.repeat() != 1 {
                return Err(Error::InvalidFormat(
                    "setting one logical cell requires a non-repeated Cell".to_string(),
                ));
            }
        }
        changes.sort_by_key(|change| (change.row, change.column));
        for repeated in changes.windows(2) {
            if repeated[0].row == repeated[1].row && repeated[0].column == repeated[1].column {
                return Err(Error::InvalidFormat(format!(
                    "ODS cell batch repeats logical coordinate ({}, {})",
                    repeated[0].row, repeated[0].column
                )));
            }
        }

        let effective = effective_cell_changes(&self.draft[index], changes);
        let changed = effective.len();
        if changed == 0 {
            return Ok(Some(0));
        }
        let mut candidate = self.draft.clone();
        candidate[index].set_cells_prevalidated(
            effective
                .into_iter()
                .map(|change| (change.row, change.column, change.cell))
                .collect(),
        )?;
        self.publish(candidate)?;
        Ok(Some(changed))
    }

    /// Clear one logical cell while retaining its direct style.
    ///
    /// # Errors
    ///
    /// Returns an error when selection, mutation, or graph validation fails.
    pub fn clear_cell<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
    ) -> Result<Option<()>> {
        self.update(selector.into(), |sheet| sheet.clear_cell(row, column))
    }

    /// Set one inert cell formula.
    ///
    /// # Errors
    ///
    /// Returns an error when selection, mutation, or graph validation fails.
    pub fn set_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<Option<()>> {
        let formula_text = formula.into();
        self.update(selector.into(), |sheet| {
            sheet.set_formula(row, column, formula_text)
        })
    }

    /// Set one direct cell style reference.
    ///
    /// # Errors
    ///
    /// Returns an error when selection, mutation, or graph validation fails.
    pub fn set_cell_style<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        row: usize,
        column: usize,
        style_name: impl Into<String>,
    ) -> Result<Option<()>> {
        let style_name_text = style_name.into();
        self.update(selector.into(), |sheet| {
            sheet.set_cell_style(row, column, style_name_text)
        })
    }

    /// Replace the complete ordered worksheet graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate graph is invalid.
    pub fn replace_all(&mut self, sheets: Vec<Sheet>) -> Result<()> {
        self.publish(sheets)
    }

    pub fn rollback(&mut self) {
        self.draft = self.before.sheets.to_vec();
    }

    /// Validate, rewrite compact content XML, reparse, and publish atomically.
    ///
    /// # Errors
    ///
    /// Returns an error when validation, compactness, rebuilding, or readback fails.
    pub fn commit(self) -> Result<Commit> {
        validation::validate_sheets(&self.draft)?;
        if self.draft.as_slice() == self.before.sheets() {
            return Ok(Commit::unchanged(self.before));
        }
        let package = Arc::clone(&self.before.package);
        let row_local = super::package::try_replace_changed_rows_spliced(
            package.package(),
            self.before.sheets(),
            &self.draft,
            validation::MAX_CONTENT_XML_BYTES,
        )?;
        let (content, target_package, provenance_spliced) = match row_local {
            Some(changed)
                if litchi_odf_common::compact_xml::validate(changed.content.as_bytes()).is_ok() =>
            {
                let target =
                    package.replace_spliced_content_xml(&changed.content, changed.publication)?;
                (changed.content, target, true)
            },
            Some(_) | None => {
                let content = super::package::replace_tables(package.content_xml(), &self.draft)?;
                let target = package.replace_content_xml(&content)?;
                (content, target, false)
            },
        };
        litchi_odf_common::compact_xml::validate(content.as_bytes()).map_err(Error::from)?;
        let target = Snapshot::from_package(target_package)?;
        if target.sheets() != self.draft {
            return Err(Error::InvalidFormat(
                "ODS worksheet typed readback does not match the staged edit".to_string(),
            ));
        }
        Ok(Commit {
            patch: Patch {
                source: self.before.source.clone(),
                target: target.source.clone(),
            },
            snapshot: target,
            provenance_spliced,
        })
    }

    fn update(
        &mut self,
        selector: Selector<'_>,
        operation: impl FnOnce(&mut Sheet) -> Result<()>,
    ) -> Result<Option<()>> {
        let Some(index) = select(&self.draft, selector)? else {
            return Ok(None);
        };
        let mut candidate = self.draft.clone();
        operation(&mut candidate[index])?;
        self.publish(candidate)?;
        Ok(Some(()))
    }

    fn publish(&mut self, candidate: Vec<Sheet>) -> Result<()> {
        validation::validate_sheets(&candidate)?;
        self.draft = candidate;
        Ok(())
    }
}

/// An exact-source, reversible worksheet patch.
#[derive(Clone)]
pub struct Patch {
    source: Arc<Vec<u8>>,
    target: Arc<Vec<u8>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_bytes", &self.source.len())
            .field("target_bytes", &self.target.len())
            .finish()
    }
}

impl Patch {
    #[must_use]
    pub fn changed(&self) -> bool {
        self.source != self.target
    }

    #[must_use]
    pub fn is_applicable_to(&self, snapshot: &Snapshot) -> bool {
        self.source.as_slice() == snapshot.as_bytes()
    }

    /// Apply this patch to its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale source or invalid target.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if !self.is_applicable_to(snapshot) {
            return Err(Error::InvalidFormat(
                "ODS worksheet patch source snapshot does not match".to_string(),
            ));
        }
        Ok(Commit {
            snapshot: Snapshot::from_arc(self.target.clone())?,
            patch: self.clone(),
            provenance_spliced: false,
        })
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }
}

/// A validated worksheet publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    provenance_spliced: bool,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        let source = snapshot.source.clone();
        Self {
            snapshot,
            patch: Patch {
                source: source.clone(),
                target: source,
            },
            provenance_spliced: false,
        }
    }

    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.changed()
    }

    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    pub(crate) const fn content_provenance_spliced(&self) -> bool {
        self.provenance_spliced
    }

    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

impl Snapshot {
    pub(crate) fn into_shared_bytes(self) -> Arc<Vec<u8>> {
        self.source
    }

    pub(crate) fn package_owner(&self) -> Arc<Package> {
        Arc::clone(&self.package)
    }
}

fn select(sheets: &[Sheet], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Position(position) => {
            Ok((position.get() < sheets.len()).then_some(position.get()))
        },
        Selector::Name(name) => {
            let mut matches = sheets
                .iter()
                .enumerate()
                .filter(|(_, sheet)| sheet.name == name);
            let selected = matches.next().map(|(index, _)| index);
            if selected.is_some() && matches.next().is_some() {
                return Err(Error::InvalidFormat(
                    "ODS worksheet selector is ambiguous".to_string(),
                ));
            }
            Ok(selected)
        },
    }
}

fn effective_cell_changes(sheet: &Sheet, changes: Vec<CellChange>) -> Vec<CellChange> {
    let mut effective = Vec::with_capacity(changes.len());
    let mut row_index = 0usize;
    let mut row_start = 0usize;
    let mut logical_row = None;
    let mut cell_index = 0usize;
    let mut cell_start = 0usize;

    for change in changes {
        while let Some(row) = sheet.rows.get(row_index) {
            let row_end = row_start.saturating_add(row.repeat());
            if change.row < row_end {
                break;
            }
            row_start = row_end;
            row_index += 1;
        }
        if logical_row != Some(change.row) {
            logical_row = Some(change.row);
            cell_index = 0;
            cell_start = 0;
        }

        let existing = sheet.rows.get(row_index).and_then(|row| {
            while let Some(cell) = row.cells.get(cell_index) {
                let cell_end = cell_start.saturating_add(cell.repeat());
                if change.column < cell_end {
                    return Some(cell);
                }
                cell_start = cell_end;
                cell_index += 1;
            }
            None
        });
        if existing.is_none_or(|cell| !cell.equivalent_run(&change.cell)) {
            effective.push(change);
        }
    }
    effective
}

fn validate_aggregate_cell_payload(changes: &[CellChange]) -> Result<()> {
    let mut total = 0usize;
    for change in changes {
        checked_add_payload_bytes(&mut total, change.cell.text.len())?;
        checked_add_payload_bytes(
            &mut total,
            change.cell.formula.as_ref().map_or(0, String::len),
        )?;
        checked_add_payload_bytes(
            &mut total,
            change.cell.style_name.as_ref().map_or(0, String::len),
        )?;
        match &change.cell.value {
            super::CellValue::Text(value)
            | super::CellValue::Date(value)
            | super::CellValue::Time(value) => {
                checked_add_payload_bytes(&mut total, value.len())?;
            },
            super::CellValue::Currency { currency, .. } => {
                checked_add_payload_bytes(&mut total, currency.len())?;
            },
            super::CellValue::Unknown { kind, value } => {
                checked_add_payload_bytes(&mut total, kind.len())?;
                checked_add_payload_bytes(&mut total, value.as_ref().map_or(0, String::len))?;
            },
            super::CellValue::Empty
            | super::CellValue::Number(_)
            | super::CellValue::Percentage(_)
            | super::CellValue::Boolean(_) => {},
        }
    }
    Ok(())
}

fn checked_add_payload_bytes(total: &mut usize, bytes: usize) -> Result<()> {
    *total = total.checked_add(bytes).ok_or_else(|| {
        Error::InvalidFormat("ODS cell batch retained string payload overflows".to_string())
    })?;
    if *total > validation::MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODS cell batch retained string payload exceeds the {}-byte safety limit",
            validation::MAX_CONTENT_XML_BYTES
        )));
    }
    Ok(())
}

fn bounds(position: usize, length: usize) -> Error {
    Error::InvalidFormat(format!(
        "ODS worksheet position {position} is outside sheet count {length}"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Builder, CellValue};

    #[test]
    fn snapshot_and_package_share_the_exact_archive_allocation() -> Result<()> {
        let mut sheet = Sheet::new("Data")?;
        sheet.set_cell(
            0,
            0,
            Cell::new(CellValue::Text("before".to_string()), "before"),
        )?;
        let mut builder = Builder::new();
        builder.add_sheet(sheet)?;
        let source = builder.build()?;
        let source_pointer = source.as_ptr();

        let snapshot = Snapshot::from_bytes(source)?;
        assert_eq!(snapshot.source.as_slice().as_ptr(), source_pointer);
        let package = Package::from_shared_bytes(Arc::clone(&snapshot.source))?;
        assert!(Arc::ptr_eq(&snapshot.source, &package.shared_bytes()));

        let package_identity = package.prepared_index_identity();
        let adopted = Snapshot::from_package(package)?;
        assert_eq!(adopted.package.prepared_index_identity(), package_identity);

        let mut edit = snapshot.edit();
        assert!(
            edit.set_cell(
                "Data",
                0,
                0,
                Cell::new(CellValue::Text("after".to_string()), "after"),
            )?
            .is_some()
        );
        let commit = edit.commit()?;
        assert!(commit.content_provenance_spliced());
        Ok(())
    }

    #[test]
    fn aggregate_cell_payload_accepts_exact_limit_and_rejects_next_byte() -> Result<()> {
        let mut exact = 0;
        checked_add_payload_bytes(&mut exact, validation::MAX_CONTENT_XML_BYTES)?;
        assert_eq!(exact, validation::MAX_CONTENT_XML_BYTES);

        let error = checked_add_payload_bytes(&mut exact, 1);
        assert!(matches!(
            error,
            Err(Error::InvalidFormat(message)) if message.contains("payload exceeds")
        ));

        let mut overflow = usize::MAX;
        assert!(checked_add_payload_bytes(&mut overflow, 1).is_err());
        Ok(())
    }
}
