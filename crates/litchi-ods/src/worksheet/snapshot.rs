//! Package-bound worksheet snapshots and reversible edits.

use std::{fmt, sync::Arc};

use litchi_core::{Error, Position, Result};

use super::{Cell, Sheet, validation};
use crate::package::Package;

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
    source: Arc<[u8]>,
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
        Self::from_arc(Arc::from(source))
    }

    fn from_arc(source: Arc<[u8]>) -> Result<Self> {
        let package = Package::from_bytes(source.as_ref().to_vec())?;
        Ok(Self {
            source,
            sheets: Arc::from(package.sheets()?),
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
        let package = Package::from_bytes(self.before.source.as_ref().to_vec())?;
        let content = match super::package::try_replace_changed_rows(
            package.content_xml(),
            self.before.sheets(),
            &self.draft,
            validation::MAX_CONTENT_XML_BYTES,
        )? {
            Some(content)
                if litchi_odf_common::compact_xml::validate(content.as_bytes()).is_ok() =>
            {
                content
            },
            Some(_) | None => super::package::replace_tables(package.content_xml(), &self.draft)?,
        };
        litchi_odf_common::compact_xml::validate(content.as_bytes()).map_err(Error::from)?;
        let target = Snapshot::from_bytes(package.replace_content_xml(&content)?.into_bytes())?;
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
    source: Arc<[u8]>,
    target: Arc<[u8]>,
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
        self.source.as_ref() == snapshot.as_bytes()
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

    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
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

fn bounds(position: usize, length: usize) -> Error {
    Error::InvalidFormat(format!(
        "ODS worksheet position {position} is outside sheet count {length}"
    ))
}
