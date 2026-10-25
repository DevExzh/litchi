//! Immutable annotation snapshots, failure-atomic transactions, and patches.

use super::{codec, model, validation};
use litchi_core::{Error, Result};
use litchi_odf_common::{XmlStreamLimits, annotation::Annotation};

/// An immutable, source-checked ODS cell-annotation snapshot.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(super) source: String,
    pub(super) entries: Vec<model::Entry>,
    pub(super) xml_limits: XmlStreamLimits,
}

impl Snapshot {
    /// Parse one complete ODS `content.xml` source.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(source: impl Into<String>, limits: &XmlStreamLimits) -> Result<Self> {
        let source = source.into();
        let parsed = codec::parse(&source, limits)?;
        validation::validate_entries(&parsed.entries)?;
        Ok(Self {
            source,
            entries: parsed.entries,
            xml_limits: *limits,
        })
    }

    /// The exact source XML captured by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source
    }

    /// Annotation entries in source order.
    #[must_use]
    pub fn entries(&self) -> &[model::Entry] {
        &self.entries
    }

    /// Checked source-order lookup.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn at(&self, index: usize) -> Result<&model::Entry> {
        self.entries.get(index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODS annotation index {index} is out of bounds for {} entries",
                self.entries.len()
            ))
        })
    }

    /// Find the annotation attached to an exact logical sheet/cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn cell(&self, sheet: &str, row: usize, column: usize) -> Result<Option<&model::Entry>> {
        let cell = model::Cell::new(sheet, row, column)?;
        Ok(self.entries.iter().find(|entry| entry.cell() == &cell))
    }

    /// Alias emphasizing the contextual sheet/cell selection.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn find(&self, sheet: &str, row: usize, column: usize) -> Result<Option<&model::Entry>> {
        self.cell(sheet, row, column)
    }

    /// Find an annotation by its optional ODF `office:name` identifier.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn named(&self, name: &str) -> Result<Option<&model::Entry>> {
        if name.is_empty() {
            return Err(Error::InvalidFormat(
                "ODS annotation name cannot be empty".to_string(),
            ));
        }
        Ok(self
            .entries
            .iter()
            .find(|entry| entry.annotation().name() == Some(name)))
    }

    /// Begin a failure-atomic transaction derived from this snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.clone(),
        }
    }
}

/// A staged annotation edit, handed to the closure of
/// [`crate::document::Edit::annotations`] and
/// [`crate::builder::Workbook::edit_annotations`]. The source snapshot is
/// never mutated.
pub struct Edit {
    before: Snapshot,
    draft: Snapshot,
}

impl Edit {
    /// Current candidate snapshot.
    #[must_use]
    pub fn current(&self) -> &Snapshot {
        &self.draft
    }

    /// Original source snapshot used for source checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Add a new annotation to an existing, non-repeated logical cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add(
        &mut self,
        sheet: &str,
        row: usize,
        column: usize,
        annotation: Annotation,
    ) -> Result<()> {
        let cell = model::Cell::new(sheet, row, column)?;
        self.add_at(cell, annotation)
    }

    /// Add a new annotation using a checked contextual cell selector.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_at(&mut self, cell: model::Cell, annotation: Annotation) -> Result<()> {
        validation::validate_annotation(&annotation)?;
        if self
            .draft
            .cell(cell.sheet(), cell.row(), cell.column())?
            .is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "ODS annotation already exists at '{}!R{}C{}'",
                cell.sheet(),
                cell.row().saturating_add(1),
                cell.column().saturating_add(1)
            )));
        }
        let parsed = codec::parse(self.draft.source_xml(), &self.draft.xml_limits)?;
        let site = codec::find_site(&parsed.sites, &cell).ok_or_else(|| model::not_found(&cell))?;
        ensure_unique_site(site, &cell)?;
        let fragment = codec::serialize(&annotation, &self.draft.xml_limits)?;
        let source = codec::insert(self.draft.source_xml(), site, &fragment)?;
        self.publish(source, Some(&cell))
    }

    /// Set a cell annotation, adding it when absent or replacing it in place.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set(
        &mut self,
        sheet: &str,
        row: usize,
        column: usize,
        annotation: Annotation,
    ) -> Result<()> {
        let cell = model::Cell::new(sheet, row, column)?;
        self.set_at(cell, annotation)
    }

    /// Set a cell annotation using a checked contextual selector.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_at(&mut self, cell: model::Cell, annotation: Annotation) -> Result<()> {
        if let Some(entry) = self.draft.cell(cell.sheet(), cell.row(), cell.column())? {
            let index = entry.index();
            self.replace(index, annotation)
        } else {
            self.add_at(cell, annotation)
        }
    }

    /// Replace an annotation by its checked source-order index.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace(&mut self, index: usize, annotation: Annotation) -> Result<()> {
        let entry = self.draft.at(index)?.clone();
        self.replace_at(entry.cell().clone(), annotation)
    }

    /// Replace an annotation at an exact contextual cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace_at(&mut self, cell: model::Cell, annotation: Annotation) -> Result<()> {
        validation::validate_annotation(&annotation)?;
        self.draft
            .cell(cell.sheet(), cell.row(), cell.column())?
            .ok_or_else(|| model::not_found(&cell))?;
        let parsed = codec::parse(self.draft.source_xml(), &self.draft.xml_limits)?;
        let site = codec::find_site(&parsed.sites, &cell).ok_or_else(|| model::not_found(&cell))?;
        ensure_unique_site(site, &cell)?;
        let fragment = codec::serialize(&annotation, &self.draft.xml_limits)?;
        let span = annotation_span(&parsed, &cell)?;
        let source = codec::replace(self.draft.source_xml(), span, &fragment)?;
        self.publish(source, Some(&cell))
    }

    /// Remove an annotation by source-order index and return its value.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove(&mut self, index: usize) -> Result<Annotation> {
        let entry = self.draft.at(index)?.clone();
        let annotation = entry.annotation().clone();
        self.remove_at(entry.cell().clone())?;
        Ok(annotation)
    }

    /// Remove an annotation at an exact contextual cell and return its value.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove_at(&mut self, cell: model::Cell) -> Result<Annotation> {
        let entry = self
            .draft
            .cell(cell.sheet(), cell.row(), cell.column())?
            .ok_or_else(|| model::not_found(&cell))?
            .clone();
        let parsed = codec::parse(self.draft.source_xml(), &self.draft.xml_limits)?;
        let span = annotation_span(&parsed, &cell)?;
        let source = codec::remove(self.draft.source_xml(), span)?;
        let annotation = entry.annotation().clone();
        self.publish(source, None)?;
        Ok(annotation)
    }

    /// Remove an annotation by sheet name and zero-based coordinates.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear(&mut self, sheet: &str, row: usize, column: usize) -> Result<Annotation> {
        let cell = model::Cell::new(sheet, row, column)?;
        self.remove_at(cell)
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.draft = self.before.clone();
    }

    /// Validate and atomically materialize the candidate.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn commit(self) -> Result<Commit> {
        validation::validate_entries(&self.draft.entries)?;
        let changed = self.before.source != self.draft.source;
        Ok(Commit {
            snapshot: self.draft,
            changed,
        })
    }

    /// Adopt `source` as the candidate; an added or replaced annotation must
    /// read back at `written`.
    fn publish(&mut self, source: String, written: Option<&model::Cell>) -> Result<()> {
        let draft = Snapshot::parse(source, &self.before.xml_limits)?;
        if let Some(cell) = written {
            draft
                .cell(cell.sheet(), cell.row(), cell.column())?
                .ok_or_else(|| model::not_found(cell))?;
        }
        self.draft = draft;
        Ok(())
    }
}

/// A validated annotation publication.
#[derive(Debug)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    changed: bool,
}

impl Commit {
    /// Whether the source XML changed.
    #[must_use]
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }

    #[cfg(test)]
    /// Resulting immutable annotation snapshot.
    #[must_use]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Resulting source XML.
    #[must_use]
    pub(crate) fn content_xml(&self) -> &str {
        self.snapshot.source_xml()
    }
}

fn ensure_unique_site(site: &codec::Site, cell: &model::Cell) -> Result<()> {
    if site.rows_repeated != 1 || site.columns_repeated != 1 {
        return Err(Error::InvalidFormat(format!(
            "ODS annotation cell '{}!R{}C{}' is represented by a repeated table run",
            cell.sheet(),
            cell.row().saturating_add(1),
            cell.column().saturating_add(1)
        )));
    }
    Ok(())
}

fn annotation_span<'a>(parsed: &'a codec::Parsed, cell: &model::Cell) -> Result<&'a codec::Span> {
    let entry = parsed
        .entries
        .iter()
        .find(|entry| entry.cell() == cell)
        .ok_or_else(|| model::not_found(cell))?;
    parsed
        .annotation_spans
        .get(entry.index())
        .ok_or_else(|| Error::InvalidFormat("ODS annotation source span is missing".to_string()))
}
