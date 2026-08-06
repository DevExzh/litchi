//! Immutable annotation snapshots, failure-atomic transactions, and patches.

use super::{codec, model, validation};
use litchi_core::{Error, Result};
use litchi_odf_common::annotation::Annotation;

/// An immutable, source-checked ODS cell-annotation snapshot.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(super) source: String,
    pub(super) entries: Vec<model::Entry>,
}

impl Snapshot {
    /// Parse one complete ODS `content.xml` source.
    pub fn parse(source: impl Into<String>) -> Result<Self> {
        let source = source.into();
        let parsed = codec::parse(&source)?;
        validation::validate_entries(&parsed.entries)?;
        Ok(Self {
            source,
            entries: parsed.entries,
        })
    }

    /// The exact source XML captured by this snapshot.
    pub fn source_xml(&self) -> &str {
        &self.source
    }

    /// Annotation entries in source order.
    pub fn entries(&self) -> &[model::Entry] {
        &self.entries
    }

    /// Checked source-order lookup.
    pub fn at(&self, index: usize) -> Result<&model::Entry> {
        self.entries.get(index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODS annotation index {index} is out of bounds for {} entries",
                self.entries.len()
            ))
        })
    }

    /// Find the annotation attached to an exact logical sheet/cell.
    pub fn cell(&self, sheet: &str, row: usize, column: usize) -> Result<Option<&model::Entry>> {
        let cell = model::Cell::new(sheet, row, column)?;
        Ok(self.entries.iter().find(|entry| entry.cell() == &cell))
    }

    /// Alias emphasizing the contextual sheet/cell selection.
    pub fn find(&self, sheet: &str, row: usize, column: usize) -> Result<Option<&model::Entry>> {
        self.cell(sheet, row, column)
    }

    /// Find an annotation by its optional ODF `office:name` identifier.
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
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            draft: self.clone(),
            operations: Vec::new(),
        }
    }
}

/// One semantic change recorded by an annotation transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Operation {
    /// Insert an annotation at an otherwise unannotated cell.
    Add {
        cell: model::Cell,
        annotation: Annotation,
    },
    /// Replace the complete common annotation value at a cell.
    Replace {
        cell: model::Cell,
        before: Annotation,
        after: Annotation,
    },
    /// Remove the annotation at a cell.
    Remove {
        cell: model::Cell,
        annotation: Annotation,
    },
}

/// A source-checked, reversible sequence of semantic annotation operations.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    operations: Vec<Operation>,
}

impl Patch {
    /// Operations in transaction order.
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    /// Whether the patch contains no semantic changes.
    pub fn is_empty(&self) -> bool {
        self.operations.is_empty()
    }

    /// Fingerprint of the source XML this patch expects.
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the inverse patch, checked against the committed target source.
    pub fn inverse(&self) -> Self {
        Self {
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            operations: self
                .operations
                .iter()
                .rev()
                .map(Operation::inverse)
                .collect(),
        }
    }

    /// Apply this patch only to the exact source snapshot from which it came.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if codec::fingerprint(snapshot.source_xml()) != self.source_fingerprint {
            return Err(Error::InvalidFormat(
                "ODS annotation patch source snapshot does not match".to_string(),
            ));
        }
        let mut transaction = snapshot.edit();
        for operation in &self.operations {
            transaction.apply_operation(operation)?;
        }
        transaction.commit()
    }
}

impl Operation {
    fn inverse(&self) -> Self {
        match self {
            Self::Add { cell, annotation } => Self::Remove {
                cell: cell.clone(),
                annotation: annotation.clone(),
            },
            Self::Replace {
                cell,
                before,
                after,
            } => Self::Replace {
                cell: cell.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::Remove { cell, annotation } => Self::Add {
                cell: cell.clone(),
                annotation: annotation.clone(),
            },
        }
    }
}

/// A staged annotation edit. The source snapshot is never mutated.
pub struct Transaction {
    before: Snapshot,
    draft: Snapshot,
    operations: Vec<Operation>,
}

impl Transaction {
    /// Create a transaction from an owned snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot) -> Self {
        snapshot.edit()
    }

    /// Current candidate snapshot.
    pub fn snapshot(&self) -> &Snapshot {
        &self.draft
    }

    /// Original source snapshot used for source checks and inverse patches.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Add a new annotation to an existing, non-repeated logical cell.
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
        let parsed = codec::parse(self.draft.source_xml())?;
        let site = codec::find_site(&parsed.sites, &cell).ok_or_else(|| model::not_found(&cell))?;
        ensure_unique_site(site, &cell)?;
        let fragment = codec::serialize(&annotation)?;
        let source = codec::insert(self.draft.source_xml(), site, &fragment)?;
        self.publish(source, Operation::Add { cell, annotation })
    }

    /// Set a cell annotation, adding it when absent or replacing it in place.
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
    pub fn set_at(&mut self, cell: model::Cell, annotation: Annotation) -> Result<()> {
        if let Some(entry) = self.draft.cell(cell.sheet(), cell.row(), cell.column())? {
            let index = entry.index();
            self.replace(index, annotation)
        } else {
            self.add_at(cell, annotation)
        }
    }

    /// Replace an annotation by its checked source-order index.
    pub fn replace(&mut self, index: usize, annotation: Annotation) -> Result<()> {
        let entry = self.draft.at(index)?.clone();
        self.replace_at(entry.cell().clone(), annotation)
    }

    /// Replace an annotation at an exact contextual cell.
    pub fn replace_at(&mut self, cell: model::Cell, annotation: Annotation) -> Result<()> {
        validation::validate_annotation(&annotation)?;
        let entry = self
            .draft
            .cell(cell.sheet(), cell.row(), cell.column())?
            .ok_or_else(|| model::not_found(&cell))?
            .clone();
        let parsed = codec::parse(self.draft.source_xml())?;
        let site = codec::find_site(&parsed.sites, &cell).ok_or_else(|| model::not_found(&cell))?;
        ensure_unique_site(site, &cell)?;
        let fragment = codec::serialize(&annotation)?;
        let span = annotation_span(&parsed, &cell)?;
        let source = codec::replace(self.draft.source_xml(), span, &fragment)?;
        self.publish(
            source,
            Operation::Replace {
                cell,
                before: entry.annotation().clone(),
                after: annotation,
            },
        )
    }

    /// Remove an annotation by source-order index and return its value.
    pub fn remove(&mut self, index: usize) -> Result<Annotation> {
        let entry = self.draft.at(index)?.clone();
        let annotation = entry.annotation().clone();
        self.remove_at(entry.cell().clone())?;
        Ok(annotation)
    }

    /// Remove an annotation at an exact contextual cell and return its value.
    pub fn remove_at(&mut self, cell: model::Cell) -> Result<Annotation> {
        let entry = self
            .draft
            .cell(cell.sheet(), cell.row(), cell.column())?
            .ok_or_else(|| model::not_found(&cell))?
            .clone();
        let parsed = codec::parse(self.draft.source_xml())?;
        let span = annotation_span(&parsed, &cell)?;
        let source = codec::remove(self.draft.source_xml(), span)?;
        let annotation = entry.annotation().clone();
        self.publish(
            source,
            Operation::Remove {
                cell,
                annotation: annotation.clone(),
            },
        )?;
        Ok(annotation)
    }

    /// Remove an annotation by sheet name and zero-based coordinates.
    pub fn clear(&mut self, sheet: &str, row: usize, column: usize) -> Result<Annotation> {
        let cell = model::Cell::new(sheet, row, column)?;
        self.remove_at(cell)
    }

    /// Restore the transaction's original source and semantic inventory.
    pub fn rollback(&mut self) {
        self.draft = self.before.clone();
        self.operations.clear();
    }

    /// Validate and atomically materialize the candidate.
    pub fn commit(self) -> Result<Commit> {
        validation::validate_entries(&self.draft.entries)?;
        let changed = self.before.source != self.draft.source;
        let patch = if changed {
            Patch {
                source_fingerprint: codec::fingerprint(self.before.source_xml()),
                target_fingerprint: codec::fingerprint(self.draft.source_xml()),
                operations: self.operations,
            }
        } else {
            Patch {
                source_fingerprint: codec::fingerprint(self.before.source_xml()),
                target_fingerprint: codec::fingerprint(self.before.source_xml()),
                operations: Vec::new(),
            }
        };
        Ok(Commit {
            snapshot: self.draft,
            patch,
            changed,
        })
    }

    fn apply_operation(&mut self, operation: &Operation) -> Result<()> {
        match operation {
            Operation::Add { cell, annotation } => self.add_at(cell.clone(), annotation.clone()),
            Operation::Replace {
                cell,
                before,
                after,
            } => {
                let current = self
                    .draft
                    .cell(cell.sheet(), cell.row(), cell.column())?
                    .ok_or_else(|| model::not_found(cell))?;
                if current.annotation() != before {
                    return Err(Error::InvalidFormat(
                        "ODS annotation patch expected-state conflict".to_string(),
                    ));
                }
                self.replace_at(cell.clone(), after.clone())
            },
            Operation::Remove { cell, annotation } => {
                let current = self
                    .draft
                    .cell(cell.sheet(), cell.row(), cell.column())?
                    .ok_or_else(|| model::not_found(cell))?;
                if current.annotation() != annotation {
                    return Err(Error::InvalidFormat(
                        "ODS annotation patch expected-state conflict".to_string(),
                    ));
                }
                self.remove_at(cell.clone()).map(|_| ())
            },
        }
    }

    fn publish(&mut self, source: String, operation: Operation) -> Result<()> {
        let draft = Snapshot::parse(source)?;
        let operation = match operation {
            Operation::Add { cell, .. } => {
                let annotation = draft
                    .cell(cell.sheet(), cell.row(), cell.column())?
                    .ok_or_else(|| model::not_found(&cell))?
                    .annotation()
                    .clone();
                Operation::Add { cell, annotation }
            },
            Operation::Replace { cell, before, .. } => {
                let after = draft
                    .cell(cell.sheet(), cell.row(), cell.column())?
                    .ok_or_else(|| model::not_found(&cell))?
                    .annotation()
                    .clone();
                Operation::Replace {
                    cell,
                    before,
                    after,
                }
            },
            operation @ Operation::Remove { .. } => operation,
        };
        self.draft = draft;
        self.operations.push(operation);
        Ok(())
    }
}

/// A validated annotation publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the source XML changed.
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting immutable annotation snapshot.
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Resulting source XML.
    pub fn content_xml(&self) -> &str {
        self.snapshot.source_xml()
    }

    /// Reversible semantic patch produced by this commit.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its resulting snapshot.
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the commit into its resulting XML source.
    pub fn into_owned_xml(self) -> String {
        self.snapshot.source
    }

    /// Consume the commit into its resulting patch.
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// Apply one concise closure-based annotation edit.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit>
where
    F: FnOnce(&mut Transaction) -> Result<()>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
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
