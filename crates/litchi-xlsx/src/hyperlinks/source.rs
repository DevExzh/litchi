//! Guarded source-backed publication for one XLSX worksheet hyperlink owner.
//!
//! [`SourceBackedWorkbook::edit_hyperlinks`] stages one worksheet without
//! materializing any unselected part body; [`Edit::commit`] validates the
//! edit and returns a [`Commit`] holding the new workbook, whose output is the
//! byte-preserving splice the former stream publisher wrote (the selected
//! worksheet XML is the only replaced member; every other ZIP member is
//! raw-copied). An exact no-op returns the source workbook unchanged; a
//! changed edit of a signed source is refused at commit.
//! [`SourceBackedWorkbook::apply_hyperlinks`] replays a [`Patch`] onto a workbook
//! holding its before state (record 0360).

use std::io::Write;

use litchi_opc::SourceBackedPackage;

use super::snapshot::{Snapshot, SourceEntry};
use crate::error::{Result, allocation, invalid};
use crate::{Selector, SourceBackedWorkbook};
use litchi_opc::PackURI;
use std::collections::{HashMap, HashSet};

const MAX_HYPERLINKS: usize = 1_048_576;
const MAX_INCREMENTAL_HYPERLINKS: usize = 65_536;
const MAX_EDIT_MUTATIONS: usize = 256;
const EXECUTION_CHECK_INTERVAL: usize = 256;

crate::workbook::source::deferred::source_family!("hyperlink", super::Patch);

impl Patch {
    /// The exact worksheet state the patch requires.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact worksheet state the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        self.0.after()
    }
}

impl crate::workbook::source::deferred::Change for super::Patch {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        super::Patch::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        // A replay compares content: the selected worksheet, its owner and
        // relationships must hold the before state (record 0360).
        let current = Snapshot::load_source_backed(package, self.before().sheet_position())?;
        if !current.same_source(self.before()) {
            return Err(Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        Ok(())
    }

    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        package.check_execution()?;
        crate::workbook::source::deferred::Change::check(self, package)?;
        validate_effective_publication_source(package, self.after().worksheet_part_name())?;
        package.clone().write_part_overlay_shared_to_stream(
            writer,
            self.after().worksheet_part_name(),
            self.after().source_arc()?,
        )?;
        Ok(())
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        package.check_execution()?;
        validate_effective_publication_source(package, self.after().worksheet_part_name())?;
        package.clone().write_part_overlay_shared_to_stream(
            writer,
            self.after().worksheet_part_name(),
            self.after().source_arc()?,
        )?;
        Ok(())
    }
}

impl SourceBackedWorkbook {
    /// Begin an isolated hyperlink edit of one worksheet without materializing
    /// any unselected part body.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector does not name a worksheet, its XML
    /// is invalid, or the source cannot be read.
    pub fn edit_hyperlinks<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        if self.edit_package()?.has_encrypted_entries() {
            return Err(Error::Unsupported {
                feature: "encrypted XLSX source-backed hyperlink editing",
            });
        }
        let before = Snapshot::load_source_backed(self.edit_package()?, selector)?;
        Edit::new(self.clone(), before)
    }

    /// Replay a hyperlink patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected worksheet does not
    /// hold the patch's before state, and the signed-source refusal for a
    /// changed patch on a signed source.
    pub fn apply_hyperlinks(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated hyperlink edit over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    staged: Vec<SourceEntry>,
    mutations: usize,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, before: Snapshot) -> Result<Self> {
        let staged = before.source_entries()?;
        validate_entries(&staged, &before)?;
        Ok(Self {
            workbook,
            before,
            staged,
            mutations: 0,
        })
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged direct worksheet hyperlinks.
    #[must_use]
    pub fn hyperlinks(&self) -> impl ExactSizeIterator<Item = &super::Hyperlink> {
        self.staged.iter().map(|entry| &entry.value)
    }

    /// Borrow the currently staged direct worksheet hyperlinks.
    pub fn values(&self) -> Result<Vec<super::Hyperlink>> {
        clone_values(&self.staged, &self.before)
    }

    /// Insert or replace one hyperlink by its checked reference.
    ///
    /// New external relationships are rejected. Existing external links may
    /// only retain their exact target and private relationship ID.
    pub fn put(&mut self, value: super::Hyperlink) -> Result<bool> {
        self.replace_by_reference(value)
    }

    /// Replace one existing hyperlink, allowing its reference to change.
    pub fn replace_at(
        &mut self,
        reference: super::HyperlinkReference,
        value: super::Hyperlink,
    ) -> Result<bool> {
        self.before.check_execution()?;
        self.ensure_incremental_mutation()?;
        let index = self
            .staged
            .iter()
            .position(|entry| entry.value.reference().range() == reference.range());
        let Some(index) = index else {
            return Err(invalid("hyperlink replacement reference did not resolve"));
        };
        if self.staged[index].value == value {
            return Ok(false);
        }
        let replacement_key = value.reference().range();
        if replacement_key != reference.range()
            && self
                .staged
                .iter()
                .any(|entry| entry.value.reference().range() == replacement_key)
        {
            return Err(invalid("XLSX worksheet has duplicate hyperlink references"));
        }
        self.staged[index].value = value;
        self.record_mutation();
        self.before.check_execution()?;
        Ok(true)
    }

    /// Remove one hyperlink by its checked reference.
    ///
    /// Removing an external hyperlink is stageable, but commit refuses it
    /// because this focused overlay does not mutate worksheet relationships.
    pub fn remove(
        &mut self,
        reference: super::HyperlinkReference,
    ) -> Result<Option<super::Hyperlink>> {
        self.before.check_execution()?;
        self.ensure_incremental_mutation()?;
        let Some(index) = self
            .staged
            .iter()
            .position(|entry| entry.value.reference().range() == reference.range())
        else {
            return Ok(None);
        };
        let removed = self.staged.remove(index).value;
        self.record_mutation();
        self.before.check_execution()?;
        Ok(Some(removed))
    }

    /// Replace the complete staged hyperlink list.
    pub fn set(&mut self, values: Vec<super::Hyperlink>) -> Result<bool> {
        self.before.check_execution()?;
        let candidate = assign_entries(&self.before, &self.staged, values)?;
        validate_entries(&candidate, &self.before)?;
        if candidate == self.staged {
            return Ok(false);
        }
        self.ensure_mutation_slot()?;
        self.staged = candidate;
        self.record_mutation();
        self.before.check_execution()?;
        Ok(true)
    }

    /// Remove every hyperlink.
    ///
    /// A commit still refuses the edit if this would remove an external
    /// relationship owner.
    pub fn clear(&mut self) -> Result<bool> {
        if self.staged.is_empty() {
            return Ok(false);
        }
        self.before.check_execution()?;
        self.ensure_mutation_slot()?;
        self.staged.clear();
        self.record_mutation();
        self.before.check_execution()?;
        Ok(true)
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        let changed = self.mutations != 0 && !self.before.matches_entries_checked(&self.staged)?;
        if !changed {
            let patch = super::Patch::new(self.before.clone(), self.before);
            return commit_change(&self.workbook, patch);
        }
        if self.before.protected() {
            return Err(Error::Unsupported {
                feature: "protected XLSX hyperlink editing",
            });
        }
        validate_entries(&self.staged, &self.before)?;
        enforce_relationship_policy(&self.before, &self.staged)?;
        let values = clone_values(&self.staged, &self.before)?;
        let mut relationship_ids = Vec::new();
        relationship_ids
            .try_reserve_exact(self.staged.len())
            .map_err(|source| allocation("source-backed hyperlink relationship IDs", source))?;
        for (index, entry) in self.staged.iter().enumerate() {
            if index % EXECUTION_CHECK_INTERVAL == 0 {
                self.before.check_execution()?;
            }
            relationship_ids.push(entry.relationship_id.as_deref());
        }
        let output = super::codec::rewrite_hyperlinks_checked(
            self.before.source_xml(),
            &values,
            &relationship_ids,
            || self.before.check_execution(),
        )?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output, &self.staged)?;
        snapshot.check_execution()?;
        commit_change(&self.workbook, super::Patch::new(self.before, snapshot))
    }

    fn replace_by_reference(&mut self, value: super::Hyperlink) -> Result<bool> {
        self.before.check_execution()?;
        self.ensure_incremental_mutation()?;
        let key = value.reference().range();
        let Some(index) = self
            .staged
            .iter()
            .position(|entry| entry.value.reference().range() == key)
        else {
            if self.staged.len() >= MAX_HYPERLINKS {
                return Err(invalid(format!(
                    "XLSX worksheet exceeds the {MAX_HYPERLINKS} hyperlink safety limit"
                )));
            }
            self.staged
                .try_reserve(1)
                .map_err(|source| allocation("source-backed worksheet hyperlinks", source))?;
            let relationship_id = self
                .before
                .source_entry_refs()
                .find(|entry| entry.value.reference().range() == key)
                .and_then(|entry| entry.relationship_id.map(copy_boxed_relationship_id))
                .transpose()?;
            self.staged.push(SourceEntry {
                value,
                relationship_id,
            });
            self.record_mutation();
            self.before.check_execution()?;
            return Ok(true);
        };
        if self.staged[index].value == value {
            return Ok(false);
        }
        self.staged[index].value = value;
        self.record_mutation();
        self.before.check_execution()?;
        Ok(true)
    }

    fn ensure_mutation_slot(&self) -> Result<()> {
        if self.mutations >= MAX_EDIT_MUTATIONS {
            return Err(Error::Unsupported {
                feature: "more than 256 mutations in one source-backed hyperlink edit",
            });
        }
        Ok(())
    }

    fn ensure_incremental_mutation(&self) -> Result<()> {
        self.ensure_mutation_slot()?;
        if self.before.hyperlinks().len().max(self.staged.len()) > MAX_INCREMENTAL_HYPERLINKS {
            return Err(Error::Unsupported {
                feature: "incremental hyperlink mutation above 65,536 entries; use set instead",
            });
        }
        Ok(())
    }

    fn record_mutation(&mut self) {
        self.mutations += 1;
    }
}

fn clone_values(entries: &[SourceEntry], before: &Snapshot) -> Result<Vec<super::Hyperlink>> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(entries.len())
        .map_err(|source| allocation("source-backed worksheet hyperlinks", source))?;
    for (index, entry) in entries.iter().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        values.push(entry.value.clone());
    }
    before.check_execution()?;
    Ok(values)
}

fn validate_entries(entries: &[SourceEntry], before: &Snapshot) -> Result<()> {
    if entries.len() > MAX_HYPERLINKS {
        return Err(invalid(format!(
            "XLSX worksheet exceeds the {MAX_HYPERLINKS} hyperlink safety limit"
        )));
    }
    let mut ranges = HashSet::new();
    ranges
        .try_reserve(entries.len())
        .map_err(|source| allocation("source-backed hyperlink range index", source))?;
    for (index, entry) in entries.iter().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        if !ranges.insert(entry.value.reference().range()) {
            return Err(invalid("XLSX worksheet has duplicate hyperlink references"));
        }
    }
    before.check_execution()?;
    Ok(())
}

fn enforce_relationship_policy(before: &Snapshot, candidate: &[SourceEntry]) -> Result<()> {
    let mut original_external = HashMap::new();
    original_external
        .try_reserve(before.hyperlinks().len())
        .map_err(|source| allocation("source-backed hyperlink relationship index", source))?;
    for (index, entry) in before.source_entry_refs().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        if let Some(id) = entry.relationship_id {
            let target = entry.value.external_target();
            let (_, count) = original_external.entry(id).or_insert((target, 0usize));
            *count = count
                .checked_add(1)
                .ok_or_else(|| invalid("hyperlink relationship reference count overflows usize"))?;
        }
    }
    let mut candidate_external = HashMap::new();
    candidate_external
        .try_reserve(candidate.len())
        .map_err(|source| allocation("source-backed hyperlink relationship index", source))?;
    for (index, entry) in candidate.iter().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        if entry.relationship_id.is_none() && entry.value.external_target().is_some() {
            return Err(relationship_policy_refusal());
        }
        if let Some(id) = entry.relationship_id.as_deref() {
            if entry.value.external_target().is_none() {
                return Err(relationship_policy_refusal());
            }
            let Some((original_target, _)) = original_external.get(id) else {
                return Err(relationship_policy_refusal());
            };
            if *original_target != entry.value.external_target() {
                return Err(relationship_policy_refusal());
            }
            let count = candidate_external.entry(id).or_insert(0usize);
            *count = count
                .checked_add(1)
                .ok_or_else(|| invalid("hyperlink relationship reference count overflows usize"))?;
        }
    }
    if candidate_external.len() != original_external.len()
        || original_external
            .iter()
            .any(|(id, (_, original_count))| candidate_external.get(id) != Some(original_count))
    {
        return Err(relationship_policy_refusal());
    }
    before.check_execution()?;
    Ok(())
}

fn relationship_policy_refusal() -> Error {
    Error::Unsupported {
        feature: "XLSX source-backed hyperlink relationship topology changes",
    }
}

fn assign_entries(
    before: &Snapshot,
    current: &[SourceEntry],
    values: Vec<super::Hyperlink>,
) -> Result<Vec<SourceEntry>> {
    if values.len() > MAX_HYPERLINKS {
        return Err(invalid(format!(
            "XLSX worksheet exceeds the {MAX_HYPERLINKS} hyperlink safety limit"
        )));
    }
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(values.len())
        .map_err(|source| allocation("source-backed worksheet hyperlinks", source))?;
    let mut baseline_ids = HashMap::new();
    baseline_ids
        .try_reserve(before.hyperlinks().len())
        .map_err(|source| allocation("source-backed hyperlink baseline index", source))?;
    for (index, entry) in before.source_entry_refs().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        baseline_ids.insert(entry.value.reference().range(), entry.relationship_id);
    }
    let mut current_ids = HashMap::new();
    current_ids
        .try_reserve(current.len())
        .map_err(|source| allocation("source-backed hyperlink staged index", source))?;
    for (index, entry) in current.iter().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        current_ids.insert(
            entry.value.reference().range(),
            entry.relationship_id.as_deref(),
        );
    }
    for (index, value) in values.into_iter().enumerate() {
        if index % EXECUTION_CHECK_INTERVAL == 0 {
            before.check_execution()?;
        }
        let range = value.reference().range();
        let relationship_id = current_ids
            .get(&range)
            .copied()
            .unwrap_or_else(|| baseline_ids.get(&range).copied().flatten())
            .map(copy_boxed_relationship_id)
            .transpose()?;
        candidate.push(SourceEntry {
            value,
            relationship_id,
        });
    }
    before.check_execution()?;
    Ok(candidate)
}

fn copy_boxed_relationship_id(value: &str) -> Result<Box<str>> {
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("source-backed hyperlink relationship ID", source))?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}

fn validate_effective_publication_source(
    package: &SourceBackedPackage,
    worksheet_part_name: &PackURI,
) -> Result<()> {
    package.validate_topology_source_boundary()?;
    let expected = worksheet_part_name
        .as_str()
        .strip_prefix('/')
        .ok_or_else(|| invalid("worksheet Part name is not package-absolute"))?;
    let mut exact_member = false;
    for name in package.physical_member_names() {
        if name.eq_ignore_ascii_case(expected) {
            if name != expected || exact_member {
                return Err(Error::Package(
                    litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                        reason: format!(
                            "hyperlink publication refuses case-equivalent or duplicate physical member aliases for {expected}"
                        ),
                    },
                ));
            }
            exact_member = true;
        }
    }
    if !exact_member {
        return Err(Error::Package(
            litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: format!(
                    "hyperlink publication requires the exact physical member {expected}"
                ),
            },
        ));
    }
    Ok(())
}
