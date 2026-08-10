//! Failure-atomic, source-bound edits for workbook external links.

use litchi_opc::OpcPackage;

use super::model::Link;
use super::package::apply_entries;
use super::patch::{Commit, Patch};
use super::snapshot::{Snapshot, next_part_uri};
use super::validation;
use crate::error::{Result, invalid};

/// A typed transaction over the workbook's external-link catalog.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Vec<super::package::Entry>,
    next_relationship_id: String,
    next_part_uri: litchi_opc::PackURI,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating and capturing the workbook graph.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        let next_part_uri = next_part_uri(target)?;
        Ok(Self {
            target,
            draft: before.entries().to_vec(),
            next_relationship_id: before.next_relationship_id().to_owned(),
            next_part_uri,
            before,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged external-link entries.
    #[must_use]
    pub fn entries(&self) -> &[super::package::Entry] {
        &self.draft
    }

    /// Contextual alias for [`Self::entries`].
    #[must_use]
    pub fn links(&self) -> &[super::package::Entry] {
        self.entries()
    }

    /// Apply a checked mutation to one existing inert link.
    pub fn edit(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut Link) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self.draft.clone();
        let entry = draft
            .get_mut(index)
            .ok_or_else(|| invalid(format!("external-link index {index} is absent")))?;
        edit(&mut entry.link)?;
        validation::link(&entry.link, self.before.conformance())?;
        if draft == self.draft {
            return Ok(false);
        }
        self.draft = draft;
        Ok(true)
    }

    /// Replace one existing link while retaining its workbook relationship
    /// and external-link part identity.
    pub fn set(&mut self, index: usize, link: Link) -> Result<bool> {
        self.edit(index, |value| {
            *value = link;
            Ok(())
        })
    }

    /// Insert a new link with an allocated workbook relationship and part URI.
    pub fn insert(&mut self, link: Link) -> Result<usize> {
        validation::link(&link, self.before.conformance())?;
        let mut relationship_id = self.next_relationship_id.clone();
        while self
            .draft
            .iter()
            .any(|entry| entry.relationship_id == relationship_id)
        {
            relationship_id = next_relationship_id_after(&relationship_id);
        }
        self.next_relationship_id = next_relationship_id_after(&relationship_id);
        let part_uri = self.next_part_uri.clone();
        self.next_part_uri = next_part_uri_after(&part_uri, self.target)?;
        let index = self.draft.len();
        let wire_index =
            u32::try_from(index).map_err(|_source| invalid("external-link index exceeds u32"))?;
        self.draft.push(super::package::Entry {
            index: wire_index,
            relationship_id,
            part_uri,
            link,
        });
        validation::entries(&self.draft, self.before.conformance())?;
        Ok(index)
    }

    /// Remove one link from the staged catalog.
    pub fn remove(&mut self, index: usize) -> Result<Option<super::package::Entry>> {
        if index >= self.draft.len() {
            return Ok(None);
        }
        let removed = self.draft.remove(index);
        for (index, entry) in self.draft.iter_mut().enumerate() {
            entry.index = u32::try_from(index)
                .map_err(|_source| invalid("external-link index exceeds u32"))?;
        }
        validation::entries(&self.draft, self.before.conformance())?;
        Ok(Some(removed))
    }

    /// Whether staged entries differ from the captured source graph.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.entries() != self.draft.as_slice()
    }

    /// Validate and atomically publish the staged graph.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }

        let mut candidate = self.target.clone();
        let current = Snapshot::load(&candidate)?;
        if !current.same_source(&self.before) {
            return Err(crate::error::Error::PatchConflict {
                part: self.before.workbook_part_name().to_owned(),
            });
        }
        apply_entries(
            &mut candidate,
            self.before.entries(),
            &self.draft,
            self.before.conformance(),
        )?;
        let snapshot = Snapshot::load(&candidate)?;
        if snapshot.entries() != self.draft.as_slice() {
            return Err(invalid(
                "external-link publication changed the staged graph",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

fn next_relationship_id_after(current: &str) -> String {
    let number = current
        .strip_prefix("rId")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(0)
        .saturating_add(1);
    format!("rId{number}")
}

fn next_part_uri_after(
    current: &litchi_opc::PackURI,
    package: &OpcPackage,
) -> Result<litchi_opc::PackURI> {
    let current_name = current.as_str();
    let number = current_name
        .strip_prefix("/xl/externalLinks/externalLink")
        .and_then(|value| value.strip_suffix(".xml"))
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(0);
    let mut candidate = number.saturating_add(1);
    loop {
        let uri =
            litchi_opc::PackURI::new(format!("/xl/externalLinks/externalLink{candidate}.xml"))
                .map_err(invalid)?;
        if package.get_part(&uri).is_err() {
            return Ok(uri);
        }
        candidate = candidate
            .checked_add(1)
            .ok_or_else(|| invalid("external-link part name space exhausted"))?;
    }
}
