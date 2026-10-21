//! Failure-atomic, source-bound edits for workbook external links.

use litchi_opc::OpcPackage;

use super::model::{AlternateUrls, Link};
use super::package::apply_entries;
use super::patch::{Commit, Diagnostics, Patch};
use super::snapshot::{Snapshot, next_part_uri};
use super::validation;
use crate::error::{Result, invalid};

/// A staged, source-bound edit of the workbook's external-link catalog of
/// one [`crate::Package`]. Dropping it discards it; [`Self::commit`]
/// publishes a new package.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    before: Snapshot,
    draft: Vec<super::package::Entry>,
    next_relationship_id: String,
    next_part_uri: litchi_opc::PackURI,
    first_part_uri: litchi_opc::PackURI,
}

impl Edit {
    /// Start an edit after validating and capturing the workbook graph.
    fn start(target: crate::record::Source) -> Result<Self> {
        let before = Snapshot::load(&target)?;
        let next_part_uri = next_part_uri(&target)?;
        Ok(Self {
            target,
            draft: before.entries().to_vec(),
            next_relationship_id: before.next_relationship_id().to_owned(),
            first_part_uri: next_part_uri.clone(),
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

    /// Replace or remove one workbook link's inert `alternateUrls` extension.
    ///
    /// DDE and OLE links do not carry this SpreadsheetML extension and are
    /// rejected before the staged catalog changes.
    pub fn set_alternate_urls(
        &mut self,
        index: usize,
        alternate_urls: Option<AlternateUrls>,
    ) -> Result<bool> {
        self.edit(index, |value| {
            let Link::Workbook(workbook) = value else {
                return Err(invalid(
                    "alternateUrls can only be edited on an external workbook link",
                ));
            };
            workbook.alternate_urls = alternate_urls;
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
        self.next_part_uri = next_part_uri_after(&part_uri, &self.target)?;
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

    fn changed(&self) -> bool {
        self.before.entries() != self.draft.as_slice()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.draft = self.before.entries().to_vec();
        self.before
            .next_relationship_id()
            .clone_into(&mut self.next_relationship_id);
        self.next_part_uri.clone_from(&self.first_part_uri);
    }

    /// Validate and publish the staged graph as a new package. An unchanged
    /// edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when the source changed and
    /// the publication's validation errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    pub(super) fn publish(self) -> Result<(Commit, Snapshot)> {
        if !self.changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                self.before,
            ));
        }

        let mut candidate = OpcPackage::clone(&self.target);
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
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
    }
}

impl crate::Package {
    /// Start an edit of the workbook's external-link catalog.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, and the validation errors of the external-link graph.
    pub fn edit_external_links(&self) -> Result<Edit> {
        Edit::start(crate::record::Source::new(self, "edit_external_links")?)
    }

    /// Apply an exact external-link patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_external_links(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_external_links")
    }
}

crate::record::raw_test_route!(edit_external_links);

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
        match package.get_part(&uri) {
            Ok(_) => {},
            Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(uri),
            Err(error) => return Err(error.into()),
        }
        candidate = candidate
            .checked_add(1)
            .ok_or_else(|| invalid("external-link part name space exhausted"))?;
    }
}
