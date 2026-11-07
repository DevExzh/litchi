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
/// publishes a new package. Starting one retains only the source package;
/// the graph and semantic draft are materialized by the first getter,
/// mutation or commit.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    loaded: Option<Loaded>,
}

#[derive(Debug, Clone)]
struct Loaded {
    before: Snapshot,
    draft: Vec<super::package::Entry>,
    next_relationship_id: String,
    next_part_uri: litchi_opc::PackURI,
    first_part_uri: litchi_opc::PackURI,
}

impl Edit {
    /// Start an edit retaining only the source package. Graph validation and
    /// source capture are deferred to the first getter, mutation or commit.
    fn start(target: crate::record::Source) -> Result<Self> {
        Ok(Self {
            target,
            loaded: None,
        })
    }

    fn ensure_loaded(&mut self) -> Result<&mut Loaded> {
        if self.loaded.is_none() {
            let limits = self.target.read_limits();
            let before = Snapshot::load(&self.target, &limits)?;
            let next_part_uri = next_part_uri(&self.target)?;
            self.loaded = Some(Loaded {
                draft: before.entries().to_vec(),
                next_relationship_id: before.next_relationship_id().to_owned(),
                first_part_uri: next_part_uri.clone(),
                next_part_uri,
                before,
            });
        }
        Ok(self.loaded.as_mut().expect("external-link graph loaded"))
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    pub fn before(&mut self) -> Result<&Snapshot> {
        Ok(&self.ensure_loaded()?.before)
    }

    /// Borrow the currently staged external-link entries.
    pub fn entries(&mut self) -> Result<&[super::package::Entry]> {
        Ok(&self.ensure_loaded()?.draft)
    }

    /// Apply a checked mutation to one existing inert link.
    pub fn edit(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut Link) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self.ensure_loaded()?.draft.clone();
        let entry = draft
            .get_mut(index)
            .ok_or_else(|| invalid(format!("external-link index {index} is absent")))?;
        edit(&mut entry.link)?;
        let limits = self.target.read_limits();
        let conformance = self
            .loaded
            .as_ref()
            .expect("external-link graph loaded")
            .before
            .conformance();
        validation::link(&entry.link, conformance, &limits)?;
        if draft
            == self
                .loaded
                .as_ref()
                .expect("external-link graph loaded")
                .draft
        {
            return Ok(false);
        }
        self.loaded
            .as_mut()
            .expect("external-link graph loaded")
            .draft = draft;
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

    /// Append a new link with an allocated workbook relationship and part URI.
    /// Commit registers its ordered workbook externalReference as well.
    pub fn insert(&mut self, link: Link) -> Result<usize> {
        self.ensure_loaded()?;
        let limits = self.target.read_limits();
        let loaded = self.loaded.as_ref().expect("external-link graph loaded");
        validation::link(&link, loaded.before.conformance(), &limits)?;
        let mut relationship_id = loaded.next_relationship_id.clone();
        while loaded
            .draft
            .iter()
            .any(|entry| entry.relationship_id == relationship_id)
        {
            relationship_id = next_relationship_id_after(&relationship_id);
        }
        let next_relationship_id = next_relationship_id_after(&relationship_id);
        let part_uri = loaded.next_part_uri.clone();
        let next_part_uri = next_part_uri_after(&part_uri, &self.target)?;
        let index = loaded.draft.len();
        let wire_index =
            u32::try_from(index).map_err(|_source| invalid("external-link index exceeds u32"))?;
        let mut draft = loaded.draft.clone();
        draft.push(super::package::Entry {
            index: wire_index,
            relationship_id,
            part_uri,
            link,
        });
        let limits = self.target.read_limits();
        validation::entries(&draft, loaded.before.conformance(), &limits)?;
        let loaded = self.loaded.as_mut().expect("external-link graph loaded");
        loaded.next_relationship_id = next_relationship_id;
        loaded.next_part_uri = next_part_uri;
        loaded.draft = draft;
        Ok(index)
    }

    /// Remove one link from the staged catalog. Commit refuses removal or
    /// renumbering when XML formula consumers may depend on an affected slot;
    /// this owner does not rewrite formula references.
    pub fn remove(&mut self, index: usize) -> Result<Option<super::package::Entry>> {
        self.ensure_loaded()?;
        let mut draft = self
            .loaded
            .as_ref()
            .expect("external-link graph loaded")
            .draft
            .clone();
        if index >= draft.len() {
            return Ok(None);
        }
        let removed = draft.remove(index);
        for (index, entry) in draft.iter_mut().enumerate() {
            entry.index = u32::try_from(index)
                .map_err(|_source| invalid("external-link index exceeds u32"))?;
        }
        let limits = self.target.read_limits();
        let loaded = self.loaded.as_ref().expect("external-link graph loaded");
        validation::entries(&draft, loaded.before.conformance(), &limits)?;
        self.loaded
            .as_mut()
            .expect("external-link graph loaded")
            .draft = draft;
        Ok(Some(removed))
    }

    fn changed(loaded: &Loaded) -> bool {
        loaded.before.entries() != loaded.draft.as_slice()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        let loaded = self.ensure_loaded()?;
        loaded.draft = loaded.before.entries().to_vec();
        loaded
            .before
            .next_relationship_id()
            .clone_into(&mut loaded.next_relationship_id);
        loaded.next_part_uri.clone_from(&loaded.first_part_uri);
        Ok(())
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
    pub(super) fn publish(mut self) -> Result<(Commit, Snapshot)> {
        let unchanged = {
            let loaded = self.ensure_loaded()?;
            !Self::changed(loaded)
        };
        if unchanged {
            let loaded = self.loaded.take().expect("external-link graph loaded");
            let patch = Patch::new(
                loaded.before.clone(),
                loaded.before.clone(),
                crate::part_edit::PartPatch::default(),
            );
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                loaded.before,
            ));
        }

        let loaded = self.loaded.take().expect("external-link graph loaded");
        let Loaded { before, draft, .. } = loaded;

        let mut candidate = OpcPackage::clone(&self.target);
        let limits = candidate.read_limits();
        let current = Snapshot::load(&candidate, &limits)?;
        if !current.same_source(&before) {
            return Err(crate::error::Error::PatchConflict {
                part: before.workbook_part_name().to_owned(),
            });
        }
        apply_entries(
            &mut candidate,
            before.entries(),
            &draft,
            before.conformance(),
        )?;
        let snapshot = Snapshot::load(&candidate, &limits)?;
        if snapshot.entries() != draft.as_slice() {
            return Err(invalid(
                "external-link publication changed the staged graph",
            ));
        }
        let names = std::iter::once(before.workbook_part_name().to_owned())
            .chain(
                before
                    .entries()
                    .iter()
                    .chain(snapshot.entries())
                    .map(|e| e.part_uri.to_string()),
            )
            .collect();
        let image = crate::part_edit::PartPatch::capture(&self.target, &candidate, &names)?;
        let patch = Patch::new(before, snapshot.clone(), image);
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
    /// refused. External-link graph, XML and source-limit errors are deferred
    /// to the first getter, mutation or commit.
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
