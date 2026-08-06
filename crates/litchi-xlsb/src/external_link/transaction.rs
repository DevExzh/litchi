//! Source-checked, inert edits for one XLSB external-link stream.

use std::sync::Arc;

use super::codec;
use super::model::{
    DdeItem, DefinedName, Entries, Kind, Link, OleItem, UnknownRecord, ValueMatrix,
};
use super::{Error, Result, package, validation};

/// An immutable external-link snapshot bound to the exact source stream.
#[derive(Debug, Clone)]
pub struct Snapshot {
    link: Link,
    source: Arc<SourceState>,
}

impl Snapshot {
    /// Parse and validate one complete inert External Link part stream.
    pub fn read(data: &[u8]) -> Result<Self> {
        let source = codec::parse_source(data)?;
        validation::validate_relationship(source.parsed.link(), source.parsed.relationship_id())?;
        Ok(Self::from_source(source))
    }

    fn from_source(source: codec::Source) -> Self {
        let link = source.parsed.clone().into_link();
        Self {
            link,
            source: Arc::new(SourceState {
                bytes: source.bytes,
                relationship_id: source.parsed.relationship_id().map(str::to_owned),
                unknown_records: source.unknown_records,
            }),
        }
    }

    /// Borrow the unresolved semantic link metadata.
    ///
    /// For workbook and OLE links, `source()` is the relationship identifier
    /// stored in `BrtBeginSupBook`; it is never followed by this API.
    #[must_use]
    pub const fn link(&self) -> &Link {
        &self.link
    }

    /// Return the relationship identifier stored by the stream, if any.
    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.source.relationship_id.as_deref()
    }

    /// Resolve a workbook/OLE relationship in the caller's already-validated
    /// host context without contacting the target.
    pub fn resolved_link(&self, source: impl Into<String>) -> Result<Link> {
        if self.relationship_id().is_none() {
            return Err(Error::InvalidFormula(
                "DDE external links do not have a relationship source".to_string(),
            ));
        }
        let mut link = self.link.clone();
        link.source = source.into();
        validation::validate_link(&link)?;
        Ok(link)
    }

    /// Borrow opaque records in source order.
    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.source.unknown_records
    }

    /// Borrow the exact source bytes used by stale-source checks.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source.bytes
    }

    /// A stream snapshot is always source-bound.
    #[must_use]
    pub const fn is_source_bound(&self) -> bool {
        true
    }

    /// Start a detached, failure-atomic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            link: self.link.clone(),
        }
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source_bytes() == other.source_bytes()
    }
}

impl Eq for Snapshot {}

/// A detached transaction over typed link metadata and inert caches.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    link: Link,
}

impl Transaction {
    /// Borrow the currently staged link.
    #[must_use]
    pub const fn link(&self) -> &Link {
        &self.link
    }

    /// Replace the complete typed link graph.
    pub fn replace(&mut self, link: Link) -> Result<()> {
        self.apply_candidate(link)
    }

    /// Set the inert source string or unresolved relationship identifier.
    pub fn set_source(&mut self, source: impl Into<String>) -> Result<()> {
        let mut candidate = self.link.clone();
        candidate.source = source.into();
        self.apply_candidate(candidate)
    }

    /// Alias for editing the unresolved workbook/OLE relationship identifier.
    pub fn set_relationship_id(&mut self, relationship_id: impl Into<String>) -> Result<()> {
        self.set_source(relationship_id)
    }

    /// Replace the DDE topic without contacting the DDE server.
    pub fn set_dde_topic(&mut self, topic: impl Into<String>) -> Result<()> {
        let mut candidate = self.link.clone();
        if candidate.kind() != Kind::Dde {
            return Err(invalid("DDE topic can only be edited on a DDE link"));
        }
        candidate.detail = Some(topic.into());
        self.apply_candidate(candidate)
    }

    /// Replace the OLE ProgID without instantiating an OLE class.
    pub fn set_ole_program_id(&mut self, program_id: impl Into<String>) -> Result<()> {
        let mut candidate = self.link.clone();
        if candidate.kind() != Kind::Ole {
            return Err(invalid("OLE ProgID can only be edited on an OLE link"));
        }
        candidate.detail = Some(program_id.into());
        self.apply_candidate(candidate)
    }

    /// Replace external workbook sheet names and revalidate name scopes.
    pub fn set_sheet_names(&mut self, sheet_names: Vec<String>) -> Result<()> {
        let mut candidate = self.link.clone();
        if candidate.kind() != Kind::Workbook {
            return Err(invalid(
                "external workbook sheet names can only be edited on a workbook link",
            ));
        }
        candidate.sheet_names = sheet_names;
        self.apply_candidate(candidate)
    }

    /// Replace all external defined names.
    pub fn set_defined_names(&mut self, names: Vec<DefinedName>) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Workbook(entries) = &mut candidate.entries else {
            return Err(invalid(
                "external defined names can only be edited on a workbook link",
            ));
        };
        *entries = names;
        self.apply_candidate(candidate)
    }

    /// Add or replace one external defined name by case-insensitive name.
    pub fn upsert_defined_name(&mut self, name: DefinedName) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Workbook(entries) = &mut candidate.entries else {
            return Err(invalid(
                "external defined names can only be edited on a workbook link",
            ));
        };
        if let Some(index) = entries
            .iter()
            .position(|existing| same_name(existing.name(), name.name()))
        {
            entries[index] = name;
        } else {
            entries.push(name);
        }
        self.apply_candidate(candidate)
    }

    /// Remove one external defined name by case-insensitive name.
    pub fn remove_defined_name(&mut self, name: &str) -> Result<Option<DefinedName>> {
        let mut candidate = self.link.clone();
        let Entries::Workbook(entries) = &mut candidate.entries else {
            return Err(invalid(
                "external defined names can only be edited on a workbook link",
            ));
        };
        let Some(index) = entries
            .iter()
            .position(|entry| same_name(entry.name(), name))
        else {
            return Ok(None);
        };
        let removed = entries.remove(index);
        self.apply_candidate(candidate)?;
        Ok(Some(removed))
    }

    /// Add or replace one DDE item by case-insensitive item name.
    pub fn upsert_dde_item(&mut self, item: DdeItem) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Dde(entries) = &mut candidate.entries else {
            return Err(invalid("DDE items can only be edited on a DDE link"));
        };
        if let Some(index) = entries
            .iter()
            .position(|existing| same_name(existing.name(), item.name()))
        {
            entries[index] = item;
        } else {
            entries.push(item);
        }
        self.apply_candidate(candidate)
    }

    /// Remove one DDE item without contacting the DDE server.
    pub fn remove_dde_item(&mut self, name: &str) -> Result<Option<DdeItem>> {
        let mut candidate = self.link.clone();
        let Entries::Dde(entries) = &mut candidate.entries else {
            return Err(invalid("DDE items can only be edited on a DDE link"));
        };
        let Some(index) = entries
            .iter()
            .position(|entry| same_name(entry.name(), name))
        else {
            return Ok(None);
        };
        let removed = entries.remove(index);
        self.apply_candidate(candidate)?;
        Ok(Some(removed))
    }

    /// Edit DDE item flags while retaining its inert cached values.
    pub fn set_dde_flags(
        &mut self,
        name: &str,
        want_advise: bool,
        want_picture: bool,
        supports_ole: bool,
    ) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Dde(entries) = &mut candidate.entries else {
            return Err(invalid("DDE flags can only be edited on a DDE link"));
        };
        let item = entries
            .iter_mut()
            .find(|entry| same_name(entry.name(), name))
            .ok_or_else(|| invalid("DDE item was not found"))?;
        *item = item
            .clone()
            .with_advise(want_advise)
            .with_picture(want_picture)
            .with_ole_support(supports_ole);
        self.apply_candidate(candidate)
    }

    /// Replace or clear one DDE item's inert cached matrix.
    pub fn set_dde_cache(&mut self, name: &str, cache: Option<ValueMatrix>) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Dde(entries) = &mut candidate.entries else {
            return Err(invalid("DDE caches can only be edited on a DDE link"));
        };
        let item = entries
            .iter_mut()
            .find(|entry| same_name(entry.name(), name))
            .ok_or_else(|| invalid("DDE item was not found"))?;
        *item = match cache {
            Some(cache) => item.clone().with_cached_values(cache),
            None => item.clone().without_cached_values(),
        };
        self.apply_candidate(candidate)
    }

    /// Add or replace one OLE item by case-insensitive item name.
    pub fn upsert_ole_item(&mut self, item: OleItem) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Ole(entries) = &mut candidate.entries else {
            return Err(invalid("OLE items can only be edited on an OLE link"));
        };
        if let Some(index) = entries
            .iter()
            .position(|existing| same_name(existing.name(), item.name()))
        {
            entries[index] = item;
        } else {
            entries.push(item);
        }
        self.apply_candidate(candidate)
    }

    /// Remove one OLE item without instantiating an OLE class.
    pub fn remove_ole_item(&mut self, name: &str) -> Result<Option<OleItem>> {
        let mut candidate = self.link.clone();
        let Entries::Ole(entries) = &mut candidate.entries else {
            return Err(invalid("OLE items can only be edited on an OLE link"));
        };
        let Some(index) = entries
            .iter()
            .position(|entry| same_name(entry.name(), name))
        else {
            return Ok(None);
        };
        let removed = entries.remove(index);
        self.apply_candidate(candidate)?;
        Ok(Some(removed))
    }

    /// Edit OLE item flags while retaining its inert cached values.
    pub fn set_ole_flags(
        &mut self,
        name: &str,
        want_advise: bool,
        want_picture: bool,
        display_as_icon: bool,
    ) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Ole(entries) = &mut candidate.entries else {
            return Err(invalid("OLE flags can only be edited on an OLE link"));
        };
        let item = entries
            .iter_mut()
            .find(|entry| same_name(entry.name(), name))
            .ok_or_else(|| invalid("OLE item was not found"))?;
        *item = item
            .clone()
            .with_advise(want_advise)
            .with_picture(want_picture)
            .with_icon(display_as_icon);
        self.apply_candidate(candidate)
    }

    /// Replace or clear one OLE item's inert cached matrix.
    pub fn set_ole_cache(&mut self, name: &str, cache: Option<ValueMatrix>) -> Result<()> {
        let mut candidate = self.link.clone();
        let Entries::Ole(entries) = &mut candidate.entries else {
            return Err(invalid("OLE caches can only be edited on an OLE link"));
        };
        let item = entries
            .iter_mut()
            .find(|entry| same_name(entry.name(), name))
            .ok_or_else(|| invalid("OLE item was not found"))?;
        *item = match cache {
            Some(cache) => item.clone().with_cached_values(cache),
            None => item.clone().without_cached_values(),
        };
        self.apply_candidate(candidate)
    }

    /// Validate and publish an immutable snapshot plus reversible patch.
    pub fn commit(self) -> Result<Commit> {
        validation::validate_link(&self.link)?;
        if self.link == self.base.link {
            let source = Arc::clone(&self.base.source.bytes);
            return Ok(Commit {
                snapshot: self.base.clone(),
                patch: Patch {
                    before: Arc::clone(&source),
                    after: source,
                },
            });
        }

        let relationship_id = match self.link.kind() {
            Kind::Dde => None,
            Kind::Workbook | Kind::Ole => Some(self.link.source()),
        };
        let bytes = package::write_external_link_stream_with_unknown(
            &self.link,
            relationship_id,
            self.base.unknown_records(),
        )?;
        let after = Arc::<[u8]>::from(bytes);
        if after.as_ref() == self.base.source_bytes() {
            return Ok(Commit {
                snapshot: self.base.clone(),
                patch: Patch {
                    before: Arc::clone(&self.base.source.bytes),
                    after,
                },
            });
        }
        let snapshot = Snapshot::read(&after)?;
        Ok(Commit {
            snapshot,
            patch: Patch {
                before: Arc::clone(&self.base.source.bytes),
                after,
            },
        })
    }

    fn apply_candidate(&mut self, candidate: Link) -> Result<()> {
        validation::validate_link(&candidate)?;
        self.link = candidate;
        Ok(())
    }
}

/// A successful immutable transaction result.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the committed snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible patch guarded by the exact source stream bytes.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    /// Whether this patch leaves the stream byte-identical.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    /// Borrow the exact before-image.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Borrow the exact after-image.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Apply only to the exact source stream used to create this patch.
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before.as_ref() {
            return Err(invalid(
                "external-link patch source snapshot does not match",
            ));
        }
        Ok(self.after.to_vec())
    }

    /// Alias for transaction pipelines.
    pub fn commit(&self, source: &[u8]) -> Result<Vec<u8>> {
        self.apply(source)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Read one complete inert external-link stream.
pub fn read(data: &[u8]) -> Result<Snapshot> {
    Snapshot::read(data)
}

/// Apply a previously committed patch to one complete external-link stream.
pub fn apply(data: &[u8], patch: &Patch) -> Result<Vec<u8>> {
    patch.apply(data)
}

#[derive(Debug, Clone)]
struct SourceState {
    bytes: Arc<[u8]>,
    relationship_id: Option<String>,
    unknown_records: Vec<UnknownRecord>,
}

fn same_name(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormula(message.into())
}
