//! Failure-atomic typed Custom XML store transactions.

use std::sync::Arc;

use super::codec;
use super::model::{Item, Promotion, Properties, Result, Store, invalid};
use super::patch::Patch;
use super::snapshot::Snapshot;

/// An isolated typed edit over one immutable Custom XML store snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    draft: Store,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            draft: source.store().clone(),
            source,
        }
    }

    /// Borrow the immutable source used to start this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrow the currently staged typed store.
    #[must_use]
    pub fn store(&self) -> &Store {
        &self.draft
    }

    /// Borrow the currently staged item projection.
    #[must_use]
    pub fn items(&self) -> &[Item] {
        self.draft.items()
    }

    /// Return the staged promotion state.
    #[must_use]
    pub fn promotion(&self) -> Promotion {
        self.draft.promotion
    }

    /// Whether the staged typed or opaque source state differs from its base.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.draft != *self.source.store()
    }

    /// Change the store's promotion marker state.
    ///
    /// # Errors
    ///
    /// Returns an error if the resulting store violates its source limits.
    pub fn set_promotion(&mut self, promotion: Promotion) -> Result<&mut Self> {
        self.stage(|store| {
            store.promotion = promotion;
            Ok(())
        })?;
        Ok(self)
    }

    /// Replace one inert Item XML payload after bounded validation.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent, `xml` is malformed, or the
    /// resulting store violates its source limits.
    pub fn set_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        self.stage(|store| {
            let item = item_mut(store, index)?;
            item.set_xml(xml)
        })
    }

    /// Contextual alias for [`Self::set_xml`].
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::set_xml`].
    pub fn set_item_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        self.set_xml(index, xml)
    }

    /// Replace one Custom XML sub-storage name without touching either stream.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent, `storage_name` is invalid, or
    /// the resulting store violates its source limits.
    pub fn set_storage_name(&mut self, index: usize, storage_name: String) -> Result<bool> {
        self.stage(|store| {
            codec::validate_storage_name(&storage_name)?;
            let item = item_mut(store, index)?;
            if item.storage_name == storage_name {
                return Ok(());
            }
            item.storage_name = storage_name;
            Ok(())
        })
    }

    /// Replace one typed Properties projection while retaining source XML
    /// around known attributes whenever the schema-reference topology allows
    /// a lossless splice.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent, the retained XML is stale, or
    /// the resulting properties violate the source limits.
    pub fn set_properties(&mut self, index: usize, properties: Properties) -> Result<bool> {
        let limits = self.source.limits();
        self.stage(|store| {
            let item = item_mut(store, index)?;
            if item.properties == properties {
                return Ok(());
            }
            let properties_xml = codec::rewrite_properties(
                item.properties_xml(),
                &item.properties,
                &properties,
                &limits,
            )?;
            item.properties = properties;
            item.properties_xml = Arc::from(properties_xml);
            Ok(())
        })
    }

    /// Replace only an item's typed identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent or the resulting store violates
    /// its source limits.
    pub fn set_item_id(&mut self, index: usize, item_id: super::model::ItemId) -> Result<bool> {
        let mut properties = self
            .items()
            .get(index)
            .ok_or_else(|| missing_item(index))?
            .properties()
            .clone();
        properties.item_id = item_id;
        self.set_properties(index, properties)
    }

    /// Replace only an item's schema-reference list.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent or the resulting properties
    /// violate the source limits.
    pub fn set_schema_references<I, S>(&mut self, index: usize, references: I) -> Result<bool>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let mut properties = self
            .items()
            .get(index)
            .ok_or_else(|| missing_item(index))?
            .properties()
            .clone();
        properties.schema_references = references.into_iter().map(Into::into).collect();
        self.set_properties(index, properties)
    }

    /// Replace one Properties stream with caller-provided XML. The complete
    /// source bytes are retained, while the typed projection is parsed and
    /// validated before the candidate is published.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent, `xml` is malformed, or the
    /// resulting store violates its source limits.
    pub fn set_properties_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        let limits = self.source.limits();
        self.stage(|store| {
            let properties = codec::parse_properties_with_limits(&xml, &limits)?;
            let item = item_mut(store, index)?;
            if item.properties_xml() == xml.as_slice() {
                return Ok(());
            }
            item.properties = properties;
            item.properties_xml = Arc::from(xml);
            Ok(())
        })
    }

    /// Insert one already validated item at the end of the source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the resulting store violates its source limits.
    pub fn insert(&mut self, item: Item) -> Result<usize> {
        let index = self.draft.items().len();
        self.stage(|store| {
            store.items.push(item);
            Ok(())
        })?;
        Ok(index)
    }

    /// Contextual alias for [`Self::insert`].
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::insert`].
    pub fn add(&mut self, item: Item) -> Result<usize> {
        self.insert(item)
    }

    /// Remove one item and return its source-preserving value.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent or the resulting store violates
    /// its source limits.
    pub fn remove(&mut self, index: usize) -> Result<Item> {
        let mut removed = None;
        self.stage(|store| {
            if index >= store.items.len() {
                return Err(missing_item(index));
            }
            removed = Some(store.items.remove(index));
            Ok(())
        })?;
        removed.ok_or_else(|| invalid("custom XML item removal did not produce an item"))
    }

    /// Replace one item and return the old source-preserving value.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent or the resulting store violates
    /// its source limits.
    pub fn replace(&mut self, index: usize, item: Item) -> Result<Item> {
        let mut replaced = None;
        self.stage(|store| {
            let current = store
                .items
                .get_mut(index)
                .ok_or_else(|| missing_item(index))?;
            replaced = Some(std::mem::replace(current, item));
            Ok(())
        })?;
        replaced.ok_or_else(|| invalid("custom XML item replacement did not produce an item"))
    }

    /// Move one item to a checked source-order position.
    ///
    /// # Errors
    ///
    /// Returns an error if either index is absent or the resulting store
    /// violates its source limits.
    pub fn move_item(&mut self, from: usize, to: usize) -> Result<bool> {
        self.stage(|store| {
            if from >= store.items.len() || to >= store.items.len() {
                return Err(invalid(format!(
                    "custom XML item move index {from}->{to} is outside the store"
                )));
            }
            let item = store.items.remove(from);
            store.items.insert(to, item);
            Ok(())
        })
    }

    /// Apply a custom store edit to a cloned candidate and validate the whole
    /// dependency closure before publishing it into this transaction.
    ///
    /// # Errors
    ///
    /// Returns an error from `edit` or if the resulting store violates its
    /// source limits.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self>
    where
        F: FnOnce(&mut Store) -> Result<()>,
    {
        self.stage(edit)?;
        Ok(self)
    }

    /// Capture the current validated candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the staged store violates its source limits.
    pub fn snapshot(&self) -> Result<Snapshot> {
        self.materialize()
    }

    /// Discard this transaction and recover its immutable source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validate and publish a reversible source-checked commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the staged store violates its source limits.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = self.materialize()?;
        let patch = Patch::new(self.source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }

    fn stage<F>(&mut self, edit: F) -> Result<bool>
    where
        F: FnOnce(&mut Store) -> Result<()>,
    {
        let mut candidate = self.draft.clone();
        edit(&mut candidate)?;
        codec::validate_store(&candidate, &self.source.limits())?;
        let changed = candidate != self.draft;
        if changed {
            self.draft = candidate;
        }
        Ok(changed)
    }

    fn materialize(&self) -> Result<Snapshot> {
        codec::validate_store(&self.draft, &self.source.limits())?;
        if self.draft == *self.source.store() {
            return Ok(self.source.clone());
        }
        Ok(Snapshot::from_validated(
            self.draft.clone(),
            self.source.limits(),
        ))
    }
}

/// The successful result of a typed Custom XML store publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether the exact retained store source changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrow the immutable published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the commit into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Split the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Run one isolated edit and publish it atomically.
///
/// # Errors
///
/// Returns an error from `edit` or if the resulting store violates its source
/// limits.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit>
where
    F: FnOnce(&mut Transaction) -> Result<()>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}

fn item_mut(store: &mut Store, index: usize) -> Result<&mut Item> {
    store
        .items
        .get_mut(index)
        .ok_or_else(|| missing_item(index))
}

fn missing_item(index: usize) -> super::model::Error {
    invalid(format!("custom XML item index {index} is absent"))
}
