//! Immutable, source-preserving smart-tag property-bag state.

use super::model::{Error, Limits, Property, PropertyBag, PropertyBagStore};
use super::transaction::{Revision, Transaction};
use super::validation;
use litchi_codepage::Ansi;
use std::sync::Arc;

#[derive(Debug, Clone, PartialEq, Eq)]
struct State {
    bytes: Arc<[u8]>,
    store: PropertyBagStore,
    bags: Vec<PropertyBag>,
    limits: Limits,
    revision: Revision,
}

/// An immutable, cheaply clonable view of one complete [MS-OSHARED]
/// `PropertyBagStore` followed by its property bags.
///
/// The exact source allocation is retained, while the store and every bag are
/// projected into the shared typed model. No smart-tag is recognized and no
/// URI, download URL, or property value is resolved or loaded.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    state: Arc<State>,
}

impl Snapshot {
    /// Parses a complete payload whose property bags extend to the end.
    pub fn parse(data: &[u8], ansi: Ansi, limits: Limits) -> Result<Self, Error> {
        Self::parse_to_end_shared(Arc::<[u8]>::from(data), ansi, limits)
    }

    /// Parses a complete payload with an exact property-bag count.
    pub fn parse_bags(
        data: &[u8],
        count: usize,
        ansi: Ansi,
        limits: Limits,
    ) -> Result<Self, Error> {
        Self::parse_bags_shared(Arc::<[u8]>::from(data), count, ansi, limits)
    }

    /// Alias for [`Self::parse_bags`] that makes the trailing-byte policy
    /// explicit at call sites.
    pub fn parse_exact(
        data: &[u8],
        count: usize,
        ansi: Ansi,
        limits: Limits,
    ) -> Result<Self, Error> {
        Self::parse_bags(data, count, ansi, limits)
    }

    /// Parses a complete payload whose property bags extend to the end,
    /// sharing the caller's source allocation.
    pub fn parse_to_end_shared(
        bytes: Arc<[u8]>,
        ansi: Ansi,
        limits: Limits,
    ) -> Result<Self, Error> {
        Self::parse_shared(bytes, None, ansi, limits)
    }

    /// Parses a complete payload with an exact bag count, sharing its source
    /// allocation and rejecting any trailing bytes.
    pub fn parse_bags_shared(
        bytes: Arc<[u8]>,
        count: usize,
        ansi: Ansi,
        limits: Limits,
    ) -> Result<Self, Error> {
        Self::parse_shared(bytes, Some(count), ansi, limits)
    }

    /// Creates a canonical source-preserving snapshot from the complete typed
    /// model. Parsed snapshots should be preferred when original bytes matter.
    pub fn from_store(
        store: PropertyBagStore,
        bags: Vec<PropertyBag>,
        limits: Limits,
    ) -> Result<Self, Error> {
        let bytes = validation::encode(&store, &bags, limits)?;
        Ok(Self::from_parts(bytes.into(), store, bags, limits))
    }

    /// Borrows the shared type and string tables.
    #[must_use]
    pub fn store(&self) -> &PropertyBagStore {
        &self.state.store
    }

    /// Borrows every property bag in source order.
    #[must_use]
    pub fn bags(&self) -> &[PropertyBag] {
        &self.state.bags
    }

    /// Returns one property bag by source order.
    #[must_use]
    pub fn bag(&self, index: usize) -> Option<&PropertyBag> {
        self.state.bags.get(index)
    }

    /// Returns one raw key/value index pair by source order.
    #[must_use]
    pub fn property(&self, bag: usize, property: usize) -> Option<Property> {
        self.state.bags.get(bag)?.properties.get(property).copied()
    }

    /// Resolves one property pair through the retained shared string table.
    #[must_use]
    pub fn resolved_property(&self, bag: usize, property: usize) -> Option<(&str, &str)> {
        self.store().resolve_property(self.property(bag, property)?)
    }

    /// Returns the exact serialized payload retained by this snapshot.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.state.bytes
    }

    /// Returns shared ownership of the exact serialized payload.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.state.bytes)
    }

    /// Returns the source fingerprint used by source-checked patches.
    #[must_use]
    pub fn revision(&self) -> Revision {
        self.state.revision
    }

    /// Returns the compact source fingerprint.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        self.state.revision.value()
    }

    /// Returns the limits used to validate this snapshot and its edits.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.state.limits
    }

    /// Starts an isolated typed edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Alias for [`Self::edit`] for callers that prefer transactional naming.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        self.edit()
    }

    /// Consumes the snapshot into its complete typed projection.
    #[must_use]
    pub fn into_parts(self) -> (PropertyBagStore, Vec<PropertyBag>) {
        match Arc::try_unwrap(self.state) {
            Ok(state) => (state.store, state.bags),
            Err(state) => (state.store.clone(), state.bags.clone()),
        }
    }

    pub(super) fn from_parts(
        bytes: Arc<[u8]>,
        store: PropertyBagStore,
        bags: Vec<PropertyBag>,
        limits: Limits,
    ) -> Self {
        let revision = Revision::of(&bytes);
        Self {
            state: Arc::new(State {
                bytes,
                store,
                bags,
                limits,
                revision,
            }),
        }
    }

    fn parse_shared(
        bytes: Arc<[u8]>,
        count: Option<usize>,
        ansi: Ansi,
        limits: Limits,
    ) -> Result<Self, Error> {
        if bytes.len() > limits.max_bytes {
            return Err(Error::new(
                "smart-tag serialized payload exceeds the configured limit",
            ));
        }
        let (store, consumed) = PropertyBagStore::parse_prefix(&bytes, ansi, limits)?;
        let remainder = bytes
            .get(consumed..)
            .ok_or_else(|| Error::new("smart-tag store offset is outside its source"))?;
        let bags = match count {
            Some(count) => store.parse_bags(remainder, count, limits)?,
            None => store.parse_bags_to_end(remainder, limits)?,
        };
        Ok(Self::from_parts(bytes, store, bags, limits))
    }
}
