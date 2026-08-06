//! Immutable, source-preserving Custom XML store snapshots.

use std::sync::Arc;

use litchi_cfb::OleFile;

use super::codec;
use super::model::{Limits, Result, Store};
use super::transaction::Transaction;

/// A deterministic identity for one validated Custom XML store source.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(super) fn of(store: &Store) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        feed(&mut value, &[store.promotion as u8]);
        for item in store.items() {
            feed(&mut value, item.storage_name().as_bytes());
            feed(&mut value, &(item.xml().len() as u64).to_le_bytes());
            feed(&mut value, item.xml());
            feed(
                &mut value,
                &(item.properties_xml().len() as u64).to_le_bytes(),
            );
            feed(&mut value, item.properties_xml());
        }
        Self(value)
    }

    /// Returns the compact source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

fn feed(value: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        *value ^= u64::from(*byte);
        *value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
}

/// An immutable, cheaply clonable view of one complete MS-OSHARED Custom XML
/// data store. Item and Properties stream allocations are shared across
/// snapshots; edits copy only the changed stream and bounded store metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    store: Arc<Store>,
    limits: Limits,
    revision: Revision,
}

impl Snapshot {
    /// Validates a store with the default resource profile.
    pub fn from_store(store: Store) -> Result<Self> {
        Self::from_store_with_limits(store, Limits::default())
    }

    /// Validates a store with caller-selected bounded resource limits.
    pub fn from_store_with_limits(store: Store, limits: Limits) -> Result<Self> {
        codec::validate_store(&store, &limits)?;
        Ok(Self::from_validated(store, limits))
    }

    /// Inspects a complete OLE file and captures its Custom XML store.
    pub fn load<R: std::io::Read + std::io::Seek>(ole: &mut OleFile<R>) -> Result<Option<Self>> {
        Self::load_with_limits(ole, Limits::default())
    }

    /// Inspects a complete OLE file with caller-selected resource limits.
    pub fn load_with_limits<R: std::io::Read + std::io::Seek>(
        ole: &mut OleFile<R>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Ok(codec::inspect_with_limits(ole, limits)?
            .map(|store| Self::from_validated(store, limits)))
    }

    /// Alias for [`Self::load`] that emphasizes source-bound parsing.
    pub fn read<R: std::io::Read + std::io::Seek>(ole: &mut OleFile<R>) -> Result<Option<Self>> {
        Self::load(ole)
    }

    /// Borrow the complete typed store projection.
    #[must_use]
    pub fn store(&self) -> &Store {
        &self.store
    }

    /// Borrow items in the source storage order captured by the codec.
    #[must_use]
    pub fn items(&self) -> &[super::model::Item] {
        self.store.items()
    }

    /// Borrow one item by checked source position.
    #[must_use]
    pub fn item(&self, index: usize) -> Option<&super::model::Item> {
        self.items().get(index)
    }

    /// Return the promotion state of this store.
    #[must_use]
    pub fn promotion(&self) -> super::model::Promotion {
        self.store.promotion
    }

    /// Whether the store has no Custom XML items.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.items().is_empty()
    }

    /// Return the limits used to validate this snapshot and its edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the exact source fingerprint used by patches.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Alias for [`Self::revision`].
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Start an isolated bounded edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Alias for [`Self::edit`] for transaction-oriented callers.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        self.edit()
    }

    /// Consume the snapshot into its typed store projection.
    #[must_use]
    pub fn into_store(self) -> Store {
        match Arc::try_unwrap(self.store) {
            Ok(store) => store,
            Err(store) => (*store).clone(),
        }
    }

    pub(super) fn from_validated(store: Store, limits: Limits) -> Self {
        let revision = Revision::of(&store);
        Self {
            store: Arc::new(store),
            limits,
            revision,
        }
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.revision == other.revision && self.store == other.store
    }
}
