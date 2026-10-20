//! Failure-atomic typed smart-tag property-bag transactions.

use super::model::{Error, Property, PropertyBag, PropertyBagStore, PropertyBagString, Type};
use super::patch::Patch;
use super::snapshot::Snapshot;
use super::validation;

/// A deterministic identity for one exact serialized property-bag payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(super) fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Returns the raw source fingerprint.
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

/// An isolated edit over one source snapshot.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Snapshot,
    store: PropertyBagStore,
    bags: Vec<PropertyBag>,
}

impl Edit {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            store: source.store().clone(),
            bags: source.bags().to_vec(),
            source,
        }
    }

    /// Borrows the immutable source used to start this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrows the current shared type and string-table draft.
    #[must_use]
    pub const fn store(&self) -> &PropertyBagStore {
        &self.store
    }

    /// Borrows the current property-bag draft in source order.
    #[must_use]
    pub fn bags(&self) -> &[PropertyBag] {
        &self.bags
    }

    /// Whether the typed draft differs from its source projection.
    #[must_use]
    fn has_changes(&self) -> bool {
        self.store != *self.source.store() || self.bags != self.source.bags()
    }

    /// Replaces one shared string-table entry after full candidate validation.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent or the resulting payload violates
    /// the source limits.
    pub fn set_string(
        &mut self,
        index: usize,
        value: PropertyBagString,
    ) -> Result<&mut Self, Error> {
        self.update(move |store, _| {
            let entry = store
                .strings
                .get_mut(index)
                .ok_or_else(|| Error::new("smart-tag string index is outside the table"))?;
            *entry = value;
            Ok(())
        })
    }

    /// Replaces one declared type while retaining every other type and table
    /// entry. The replacement must retain the selected stable identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the type identifier is absent or the resulting
    /// payload violates the source limits.
    pub fn replace_type(&mut self, value: Type) -> Result<&mut Self, Error> {
        self.update(move |store, _| {
            let kind = store
                .types
                .iter_mut()
                .find(|kind| kind.id == value.id)
                .ok_or_else(|| Error::new("smart-tag type id is not declared"))?;
            *kind = value;
            Ok(())
        })
    }

    /// Replaces one complete property-bag while retaining the other bags.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is absent or the resulting payload violates
    /// the source limits.
    pub fn replace_bag(&mut self, index: usize, value: PropertyBag) -> Result<&mut Self, Error> {
        self.update(move |_, bags| {
            let bag = bags
                .get_mut(index)
                .ok_or_else(|| Error::new("smart-tag bag index is outside the store"))?;
            *bag = value;
            Ok(())
        })
    }

    /// Appends one property bag after validating its type and all indexes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bag or resulting payload violates the source
    /// limits.
    pub fn append_bag(&mut self, value: PropertyBag) -> Result<&mut Self, Error> {
        self.update(move |_, bags| {
            bags.push(value);
            Ok(())
        })
    }

    /// Replaces one raw key/value index pair in a property bag.
    ///
    /// # Errors
    ///
    /// Returns an error if either index is absent or the resulting payload
    /// violates the source limits.
    pub fn set_property(
        &mut self,
        bag: usize,
        property: usize,
        value: Property,
    ) -> Result<&mut Self, Error> {
        self.update(move |_, bags| {
            let entry = bags
                .get_mut(bag)
                .ok_or_else(|| Error::new("smart-tag bag index is outside the store"))?
                .properties
                .get_mut(property)
                .ok_or_else(|| Error::new("smart-tag property index is outside the bag"))?;
            *entry = value;
            Ok(())
        })
    }

    /// Replaces the key string referenced by one property.
    ///
    /// # Errors
    ///
    /// Returns an error if the bag, property, or referenced key string is
    /// absent, or the resulting payload violates the source limits.
    pub fn set_property_key(
        &mut self,
        bag: usize,
        property: usize,
        value: PropertyBagString,
    ) -> Result<&mut Self, Error> {
        self.update(move |store, bags| {
            let string_index = bags
                .get(bag)
                .ok_or_else(|| Error::new("smart-tag bag index is outside the store"))?
                .properties
                .get(property)
                .ok_or_else(|| Error::new("smart-tag property index is outside the bag"))?
                .key_index;
            let string_slot = usize::try_from(string_index).map_err(|_conversion_error| {
                Error::new("smart-tag key string index overflows usize")
            })?;
            let entry = store
                .strings
                .get_mut(string_slot)
                .ok_or_else(|| Error::new("smart-tag key string index is outside the table"))?;
            *entry = value;
            Ok(())
        })
    }

    /// Replaces the value string referenced by one property.
    ///
    /// # Errors
    ///
    /// Returns an error if the bag, property, or referenced value string is
    /// absent, or the resulting payload violates the source limits.
    pub fn set_property_value(
        &mut self,
        bag: usize,
        property: usize,
        value: PropertyBagString,
    ) -> Result<&mut Self, Error> {
        self.update(move |store, bags| {
            let string_index = bags
                .get(bag)
                .ok_or_else(|| Error::new("smart-tag bag index is outside the store"))?
                .properties
                .get(property)
                .ok_or_else(|| Error::new("smart-tag property index is outside the bag"))?
                .value_index;
            let string_slot = usize::try_from(string_index).map_err(|_conversion_error| {
                Error::new("smart-tag value string index overflows usize")
            })?;
            let entry = store
                .strings
                .get_mut(string_slot)
                .ok_or_else(|| Error::new("smart-tag value string index is outside the table"))?;
            *entry = value;
            Ok(())
        })
    }

    /// Changes the ignored `cfactoid` value without interpreting it.
    pub fn set_reserved_factoid_count(&mut self, value: u32) -> &mut Self {
        self.store.reserved_factoid_count = value;
        self
    }

    /// Applies a custom inert edit to cloned store and bag drafts.
    ///
    /// The candidate is published only after all count, index, encoding, and
    /// serialized-size constraints pass. A failed closure or validation
    /// leaves the transaction unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error from `edit` or if the resulting payload violates the
    /// source limits.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, Error>
    where
        F: FnOnce(&mut PropertyBagStore, &mut Vec<PropertyBag>) -> Result<(), Error>,
    {
        let mut store = self.store.clone();
        let mut bags = self.bags.clone();
        edit(&mut store, &mut bags)?;
        validation::encode(&store, &bags, self.source.limits())?;
        self.store = store;
        self.bags = bags;
        Ok(self)
    }

    /// Validates and publishes the draft as a reversible source-checked edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the draft cannot be encoded under the source
    /// limits.
    pub fn commit(self) -> Result<Commit, Error> {
        let snapshot = self.materialize()?;
        let patch = Patch::new(self.source, snapshot.clone());
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(snapshot, patch, diagnostics))
    }

    fn materialize(&self) -> Result<Snapshot, Error> {
        let bytes = validation::encode(&self.store, &self.bags, self.source.limits())?;
        if !self.has_changes() {
            return Ok(self.source.clone());
        }
        Ok(Snapshot::from_parts(
            bytes.into(),
            self.store.clone(),
            self.bags.clone(),
            self.source.limits(),
        ))
    }
}

/// A published edit: the new snapshot, its reversible patch and
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<Snapshot, Patch>;

/// Content-free diagnostics of one commit.
pub use litchi_core::edit::Diagnostics;

/// Runs one isolated edit and publishes it atomically.
///
/// # Errors
///
/// Returns an error from `edit` or if the resulting payload violates the
/// source limits.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, Error>
where
    F: FnOnce(&mut Edit) -> Result<(), Error>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}
