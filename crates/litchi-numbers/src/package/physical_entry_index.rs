//! Fallible name lookup for a physical package catalog.
//!
//! Focused package transactions repeatedly compare source and candidate
//! members by their normalized physical name.  Keeping those comparisons as
//! iterator scans makes locality verification quadratic in the number of
//! package entries.  This private index owns one bounded, source-order pass
//! and keeps the duplicate-name check explicit instead of allowing a map
//! insertion to silently select one ambiguous member.

use std::collections::{HashMap, hash_map::Entry as HashMapEntry};

use litchi_iwa_archive::package::{Catalog, Entry};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Error {
    /// The backing lookup table could not reserve its bounded entry count.
    Allocation { amount: usize },
    /// More than one physical member has the same normalized name.
    Duplicate,
}

/// One immutable physical package entry lookup table.
#[derive(Debug)]
pub(super) struct PhysicalEntryIndex<'catalog> {
    entries: HashMap<&'catalog str, &'catalog Entry>,
}

impl<'catalog> PhysicalEntryIndex<'catalog> {
    /// Build an index in one source-order pass.
    pub(super) fn new(catalog: &'catalog Catalog) -> Result<Self, Error> {
        let count = catalog.iter().count();
        let mut entries = HashMap::new();
        entries
            .try_reserve(count)
            .map_err(|_| Error::Allocation { amount: count })?;
        for entry in catalog.iter() {
            insert_unique(&mut entries, entry.name(), entry)?;
        }
        Ok(Self { entries })
    }

    /// Return the unique physical entry with this normalized name.
    #[must_use]
    pub(super) fn get(&self, name: &str) -> Option<&'catalog Entry> {
        self.entries.get(name).copied()
    }
}

fn insert_unique<'catalog, Value>(
    entries: &mut HashMap<&'catalog str, Value>,
    name: &'catalog str,
    value: Value,
) -> Result<(), Error> {
    match entries.entry(name) {
        HashMapEntry::Occupied(_) => Err(Error::Duplicate),
        HashMapEntry::Vacant(slot) => {
            slot.insert(value);
            Ok(())
        },
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, PhysicalEntryIndex, insert_unique};
    use std::collections::HashMap;

    #[test]
    fn duplicate_names_are_rejected() {
        let mut index = HashMap::new();
        insert_unique(&mut index, "Index/a.iwa", 1).expect("first name is unique");
        assert_eq!(
            insert_unique(&mut index, "Index/a.iwa", 2),
            Err(Error::Duplicate)
        );
        assert_eq!(index.get("Index/a.iwa"), Some(&1));
    }

    #[test]
    fn every_catalog_entry_is_retrievable_by_name() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers");
        let bytes = std::fs::read(path).expect("Numbers fixture");
        let catalog = litchi_iwa_archive::package::Catalog::from_bytes(&bytes)
            .expect("physical Numbers catalog");
        let index = PhysicalEntryIndex::new(&catalog).expect("physical entry index");
        for entry in catalog.iter() {
            assert!(std::ptr::eq(
                index.get(entry.name()).expect("indexed entry"),
                entry
            ));
        }
    }
}
