//! Persist ID to byte offset mapping manager.
//!
//! Idiomatic Rust implementation using iterator chaining and functional patterns.

use super::ptr_holder::PersistPtrHolder;
use crate::consts::RecordType;
use crate::records::Record;
use std::collections::HashMap;

/// Consolidated mapping from persist IDs to byte offsets.
///
/// Efficiently merges multiple `PersistPtrHolder` records with later entries overriding earlier ones.
#[derive(Debug, Clone, Default)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`PersistMapping` is the established public API name re-exported from the crate root; renaming it would break downstream crates"
)]
pub struct PersistMapping {
    /// Consolidated mapping from persist ID to byte offset
    mappings: HashMap<u32, u32>,
}

impl PersistMapping {
    /// Create a new empty mapping.
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Build mapping from all `PersistPtrHolder` records using functional iterator patterns.
    ///
    /// # Performance
    ///
    /// - Filters records using iterator chaining
    /// - Pre-allocates `HashMap` based on record count
    /// - Later records override earlier ones (most recent wins)
    #[must_use]
    pub fn build_from_records(records: &[Record]) -> Self {
        // Count PersistPtrHolder records for capacity estimation
        let ptr_holder_count = records
            .iter()
            .filter(|r| r.record_type == RecordType::PersistPtrHolder)
            .count();

        // Pre-allocate with estimated capacity (assume ~10 slides per holder on average)
        let mut mappings = HashMap::with_capacity(ptr_holder_count * 10);

        // Process all PersistPtrHolder records in order
        records
            .iter()
            .filter(|r| r.record_type == RecordType::PersistPtrHolder)
            .filter_map(|r| PersistPtrHolder::parse(r).ok())
            .for_each(|holder| {
                // Extend mappings (later entries override earlier ones)
                mappings.extend(holder.slide_locations().iter().map(|(&k, &v)| (k, v)));
            });

        Self { mappings }
    }

    /// Build mapping from record references (zero-copy version).
    ///
    /// # Performance
    ///
    /// - Zero-copy: works with references instead of owned records
    /// - Avoids cloning large record data (`Vec<u8>`)
    /// - Same logic as `build_from_records` but more efficient
    #[must_use]
    pub fn build_from_records_ref(records: &[&Record]) -> Self {
        // Count PersistPtrHolder records for capacity estimation
        let ptr_holder_count = records
            .iter()
            .filter(|r| r.record_type == RecordType::PersistPtrHolder)
            .count();

        // Pre-allocate with estimated capacity (assume ~10 slides per holder on average)
        let mut mappings = HashMap::with_capacity(ptr_holder_count * 10);

        // Process all PersistPtrHolder records in order
        records
            .iter()
            .filter(|r| r.record_type == RecordType::PersistPtrHolder)
            .filter_map(|r| PersistPtrHolder::parse(r).ok())
            .for_each(|holder| {
                // Extend mappings (later entries override earlier ones)
                mappings.extend(holder.slide_locations().iter().map(|(&k, &v)| (k, v)));
            });

        Self { mappings }
    }

    /// Get the byte offset for a given persist ID.
    #[inline]
    #[must_use]
    pub fn get_offset(&self, persist_id: u32) -> Option<u32> {
        self.mappings.get(&persist_id).copied()
    }

    /// Get all known persist IDs in sorted order.
    #[must_use]
    pub fn get_persist_ids(&self) -> Vec<u32> {
        let mut ids: Vec<u32> = self.mappings.keys().copied().collect();
        ids.sort_unstable();
        ids
    }

    /// Iterator over all (`persist_id`, offset) pairs.
    pub fn iter(&self) -> impl Iterator<Item = (&u32, &u32)> {
        self.mappings.iter()
    }

    /// Get the number of mappings.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.mappings.len()
    }

    /// Check if the mapping is empty.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.mappings.is_empty()
    }

    /// Add a mapping manually.
    #[inline]
    pub fn add_mapping(&mut self, persist_id: u32, offset: u32) {
        self.mappings.insert(persist_id, offset);
    }

    /// Get immutable reference to all mappings.
    #[inline]
    #[must_use]
    pub fn mappings(&self) -> &HashMap<u32, u32> {
        &self.mappings
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn test_persist_mapping_from_records() {
        // Create test records with two PersistPtrHolder records
        let mut data1 = Vec::new();
        data1.extend_from_slice(&0x0020_0000_u32.to_le_bytes()); // base=0, count=2
        data1.extend_from_slice(&1000u32.to_le_bytes());
        data1.extend_from_slice(&2000u32.to_le_bytes());

        let first_record = Record {
            record_type: RecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data1.len()).unwrap(),
            data: data1,
            children: vec![],
        };

        // Second holder updates persist_id=0
        let mut data2 = Vec::new();
        data2.extend_from_slice(&0x0010_0000_u32.to_le_bytes()); // base=0, count=1
        data2.extend_from_slice(&1500u32.to_le_bytes()); // updated offset

        let second_record = Record {
            record_type: RecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data2.len()).unwrap(),
            data: data2,
            children: vec![],
        };

        let records = vec![first_record, second_record];
        let mapping = PersistMapping::build_from_records(&records);

        // persist_id=0 should have the updated offset from record2
        assert_eq!(mapping.get_offset(0), Some(1500));
        // persist_id=1 should still have the original offset
        assert_eq!(mapping.get_offset(1), Some(2000));
    }
}
