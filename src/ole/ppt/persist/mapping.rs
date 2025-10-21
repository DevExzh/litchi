//! Persist ID to byte offset mapping manager.
//!
//! Idiomatic Rust implementation using iterator chaining and functional patterns.

use super::ptr_holder::PersistPtrHolder;
use crate::ole::consts::PptRecordType;
use crate::ole::ppt::records::PptRecord;
use std::collections::HashMap;

/// Consolidated mapping from persist IDs to byte offsets.
///
/// Efficiently merges multiple PersistPtrHolder records with later entries overriding earlier ones.
#[derive(Debug, Clone, Default)]
pub struct PersistMapping {
    /// Consolidated mapping from persist ID to byte offset
    mappings: HashMap<u32, u32>,
}

impl PersistMapping {
    /// Create a new empty mapping.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Build mapping from all PersistPtrHolder records using functional iterator patterns.
    ///
    /// # Performance
    ///
    /// - Filters records using iterator chaining
    /// - Pre-allocates HashMap based on record count
    /// - Later records override earlier ones (most recent wins)
    pub fn build_from_records(records: &[PptRecord]) -> Self {
        // Count PersistPtrHolder records for capacity estimation
        let ptr_holder_count = records
            .iter()
            .filter(|r| r.record_type == PptRecordType::PersistPtrHolder)
            .count();

        // Pre-allocate with estimated capacity (assume ~10 slides per holder on average)
        let mut mappings = HashMap::with_capacity(ptr_holder_count * 10);

        // Process all PersistPtrHolder records in order
        records
            .iter()
            .filter(|r| r.record_type == PptRecordType::PersistPtrHolder)
            .filter_map(|r| PersistPtrHolder::parse(r).ok())
            .for_each(|holder| {
                // Extend mappings (later entries override earlier ones)
                mappings.extend(holder.slide_locations().iter().map(|(&k, &v)| (k, v)));
            });

        Self { mappings }
    }

    /// Get the byte offset for a given persist ID.
    #[inline]
    pub fn get_offset(&self, persist_id: u32) -> Option<u32> {
        self.mappings.get(&persist_id).copied()
    }

    /// Get all known persist IDs in sorted order.
    pub fn get_persist_ids(&self) -> Vec<u32> {
        let mut ids: Vec<u32> = self.mappings.keys().copied().collect();
        ids.sort_unstable();
        ids
    }

    /// Iterator over all (persist_id, offset) pairs.
    pub fn iter(&self) -> impl Iterator<Item = (&u32, &u32)> {
        self.mappings.iter()
    }

    /// Get the number of mappings.
    #[inline]
    pub fn len(&self) -> usize {
        self.mappings.len()
    }

    /// Check if the mapping is empty.
    #[inline]
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
    pub fn mappings(&self) -> &HashMap<u32, u32> {
        &self.mappings
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_persist_mapping_from_records() {
        // Create test records with two PersistPtrHolder records
        let mut data1 = Vec::new();
        data1.extend_from_slice(&0x00200000u32.to_le_bytes()); // base=0, count=2
        data1.extend_from_slice(&1000u32.to_le_bytes());
        data1.extend_from_slice(&2000u32.to_le_bytes());

        let record1 = PptRecord {
            record_type: PptRecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: data1.len() as u32,
            data: data1,
            children: vec![],
        };

        // Second holder updates persist_id=0
        let mut data2 = Vec::new();
        data2.extend_from_slice(&0x00100000u32.to_le_bytes()); // base=0, count=1
        data2.extend_from_slice(&1500u32.to_le_bytes()); // updated offset

        let record2 = PptRecord {
            record_type: PptRecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: data2.len() as u32,
            data: data2,
            children: vec![],
        };

        let records = vec![record1, record2];
        let mapping = PersistMapping::build_from_records(&records);

        // persist_id=0 should have the updated offset from record2
        assert_eq!(mapping.get_offset(0), Some(1500));
        // persist_id=1 should still have the original offset
        assert_eq!(mapping.get_offset(1), Some(2000));
    }
}
