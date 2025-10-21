//! PersistPtrHolder parsing - maps persist IDs to byte offsets.
//!
//! Idiomatic Rust implementation with zero-copy parsing and high performance.

use crate::ole::consts::PptRecordType;
use crate::ole::ppt::package::{PptError, Result};
use crate::ole::ppt::records::PptRecord;
use std::collections::HashMap;

/// Holder for persist pointer mappings from persist IDs to byte offsets.
///
/// Uses efficient iteration over 4-byte chunks for high-performance parsing.
#[derive(Debug, Clone)]
pub struct PersistPtrHolder {
    /// Map from persist ID (slide ID) to byte offset in the document stream
    slide_locations: HashMap<u32, u32>,
}

impl PersistPtrHolder {
    /// Parse a PersistPtrHolder from a PPT record using idiomatic Rust.
    ///
    /// # Data Format
    ///
    /// Repeating pattern:
    /// - 32-bit info: [lower 20 bits: base_id] [upper 12 bits: count]
    /// - count × 32-bit offsets
    ///
    /// # Performance
    ///
    /// - Uses `chunks_exact(4)` for efficient 4-byte iteration
    /// - Pre-allocates HashMap with estimated capacity
    /// - Zero-copy: reads directly from slice without intermediate allocations
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::PersistPtrHolder {
            return Err(PptError::InvalidFormat(format!(
                "Expected PersistPtrHolder, got {:?}",
                record.record_type
            )));
        }

        Self::parse_data(&record.data)
    }

    /// Parse from raw data (useful for testing and direct parsing).
    fn parse_data(data: &[u8]) -> Result<Self> {
        // Estimate capacity: each group has 1 info + n offsets, minimum 2 u32s per group
        let estimated_capacity = data.len() / 8;
        let mut slide_locations = HashMap::with_capacity(estimated_capacity);

        let mut chunks = data.chunks_exact(4);

        while let Some(info_bytes) = chunks.next() {
            let info = u32::from_le_bytes(info_bytes.try_into().unwrap());

            // Bit manipulation for decoding
            let base_persist_id = info & 0x000F_FFFF;
            let entry_count = (info >> 20) & 0x0FFF;

            // Read offset entries for this group
            for i in 0..entry_count {
                if let Some(offset_bytes) = chunks.next() {
                    let offset = u32::from_le_bytes(offset_bytes.try_into().unwrap());
                    slide_locations.insert(base_persist_id + i, offset);
                } else {
                    // Truncated data - stop parsing
                    break;
                }
            }
        }

        Ok(Self { slide_locations })
    }

    /// Get the byte offset for a given persist ID.
    #[inline]
    pub fn get_slide_location(&self, persist_id: u32) -> Option<u32> {
        self.slide_locations.get(&persist_id).copied()
    }

    /// Get all known persist IDs in sorted order.
    pub fn get_known_slide_ids(&self) -> Vec<u32> {
        let mut ids: Vec<u32> = self.slide_locations.keys().copied().collect();
        ids.sort_unstable();
        ids
    }

    /// Get immutable reference to the slide locations map.
    #[inline]
    pub fn slide_locations(&self) -> &HashMap<u32, u32> {
        &self.slide_locations
    }

    /// Get the number of slides tracked by this holder.
    #[inline]
    pub fn slide_count(&self) -> usize {
        self.slide_locations.len()
    }

    /// Check if this holder is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.slide_locations.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_persist_ptr_holder_parsing() {
        // Create a test record with one group: base_id=0, count=2
        // info = 0 | (2 << 20) = 0x00200000
        // offsets: 1000, 2000
        let mut data = Vec::new();
        data.extend_from_slice(&0x00200000u32.to_le_bytes()); // info: base=0, count=2
        data.extend_from_slice(&1000u32.to_le_bytes()); // offset for persist_id=0
        data.extend_from_slice(&2000u32.to_le_bytes()); // offset for persist_id=1

        let record = PptRecord {
            record_type: PptRecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children: vec![],
        };

        let holder = PersistPtrHolder::parse(&record).unwrap();

        assert_eq!(holder.slide_count(), 2);
        assert_eq!(holder.get_slide_location(0), Some(1000));
        assert_eq!(holder.get_slide_location(1), Some(2000));
        assert_eq!(holder.get_slide_location(2), None);
    }

    #[test]
    fn test_persist_ptr_holder_multiple_groups() {
        // Group 1: base_id=0, count=2
        // Group 2: base_id=10, count=1
        let mut data = Vec::new();

        // Group 1
        data.extend_from_slice(&0x00200000u32.to_le_bytes()); // base=0, count=2
        data.extend_from_slice(&1000u32.to_le_bytes());
        data.extend_from_slice(&2000u32.to_le_bytes());

        // Group 2
        data.extend_from_slice(&0x0010000Au32.to_le_bytes()); // base=10, count=1
        data.extend_from_slice(&3000u32.to_le_bytes());

        let record = PptRecord {
            record_type: PptRecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children: vec![],
        };

        let holder = PersistPtrHolder::parse(&record).unwrap();

        assert_eq!(holder.slide_count(), 3);
        assert_eq!(holder.get_slide_location(0), Some(1000));
        assert_eq!(holder.get_slide_location(1), Some(2000));
        assert_eq!(holder.get_slide_location(10), Some(3000));
    }
}
