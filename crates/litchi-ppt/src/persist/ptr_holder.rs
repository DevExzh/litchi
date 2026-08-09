//! `PersistPtrHolder` parsing - maps persist IDs to byte offsets.
//!
//! Idiomatic Rust implementation with zero-copy parsing and high performance.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashMap;

/// Holder for persist pointer mappings from persist IDs to byte offsets.
///
/// Uses efficient iteration over 4-byte chunks for high-performance parsing.
#[derive(Debug, Clone)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`PersistPtrHolder` is the established public API name re-exported from the crate root; renaming it would break downstream crates"
)]
pub struct PersistPtrHolder {
    /// Map from persist ID (slide ID) to byte offset in the document stream
    slide_locations: HashMap<u32, u32>,
}

impl PersistPtrHolder {
    /// Parse a `PersistPtrHolder` from a PPT record using idiomatic Rust.
    ///
    /// # Data Format
    ///
    /// Repeating pattern:
    /// - 32-bit info: [lower 20 bits: `base_id`] [upper 12 bits: count]
    /// - count × 32-bit offsets
    ///
    /// # Performance
    ///
    /// - Uses `chunks_exact(4)` for efficient 4-byte iteration
    /// - Pre-allocates `HashMap` with estimated capacity
    /// - Zero-copy: reads directly from slice without intermediate allocations
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::PersistPtrHolder {
            return Err(Error::InvalidFormat(format!(
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
            let info =
                u32::from_le_bytes([info_bytes[0], info_bytes[1], info_bytes[2], info_bytes[3]]);

            // Bit manipulation for decoding
            let base_persist_id = info & 0x000F_FFFF;
            let entry_count = (info >> 20) & 0x0FFF;

            // Read offset entries for this group
            if entry_count == 0 {
                return Err(Error::Corrupted(
                    "PersistDirectoryEntry has a zero entry count".to_string(),
                ));
            }
            for i in 0..entry_count {
                let offset_bytes = chunks.next().ok_or_else(|| {
                    Error::Corrupted("truncated PersistDirectoryEntry".to_string())
                })?;
                let offset = u32::from_le_bytes([
                    offset_bytes[0],
                    offset_bytes[1],
                    offset_bytes[2],
                    offset_bytes[3],
                ]);
                if slide_locations
                    .insert(base_persist_id + i, offset)
                    .is_some()
                {
                    return Err(Error::Corrupted(format!(
                        "duplicate persist identifier {}",
                        base_persist_id + i
                    )));
                }
            }
        }

        if !chunks.remainder().is_empty() {
            return Err(Error::Corrupted(
                "persist directory is not aligned to 4 bytes".to_string(),
            ));
        }

        Ok(Self { slide_locations })
    }

    /// Get the byte offset for a given persist ID.
    #[inline]
    #[must_use]
    pub fn get_slide_location(&self, persist_id: u32) -> Option<u32> {
        self.slide_locations.get(&persist_id).copied()
    }

    /// Get all known persist IDs in sorted order.
    #[must_use]
    pub fn get_known_slide_ids(&self) -> Vec<u32> {
        let mut ids: Vec<u32> = self.slide_locations.keys().copied().collect();
        ids.sort_unstable();
        ids
    }

    /// Get immutable reference to the slide locations map.
    #[inline]
    #[must_use]
    pub fn slide_locations(&self) -> &HashMap<u32, u32> {
        &self.slide_locations
    }

    /// Get the number of slides tracked by this holder.
    #[inline]
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slide_locations.len()
    }

    /// Check if this holder is empty.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.slide_locations.is_empty()
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
    fn test_persist_ptr_holder_parsing() {
        // Create a test record with one group: base_id=0, count=2
        // info = 0 | (2 << 20) = 0x00200000
        // offsets: 1000, 2000
        let mut data = Vec::new();
        data.extend_from_slice(&0x0020_0000_u32.to_le_bytes()); // info: base=0, count=2
        data.extend_from_slice(&1000u32.to_le_bytes()); // offset for persist_id=0
        data.extend_from_slice(&2000u32.to_le_bytes()); // offset for persist_id=1

        let record = Record {
            record_type: RecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data.len()).unwrap(),
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
        data.extend_from_slice(&0x0020_0000_u32.to_le_bytes()); // base=0, count=2
        data.extend_from_slice(&1000u32.to_le_bytes());
        data.extend_from_slice(&2000u32.to_le_bytes());

        // Group 2
        data.extend_from_slice(&0x0010_000A_u32.to_le_bytes()); // base=10, count=1
        data.extend_from_slice(&3000u32.to_le_bytes());

        let record = Record {
            record_type: RecordType::PersistPtrHolder,
            record_type_raw: 6001,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data.len()).unwrap(),
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
