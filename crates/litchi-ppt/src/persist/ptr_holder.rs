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
    /// - Uses `as_chunks::<4>()` for efficient 4-byte iteration
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

        Self::parse_data(&record.data, None)
    }

    /// Parse a `PersistPtrHolder`, repairing per-entry defects into
    /// `report` instead of rejecting the whole holder (finding F5.7).
    ///
    /// A zero-count entry header is skipped
    /// ([`crate::PersistDefect::ZeroCountPersistRun`]) and a persist ID
    /// repeated within the holder keeps its first offset
    /// ([`crate::PersistDefect::DuplicatePersistId`]), as LibreOffice reads
    /// them. A truncated or misaligned directory is still an error.
    ///
    /// # Errors
    ///
    /// Returns an error if the record is not a `PersistPtrHolder` or its
    /// entries are truncated or misaligned.
    pub fn parse_with_report(
        record: &Record,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Self> {
        if record.record_type != RecordType::PersistPtrHolder {
            return Err(Error::InvalidFormat(format!(
                "Expected PersistPtrHolder, got {:?}",
                record.record_type
            )));
        }

        Self::parse_data(&record.data, Some(report))
    }

    /// Parse from raw data (useful for testing and direct parsing). Every
    /// deviation is an error unless a report is threaded.
    fn parse_data(
        data: &[u8],
        mut report: Option<&mut crate::leniency::ToleranceReport>,
    ) -> Result<Self> {
        use crate::leniency::{DefectLocation, PersistDefect};

        // Estimate capacity: each group has 1 info + n offsets, minimum 2 u32s per group
        let estimated_capacity = data.len() / 8;
        let mut slide_locations = HashMap::with_capacity(estimated_capacity);

        let (chunks, remainder) = data.as_chunks::<4>();
        let mut chunks = chunks.iter();

        while let Some(info_bytes) = chunks.next() {
            let info =
                u32::from_le_bytes([info_bytes[0], info_bytes[1], info_bytes[2], info_bytes[3]]);

            // Bit manipulation for decoding
            let base_persist_id = info & 0x000F_FFFF;
            let entry_count = (info >> 20) & 0x0FFF;

            // Read offset entries for this group
            if entry_count == 0 {
                let Some(report) = report.as_deref_mut() else {
                    return Err(Error::Corrupted(
                        "PersistDirectoryEntry has a zero entry count".to_string(),
                    ));
                };
                report.record(PersistDefect::ZeroCountPersistRun, base_persist_id);
                continue;
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
                let persist_id = base_persist_id + i;
                if slide_locations.contains_key(&persist_id) {
                    let Some(report) = report.as_deref_mut() else {
                        return Err(Error::Corrupted(format!(
                            "duplicate persist identifier {persist_id}"
                        )));
                    };
                    report.record_at(
                        PersistDefect::DuplicatePersistId,
                        persist_id,
                        DefectLocation {
                            persist_id: Some(persist_id),
                            shape_id: None,
                        },
                    );
                    continue;
                }
                slide_locations.insert(persist_id, offset);
            }
        }

        if !remainder.is_empty() {
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
            data: data.into(),
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
            data: data.into(),
            children: vec![],
        };

        let holder = PersistPtrHolder::parse(&record).unwrap();

        assert_eq!(holder.slide_count(), 3);
        assert_eq!(holder.get_slide_location(0), Some(1000));
        assert_eq!(holder.get_slide_location(1), Some(2000));
        assert_eq!(holder.get_slide_location(10), Some(3000));
    }

    fn holder_record(words: &[u32]) -> Record {
        let data: Vec<u8> = words.iter().flat_map(|word| word.to_le_bytes()).collect();
        Record {
            record_type: RecordType::PersistPtrHolder,
            record_type_raw: 6002,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data.len()).unwrap(),
            data: data.into(),
            children: vec![],
        }
    }

    /// Zero-count entries are skipped and repeated persist IDs keep their
    /// first offset, recorded, instead of voiding the holder (F5.7).
    #[test]
    fn repairs_zero_count_and_duplicate_entries_per_entry() {
        use crate::{Defect, PersistDefect};

        let record = holder_record(&[
            0x0000_0007, // base 7, count 0
            0x0020_0001, // base 1, count 2
            100,
            200,
            0x0010_0002, // base 2 again
            300,
        ]);
        assert!(PersistPtrHolder::parse(&record).is_err());

        let mut report = crate::leniency::ToleranceReport::default();
        let holder = PersistPtrHolder::parse_with_report(&record, &mut report).unwrap();
        assert_eq!(holder.get_slide_location(1), Some(100));
        assert_eq!(holder.get_slide_location(2), Some(200));
        assert_eq!(holder.slide_count(), 2);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|defect| (defect.defect, defect.context, defect.location.persist_id))
                .collect::<Vec<_>>(),
            [
                (Defect::Persist(PersistDefect::ZeroCountPersistRun), 7, None),
                (
                    Defect::Persist(PersistDefect::DuplicatePersistId),
                    2,
                    Some(2)
                ),
            ]
        );

        // The public mapping keeps the repaired holder instead of dropping
        // it, and drops a misaligned holder without recording anything.
        let misaligned = Record {
            data: vec![1, 2, 3].into(),
            ..holder_record(&[])
        };
        let mut report = crate::leniency::ToleranceReport::default();
        let mapping = crate::PersistMapping::build_from_records_with_report(
            &[record.clone(), misaligned],
            &mut report,
        );
        assert_eq!(mapping.get_offset(2), Some(200));
        assert_eq!(report.total(), 2);
        assert_eq!(
            crate::PersistMapping::build_from_records(&[record]).get_offset(1),
            Some(100)
        );
    }
}
