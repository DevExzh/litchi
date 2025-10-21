//! Escher record structure with zero-copy parsing.
//!
//! # Format
//!
//! Escher records have an 8-byte header:
//! - Bytes 0-1: Version and Instance (packed)
//! - Bytes 2-3: Record Type
//! - Bytes 4-7: Record Length (32-bit)
//!
//! # Performance
//!
//! - Zero-copy: Borrows from original data
//! - No intermediate allocations
//! - Direct byte access via slices

use super::types::EscherRecordType;
use crate::ole::ppt::package::{PptError, Result};
use zerocopy::{
    FromBytes,
    byteorder::{LittleEndian, U16, U32},
};

/// An Escher record with zero-copy data access.
#[derive(Debug, Clone)]
pub struct EscherRecord<'data> {
    /// Record type
    pub record_type: EscherRecordType,
    /// Raw record type value
    pub record_type_raw: u16,
    /// Version (4 bits)
    pub version: u8,
    /// Instance (12 bits)
    pub instance: u16,
    /// Length of record data (excluding header)
    pub length: u32,
    /// Record data (zero-copy borrow)
    pub data: &'data [u8],
}

impl<'data> EscherRecord<'data> {
    /// Parse an Escher record from binary data at the given offset.
    ///
    /// # Performance
    ///
    /// - Zero-copy: Returns slice references, no allocation
    /// - Single pass: Reads header once
    /// - Branch-free bit extraction
    ///
    /// # Returns
    ///
    /// `(record, bytes_consumed)` tuple
    pub fn parse(data: &'data [u8], offset: usize) -> Result<(Self, usize)> {
        // Verify we have enough data for header
        if offset + 8 > data.len() {
            return Err(PptError::Corrupted(
                "Not enough data for Escher record header".to_string(),
            ));
        }

        // Read header (8 bytes) - little-endian format
        // Bytes 0-1: Version (4 bits) | Instance (12 bits)
        let ver_inst = U16::<LittleEndian>::read_from_bytes(&data[offset..offset + 2])
            .map(|v| v.get())
            .unwrap_or(0);

        // Bytes 2-3: Record Type
        let record_type_raw = U16::<LittleEndian>::read_from_bytes(&data[offset + 2..offset + 4])
            .map(|v| v.get())
            .unwrap_or(0);

        // Bytes 4-7: Record Length
        let length = U32::<LittleEndian>::read_from_bytes(&data[offset + 4..offset + 8])
            .map(|v| v.get())
            .unwrap_or(0);

        // Extract version (lower 4 bits) and instance (upper 12 bits)
        let version = (ver_inst & 0x000F) as u8;
        let instance = (ver_inst >> 4) & 0x0FFF;

        let record_type = EscherRecordType::from(record_type_raw);

        // Verify record data is within bounds
        let data_end = offset + 8 + length as usize;
        if data_end > data.len() {
            // Allow partial reads for container records
            if record_type.is_container() && offset + 8 <= data.len() {
                let available_length = (data.len() - offset - 8) as u32;
                let record_data = &data[offset + 8..];

                return Ok((
                    Self {
                        record_type,
                        record_type_raw,
                        version,
                        instance,
                        length: available_length,
                        data: record_data,
                    },
                    8 + record_data.len(),
                ));
            }

            return Err(PptError::Corrupted(format!(
                "Escher record data extends beyond bounds: offset={}, length={}, data_len={}",
                offset,
                length,
                data.len()
            )));
        }

        // Zero-copy: Borrow slice from original data
        let record_data = &data[offset + 8..data_end];

        Ok((
            Self {
                record_type,
                record_type_raw,
                version,
                instance,
                length,
                data: record_data,
            },
            8 + length as usize,
        ))
    }

    /// Check if this is a container record (can have children).
    ///
    /// Container records have version 0xF (15).
    #[inline]
    pub fn is_container(&self) -> bool {
        self.version == 0x0F || self.record_type.is_container()
    }

    /// Check if this record can contain text.
    #[inline]
    pub fn can_contain_text(&self) -> bool {
        self.record_type.can_contain_text()
    }

    /// Get the offset of this record's data within the parent data slice.
    #[inline]
    pub fn data_offset(&self, parent_data: &[u8]) -> Option<usize> {
        let parent_ptr = parent_data.as_ptr() as usize;
        let data_ptr = self.data.as_ptr() as usize;

        if data_ptr >= parent_ptr && data_ptr < parent_ptr + parent_data.len() {
            Some(data_ptr - parent_ptr)
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_container_record() {
        // Create a test SpContainer record (0xF004)
        let data = vec![
            0x0F, 0x00, // version=0xF, instance=0
            0x04, 0xF0, // record type = 0xF004 (SpContainer)
            0x08, 0x00, 0x00, 0x00, // length = 8
            0x01, 0x02, 0x03, 0x04, // data
            0x05, 0x06, 0x07, 0x08,
        ];

        let (record, consumed) = EscherRecord::parse(&data, 0).unwrap();

        assert_eq!(record.version, 0x0F);
        assert_eq!(record.instance, 0);
        assert_eq!(record.record_type, EscherRecordType::SpContainer);
        assert_eq!(record.length, 8);
        assert_eq!(record.data.len(), 8);
        assert_eq!(consumed, 16);
        assert!(record.is_container());
    }

    #[test]
    fn test_parse_atom_record() {
        // Create a test Sp atom record (0xF00A)
        let data = vec![
            0x02, 0x00, // version=2, instance=0
            0x0A, 0xF0, // record type = 0xF00A (Sp)
            0x04, 0x00, 0x00, 0x00, // length = 4
            0xAA, 0xBB, 0xCC, 0xDD, // data
        ];

        let (record, consumed) = EscherRecord::parse(&data, 0).unwrap();

        assert_eq!(record.version, 0x02);
        assert_eq!(record.instance, 0);
        assert_eq!(record.record_type, EscherRecordType::Sp);
        assert_eq!(record.length, 4);
        assert_eq!(record.data, &[0xAA, 0xBB, 0xCC, 0xDD]);
        assert_eq!(consumed, 12);
        assert!(!record.is_container());
    }
}
