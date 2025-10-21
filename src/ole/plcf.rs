//! Property List with Character Positions (PLCF) parser.
//!
//! PLCF is a data structure used extensively in legacy Office binary formats
//! to map character positions to properties or data.

use crate::common::binary;
use bytes::Bytes;

/// Property List with Character Positions (PLCF) parser.
///
/// Based on Apache POI's PlexOfCps. PLCF is a common structure in Office files
/// that maps character positions to properties or data.
///
/// # Format
///
/// PLCF format:
/// - n+1 character positions (4 bytes each)
/// - n property elements (element_size bytes each)
///
/// # Examples
///
/// ```
/// use litchi::ole::plcf::PlcfParser;
///
/// // Create a simple PLCF with 2 elements, element_size = 2
/// // CPs: 0, 10, 20
/// // Props: [1, 2], [3, 4]
/// let data = vec![
///     0x00, 0x00, 0x00, 0x00, // CP 0
///     0x0A, 0x00, 0x00, 0x00, // CP 10
///     0x14, 0x00, 0x00, 0x00, // CP 20
///     0x01, 0x02, // Property 1
///     0x03, 0x04, // Property 2
/// ];
///
/// let plcf = PlcfParser::parse(&data, 2).unwrap();
/// assert_eq!(plcf.count(), 2);
/// assert_eq!(plcf.position(0), Some(0));
/// assert_eq!(plcf.position(1), Some(10));
/// assert_eq!(plcf.position(2), Some(20));
/// assert_eq!(plcf.range(0), Some((0, 10)));
/// assert_eq!(plcf.range(1), Some((10, 20)));
/// ```
pub struct PlcfParser {
    /// Character positions (CP array)
    positions: Vec<u32>,
    /// Property data buffer containing all property elements
    properties_data: Bytes,
    /// Offsets into properties_data for each property element
    properties_offsets: Vec<(usize, usize)>, // (offset, length) pairs
}

impl PlcfParser {
    /// Parse a PLCF structure from binary data.
    ///
    /// # Arguments
    ///
    /// * `data` - The binary data containing the PLCF
    /// * `element_size` - Size in bytes of each property element
    pub fn parse(data: &[u8], element_size: usize) -> Option<Self> {
        if data.len() < 4 {
            return None;
        }

        // Calculate number of elements
        // Formula: (data_length) / (4 + element_size) = n
        // So: n+1 CPs (4 bytes each) + n elements (element_size each)
        let n = if element_size > 0 {
            (data.len() - 4) / (4 + element_size)
        } else {
            return None;
        };

        if n == 0 {
            return Some(Self {
                positions: Vec::new(),
                properties_data: Bytes::new(),
                properties_offsets: Vec::new(),
            });
        }

        // Read character positions
        let mut positions = Vec::with_capacity(n + 1);
        for i in 0..=n {
            let offset = i * 4;
            if let Ok(cp) = binary::read_u32_le(data, offset) {
                positions.push(cp);
            } else {
                return None;
            }
        }

        // Read property data into a single Bytes buffer
        let props_start = (n + 1) * 4;
        let props_end = props_start + (n * element_size);
        if props_end > data.len() {
            return None;
        }

        let properties_data = Bytes::copy_from_slice(&data[props_start..props_end]);
        let mut properties_offsets = Vec::with_capacity(n);

        for i in 0..n {
            let offset = i * element_size;
            properties_offsets.push((offset, element_size));
        }

        Some(Self {
            positions,
            properties_data,
            properties_offsets,
        })
    }

    /// Get the number of elements in the PLCF.
    #[inline]
    pub fn count(&self) -> usize {
        self.properties_offsets.len()
    }

    /// Get character position at index.
    #[inline]
    pub fn position(&self, index: usize) -> Option<u32> {
        self.positions.get(index).copied()
    }

    /// Get property data at index.
    #[inline]
    pub fn property(&self, index: usize) -> Option<&[u8]> {
        self.properties_offsets
            .get(index)
            .map(|(offset, len)| &self.properties_data[*offset..*offset + *len])
    }

    /// Get character range for element at index.
    ///
    /// Returns (start_cp, end_cp) tuple.
    pub fn range(&self, index: usize) -> Option<(u32, u32)> {
        if index >= self.properties_offsets.len() {
            return None;
        }
        Some((self.positions[index], self.positions[index + 1]))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_plcf_parser() {
        // Create a simple PLCF with 2 elements, element_size = 2
        // CPs: 0, 10, 20
        // Props: [1, 2], [3, 4]
        let data = vec![
            0x00, 0x00, 0x00, 0x00, // CP 0
            0x0A, 0x00, 0x00, 0x00, // CP 10
            0x14, 0x00, 0x00, 0x00, // CP 20
            0x01, 0x02, // Property 1
            0x03, 0x04, // Property 2
        ];

        let plcf = PlcfParser::parse(&data, 2).unwrap();
        assert_eq!(plcf.count(), 2);
        assert_eq!(plcf.position(0), Some(0));
        assert_eq!(plcf.position(1), Some(10));
        assert_eq!(plcf.position(2), Some(20));
        assert_eq!(plcf.range(0), Some((0, 10)));
        assert_eq!(plcf.range(1), Some((10, 20)));
    }
}
