// MTEF (MathType Equation Format) Extractor for OLE Documents
//
// This module provides functionality to extract and parse MTEF binary data
// from OLE streams in legacy Office documents (.doc, .ppt, etc.).
//
// MTEF data in OLE2 documents is stored in streams named "Equation Native".
// The stream format is:
// - 28 bytes: EQNOLEFILEHDR (OLE header)
// - 5 bytes: MTEF header (version, platform, product, version major, version minor)
// - N bytes: MTEF byte stream (formula data)

use litchi_cfb::OleFile;
use litchi_formula::{MathNode, MtefParser};
use std::io::{Read, Seek};

const EQNOLE_HEADER_SIZE: usize = 28;
const MIN_MTEF_HEADER_SIZE: usize = 5;
const MAX_MTEF_OBJECT_SIZE: u32 = 16 * 1024 * 1024;

/// MTEF extractor for OLE documents
pub(crate) struct MtefExtractor<'arena> {
    #[allow(dead_code)] // Kept for future use in instance methods
    arena: &'arena bumpalo::Bump,
}

impl<'arena> MtefExtractor<'arena> {
    /// Create a new MTEF extractor
    #[allow(dead_code)] // Reserved for future formula integration
    pub(crate) fn new(arena: &'arena bumpalo::Bump) -> Self {
        Self { arena }
    }

    /// Extract MTEF data from an OLE stream path (internal use only)
    ///
    /// MTEF data is stored in streams named "Equation Native" within OLE storages.
    /// The stream format is:
    /// - 28 bytes: EQNOLEFILEHDR (OLE header)
    /// - Remaining bytes: MTEF data (5-byte header + MTEF byte stream)
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    /// * `stream_path` - Path components to the stream (e.g., &["ObjectPool", "_1234567890", "Equation Native"])
    ///
    /// # Returns
    ///
    /// Returns the MTEF binary data (including OLE header) when found and valid.
    /// Missing streams and unrecognized headers return `None`; structurally valid
    /// headers whose declared payload is truncated return an error.
    pub(crate) fn extract_mtef_from_stream<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
        stream_path: &[&str],
    ) -> Result<Option<Vec<u8>>, MtefExtractionError> {
        // Try to open the stream
        let mut data = match ole_file.open_stream(stream_path) {
            Ok(d) => d,
            Err(_) => return Ok(None),
        };

        // Validate minimum size (28-byte OLE header + 5-byte MTEF header)
        let Some(total_size) = Self::validated_stream_len(&data)? else {
            return Ok(None);
        };

        // Keep ownership of the stream allocation and discard only trailing bytes.
        data.truncate(total_size);
        Ok(Some(data))
    }

    /// Return the exact header-declared stream length after validating its bounds.
    fn validated_stream_len(data: &[u8]) -> Result<Option<usize>, MtefExtractionError> {
        if data.len() < EQNOLE_HEADER_SIZE + MIN_MTEF_HEADER_SIZE {
            return Ok(None);
        }

        if !Self::validate_ole_header(data) {
            return Ok(None);
        }

        let declared = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        let payload_len =
            usize::try_from(declared).map_err(|_| MtefExtractionError::PayloadSizeOverflow {
                declared: u64::from(declared),
            })?;
        let total_size = Self::checked_stream_len(payload_len)?;

        if total_size > data.len() {
            return Err(MtefExtractionError::TruncatedPayload {
                declared,
                available: data.len() - EQNOLE_HEADER_SIZE,
            });
        }

        Ok(Some(total_size))
    }

    fn checked_stream_len(payload_len: usize) -> Result<usize, MtefExtractionError> {
        EQNOLE_HEADER_SIZE.checked_add(payload_len).ok_or(
            MtefExtractionError::PayloadSizeOverflow {
                declared: payload_len as u64,
            },
        )
    }

    /// Validate the OLE header manually to avoid zerocopy alignment issues
    ///
    /// The OLE header (EQNOLEFILEHDR) is 28 bytes with the following structure:
    /// - Offset 0x00-0x01 (2 bytes): cb_hdr = 28 (header size)
    /// - Offset 0x02-0x05 (4 bytes): version (typically 0x00020000)
    /// - Offset 0x06-0x07 (2 bytes): format (varies, e.g., 0xC16D, 0xC19B, 0xC1C7, 0xC2D3)
    /// - Offset 0x08-0x0B (4 bytes): cbObject (size of MTEF data after header)
    /// - Offset 0x0C-0x1B (16 bytes): reserved
    fn validate_ole_header(data: &[u8]) -> bool {
        if data.len() < EQNOLE_HEADER_SIZE {
            return false;
        }

        // Parse cb_hdr (u16 little-endian at offset 0)
        let cb_hdr = u16::from_le_bytes([data[0], data[1]]);
        if cb_hdr != 28 {
            return false;
        }

        // Parse version (u32 little-endian at offset 2)
        let version = u32::from_le_bytes([data[2], data[3], data[4], data[5]]);
        // Accept both common version formats
        if version != 0x00020000 && version != 0x00000200 {
            return false;
        }

        // Parse format (u16 little-endian at offset 6)
        let format = u16::from_le_bytes([data[6], data[7]]);
        // Format can vary; common values are 0xC16D, 0xC19B, 0xC1C7, 0xC2D3
        // We accept any format in the 0xC1xx-0xC2xx range
        if !(0xC100..=0xC2FF).contains(&format) {
            return false;
        }

        // Parse cbObject (u32 little-endian at offset 8)
        let cb_object = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        // Check that the object size is reasonable
        if cb_object == 0 || cb_object > MAX_MTEF_OBJECT_SIZE {
            return false;
        }

        true
    }

    /// Extract all MTEF formulas from embedded OLE objects in ObjectPool
    ///
    /// In Word documents, embedded equations are stored as OLE objects in the ObjectPool directory.
    /// Each embedded object is a storage (directory) with a name like "_1234567890".
    /// Within each storage, there should be a stream named "Equation Native" containing the MTEF data.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    ///
    /// # Returns
    ///
    /// Returns a HashMap mapping object IDs to their MTEF binary data (including OLE header)
    pub(crate) fn extract_all_mtef_from_objectpool<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
    ) -> Result<std::collections::HashMap<String, Vec<u8>>, MtefExtractionError> {
        use std::collections::HashMap;

        let mut mtef_map = HashMap::new();

        // Check if ObjectPool directory exists
        if !ole_file.directory_exists(&["ObjectPool"]) {
            return Ok(mtef_map);
        }

        // List all entries in ObjectPool
        let entries = ole_file
            .list_directory_entries(&["ObjectPool"])
            .map_err(|e| MtefExtractionError::IoError(std::io::Error::other(e.to_string())))?;

        // Collect entry names first to avoid borrow conflicts
        let storage_names: Vec<String> = entries
            .iter()
            .filter(|entry| {
                // Skip if not a storage (embedded OLE objects are storages)
                // STGTY_STORAGE = 1
                entry.entry_type == 1 && entry.name.starts_with('_')
            })
            .map(|entry| entry.name.clone())
            .collect();

        // Process each entry in ObjectPool
        for entry_name in storage_names {
            // Try to extract "Equation Native" stream from this embedded object
            // The stream path is: ObjectPool/<object_name>/Equation Native
            if let Some(mtef_data) = Self::extract_mtef_from_stream(
                ole_file,
                &["ObjectPool", &entry_name, "Equation Native"],
            )? {
                mtef_map.insert(entry_name, mtef_data);
            }
        }

        Ok(mtef_map)
    }

    /// Extract all MTEF formulas from a PowerPoint presentation
    ///
    /// In PPT files, MTEF formulas follow a similar pattern to Word documents.
    /// Equations are stored as OLE objects in storage directories that are children of the root storage.
    /// These are typically named with patterns like "MBD[hexadecimal]" or "Equation Native".
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    ///
    /// # Returns
    ///
    /// Returns a HashMap mapping storage names to their MTEF binary data (including OLE header)
    #[allow(dead_code)]
    pub(crate) fn extract_all_mtef_from_ppt<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
    ) -> Result<std::collections::HashMap<String, Vec<u8>>, MtefExtractionError> {
        use std::collections::HashMap;

        let mut mtef_map = HashMap::new();

        // Get all root-level entries
        let entries = ole_file
            .list_directory_entries(&[])
            .map_err(|e| MtefExtractionError::IoError(std::io::Error::other(e.to_string())))?;

        // Collect storage names first to avoid borrow conflicts
        let storage_names: Vec<String> = entries
            .iter()
            .filter(|entry| {
                // Skip if not a storage
                if entry.entry_type != 1 {
                    return false;
                }
                // Look for storages that might contain equations
                // Common patterns: "MBD[hex]", "Equation Native", or names starting with "_"
                entry.name.starts_with("MBD")
                    || entry.name == "Equation Native"
                    || entry.name.starts_with('_')
            })
            .map(|entry| entry.name.clone())
            .collect();

        // Process each entry at root level
        for entry_name in storage_names {
            // Try to extract "Equation Native" stream from this storage
            if let Some(mtef_data) =
                Self::extract_mtef_from_stream(ole_file, &[&entry_name, "Equation Native"])?
            {
                mtef_map.insert(entry_name.clone(), mtef_data);
                continue;
            }

            // Try alternative stream names
            for stream_name in &["CONTENTS", "\x01Ole", "\x01Ole10Native"] {
                if let Some(mtef_data) =
                    Self::extract_mtef_from_stream(ole_file, &[&entry_name, stream_name])?
                {
                    mtef_map.insert(entry_name.clone(), mtef_data);
                    break;
                }
            }
        }

        Ok(mtef_map)
    }

    /// Parse MTEF binary data into formula AST nodes
    ///
    /// # Arguments
    ///
    /// * `mtef_data` - The MTEF binary data to parse
    ///
    /// # Returns
    ///
    /// Returns a vector of MathNode AST nodes representing the parsed formula
    #[allow(dead_code)] // Reserved for future formula integration
    pub(crate) fn parse_mtef_to_ast(
        &self,
        mtef_data: Vec<u8>,
    ) -> Result<Vec<MathNode<'arena>>, MtefExtractionError> {
        // The parser may borrow string data from its input. Copy the bytes into the
        // caller-provided arena so the input and returned nodes share the same real
        // lifetime; the source Vec can then be dropped normally.
        let arena_data = self.arena.alloc_slice_copy(&mtef_data);
        let mut parser = MtefParser::new(self.arena, arena_data);
        parser
            .parse()
            .map_err(|e| MtefExtractionError::ParseError(e.to_string()))
    }

    /// Extract and parse MTEF data from an OLE file
    ///
    /// This is a convenience method that combines extraction and parsing.
    /// It tries each stream name in order and returns the first valid MTEF data found.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    /// * `stream_names` - Possible stream names to check (e.g., &["Equation Native", "MSWordEquation"])
    ///
    /// # Returns
    ///
    /// Returns parsed formula AST nodes if MTEF data is found and valid
    #[allow(dead_code)] // Reserved for future formula integration
    pub(crate) fn extract_and_parse_mtef<R: Read + Seek>(
        &self,
        ole_file: &mut OleFile<R>,
        stream_names: &[&str],
    ) -> Result<Option<Vec<MathNode<'arena>>>, MtefExtractionError> {
        // Try each stream name in order
        for stream_name in stream_names {
            if let Some(mtef_data) = Self::extract_mtef_from_stream(ole_file, &[stream_name])? {
                // Found valid MTEF data - parse it
                return Ok(Some(self.parse_mtef_to_ast(mtef_data)?));
            }
        }

        // No valid MTEF data found in any of the stream names
        Ok(None)
    }
}

/// Errors that can occur during MTEF extraction
#[derive(Debug)]
pub(crate) enum MtefExtractionError {
    IoError(std::io::Error),
    #[allow(dead_code)] // Kept for completeness
    InvalidOleHeader,
    /// The declared payload cannot be combined with its fixed-size header.
    PayloadSizeOverflow {
        declared: u64,
    },
    /// The stream ends before the complete header-declared payload.
    TruncatedPayload {
        declared: u32,
        available: usize,
    },
    ParseError(String),
}

impl std::fmt::Display for MtefExtractionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MtefExtractionError::IoError(e) => write!(f, "IO error: {}", e),
            MtefExtractionError::InvalidOleHeader => write!(f, "Invalid OLE header"),
            MtefExtractionError::PayloadSizeOverflow { declared } => write!(
                f,
                "MTEF payload size {declared} overflows the platform stream length"
            ),
            MtefExtractionError::TruncatedPayload {
                declared,
                available,
            } => write!(
                f,
                "Truncated MTEF payload: header declares {declared} bytes, but only {available} are available"
            ),
            MtefExtractionError::ParseError(msg) => write!(f, "Parse error: {}", msg),
        }
    }
}

impl std::error::Error for MtefExtractionError {}

impl From<std::io::Error> for MtefExtractionError {
    fn from(e: std::io::Error) -> Self {
        MtefExtractionError::IoError(e)
    }
}

impl From<litchi_formula::MtefError> for MtefExtractionError {
    fn from(e: litchi_formula::MtefError) -> Self {
        MtefExtractionError::ParseError(e.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_formula::ast::Formula;

    fn equation_native_stream(declared: u32, actual_payload_len: usize) -> Vec<u8> {
        let mut data = vec![0; EQNOLE_HEADER_SIZE + actual_payload_len];
        data[0..2].copy_from_slice(&(EQNOLE_HEADER_SIZE as u16).to_le_bytes());
        data[2..6].copy_from_slice(&0x0002_0000_u32.to_le_bytes());
        data[6..8].copy_from_slice(&0xC2D3_u16.to_le_bytes());
        data[8..12].copy_from_slice(&declared.to_le_bytes());
        data
    }

    #[test]
    fn test_ole_header_validation() {
        // Valid OLE header
        let valid_data = vec![
            0x1C, 0x00, // cb_hdr = 28
            0x00, 0x00, 0x02, 0x00, // version = 0x00020000 (little endian)
            0xD3, 0xC2, // format = 0xC2D3
            0x20, 0x00, 0x00, 0x00, // size = 32
            0x00, 0x00, 0x00, 0x00, // reserved
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];

        assert!(MtefExtractor::validate_ole_header(&valid_data));

        // Invalid OLE header (wrong cb_hdr)
        let mut invalid_data = valid_data.clone();
        invalid_data[0] = 0x10; // Invalid cb_hdr
        assert!(!MtefExtractor::validate_ole_header(&invalid_data));

        // Invalid version
        let mut invalid_version = valid_data.clone();
        invalid_version[2] = 0xFF;
        assert!(!MtefExtractor::validate_ole_header(&invalid_version));
    }

    #[test]
    fn test_mtef_extractor_creation() {
        let formula = Formula::new();
        let _extractor = MtefExtractor::new(formula.arena());

        // Just test that it can be created
        // Test passed - no assertions needed
    }

    #[test]
    fn rejects_truncated_header_declared_payload() {
        let data = equation_native_stream(6, MIN_MTEF_HEADER_SIZE);

        assert!(matches!(
            MtefExtractor::validated_stream_len(&data),
            Err(MtefExtractionError::TruncatedPayload {
                declared: 6,
                available: MIN_MTEF_HEADER_SIZE,
            })
        ));
    }

    #[test]
    fn accepts_exact_and_trims_trailing_stream_lengths() {
        let exact = equation_native_stream(5, 5);
        assert_eq!(
            MtefExtractor::validated_stream_len(&exact).unwrap(),
            Some(EQNOLE_HEADER_SIZE + 5)
        );

        let trailing = equation_native_stream(5, 8);
        assert_eq!(
            MtefExtractor::validated_stream_len(&trailing).unwrap(),
            Some(EQNOLE_HEADER_SIZE + 5)
        );
    }

    #[test]
    fn reports_maximum_declared_payload_as_truncated_without_allocating_it() {
        let data = equation_native_stream(MAX_MTEF_OBJECT_SIZE, MIN_MTEF_HEADER_SIZE);

        assert!(matches!(
            MtefExtractor::validated_stream_len(&data),
            Err(MtefExtractionError::TruncatedPayload {
                declared: MAX_MTEF_OBJECT_SIZE,
                available: MIN_MTEF_HEADER_SIZE,
            })
        ));
    }

    #[test]
    fn checks_stream_size_at_usize_boundary() {
        let largest_payload = usize::MAX - EQNOLE_HEADER_SIZE;
        assert_eq!(
            MtefExtractor::checked_stream_len(largest_payload).unwrap(),
            usize::MAX
        );

        let overflowing_payload = largest_payload + 1;
        assert!(matches!(
            MtefExtractor::checked_stream_len(overflowing_payload),
            Err(MtefExtractionError::PayloadSizeOverflow { declared })
                if declared == overflowing_payload as u64
        ));
    }

    fn minimal_mtef() -> Vec<u8> {
        vec![
            0x1C, 0x00, 0x00, 0x00, 0x02, 0x00, 0xD3, 0xC2, 0x0B, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x28, 0x04, 0x6D, 0x74, 0x05, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x09, 0x00,
        ]
    }

    #[test]
    fn parsed_nodes_are_tied_to_the_arena_not_the_source_vec() {
        let formula = Formula::new();
        let extractor = MtefExtractor::new(formula.arena());
        let nodes = extractor
            .parse_mtef_to_ast(minimal_mtef())
            .expect("arena-owned MTEF parses");
        // The input Vec was consumed and dropped inside parse_mtef_to_ast. Accessing
        // the returned nodes remains sound because their borrowed data is in `formula`.
        assert!(
            nodes
                .iter()
                .all(|node| format!("{node:?}").len() < 1_000_000)
        );
    }

    #[test]
    fn malformed_input_is_inert_and_does_not_escape_a_borrow() {
        let formula = Formula::new();
        let extractor = MtefExtractor::new(formula.arena());
        let nodes = extractor
            .parse_mtef_to_ast(vec![0xFF; 12])
            .expect("invalid MTEF uses the parser's empty fallback");
        assert!(nodes.is_empty());
    }
}
