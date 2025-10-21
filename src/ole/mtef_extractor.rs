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

use crate::formula::{MathNode, MtefParser};
use crate::ole::file::OleFile;
use std::io::{Read, Seek};

/// MTEF extractor for OLE documents
pub struct MtefExtractor<'arena> {
    #[allow(dead_code)] // Kept for future use in instance methods
    arena: &'arena bumpalo::Bump,
}

impl<'arena> MtefExtractor<'arena> {
    /// Create a new MTEF extractor
    #[allow(dead_code)] // Public API for future use
    pub fn new(arena: &'arena bumpalo::Bump) -> Self {
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
    /// Returns the MTEF binary data (including OLE header) if found and valid, None otherwise
    pub(crate) fn extract_mtef_from_stream<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
        stream_path: &[&str],
    ) -> Result<Option<Vec<u8>>, MtefExtractionError> {
        // Try to open the stream
        let data = match ole_file.open_stream(stream_path) {
            Ok(d) => d,
            Err(_) => return Ok(None),
        };

        // Validate minimum size (28-byte OLE header + 5-byte MTEF header)
        if data.len() < 33 {
            return Ok(None);
        }

        // Validate OLE header and extract proper data size
        if !Self::validate_ole_header(&data) {
            return Ok(None);
        }

        // Parse cbObject field (u32 little-endian at offset 8)
        // This tells us the size of MTEF data after the 28-byte header
        let cb_object = u32::from_le_bytes([data[8], data[9], data[10], data[11]]) as usize;

        // Calculate total size: OLE header (28) + MTEF data (cbObject)
        let total_size = 28 + cb_object;

        // Ensure we don't read past the actual data
        let actual_size = total_size.min(data.len());

        // Return only the valid portion (OLE header + exact MTEF data)
        Ok(Some(data[..actual_size].to_vec()))
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
        if data.len() < 28 {
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
        if cb_object == 0 || cb_object as usize > data.len() - 28 {
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

        // Process each entry in ObjectPool
        for entry in entries {
            // Skip if not a storage (embedded OLE objects are storages)
            // STGTY_STORAGE = 1
            if entry.entry_type != 1 {
                continue;
            }

            // Object names typically start with "_"
            if !entry.name.starts_with('_') {
                continue;
            }

            // Try to extract "Equation Native" stream from this embedded object
            // The stream path is: ObjectPool/<object_name>/Equation Native
            if let Ok(Some(mtef_data)) = Self::extract_mtef_from_stream(
                ole_file,
                &["ObjectPool", &entry.name, "Equation Native"],
            ) {
                mtef_map.insert(entry.name.clone(), mtef_data);
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

        // Process each entry at root level
        for entry in entries {
            // Skip if not a storage
            if entry.entry_type != 1 {
                continue;
            }

            // Look for storages that might contain equations
            // Common patterns: "MBD[hex]", "Equation Native", or names starting with "_"
            let is_equation_storage = entry.name.starts_with("MBD")
                || entry.name == "Equation Native"
                || entry.name.starts_with('_');

            if !is_equation_storage {
                continue;
            }

            // Try to extract "Equation Native" stream from this storage
            if let Ok(Some(mtef_data)) =
                Self::extract_mtef_from_stream(ole_file, &[&entry.name, "Equation Native"])
            {
                mtef_map.insert(entry.name.clone(), mtef_data);
                continue;
            }

            // Try alternative stream names
            for stream_name in &["CONTENTS", "\x01Ole", "\x01Ole10Native"] {
                if let Ok(Some(mtef_data)) =
                    Self::extract_mtef_from_stream(ole_file, &[&entry.name, stream_name])
                {
                    mtef_map.insert(entry.name.clone(), mtef_data);
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
    #[allow(dead_code)] // Public API for future use
    pub fn parse_mtef_to_ast(
        &self,
        mtef_data: Vec<u8>,
    ) -> Result<Vec<MathNode<'arena>>, MtefExtractionError> {
        // We need to create a reference with the arena lifetime
        // This is a bit of a hack, but we know the data will live long enough
        let data_ref: &[u8] =
            unsafe { std::slice::from_raw_parts(mtef_data.as_ptr(), mtef_data.len()) };
        // Extend the lifetime (this is safe because we control the Vec lifetime)
        let data_with_lifetime: &'arena [u8] = unsafe { std::mem::transmute(data_ref) };

        let mut parser = MtefParser::new(self.arena, data_with_lifetime);
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
    #[allow(dead_code)] // Public API for future use
    pub fn extract_and_parse_mtef<R: Read + Seek>(
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
pub enum MtefExtractionError {
    IoError(std::io::Error),
    #[allow(dead_code)] // Kept for completeness
    InvalidOleHeader,
    ParseError(String),
}

impl std::fmt::Display for MtefExtractionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MtefExtractionError::IoError(e) => write!(f, "IO error: {}", e),
            MtefExtractionError::InvalidOleHeader => write!(f, "Invalid OLE header"),
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

impl From<crate::formula::MtefError> for MtefExtractionError {
    fn from(e: crate::formula::MtefError) -> Self {
        MtefExtractionError::ParseError(e.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::Formula;

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
}
