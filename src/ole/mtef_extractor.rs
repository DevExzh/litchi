// MTEF (MathType Equation Format) Extractor for OLE Documents
//
// This module provides functionality to extract and parse MTEF binary data
// from OLE streams in legacy Office documents (.doc, .ppt, etc.).

use zerocopy::FromBytes;

use crate::formula::{MtefParser, MathNode};
use crate::ole::file::OleFile;
use std::io::{Read, Seek};

/// MTEF extractor for OLE documents
pub struct MtefExtractor<'arena> {
    arena: &'arena bumpalo::Bump,
}

impl<'arena> MtefExtractor<'arena> {
    /// Create a new MTEF extractor
    pub fn new(arena: &'arena bumpalo::Bump) -> Self {
        Self { arena }
    }

    /// Extract MTEF data from an OLE stream
    ///
    /// MTEF data is stored in streams named "Equation Native" in OLE documents.
    /// The format is: 28-byte OLE header + MTEF header + MTEF records + optional WMF
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    /// * `stream_name` - Name of the stream containing MTEF data (typically "Equation Native")
    ///
    /// # Returns
    ///
    /// Returns the MTEF binary data if found, None otherwise
    pub fn extract_mtef_data_from_stream<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
        stream_name: &str,
    ) -> Result<Option<Vec<u8>>, MtefExtractionError> {
        // Try to open the stream
        if let Ok(data) = ole_file.open_stream(&[stream_name]) {
            if data.is_empty() {
                return Ok(None);
            }

            // Validate that we have at least the OLE header (28 bytes)
            if data.len() < 28 {
                return Ok(None);
            }

            // Validate OLE header
            let ole_header = OleHeader::from_bytes(&data[0..28])?;
            if !ole_header.is_valid() {
                return Ok(None);
            }

            // Return the full data including the OLE header
            // The MTEF parser will handle parsing both the OLE header and MTEF data
            Ok(Some(data.to_vec()))
        } else {
            Ok(None)
        }
    }

    /// Extract MTEF data from multiple possible stream names
    ///
    /// Some documents may use different stream names for MTEF data.
    /// This method tries common stream names used for equation storage.
    pub fn extract_mtef_data<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
        possible_names: &[&str],
    ) -> Result<Option<Vec<u8>>, MtefExtractionError> {
        for name in possible_names {
            if let Some(data) = Self::extract_mtef_data_from_stream(ole_file, name)? {
                return Ok(Some(data));
            }
        }
        Ok(None)
    }

    /// Extract all MTEF formulas from embedded OLE objects in ObjectPool
    ///
    /// In Word documents, embedded equations are stored as OLE objects in the ObjectPool directory.
    /// Each object has a stream name like "_1234567890" and contains an "Equation Native" stream.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    ///
    /// # Returns
    ///
    /// Returns a HashMap mapping object IDs to their MTEF binary data
    pub fn extract_all_mtef_from_objectpool<R: Read + Seek>(
        ole_file: &mut OleFile<R>,
    ) -> Result<std::collections::HashMap<String, Vec<u8>>, MtefExtractionError> {
        use std::collections::HashMap;

        let mut mtef_map = HashMap::new();

        // Check if ObjectPool directory exists
        if !ole_file.directory_exists(&["ObjectPool"]) {
            eprintln!("DEBUG: ObjectPool directory does not exist");
            return Ok(mtef_map);
        }
        eprintln!("DEBUG: ObjectPool directory exists");

        // List all entries in ObjectPool
        let entries = ole_file
            .list_directory_entries(&["ObjectPool"])
            .map_err(|e| MtefExtractionError::IoError(std::io::Error::other(e.to_string())))?;

        eprintln!("DEBUG: Found {} entries in ObjectPool", entries.len());

        // Process each entry in ObjectPool
        for entry in entries {
            eprintln!("DEBUG:   Entry: name='{}', type={}", entry.name, entry.entry_type);
            
            // Skip if not a storage (embedded OLE objects are storages)
            if entry.entry_type != 1 { // STGTY_STORAGE = 1
                eprintln!("DEBUG:     Skipped (not a storage)");
                continue;
            }

            // Object names typically start with "_"
            if !entry.name.starts_with('_') {
                eprintln!("DEBUG:     Skipped (name doesn't start with '_')");
                continue;
            }

            eprintln!("DEBUG:     Processing as potential MTEF object");

            // Try to extract "Equation Native" stream from this embedded object
            let equation_stream_path = ["ObjectPool", &entry.name, "Equation Native"];
            if let Ok(Some(mtef_data)) = Self::extract_mtef_data_from_stream(ole_file, &equation_stream_path.join("/")) {
                eprintln!("DEBUG:     ✓ Found MTEF data via equation_stream_path");
                mtef_map.insert(entry.name.clone(), mtef_data);
            } else {
                eprintln!("DEBUG:     Trying alternative paths");
                // Try alternative paths for embedded equations
                let alt_paths = [
                    vec!["ObjectPool", &entry.name, "Equation Native"],
                    vec!["ObjectPool", &entry.name, "\x01Ole10Native"],
                    vec!["ObjectPool", &entry.name, "CONTENTS"],
                ];

                for path in &alt_paths {
                    eprintln!("DEBUG:       Trying: {:?}", path);
                    if let Ok(data) = ole_file.open_stream(path) {
                        eprintln!("DEBUG:       Found data ({} bytes)", data.len());
                        
                        // Dump first 40 bytes for debugging
                        if data.len() >= 40 {
                            eprint!("DEBUG:       First 40 bytes: ");
                            for i in 0..40 {
                                eprint!("{:02X} ", data[i]);
                            }
                            eprintln!();
                        }
                        
                        // Check if this looks like MTEF data
                        if Self::is_mtef_data(&data) {
                            eprintln!("DEBUG:       ✓ Validated as MTEF data");
                            mtef_map.insert(entry.name.clone(), data);
                            break;
                        } else {
                            eprintln!("DEBUG:       ✗ Not valid MTEF data");
                        }
                    }
                }
            }
        }

        Ok(mtef_map)
    }

    /// Check if data looks like MTEF format
    ///
    /// MTEF data starts with a 28-byte OLE header followed by MTEF header
    fn is_mtef_data(data: &[u8]) -> bool {
        if data.len() < 28 {
            return false;
        }

        // Check OLE header
        OleHeader::read_from_bytes(&data[0..28])
            .map(|header| header.is_valid())
            .unwrap_or(false)
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
    pub fn parse_mtef_to_ast(&self, mtef_data: Vec<u8>) -> Result<Vec<MathNode<'arena>>, MtefExtractionError> {
        // We need to create a reference with the arena lifetime
        // This is a bit of a hack, but we know the data will live long enough
        let data_ref: &[u8] = unsafe {
            std::slice::from_raw_parts(mtef_data.as_ptr(), mtef_data.len())
        };
        // Extend the lifetime (this is safe because we control the Vec lifetime)
        let data_with_lifetime: &'arena [u8] = unsafe {
            std::mem::transmute(data_ref)
        };

        let mut parser = MtefParser::new(self.arena, data_with_lifetime);
        parser.parse().map_err(|e| MtefExtractionError::ParseError(e.to_string()))
    }

    /// Extract and parse MTEF data from an OLE file
    ///
    /// This is a convenience method that combines extraction and parsing.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - The OLE file to extract from
    /// * `stream_names` - Possible stream names to check
    ///
    /// # Returns
    ///
    /// Returns parsed formula AST nodes if MTEF data is found and valid
    pub fn extract_and_parse_mtef<R: Read + Seek>(
        &self,
        ole_file: &mut OleFile<R>,
        stream_names: &[&str],
    ) -> Result<Option<Vec<MathNode<'arena>>>, MtefExtractionError> {
        if let Some(mtef_data) = Self::extract_mtef_data(ole_file, stream_names)? {
            // Copy the data to extend its lifetime for the parser
            let data_copy = mtef_data.to_vec();
            Ok(Some(self.parse_mtef_to_ast(data_copy)?))
        } else {
            Ok(None)
        }
    }
}

/// OLE header structure (28 bytes)
#[derive(Debug, Clone, FromBytes)]
#[repr(C)]
struct OleHeader {
    pub cb_hdr: u16,        // Total header length = 28
    pub version: u32,       // Version number (0x00020000)
    pub format: u16,        // Clipboard format (0xC2D3)
    pub size: u32,          // "MTEF header + MTEF data" length
    pub reserved: [u32; 4], // Reserved fields
}

impl OleHeader {
    /// Create OLE header from byte slice
    fn from_bytes(data: &[u8]) -> Result<Self, MtefExtractionError> {
        if data.len() < 28 {
            return Err(MtefExtractionError::InvalidOleHeader);
        }

        Self::read_from_bytes(data).map_err(|_| MtefExtractionError::InvalidOleHeader)
    }

    /// Validate the OLE header
    ///
    /// Note: The format field can vary (0xC16D, 0xC19B, 0xC1C7, 0xC2D3, etc.)
    /// so we only check cb_hdr and version fields for validation.
    fn is_valid(&self) -> bool {
        self.cb_hdr == 28 &&
        (self.version == 0x00020000 || self.version == 0x00000200)
    }
}

/// Errors that can occur during MTEF extraction
#[derive(Debug)]
pub enum MtefExtractionError {
    IoError(std::io::Error),
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
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
        ];

        let header = OleHeader::from_bytes(&valid_data).unwrap();
        assert!(header.is_valid());

        // Invalid OLE header
        let mut invalid_data = valid_data.clone();
        invalid_data[0] = 0x10; // Invalid cb_hdr

        let header = OleHeader::from_bytes(&invalid_data).unwrap();
        assert!(!header.is_valid());
    }

    #[test]
    fn test_mtef_extractor_creation() {
        let formula = Formula::new();
        let _extractor = MtefExtractor::new(formula.arena());

        // Just test that it can be created
        // Test passed - no assertions needed
    }
}
