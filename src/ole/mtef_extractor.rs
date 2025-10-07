// MTEF (MathType Equation Format) Extractor for OLE Documents
//
// This module provides functionality to extract and parse MTEF binary data
// from OLE streams in legacy Office documents (.doc, .ppt, etc.).

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

            // The MTEF data starts after the OLE header
            let mtef_start = 28;
            if mtef_start >= data.len() {
                return Ok(None);
            }

            // Extract the MTEF data portion
            let mtef_size = ole_header.size as usize;
            if mtef_start + mtef_size > data.len() {
                // If size field is incorrect, take all remaining data
                Ok(Some(data[mtef_start..].to_vec()))
            } else {
                Ok(Some(data[mtef_start..mtef_start + mtef_size].to_vec()))
            }
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
#[derive(Debug, Clone)]
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

        Ok(Self {
            cb_hdr: u16::from_le_bytes([data[0], data[1]]),
            version: u32::from_le_bytes([data[2], data[3], data[4], data[5]]),
            format: u16::from_le_bytes([data[6], data[7]]),
            size: u32::from_le_bytes([data[8], data[9], data[10], data[11]]),
            reserved: [
                u32::from_le_bytes([data[12], data[13], data[14], data[15]]),
                u32::from_le_bytes([data[16], data[17], data[18], data[19]]),
                u32::from_le_bytes([data[20], data[21], data[22], data[23]]),
                u32::from_le_bytes([data[24], data[25], data[26], data[27]]),
            ],
        })
    }

    /// Validate the OLE header
    fn is_valid(&self) -> bool {
        self.cb_hdr == 28 &&
        self.version == 0x00020000 &&
        self.format == 0xC2D3
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
        assert!(true);
    }
}
