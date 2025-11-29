//! Shared utilities for parsing IWA files from ZIP archives.
//!
//! This module provides common functionality for reading and parsing IWA files
//! from ZIP archives, avoiding code duplication across the codebase.
//!
//! Uses soapberry-zip for high-performance ZIP reading.

use std::collections::HashMap;
use std::io::Cursor;

use soapberry_zip::office::ArchiveReader;

use crate::iwa::archive::Archive;
use crate::iwa::snappy::SnappyStream;
use crate::iwa::{Error, Result};

/// Parse all IWA files from a ZIP archive and return parsed Archives.
///
/// This function iterates through all files in a ZIP archive, identifies IWA files,
/// decompresses them using Snappy, and parses them into Archive structures.
///
/// # Arguments
///
/// * `archive` - A reference to an ArchiveReader to parse
///
/// # Returns
///
/// * `Result<HashMap<String, Archive>>` - Map of archive names to parsed Archives
///
/// # Examples
///
/// ```rust,no_run
/// use soapberry_zip::office::ArchiveReader;
/// use litchi::iwa::zip_utils::parse_iwa_files_from_archive;
///
/// let data = std::fs::read("document.pages")?;
/// let archive = ArchiveReader::new(&data)?;
/// let archives = parse_iwa_files_from_archive(&archive)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn parse_iwa_files_from_archive(
    archive: &ArchiveReader<'_>,
) -> Result<HashMap<String, Archive>> {
    let mut archives = HashMap::new();

    for name in archive.file_names() {
        if name.ends_with(".iwa") {
            let compressed_data = archive
                .read(name)
                .map_err(|e| Error::Bundle(format!("Failed to read zip entry: {}", e)))?;

            // Decompress IWA file
            let mut cursor = Cursor::new(&compressed_data);
            let decompressed = SnappyStream::decompress(&mut cursor)?;

            // Parse archive
            let iwa_archive = Archive::parse(decompressed.data())?;
            archives.insert(name.to_string(), iwa_archive);
        }
    }

    Ok(archives)
}

/// Extract message types from all IWA files in a ZIP archive.
///
/// This is a lightweight alternative to `parse_iwa_files_from_archive` that only extracts
/// message types without fully parsing the archives. Useful for format detection.
///
/// # Arguments
///
/// * `archive` - A reference to an ArchiveReader to analyze
///
/// # Returns
///
/// * `Result<Vec<u32>>` - List of all message types found in the archive
///
/// # Examples
///
/// ```rust,no_run
/// use soapberry_zip::office::ArchiveReader;
/// use litchi::iwa::zip_utils::extract_message_types_from_archive;
///
/// let data = std::fs::read("document.pages")?;
/// let archive = ArchiveReader::new(&data)?;
/// let message_types = extract_message_types_from_archive(&archive)?;
/// println!("Found {} message types", message_types.len());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn extract_message_types_from_archive(archive: &ArchiveReader<'_>) -> Result<Vec<u32>> {
    let mut all_message_types = Vec::new();

    for name in archive.file_names() {
        if name.ends_with(".iwa") {
            let Ok(compressed_data) = archive.read(name) else {
                continue; // Skip files we can't read
            };

            // Try to decompress and parse the IWA file
            let mut cursor = Cursor::new(&compressed_data);
            let Ok(decompressed) = SnappyStream::decompress(&mut cursor) else {
                continue; // Skip files we can't decompress
            };

            let Ok(iwa_archive) = Archive::parse(decompressed.data()) else {
                continue; // Skip files we can't parse
            };

            // Extract message types from all objects in this archive
            for object in &iwa_archive.objects {
                for message in &object.messages {
                    all_message_types.push(message.type_);
                }
            }
        }
    }

    Ok(all_message_types)
}

/// Analyze file structure patterns in a ZIP archive to aid in format detection.
///
/// This function examines the ZIP archive structure without fully parsing IWA files,
/// looking for characteristic file patterns that indicate specific iWork applications.
///
/// # Arguments
///
/// * `archive` - A reference to an ArchiveReader to analyze
///
/// # Returns
///
/// * `FileStructureInfo` - Summary of file structure patterns found
pub fn analyze_file_structure(archive: &ArchiveReader<'_>) -> FileStructureInfo {
    let mut info = FileStructureInfo::default();

    for name in archive.file_names() {
        // Count table files
        if name.starts_with("Index/Tables/") && name.ends_with(".iwa") {
            info.table_file_count += 1;
        }

        // Check for Numbers-specific patterns
        if name == "Index/CalculationEngine.iwa" {
            info.has_calculation_engine = true;
        }

        // Count presentation files (slides and templates)
        if (name.starts_with("Index/Slide") || name.starts_with("Index/TemplateSlide"))
            && name.ends_with(".iwa")
        {
            info.slide_file_count += 1;
        }

        // Track all file names for debugging
        info.file_names.push(name.to_string());
    }

    info
}

/// Information about file structure patterns in a ZIP archive.
#[derive(Debug, Default, Clone)]
pub struct FileStructureInfo {
    /// Number of table files found (Index/Tables/*.iwa)
    pub table_file_count: usize,
    /// Whether CalculationEngine.iwa exists (Numbers-specific)
    pub has_calculation_engine: bool,
    /// Number of slide files found (Index/Slide*.iwa)
    pub slide_file_count: usize,
    /// All file names in the archive
    pub file_names: Vec<String>,
}

impl FileStructureInfo {
    /// Check if the structure indicates a Keynote document.
    pub fn is_likely_keynote(&self) -> bool {
        // Pure Keynote: has slides but no tables or calc engine
        self.slide_file_count > 0 && self.table_file_count == 0
    }

    /// Check if the structure indicates a Numbers document.
    pub fn is_likely_numbers(&self) -> bool {
        // Has calculation engine and tables, no slides
        self.has_calculation_engine && self.slide_file_count == 0
    }

    /// Check if the structure indicates a Pages document.
    pub fn is_likely_pages(&self) -> bool {
        // Has tables but no calc engine or slides
        self.table_file_count > 0 && !self.has_calculation_engine && self.slide_file_count == 0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_file_structure_detection() {
        // Test Keynote detection
        let keynote_info = FileStructureInfo {
            slide_file_count: 5,
            table_file_count: 0,
            has_calculation_engine: false,
            file_names: vec![],
        };
        assert!(keynote_info.is_likely_keynote());
        assert!(!keynote_info.is_likely_numbers());
        assert!(!keynote_info.is_likely_pages());

        // Test Numbers detection
        let numbers_info = FileStructureInfo {
            slide_file_count: 0,
            table_file_count: 3,
            has_calculation_engine: true,
            file_names: vec![],
        };
        assert!(!numbers_info.is_likely_keynote());
        assert!(numbers_info.is_likely_numbers());
        assert!(!numbers_info.is_likely_pages());

        // Test Pages detection
        let pages_info = FileStructureInfo {
            slide_file_count: 0,
            table_file_count: 2,
            has_calculation_engine: false,
            file_names: vec![],
        };
        assert!(!pages_info.is_likely_keynote());
        assert!(!pages_info.is_likely_numbers());
        assert!(pages_info.is_likely_pages());
    }
}
