//! Shared utilities for parsing IWA files from ZIP archives.
//!
//! This module provides common functionality for reading and parsing IWA files
//! from ZIP archives, avoiding code duplication across the codebase.

use std::collections::HashMap;
use std::io::{Cursor, Read};

use zip::ZipArchive;

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
/// * `zip_archive` - A mutable reference to a ZipArchive to parse
///
/// # Returns
///
/// * `Result<HashMap<String, Archive>>` - Map of archive names to parsed Archives
///
/// # Examples
///
/// ```rust,no_run
/// use std::fs::File;
/// use zip::ZipArchive;
/// use litchi::iwa::zip_utils::parse_iwa_files_from_zip;
///
/// let file = File::open("Index.zip")?;
/// let mut archive = ZipArchive::new(file)?;
/// let archives = parse_iwa_files_from_zip(&mut archive)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn parse_iwa_files_from_zip<R: Read + std::io::Seek>(
    zip_archive: &mut ZipArchive<R>,
) -> Result<HashMap<String, Archive>> {
    let mut archives = HashMap::new();

    for i in 0..zip_archive.len() {
        let mut zip_file = zip_archive
            .by_index(i)
            .map_err(|e| Error::Bundle(format!("Failed to read zip entry: {}", e)))?;

        if zip_file.name().ends_with(".iwa") {
            let mut compressed_data = Vec::new();
            zip_file
                .read_to_end(&mut compressed_data)
                .map_err(Error::Io)?;

            // Decompress IWA file
            let mut cursor = Cursor::new(&compressed_data);
            let decompressed = SnappyStream::decompress(&mut cursor)?;

            // Parse archive
            let archive = Archive::parse(decompressed.data())?;
            let name = zip_file.name().to_string();
            archives.insert(name, archive);
        }
    }

    Ok(archives)
}

/// Extract message types from all IWA files in a ZIP archive.
///
/// This is a lightweight alternative to `parse_iwa_files_from_zip` that only extracts
/// message types without fully parsing the archives. Useful for format detection.
///
/// # Arguments
///
/// * `zip_archive` - A mutable reference to a ZipArchive to analyze
///
/// # Returns
///
/// * `Result<Vec<u32>>` - List of all message types found in the archive
///
/// # Examples
///
/// ```rust,no_run
/// use std::fs::File;
/// use zip::ZipArchive;
/// use litchi::iwa::zip_utils::extract_message_types_from_zip;
///
/// let file = File::open("document.pages")?;
/// let mut archive = ZipArchive::new(file)?;
/// let message_types = extract_message_types_from_zip(&mut archive)?;
/// println!("Found {} message types", message_types.len());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn extract_message_types_from_zip<R: Read + std::io::Seek>(
    zip_archive: &mut ZipArchive<R>,
) -> Result<Vec<u32>> {
    let mut all_message_types = Vec::new();

    for i in 0..zip_archive.len() {
        let mut zip_file = zip_archive
            .by_index(i)
            .map_err(|e| Error::Bundle(format!("Failed to read zip entry: {}", e)))?;

        if zip_file.name().ends_with(".iwa") {
            let mut compressed_data = Vec::new();
            if zip_file.read_to_end(&mut compressed_data).is_err() {
                continue; // Skip files we can't read
            }

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
/// * `zip_archive` - A mutable reference to a ZipArchive to analyze
///
/// # Returns
///
/// * `FileStructureInfo` - Summary of file structure patterns found
pub fn analyze_file_structure<R: Read + std::io::Seek>(
    zip_archive: &mut ZipArchive<R>,
) -> FileStructureInfo {
    let mut info = FileStructureInfo::default();

    for i in 0..zip_archive.len() {
        if let Ok(zip_file) = zip_archive.by_index(i) {
            let name = zip_file.name();

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
