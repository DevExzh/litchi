//! Shared utilities for parsing IWA files from ZIP archives.
//!
//! This module provides common functionality for reading and parsing IWA files
//! from ZIP archives, avoiding code duplication across the codebase.
//!
//! Uses soapberry-zip for high-performance ZIP reading.

use std::collections::HashMap;
use std::io::Cursor;

use soapberry_zip::office::ArchiveReader;

use crate::archive::Archive;
use crate::package::is_calculation_engine_entry_name;
use crate::snappy::SnappyStream;
use crate::{Error, Result};

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
/// use litchi_iwa::zip_utils::parse_iwa_files_from_archive;
///
/// let data = std::fs::read("document.pages")?;
/// let archive = ArchiveReader::new(&data)?;
/// let archives = parse_iwa_files_from_archive(&archive)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn parse_iwa_files_from_archive(
    archive: &ArchiveReader<'_>,
) -> Result<HashMap<String, Archive>> {
    if is_encrypted_iwork_archive(archive) {
        return Err(Error::InvalidFormat(
            "password-protected iWork documents are not supported".to_owned(),
        ));
    }

    if archive.file_names().any(|name| name.ends_with(".iwa")) {
        return parse_direct_iwa_files(archive);
    }

    let Some(index_name) = nested_index_zip_name(archive)? else {
        return Ok(HashMap::new());
    };
    let index_data = archive.read(&index_name).map_err(|error| {
        Error::Bundle(format!(
            "Failed to read legacy package index {index_name}: {error}"
        ))
    })?;
    let index = ArchiveReader::new(&index_data).map_err(|error| {
        Error::Bundle(format!(
            "Failed to open legacy package index {index_name}: {error}"
        ))
    })?;
    let archives = parse_direct_iwa_files(&index)?;
    if archives.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }
    Ok(archives)
}

fn parse_direct_iwa_files(archive: &ArchiveReader<'_>) -> Result<HashMap<String, Archive>> {
    let mut archives = HashMap::new();

    for name in archive.file_names() {
        if name.ends_with(".iwa") {
            let compressed_data = archive
                .read(name)
                .map_err(|e| Error::Bundle(format!("Failed to read zip entry: {}", e)))?;

            // OperationStorage is a separate persistence format despite its
            // `.iwa` suffix in some legacy documents (its stream starts with
            // the `bvxn` magic rather than an IWA Snappy frame). It remains a
            // raw package member for round-trip preservation, but is not part
            // of the IWA object graph.
            if name.rsplit('/').next() == Some("OperationStorage.iwa")
                && compressed_data.starts_with(b"bvxn")
            {
                continue;
            }

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

pub(crate) fn is_encrypted_iwork_archive(archive: &ArchiveReader<'_>) -> bool {
    archive
        .file_names()
        .any(|name| matches!(name.rsplit('/').next(), Some(".iwpv2") | Some(".iwph")))
}

pub(crate) fn nested_index_zip_name(archive: &ArchiveReader<'_>) -> Result<Option<String>> {
    let mut candidates = archive
        .file_names()
        .filter(|name| name.rsplit('/').next() == Some("Index.zip"))
        .map(str::to_owned);
    let first = candidates.next();
    if let Some(second) = candidates.next() {
        return Err(Error::InvalidFormat(format!(
            "iWork package contains ambiguous nested indexes: {} and {second}",
            first.as_deref().unwrap_or("Index.zip")
        )));
    }
    Ok(first)
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
/// use litchi_iwa::zip_utils::extract_message_types_from_archive;
///
/// let data = std::fs::read("document.pages")?;
/// let archive = ArchiveReader::new(&data)?;
/// let message_types = extract_message_types_from_archive(&archive)?;
/// println!("Found {} message types", message_types.len());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn extract_message_types_from_archive(archive: &ArchiveReader<'_>) -> Result<Vec<u32>> {
    let mut all_message_types = Vec::new();

    for iwa_archive in parse_iwa_files_from_archive(archive)?.values() {
        for object in &iwa_archive.objects {
            for message in &object.messages {
                all_message_types.push(message.type_);
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
        if is_calculation_engine_entry_name(name) {
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
    /// Whether a canonical or app-versioned CalculationEngine component exists.
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
    use crate::archive::{ArchiveObject, RawMessage};
    use soapberry_zip::office::StreamingArchiveWriter;

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

    #[test]
    fn versioned_calculation_engine_is_detected() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("Index/Tables/Table-1.iwa", b"table")
            .unwrap();
        writer
            .write_stored("Index/CalculationEngine-138.iwa", b"engine")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let archive = ArchiveReader::new(&bytes).unwrap();

        let info = analyze_file_structure(&archive);
        assert!(info.has_calculation_engine);
        assert!(info.is_likely_numbers());
    }

    #[test]
    fn parses_iwa_from_legacy_nested_index() {
        let iwa = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 6000,
                        data: vec![1, 2, 3],
                    }],
                )
                .unwrap(),
            ],
        };
        let compressed = SnappyStream::compress(&iwa.to_bytes().unwrap()).unwrap();
        let mut index = StreamingArchiveWriter::new();
        index
            .write_stored("Index/Document.iwa", &compressed)
            .unwrap();
        let index = index.finish_to_bytes().unwrap();
        let mut outer = StreamingArchiveWriter::new();
        outer
            .write_stored("legacy.pages/Index.zip", &index)
            .unwrap();
        let outer = outer.finish_to_bytes().unwrap();

        let zip = ArchiveReader::new(&outer).unwrap();
        let archives = parse_iwa_files_from_archive(&zip).unwrap();
        assert_eq!(
            archives["Index/Document.iwa"].objects[0].messages[0].type_,
            6000
        );
        assert_eq!(extract_message_types_from_archive(&zip).unwrap(), [6000]);
    }

    #[test]
    fn reports_encrypted_packages_explicitly() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(".iwpv2", b"encryption metadata")
            .unwrap();
        writer
            .write_stored("Index/Document.iwa", b"ciphertext")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let zip = ArchiveReader::new(&bytes).unwrap();

        let error = parse_iwa_files_from_archive(&zip).unwrap_err();
        assert!(error.to_string().contains("password-protected"));
    }

    #[test]
    fn skips_legacy_operation_storage_that_is_not_an_iwa_stream() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("Index/OperationStorage.iwa", b"bvxn payload")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let zip = ArchiveReader::new(&bytes).unwrap();

        assert!(parse_iwa_files_from_archive(&zip).unwrap().is_empty());
    }
}
