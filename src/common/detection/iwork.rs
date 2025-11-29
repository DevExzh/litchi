//! Detection utilities for iWork file formats (Pages, Numbers, Keynote).
//!
//! This module provides functionality to detect and identify iWork file formats,
//! including both modern iWork '13+ formats (.pages, .numbers, .key) and legacy formats.
//!
//! Uses soapberry-zip for high-performance ZIP reading.

use soapberry_zip::office::ArchiveReader;
use std::path::Path;

use super::FileFormat;
use crate::common::Result;
use crate::common::error::Error;

/// Detect iWork format from a ZIP archive.
///
/// This function analyzes the contents of a ZIP archive to determine which iWork
/// application created it (Pages, Numbers, or Keynote).
///
/// # Arguments
///
/// * `archive` - A reference to an ArchiveReader to analyze
///
/// # Returns
///
/// * `Result<FileFormat>` - The detected iWork format flavor, or an error if not recognized
pub fn detect_iwork_format(archive: &ArchiveReader<'_>) -> Result<FileFormat> {
    // First, check for specific marker files that definitively identify each format
    if let Ok(flavor) = detect_from_marker_files(archive) {
        return Ok(flavor);
    }

    // If marker files don't work, fall back to IWA message type analysis
    if let Ok(flavor) = detect_from_iwa_messages(archive) {
        return Ok(flavor);
    }

    // If still can't determine, analyze file structure patterns
    detect_from_file_structure(archive)
}

/// Detect from marker files that definitively identify the iWork application.
fn detect_from_marker_files(archive: &ArchiveReader<'_>) -> Result<FileFormat> {
    // Check for Keynote-specific files
    if archive.contains("Index/MasterSlide.iwa") || archive.contains("Index/Slide.iwa") {
        return Ok(FileFormat::Keynote);
    }

    // Check for Numbers-specific files
    if archive.contains("Index/CalculationEngine.iwa") {
        return Ok(FileFormat::Numbers);
    }

    // Check for document.iwa which is common to Pages
    if archive.contains("Index/Document.iwa") && !archive.contains("Index/CalculationEngine.iwa") {
        // Could be Pages - but need more verification
        // Check for absence of presentation-specific files
        let has_slides = archive.file_names().any(|n| n.contains("Slide"));
        if !has_slides {
            return Ok(FileFormat::Pages);
        }
    }

    Err(Error::InvalidFormat(
        "Could not detect iWork format from marker files".to_string(),
    ))
}

/// Detect from IWA message types.
#[cfg(feature = "iwa")]
fn detect_from_iwa_messages(archive: &ArchiveReader<'_>) -> Result<FileFormat> {
    use crate::iwa::zip_utils::extract_message_types_from_archive;

    let message_types = extract_message_types_from_archive(archive)
        .map_err(|e| Error::InvalidFormat(format!("Failed to extract message types: {}", e)))?;

    if message_types.is_empty() {
        return Err(Error::InvalidFormat(
            "No IWA message types found".to_string(),
        ));
    }

    // Use the registry to detect application
    match crate::iwa::registry::detect_application(&message_types) {
        Some(app) => match app {
            crate::iwa::registry::Application::Pages => Ok(FileFormat::Pages),
            crate::iwa::registry::Application::Keynote => Ok(FileFormat::Keynote),
            crate::iwa::registry::Application::Numbers => Ok(FileFormat::Numbers),
            crate::iwa::registry::Application::Common => Err(Error::InvalidFormat(
                "Only common message types found".to_string(),
            )),
        },
        None => Err(Error::InvalidFormat(
            "Could not detect application from message types".to_string(),
        )),
    }
}

#[cfg(not(feature = "iwa"))]
fn detect_from_iwa_messages(_archive: &ArchiveReader<'_>) -> Result<FileFormat> {
    Err(Error::InvalidFormat("IWA feature not enabled".to_string()))
}

/// Detect from file structure patterns.
#[cfg(feature = "iwa")]
fn detect_from_file_structure(archive: &ArchiveReader<'_>) -> Result<FileFormat> {
    use crate::iwa::zip_utils::analyze_file_structure;

    let info = analyze_file_structure(archive);

    // Apply detection logic based on file patterns
    if info.is_likely_keynote() {
        return Ok(FileFormat::Keynote);
    }

    if info.is_likely_numbers() {
        return Ok(FileFormat::Numbers);
    }

    if info.is_likely_pages() {
        return Ok(FileFormat::Pages);
    }

    Err(Error::InvalidFormat(
        "Could not detect iWork format from file structure".to_string(),
    ))
}

#[cfg(not(feature = "iwa"))]
fn detect_from_file_structure(_archive: &ArchiveReader<'_>) -> Result<FileFormat> {
    Err(Error::InvalidFormat("IWA feature not enabled".to_string()))
}

/// Detect iWork format from bytes.
#[cfg(feature = "iwa")]
pub fn detect_iwork_format_from_bytes(bytes: &[u8]) -> Option<FileFormat> {
    let archive = ArchiveReader::new(bytes).ok()?;

    // Check if archive contains IWA files
    let has_iwa_files = archive.file_names().any(|name| name.ends_with(".iwa"));
    if !has_iwa_files {
        return None;
    }

    detect_iwork_format(&archive).ok()
}

#[cfg(not(feature = "iwa"))]
pub fn detect_iwork_format_from_bytes(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect iWork format from file path.
/// Validates bundle structure following Apple's bundle format standard.
#[cfg(feature = "iwa")]
pub fn detect_iwork_format_from_path<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
    let path = path.as_ref();

    // First check if it's a directory (required for bundles)
    if !path.is_dir() {
        // It's a file, try ZIP-based detection
        if let Ok(data) = std::fs::read(path) {
            return detect_iwork_format_from_bytes(&data);
        }
        return None;
    }

    // Try direct file-based detection first (faster and more reliable)
    if let Some(format) = detect_iwork_format_from_bundle_files(path) {
        return Some(format);
    }

    // Fallback to IWA bundle parsing if direct detection fails
    match crate::iwa::bundle::Bundle::open(path) {
        Ok(bundle) => {
            // Extract message types from all archives to identify application
            let all_message_types: Vec<u32> = bundle
                .archives()
                .values()
                .flat_map(|archive| &archive.objects)
                .flat_map(|obj| &obj.messages)
                .map(|msg| msg.type_)
                .collect();

            // Detect application type from IWA message types
            match crate::iwa::registry::detect_application(&all_message_types) {
                Some(app) => match app {
                    crate::iwa::registry::Application::Pages => Some(FileFormat::Pages),
                    crate::iwa::registry::Application::Keynote => Some(FileFormat::Keynote),
                    crate::iwa::registry::Application::Numbers => Some(FileFormat::Numbers),
                    crate::iwa::registry::Application::Common => None,
                },
                None => None,
            }
        },
        Err(_) => None,
    }
}

#[cfg(not(feature = "iwa"))]
pub fn detect_iwork_format_from_path<P: AsRef<Path>>(_path: P) -> Option<FileFormat> {
    None
}

/// Detect iWork format by checking for identifying files in the bundle directory.
#[cfg(feature = "iwa")]
fn detect_iwork_format_from_bundle_files(bundle_path: &Path) -> Option<FileFormat> {
    // Check for Keynote: contains index.apxl file
    let keynote_index = bundle_path.join("index.apxl");
    if keynote_index.exists() && keynote_index.is_file() {
        return Some(FileFormat::Keynote);
    }

    // Check for Pages: contains index.xml file
    let pages_index = bundle_path.join("index.xml");
    if pages_index.exists() && pages_index.is_file() {
        return Some(FileFormat::Pages);
    }

    // Check for Numbers: contains index.numbers file
    let numbers_index = bundle_path.join("index.numbers");
    if numbers_index.exists() && numbers_index.is_file() {
        return Some(FileFormat::Numbers);
    }

    // Check for Keynote presentation files
    let keynote_data = bundle_path.join("Data");
    if keynote_data.exists() && keynote_data.is_dir() {
        let keynote_theme = bundle_path.join("theme-preview.jpg");
        let keynote_assets = bundle_path.join("Assets");
        if keynote_theme.exists() || (keynote_assets.exists() && keynote_assets.is_dir()) {
            return Some(FileFormat::Keynote);
        }
    }

    // Check for Numbers-specific structure
    let numbers_calc = bundle_path.join("Index");
    if numbers_calc.exists() && numbers_calc.is_dir() {
        let calc_files = std::fs::read_dir(&numbers_calc).ok()?;
        for entry in calc_files.flatten() {
            if entry.path().extension().is_some_and(|ext| ext == "iwa") {
                return Some(FileFormat::Numbers);
            }
        }
    }

    None
}
