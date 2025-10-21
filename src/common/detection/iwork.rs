//! iWork format detection (Pages, Keynote, Numbers).
//!
//! iWork files follow Apple's bundle format standard where each document
//! is stored as a directory containing multiple files including Index.zip.
//! Detection follows the iWork Archive (IWA) format specification.

use crate::common::detection::FileFormat;
#[cfg(feature = "iwa")]
use std::io::{Cursor, Read};
use std::path::Path;

/// Detect iWork formats from bytes.
/// iWork files can be ZIP archives containing IWA files.
/// Detection analyzes the ZIP structure and IWA message types.
#[cfg(feature = "iwa")]
pub fn detect_iwork_format(bytes: &[u8]) -> Option<FileFormat> {
    // Check if it starts with ZIP signature
    if bytes.len() < 4 || &bytes[0..4] != crate::common::detection::utils::ZIP_SIGNATURE {
        return None;
    }

    // Try to open as ZIP archive and check for IWA files
    let cursor = std::io::Cursor::new(bytes);
    match zip::ZipArchive::new(cursor) {
        Ok(mut archive) => {
            // Check if archive contains IWA files
            let has_iwa_files = (0..archive.len()).any(|i| {
                archive
                    .by_index(i)
                    .ok()
                    .map(|file| file.name().ends_with(".iwa"))
                    .unwrap_or(false)
            });

            if !has_iwa_files {
                return None;
            }

            // Try to extract message types and detect application
            detect_application_from_zip_archive(&mut archive)
        },
        Err(_) => None,
    }
}

#[cfg(not(feature = "iwa"))]
pub fn detect_iwork_format(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect iWork formats from reader.
/// iWork files are ZIP archives that can be detected from stream data.
#[cfg(feature = "iwa")]
pub fn detect_iwork_format_from_reader<R: std::io::Read + std::io::Seek>(
    reader: &mut R,
) -> Option<FileFormat> {
    // Read the first 4 bytes to check for ZIP signature
    let mut header = [0u8; 4];
    if reader.read_exact(&mut header).is_err() {
        return None;
    }

    // Reset to beginning
    let _ = reader.seek(std::io::SeekFrom::Start(0));

    // Check for ZIP signature
    if &header[0..4] == crate::common::detection::utils::ZIP_SIGNATURE {
        // Try to open as ZIP archive and check for IWA files
        if let Ok(mut archive) = zip::ZipArchive::new(reader) {
            // Check if archive contains IWA files
            let has_iwa_files = (0..archive.len()).any(|i| {
                archive
                    .by_index(i)
                    .ok()
                    .map(|file| file.name().ends_with(".iwa"))
                    .unwrap_or(false)
            });

            if has_iwa_files {
                // Try to extract message types and detect application
                return detect_application_from_zip_archive(&mut archive);
            }
        }
    }

    None
}

#[cfg(not(feature = "iwa"))]
pub fn detect_iwork_format_from_reader<R: std::io::Read + std::io::Seek>(
    _reader: &mut R,
) -> Option<FileFormat> {
    None
}

/// Detect iWork format from file path.
/// Validates bundle structure following Apple's bundle format standard.
/// Checks for identifying files and uses IWA message types to identify application.
pub fn detect_iwork_format_from_path<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
    let path = path.as_ref();

    // First check if it's a directory (required for bundles)
    if !path.is_dir() {
        // It's a file, try ZIP-based detection
        if let Ok(data) = std::fs::read(path) {
            return detect_iwork_format(&data);
        }
        return None;
    }

    // Try direct file-based detection first (faster and more reliable)
    if let Some(format) = detect_iwork_format_from_files(path) {
        return Some(format);
    }

    // Fallback to IWA bundle parsing if direct detection fails
    // Validate bundle structure by attempting to open with Bundle::open()
    // This follows the iWork Archive format specification
    #[cfg(feature = "iwa")]
    {
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
                    Some(app) => {
                        let format = match app {
                            crate::iwa::registry::Application::Pages => FileFormat::Pages,
                            crate::iwa::registry::Application::Keynote => FileFormat::Keynote,
                            crate::iwa::registry::Application::Numbers => FileFormat::Numbers,
                            crate::iwa::registry::Application::Common => return None, // Common alone doesn't indicate a specific format
                        };

                        Some(format)
                    },
                    None => None,
                }
            },
            Err(_) => None,
        }
    }

    #[cfg(not(feature = "iwa"))]
    None
}

/// Detect iWork format by checking for identifying files in the bundle directory.
/// Uses Apple's iWork bundle format standards to identify application types.
fn detect_iwork_format_from_files(bundle_path: &Path) -> Option<FileFormat> {
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

    // Additional checks for other common iWork identifying files

    // Check for Keynote presentation files
    let keynote_data = bundle_path.join("Data");
    if keynote_data.exists() && keynote_data.is_dir() {
        // Look for Keynote-specific files
        let keynote_theme = bundle_path.join("theme-preview.jpg");
        let keynote_assets = bundle_path.join("Assets");
        if keynote_theme.exists() || (keynote_assets.exists() && keynote_assets.is_dir()) {
            return Some(FileFormat::Keynote);
        }
    }

    // Check for Numbers-specific structure
    let numbers_calc = bundle_path.join("Index");
    if numbers_calc.exists() && numbers_calc.is_dir() {
        // Look for Numbers calculation files
        let calc_files = std::fs::read_dir(&numbers_calc).ok()?;
        for entry in calc_files.flatten() {
            if entry.path().extension().is_some_and(|ext| ext == "iwa") {
                return Some(FileFormat::Numbers);
            }
        }
    }

    None
}

/// Detect application type from a ZIP archive containing IWA files
#[cfg(feature = "iwa")]
pub fn detect_application_from_zip_archive<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Option<FileFormat> {
    let mut table_file_count = 0;
    let mut has_calculation_engine = false;
    let mut slide_file_count = 0;
    let mut file_names = Vec::new();

    // Check file structure patterns
    for i in 0..archive.len() {
        if let Ok(zip_file) = archive.by_index(i) {
            let name = zip_file.name();
            file_names.push(name.to_string());

            // Count table files
            if name.starts_with("Index/Tables/") && name.ends_with(".iwa") {
                table_file_count += 1;
            }

            // Check for Numbers-specific patterns
            if name == "Index/CalculationEngine.iwa" {
                has_calculation_engine = true;
            }

            // Count presentation files (slides and templates)
            if (name.starts_with("Index/Slide") || name.starts_with("Index/TemplateSlide"))
                && name.ends_with(".iwa")
            {
                slide_file_count += 1;
            }
        }
    }

    // Debug: print summary for complex cases
    // println!("DEBUG: Archive files: {} total, {} table files, {} slide files, calc_engine={}",
    //          file_names.len(), table_file_count, slide_file_count, has_calculation_engine);

    // Apply detection logic based on file patterns with priorities
    if slide_file_count > 0 && table_file_count == 0 {
        // Pure Keynote: has slides but no tables
        Some(FileFormat::Keynote)
    } else if slide_file_count > 0 && has_calculation_engine {
        // Document with slides and calc engine: Keynote (prioritize presentation aspect)
        Some(FileFormat::Keynote)
    } else if table_file_count > 0 && slide_file_count == 0 {
        // Pure document with tables: could be Pages or Numbers
        if has_calculation_engine {
            // Has calc engine: Numbers
            Some(FileFormat::Numbers)
        } else {
            // No calc engine: Pages
            Some(FileFormat::Pages)
        }
    } else if slide_file_count > 0 {
        // Has slides: Keynote
        Some(FileFormat::Keynote)
    } else if has_calculation_engine {
        // Has calculation engine: Numbers
        Some(FileFormat::Numbers)
    } else {
        // Fallback: try message type detection
        detect_application_from_message_types(archive)
    }
}

/// Fallback detection using message types when file patterns don't give a clear answer
#[cfg(feature = "iwa")]
fn detect_application_from_message_types<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Option<FileFormat> {
    let mut all_message_types = Vec::new();

    // Process each IWA file in the archive
    for i in 0..archive.len() {
        if let Ok(mut zip_file) = archive.by_index(i)
            && zip_file.name().ends_with(".iwa")
        {
            // Read the compressed IWA data
            let mut compressed_data = Vec::new();
            if zip_file.read_to_end(&mut compressed_data).is_err() {
                continue; // Skip files we can't read
            }

            // Try to decompress and parse the IWA file
            let Ok(decompressed) =
                crate::iwa::snappy::SnappyStream::decompress(&mut Cursor::new(&compressed_data))
            else {
                continue; // Skip files we can't read
            };
            let Ok(iwa_archive) = crate::iwa::archive::Archive::parse(decompressed.data()) else {
                continue; // Skip files we can't read
            };
            // Extract message types from all objects in this archive
            for object in &iwa_archive.objects {
                for message in &object.messages {
                    all_message_types.push(message.type_);
                }
            }
        }
    }

    // Simple heuristic: look for known message type ranges
    // This is a rough approximation based on observed patterns
    let mut pages_score = 0;
    let mut keynote_score = 0;
    let mut numbers_score = 0;

    for &type_id in &all_message_types {
        match type_id {
            10000..=19999 => pages_score += 1, // Pages tends to have higher type IDs
            200..=999 => keynote_score += 1,   // Keynote has mid-range IDs
            1..=199 => numbers_score += 1,     // Numbers has lower IDs
            _ => {},
        }
    }

    // Return the application with the highest score
    if pages_score > keynote_score && pages_score > numbers_score {
        Some(FileFormat::Pages)
    } else if keynote_score > pages_score && keynote_score > numbers_score {
        Some(FileFormat::Keynote)
    } else if numbers_score > pages_score && numbers_score > keynote_score {
        Some(FileFormat::Numbers)
    } else {
        None
    }
}
