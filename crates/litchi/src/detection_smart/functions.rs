//! Core file format detection functions.
//!
//! Uses SIMD-accelerated signature matching for high-performance format detection,
//! including parallel signature checking for maximum throughput.

use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;

use super::{iwork, ole2, ooxml};
use litchi_core::detection::FileFormat;
use litchi_core::detection::simd_utils::{check_office_signatures, signature_matches};
use litchi_core::detection::{rtf, utils};

#[cfg(feature = "odf")]
use litchi_core::detection::odf;

/// Detect file format from a file path.
///
/// This function opens the file and reads only the necessary bytes
/// to determine the format, making it very efficient.
///
/// # Arguments
///
/// * `path` - Path to the file to analyze
///
/// # Returns
///
/// * `Some(FileFormat)` if a supported format is detected
/// * `None` if the format is not recognized or file cannot be read
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::common::detection::detect_file_format;
///
/// let format = detect_file_format("document.docx");
/// if let Some(format) = format {
///     println!("Detected format: {:?}", format);
/// }
/// # Ok::<(), std::io::Error>(())
/// ```
pub fn detect_file_format<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
    let mut file = File::open(path).ok()?;
    detect_format_from_reader(&mut file)
}

/// Detect file format from a byte slice.
///
/// This function analyzes the byte signature in memory without
/// requiring file I/O, making it ideal for network data or
/// in-memory processing.
///
/// Uses SIMD-accelerated parallel signature checking for maximum performance
/// when processing multiple files or large batches.
///
/// # Arguments
///
/// * `bytes` - The file data as bytes
///
/// # Returns
///
/// * `Some(FileFormat)` if a supported format is detected
/// * `None` if the format is not recognized
///
/// # Performance
///
/// Optimized using SIMD operations for fast signature matching:
/// - Checks OLE2, ZIP, and RTF signatures **in parallel** (3-6x faster)
/// - Provides up to 6-32x speedup per signature depending on CPU capabilities
/// - Overall detection is 3-10x faster than sequential checking
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::common::detection::detect_file_format_from_bytes;
/// use std::fs;
///
/// let data = fs::read("document.docx")?;
/// let format = detect_file_format_from_bytes(&data);
/// if let Some(format) = format {
///     println!("Detected format: {:?}", format);
/// }
/// # Ok::<(), std::io::Error>(())
/// ```
pub fn detect_file_format_from_bytes(bytes: &[u8]) -> Option<FileFormat> {
    if bytes.len() < 4 {
        return None;
    }

    #[cfg(feature = "odf")]
    if let Some(result) = odf::detect_flat_odf_format(bytes) {
        return Some(result);
    }

    // Use parallel signature checking to test OLE2, ZIP, and RTF simultaneously
    // This is 3-6x faster than checking each signature individually
    let mask = check_office_signatures(bytes);

    // Check OLE2 first (if matched)
    if mask.is_ole2()
        && let Some(result) = ole2::detect_ole2_format(bytes)
    {
        return Some(result);
    }

    // Check ZIP-based formats (if matched)
    if mask.is_zip() {
        // First try OOXML detection (most common)
        if let Some(result) = ooxml::detect_zip_format(bytes) {
            return Some(result);
        }

        // Then try ODF detection
        #[cfg(feature = "odf")]
        if let Some(result) = odf::detect_odf_format(bytes) {
            return Some(result);
        }

        // Finally try iWork detection
        if let Some(result) = iwork::detect_iwork_format_from_bytes(bytes) {
            return Some(result);
        }

        return None;
    }

    // Check RTF format (if matched)
    if mask.is_rtf()
        && let Some(result) = rtf::detect_rtf_format(bytes)
    {
        return Some(result);
    }

    // Check for iWork bundle formats (non-ZIP based)
    if let Some(result) = iwork::detect_iwork_format_from_bytes(bytes) {
        return Some(result);
    }

    None
}

/// Detect file format from any reader that implements Read + Seek.
///
/// This is the core detection function used by both file path and
/// byte slice detection methods.
///
/// Uses SIMD-accelerated signature matching for improved performance.
///
/// # Arguments
///
/// * `reader` - A reader that can read and seek
///
/// # Returns
///
/// * `Some(FileFormat)` if a supported format is detected
/// * `None` if the format is not recognized
///
/// # Performance
///
/// Optimized using SIMD operations for fast signature matching,
/// especially beneficial when processing multiple files.
pub fn detect_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    // Read the first 8 bytes for signature detection
    let mut header = [0u8; 8];
    if reader.read_exact(&mut header).is_err() {
        return None;
    }

    // Reset to beginning
    let _ = reader.seek(std::io::SeekFrom::Start(0));

    // Check for OLE2 signature (legacy Office formats) using SIMD
    if signature_matches(&header[0..8], utils::OLE2_SIGNATURE) {
        return ole2::detect_ole2_format_from_reader(reader);
    }

    // Check for ZIP signature (OOXML, ODF, and iWork formats) using SIMD
    #[cfg(any(feature = "ooxml", feature = "odf", feature = "iwa"))]
    if signature_matches(&header[0..4], utils::ZIP_SIGNATURE) {
        // For ZIP files, first check if it contains IWA files (iWork indicator)
        #[cfg(feature = "iwa")]
        {
            let _ = reader.seek(std::io::SeekFrom::Start(0));
            // Read all data for ArchiveReader
            let mut data = Vec::new();
            if reader.read_to_end(&mut data).is_ok()
                && let Ok(archive) = soapberry_zip::office::ArchiveReader::new(&data)
            {
                let has_iwa_files = archive.file_names().any(|name| name.ends_with(".iwa"));

                if has_iwa_files {
                    // This is an iWork file, detect the specific type
                    if let Ok(result) = iwork::detect_iwork_format(&archive) {
                        return Some(result);
                    }
                    return None;
                }
            }
        }

        // Reset to beginning for OOXML detection
        let _ = reader.seek(std::io::SeekFrom::Start(0));

        // Try OOXML
        #[cfg(feature = "ooxml")]
        if let Some(result) = ooxml::detect_zip_format_from_reader(reader) {
            return Some(result);
        }

        // Reset to beginning for ODF detection
        let _ = reader.seek(std::io::SeekFrom::Start(0));

        // Try ODF
        #[cfg(feature = "odf")]
        if let Some(result) = odf::detect_odf_format_from_reader(reader) {
            return Some(result);
        }

        return None;
    }

    #[cfg(feature = "odf")]
    if header
        .iter()
        .copied()
        .find(|byte| !byte.is_ascii_whitespace())
        == Some(b'<')
        || header.starts_with(&[0xef, 0xbb, 0xbf])
    {
        let _ = reader.seek(std::io::SeekFrom::Start(0));
        if let Some(result) = odf::detect_odf_format_from_reader(reader) {
            return Some(result);
        }
    }

    // Check for RTF format (plain text with {\rtf signature)
    if let Some(result) = rtf::detect_rtf_format_from_reader(reader) {
        return Some(result);
    }

    // Note: iWork format detection from reader is handled via byte-based detection
    // Use detect_iwork_format_from_bytes() for iWork detection

    None
}

/// Detect iWork format from file path.
pub fn detect_iwork_format_from_path<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
    iwork::detect_iwork_format_from_path(path)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[cfg(feature = "odf")]
    fn detects_flat_odf_bytes_and_reader() {
        let xml = br#"<?xml version="1.0"?><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.chart"><o:body><o:chart/></o:body></o:document>"#;
        assert_eq!(detect_file_format_from_bytes(xml), Some(FileFormat::Odc));
        let mut reader = std::io::Cursor::new(xml);
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Odc)
        );
    }

    #[test]
    #[cfg(feature = "ooxml")]
    fn test_detect_docx_from_bytes() {
        // Create a minimal ZIP file that looks like a DOCX
        let zip_data = create_minimal_docx_zip();
        let format = detect_file_format_from_bytes(&zip_data);
        assert!(format.is_some());
        assert_eq!(format.unwrap(), FileFormat::Docx);
    }

    #[test]
    #[cfg(feature = "ooxml")]
    fn detects_xml_binary_and_macro_enabled_ooxml_families() {
        let cases = [
            (
                "word/document.xml",
                "application/vnd.ms-word.document.macroEnabled.main+xml",
                FileFormat::Docx,
            ),
            (
                "ppt/presentation.xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                FileFormat::Pptx,
            ),
            (
                "ppt/presentation.xml",
                "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
                FileFormat::Pptx,
            ),
            (
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                FileFormat::Xlsx,
            ),
            (
                "xl/workbook.xml",
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                FileFormat::Xlsx,
            ),
            (
                "xl/workbook.bin",
                "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
                FileFormat::Xlsb,
            ),
        ];

        for (part_name, content_type, expected) in cases {
            let package = create_minimal_ooxml_zip(part_name, content_type);
            assert_eq!(detect_file_format_from_bytes(&package), Some(expected));
        }
    }

    #[test]
    fn test_detect_ole2_from_bytes() {
        // Just having the OLE2 signature is not enough - we need a valid OLE file
        // This test verifies that a file with OLE2 signature that isn't a valid
        // OLE file returns None (as expected)
        let ole2_data = utils::OLE2_SIGNATURE.to_vec();
        let format = detect_file_format_from_bytes(&ole2_data);
        // Should return None because it's not a complete OLE file
        assert!(format.is_none());
    }

    // Helper function to create a minimal DOCX-like ZIP for testing
    #[cfg(feature = "ooxml")]
    fn create_minimal_docx_zip() -> Vec<u8> {
        create_minimal_ooxml_zip(
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        )
    }

    #[cfg(feature = "ooxml")]
    fn create_minimal_ooxml_zip(part_name: &str, content_type: &str) -> Vec<u8> {
        use soapberry_zip::office::StreamingArchiveWriter;

        let mut writer = StreamingArchiveWriter::new();
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/{part_name}" ContentType="{content_type}"/></Types>"#
        );

        // Add [Content_Types].xml
        writer
            .write_deflated("[Content_Types].xml", content_types.as_bytes())
            .unwrap();

        let relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{part_name}"/></Relationships>"#
        );
        writer
            .write_deflated("_rels/.rels", relationships.as_bytes())
            .unwrap();

        writer.write_deflated(part_name, b"office data").unwrap();

        writer.finish_to_bytes().unwrap()
    }
}
