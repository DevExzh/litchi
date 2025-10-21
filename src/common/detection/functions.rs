//! Core file format detection functions.

use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;

use super::types::FileFormat;
use super::{iwork, ole2, ooxml, utils};

#[cfg(feature = "odf")]
use super::odf;

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
/// # Arguments
///
/// * `bytes` - The file data as bytes
///
/// # Returns
///
/// * `Some(FileFormat)` if a supported format is detected
/// * `None` if the format is not recognized
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

    // Check for OLE2 signature (legacy Office formats)
    if bytes.len() >= 8 && &bytes[0..8] == utils::OLE2_SIGNATURE {
        return ole2::detect_ole2_format(bytes);
    }

    // Check for ZIP signature (OOXML and ODF formats)
    if &bytes[0..4] == utils::ZIP_SIGNATURE {
        return ooxml::detect_zip_format(bytes);
    }

    // Check for iWork bundle formats
    if let Some(result) = iwork::detect_iwork_format(bytes) {
        return Some(result);
    }

    None
}

/// Detect file format from any reader that implements Read + Seek.
///
/// This is the core detection function used by both file path and
/// byte slice detection methods.
///
/// # Arguments
///
/// * `reader` - A reader that can read and seek
///
/// # Returns
///
/// * `Some(FileFormat)` if a supported format is detected
/// * `None` if the format is not recognized
pub fn detect_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    // Read the first 8 bytes for signature detection
    let mut header = [0u8; 8];
    if reader.read_exact(&mut header).is_err() {
        return None;
    }

    // Reset to beginning
    let _ = reader.seek(std::io::SeekFrom::Start(0));

    // Check for OLE2 signature (legacy Office formats)
    if &header[0..8] == utils::OLE2_SIGNATURE {
        return ole2::detect_ole2_format_from_reader(reader);
    }

    // Check for ZIP signature (OOXML, ODF, and iWork formats)
    #[cfg(any(feature = "ooxml", feature = "odf", feature = "iwa"))]
    if &header[0..4] == utils::ZIP_SIGNATURE {
        // For ZIP files, first check if it contains IWA files (iWork indicator)
        #[cfg(feature = "iwa")]
        {
            let _ = reader.seek(std::io::SeekFrom::Start(0));
            if let Ok(mut zip_archive) = zip::ZipArchive::new(&mut *reader) {
                let has_iwa_files = (0..zip_archive.len()).any(|i| {
                    zip_archive
                        .by_index(i)
                        .ok()
                        .map(|file| file.name().ends_with(".iwa"))
                        .unwrap_or(false)
                });

                if has_iwa_files {
                    // This is an iWork file, detect the specific type
                    if let Some(result) =
                        iwork::detect_application_from_zip_archive(&mut zip_archive)
                    {
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

    // Check for iWork bundle formats
    if let Some(result) = iwork::detect_iwork_format_from_reader(reader) {
        return Some(result);
    }

    None
}

/// Detect iWork format from file path.
pub fn detect_iwork_format_from_path<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
    iwork::detect_iwork_format_from_path(path)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_detect_docx_from_bytes() {
        // Create a minimal ZIP file that looks like a DOCX
        let zip_data = create_minimal_docx_zip();
        let format = detect_file_format_from_bytes(&zip_data);
        assert!(format.is_some());
        assert_eq!(format.unwrap(), FileFormat::Docx);
    }

    #[test]
    fn test_detect_ole2_from_bytes() {
        let ole2_data = utils::OLE2_SIGNATURE.to_vec();
        let format = detect_file_format_from_bytes(&ole2_data);
        assert!(format.is_some());
        // Should detect as some OLE2 format
    }

    #[test]
    fn test_detect_iwork_pages() {
        // This would require creating a mock iWork bundle structure
        // For now, test the extension-based detection logic
        let mock_path = std::path::Path::new("test.pages");
        let format = detect_iwork_format_from_path(mock_path);
        // Will return None since the file doesn't exist
        assert!(format.is_none());
    }

    // Helper function to create a minimal DOCX-like ZIP for testing
    fn create_minimal_docx_zip() -> Vec<u8> {
        use std::io::Write;

        let mut buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buffer));

            let options = zip::write::SimpleFileOptions::default();

            // Add [Content_Types].xml
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(
                b"<Types><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>",
            )
            .unwrap();

            // Add word/document.xml
            zip.start_file("word/document.xml", options).unwrap();
            zip.write_all(b"<document><body><p>Hello</p></body></document>")
                .unwrap();

            zip.finish().unwrap();
        }

        buffer
    }
}
