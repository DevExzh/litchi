//! Core file format detection functions.
//!
//! Uses a combined fixed-signature check before invoking format-specific
//! detectors owned by their leaf crates.

use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

use super::{ole2, ooxml};
use litchi_core::detection::FileFormat;
use litchi_core::detection::simd_utils::{check_office_signatures, signature_matches};
use litchi_core::detection::{rtf, utils};

#[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
use litchi_odf_common::detect as odf;

/// Detect file format from a file path.
///
/// This function opens the file and delegates to the enabled format owners.
/// Container formats may require a complete package scan.
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

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    if let Some(result) = odf::flat(bytes) {
        return Some(result);
    }

    // Classify the fixed signatures together before invoking format parsers.
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
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
        if let Some(result) = odf::bytes(bytes) {
            return Some(result);
        }

        // Finally try iWork detection
        #[cfg(feature = "iwork")]
        if let Ok(Some(result)) =
            litchi_iwa::detect::bytes(bytes).map(|result| result.map(iwork_format))
        {
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

    None
}

/// Detect file format from any reader that implements Read + Seek.
///
/// Detection always starts at the beginning of the stream. The caller's
/// original cursor position is restored before returning. A failed restore is
/// reported as `None` because the cursor contract could not be upheld.
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
pub fn detect_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    let original = reader.stream_position().ok()?;
    let detected = (|| {
        reader.seek(SeekFrom::Start(0)).ok()?;
        let mut header = [0u8; 8];
        let mut header_len = 0;
        while header_len < header.len() {
            let read = reader.read(&mut header[header_len..]).ok()?;
            if read == 0 {
                break;
            }
            header_len += read;
        }

        if header_len >= utils::OLE2_SIGNATURE.len()
            && signature_matches(&header[..header_len], utils::OLE2_SIGNATURE)
        {
            reader.seek(SeekFrom::Start(0)).ok()?;
            return ole2::detect_ole2_format_from_reader(reader);
        }

        #[cfg(any(
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
            any(feature = "odt", feature = "ods", feature = "odp"),
            feature = "iwork"
        ))]
        if header_len >= utils::ZIP_SIGNATURE.len()
            && signature_matches(&header[..header_len], utils::ZIP_SIGNATURE)
        {
            #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
            {
                reader.seek(SeekFrom::Start(0)).ok()?;
                if let Some(format) = ooxml::detect_zip_format_from_reader(reader) {
                    return Some(format);
                }
            }

            #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
            {
                reader.seek(SeekFrom::Start(0)).ok()?;
                if let Some(format) = odf::reader(reader) {
                    return Some(format);
                }
            }

            #[cfg(feature = "iwork")]
            {
                reader.seek(SeekFrom::Start(0)).ok()?;
                if let Ok(Some(format)) =
                    litchi_iwa::detect::reader(reader).map(|result| result.map(iwork_format))
                {
                    return Some(format);
                }
            }

            return None;
        }

        #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
        if header[..header_len]
            .iter()
            .copied()
            .find(|byte| !byte.is_ascii_whitespace())
            == Some(b'<')
            || header[..header_len].starts_with(&[0xef, 0xbb, 0xbf])
        {
            reader.seek(SeekFrom::Start(0)).ok()?;
            if let Some(format) = odf::reader(reader) {
                return Some(format);
            }
        }

        reader.seek(SeekFrom::Start(0)).ok()?;
        rtf::detect_rtf_format_from_reader(reader)
    })();
    reader.seek(SeekFrom::Start(original)).ok()?;
    detected
}

/// Detect iWork format from file path.
#[cfg(feature = "iwork")]
pub fn detect_iwork_format_from_path<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
    litchi_iwa::detect::path(path.as_ref())
        .ok()
        .flatten()
        .map(iwork_format)
}

#[cfg(feature = "iwork")]
fn iwork_format(format: litchi_iwa::detect::Format) -> FileFormat {
    match format {
        litchi_iwa::detect::Format::Pages => FileFormat::Pages,
        litchi_iwa::detect::Format::Keynote => FileFormat::Keynote,
        litchi_iwa::detect::Format::Numbers => FileFormat::Numbers,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    fn detects_flat_odf_bytes_and_reader() {
        let xml = br#"<?xml version="1.0"?><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.chart"><o:body><o:chart/></o:body></o:document>"#;
        assert_eq!(detect_file_format_from_bytes(xml), Some(FileFormat::Odc));
        let mut reader = std::io::Cursor::new(xml);
        reader.set_position(7);
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Odc)
        );
        assert_eq!(reader.position(), 7);
    }

    #[test]
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    fn test_detect_docx_from_bytes() {
        // Create a minimal ZIP file that looks like a DOCX
        let zip_data = create_minimal_docx_zip();
        let format = detect_file_format_from_bytes(&zip_data);
        assert!(format.is_some());
        assert_eq!(format.unwrap(), FileFormat::Docx);
    }

    #[test]
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
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

    #[test]
    fn detects_short_rtf_reader_and_restores_cursor() {
        let mut reader = std::io::Cursor::new(br#"{\rtf"#.to_vec());
        reader.set_position(2);

        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Rtf)
        );
        assert_eq!(reader.position(), 2);
    }

    #[test]
    #[cfg(feature = "iwork")]
    fn iwork_leaf_detection_maps_bytes_reader_smart_and_path() {
        let bytes = iwork_package(None);
        assert_eq!(
            detect_file_format_from_bytes(&bytes),
            Some(FileFormat::Numbers)
        );

        let mut reader = std::io::Cursor::new(&bytes);
        reader.set_position(6);
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Numbers)
        );
        assert_eq!(reader.position(), 6);

        assert!(matches!(
            crate::detection_smart::detect_format_smart(bytes.clone()),
            Some(crate::detection_smart::DetectedFormat::Numbers(retained)) if retained == bytes
        ));

        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), &bytes).unwrap();
        assert_eq!(
            detect_iwork_format_from_path(file.path()),
            Some(FileFormat::Numbers)
        );
    }

    #[test]
    #[cfg(all(feature = "iwork", feature = "ods"))]
    fn odf_precedes_iwork_consistently_for_mixed_packages() {
        let bytes = iwork_package(Some("application/vnd.oasis.opendocument.spreadsheet"));
        assert_eq!(detect_file_format_from_bytes(&bytes), Some(FileFormat::Ods));

        let mut reader = std::io::Cursor::new(&bytes);
        reader.set_position(4);
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Ods)
        );
        assert_eq!(reader.position(), 4);

        assert!(matches!(
            crate::detection_smart::detect_format_smart(bytes),
            Some(crate::detection_smart::DetectedFormat::Ods(_))
        ));
    }

    #[cfg(feature = "iwork")]
    fn iwork_package(mimetype: Option<&str>) -> Vec<u8> {
        use litchi_iwa::archive::{Archive, ArchiveObject, RawMessage};
        use std::io::{Cursor, Write};

        let document = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 6,
                        // TN.DocumentArchive required references plus its shared root.
                        data: vec![
                            0x22, 0x02, 0x08, 0x01, 0x2a, 0x02, 0x08, 0x02, 0x32, 0x02, 0x08, 0x03,
                            0x42, 0x02, 0x0a, 0x00,
                        ],
                    }],
                )
                .unwrap(),
            ],
        };
        let document = litchi_iwa::SnappyStream::compress(&document.to_bytes().unwrap()).unwrap();

        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            if let Some(mimetype) = mimetype {
                writer.start_file("mimetype", options).unwrap();
                writer.write_all(mimetype.as_bytes()).unwrap();
            }
            writer.start_file("Index/Document.iwa", options).unwrap();
            writer.write_all(&document).unwrap();
            writer
                .start_file("Index/CalculationEngine-174.iwa", options)
                .unwrap();
            writer.write_all(b"iwa").unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    // Helper function to create a minimal DOCX-like ZIP for testing
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    fn create_minimal_docx_zip() -> Vec<u8> {
        create_minimal_ooxml_zip(
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        )
    }

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    fn create_minimal_ooxml_zip(part_name: &str, content_type: &str) -> Vec<u8> {
        use crate::opc::PackURI;
        use crate::opc::phys_pkg::PhysPkgWriter;

        let mut writer = PhysPkgWriter::new();
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/{part_name}" ContentType="{content_type}"/></Types>"#
        );

        writer
            .write(
                &PackURI::new("/[Content_Types].xml").unwrap(),
                content_types.as_bytes(),
            )
            .unwrap();

        let relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{part_name}"/></Relationships>"#
        );
        writer
            .write(
                &PackURI::new("/_rels/.rels").unwrap(),
                relationships.as_bytes(),
            )
            .unwrap();

        writer
            .write(
                &PackURI::new(format!("/{part_name}")).unwrap(),
                b"office data",
            )
            .unwrap();

        writer.finish().unwrap()
    }
}
