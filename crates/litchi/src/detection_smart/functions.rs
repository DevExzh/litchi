//! Core file format detection functions.
//!
//! Uses a combined fixed-signature check before invoking format-specific
//! detectors owned by their leaf crates.

use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

use super::ole2;
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
use super::ooxml;
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
        // A valid ODF catalog without an OPC content-types member cannot be
        // an OOXML package. Keep the cheap ODF path ahead of the full OPC
        // probe, while retaining OOXML-first precedence for polyglots and
        // malformed ZIPs (where the catalog probe deliberately returns None).
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        {
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
            let normal_odf = is_normal_odf_package(bytes);
            #[cfg(not(any(feature = "odt", feature = "ods", feature = "odp")))]
            let normal_odf = false;

            // First try OOXML detection (most common), except for a
            // validated ordinary ODF package with no OPC catalog marker.
            if !normal_odf {
                #[cfg(test)]
                super::record_opc_probe();
                if let Some(result) = ooxml::detect_zip_format(bytes) {
                    return Some(result);
                }
            }
        }

        // Then try ODF detection
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
        if let Some(result) = odf::bytes(bytes) {
            return Some(result);
        }

        // Finally try iWork detection
        #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
        if let Ok(Some(format)) = litchi_iwa_detect::bytes(bytes)
            && let Some(result) = iwork_format(format)
        {
            return Some(result);
        }

        return None;
    }

    // Check compressed RTF after container formats so ZIP/OLE precedence is
    // unchanged for deliberately crafted polyglots.
    #[cfg(feature = "rtf")]
    if !mask.is_ole2() && !mask.is_zip() && super::is_compressed_rtf(bytes) {
        return Some(FileFormat::Rtf);
    }

    // Check RTF format (if matched)
    if mask.is_rtf()
        && let Some(result) = rtf::detect_rtf_format(bytes)
    {
        return Some(result);
    }

    None
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
fn is_normal_odf_package(bytes: &[u8]) -> bool {
    let limits = crate::opc::ReadLimits::default();
    litchi_odf_common::detect::packaged_mime(bytes).is_some()
        && litchi_odf_common::detect::packaged_has_ooxml_catalog_with_limits(
            bytes,
            super::catalog_probe_limits(limits),
        ) == Some(false)
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
        #[cfg(feature = "rtf")]
        let mut header = [0u8; 12];
        #[cfg(not(feature = "rtf"))]
        let mut header = [0u8; 8];
        let mut header_len = 0;
        while header_len < header.len() {
            let read = reader.read(&mut header[header_len..]).ok()?;
            if read == 0 {
                break;
            }
            header_len += read;
        }

        // Keep container precedence available even in builds where the
        // corresponding owner crates are feature-elided. Compressed RTF has
        // a marker at offset eight, which may otherwise collide with a ZIP
        // local-header payload in an RTF-only build.
        #[cfg(feature = "rtf")]
        let signature_mask = check_office_signatures(&header[..header_len]);

        if header_len >= utils::OLE2_SIGNATURE.len()
            && signature_matches(&header[..header_len], utils::OLE2_SIGNATURE)
        {
            reader.seek(SeekFrom::Start(0)).ok()?;
            return ole2::detect_ole2_format_from_reader(reader);
        }

        #[cfg(any(
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
            any(feature = "odt", feature = "ods", feature = "odp"),
            feature = "pages",
            feature = "keynote",
            feature = "numbers"
        ))]
        if header_len >= utils::ZIP_SIGNATURE.len()
            && signature_matches(&header[..header_len], utils::ZIP_SIGNATURE)
        {
            #[cfg(all(
                any(feature = "odt", feature = "ods", feature = "odp"),
                any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
            ))]
            let normal_odf =
                litchi_odf_common::detect::packaged_has_ooxml_catalog_from_reader_with_limits(
                    reader,
                    super::catalog_probe_limits(crate::opc::ReadLimits::default()),
                ) == Some(false);
            #[cfg(all(
                not(any(feature = "odt", feature = "ods", feature = "odp")),
                any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
            ))]
            let normal_odf = false;

            #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
            if !normal_odf {
                reader.seek(SeekFrom::Start(0)).ok()?;
                #[cfg(test)]
                super::record_opc_probe();
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

            #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
            {
                reader.seek(SeekFrom::Start(0)).ok()?;
                if let Ok(Some(format)) = litchi_iwa_detect::reader(reader)
                    && let Some(format) = iwork_format(format)
                {
                    return Some(format);
                }
            }

            return None;
        }

        #[cfg(feature = "rtf")]
        if !signature_mask.is_ole2()
            && !signature_mask.is_zip()
            && super::is_compressed_rtf(&header[..header_len])
        {
            return Some(FileFormat::Rtf);
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

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
#[allow(
    unreachable_patterns,
    reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
)]
fn iwork_format(format: litchi_iwa_detect::Format) -> Option<FileFormat> {
    match format {
        #[cfg(feature = "pages")]
        litchi_iwa_detect::Format::Pages => Some(FileFormat::Pages),
        #[cfg(feature = "keynote")]
        litchi_iwa_detect::Format::Keynote => Some(FileFormat::Keynote),
        #[cfg(feature = "numbers")]
        litchi_iwa_detect::Format::Numbers => Some(FileFormat::Numbers),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg(all(feature = "odt", feature = "docx"))]
    use std::io::Write;

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
    #[cfg(all(
        any(feature = "odt", feature = "ods", feature = "odp"),
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    fn packaged_odf_detection_handles_ordinary_and_malformed_catalogs() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_TEXT)
            .unwrap();
        writer.add_file("content.xml", b"<content/>").unwrap();
        let ordinary = writer.finish_to_bytes().unwrap();
        super::super::reset_opc_probe_count();
        assert_eq!(
            detect_file_format_from_bytes(&ordinary),
            Some(FileFormat::Odt)
        );
        assert_eq!(super::super::opc_probe_count(), 0);
        let mut reader = std::io::Cursor::new(ordinary.clone());
        reader.set_position(3);
        super::super::reset_opc_probe_count();
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Odt)
        );
        assert_eq!(reader.position(), 3);
        assert_eq!(super::super::opc_probe_count(), 0);

        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_TEXT)
            .unwrap();
        writer.add_file("content.xml", b"<content/>").unwrap();
        writer
            .add_file("[Content_Types].xml", b"<broken/>")
            .unwrap();
        let malformed_catalog = writer.finish_to_bytes().unwrap();
        super::super::reset_opc_probe_count();
        assert_eq!(
            detect_file_format_from_bytes(&malformed_catalog),
            Some(FileFormat::Odt)
        );
        assert_eq!(super::super::opc_probe_count(), 1);

        let mut malformed_zip = ordinary;
        malformed_zip.truncate(malformed_zip.len() - 1);
        assert_eq!(detect_file_format_from_bytes(&malformed_zip), None);
    }

    #[test]
    #[cfg(all(feature = "odt", feature = "docx"))]
    fn catalog_gate_counts_only_canonical_ordinary_odf_as_a_skip() {
        fn zip_entries(entries: &[(&str, &[u8])]) -> Vec<u8> {
            let mut output = std::io::Cursor::new(Vec::new());
            let mut writer = zip::ZipWriter::new(&mut output);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for (name, value) in entries {
                writer.start_file(*name, options).unwrap();
                writer.write_all(value).unwrap();
            }
            writer.finish().unwrap();
            output.into_inner()
        }

        fn central_record(bytes: &[u8], name: &[u8]) -> usize {
            let mut offset = bytes
                .windows(4)
                .position(|signature| signature == b"PK\x01\x02")
                .unwrap();
            loop {
                let name_len = usize::from(u16::from_le_bytes(
                    bytes[offset + 28..offset + 30].try_into().unwrap(),
                ));
                let extra_len = usize::from(u16::from_le_bytes(
                    bytes[offset + 30..offset + 32].try_into().unwrap(),
                ));
                let comment_len = usize::from(u16::from_le_bytes(
                    bytes[offset + 32..offset + 34].try_into().unwrap(),
                ));
                if &bytes[offset + 46..offset + 46 + name_len] == name {
                    return offset;
                }
                offset += 46 + name_len + extra_len + comment_len;
            }
        }

        fn rename_entry(bytes: &mut [u8], old_name: &[u8], new_name: &[u8]) {
            assert_eq!(old_name.len(), new_name.len());
            let central = central_record(bytes, old_name);
            let central_name_len = usize::from(u16::from_le_bytes(
                bytes[central + 28..central + 30].try_into().unwrap(),
            ));
            assert_eq!(central_name_len, old_name.len());

            let local = usize::try_from(u32::from_le_bytes(
                bytes[central + 42..central + 46].try_into().unwrap(),
            ))
            .unwrap();
            assert_eq!(&bytes[local..local + 4], b"PK\x03\x04");
            let local_name_len = usize::from(u16::from_le_bytes(
                bytes[local + 26..local + 28].try_into().unwrap(),
            ));
            assert_eq!(local_name_len, old_name.len());
            assert_eq!(&bytes[local + 30..local + 30 + local_name_len], old_name);
            assert_eq!(
                &bytes[central + 46..central + 46 + central_name_len],
                old_name
            );

            bytes[local + 30..local + 30 + local_name_len].copy_from_slice(new_name);
            bytes[central + 46..central + 46 + central_name_len].copy_from_slice(new_name);
        }

        fn assert_probe(bytes: &[u8], expected: usize) {
            super::super::reset_opc_probe_count();
            let _ = detect_file_format_from_bytes(bytes);
            assert_eq!(super::super::opc_probe_count(), expected);
        }

        let ordinary = zip_entries(&[
            (
                "mimetype",
                litchi_odf_common::constants::ODF_TEXT.as_bytes(),
            ),
            ("content.xml", b"<content/>"),
        ]);
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&ordinary),
            Some(false)
        );
        assert_probe(&ordinary, 0);

        let alias = zip_entries(&[
            (
                "mimetype",
                litchi_odf_common::constants::ODF_TEXT.as_bytes(),
            ),
            ("./content.xml", b"<content/>"),
        ]);
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&alias),
            None
        );
        assert_probe(&alias, 1);

        // zip 8.6 rejects duplicate names at authoring time. Emit two legal,
        // equal-length names, then rewrite one entry's local and central
        // headers so the detector still receives a genuine duplicate ZIP.
        let mut duplicate = zip_entries(&[
            (
                "mimetype",
                litchi_odf_common::constants::ODF_TEXT.as_bytes(),
            ),
            ("content.xm2", b"one"),
            ("content.xml", b"two"),
        ]);
        rename_entry(&mut duplicate, b"content.xm2", b"content.xml");
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&duplicate),
            None
        );
        assert_probe(&duplicate, 1);

        let case_catalog = zip_entries(&[
            (
                "mimetype",
                litchi_odf_common::constants::ODF_TEXT.as_bytes(),
            ),
            ("[content_types].xml", b"<broken/>"),
        ]);
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&case_catalog),
            Some(true)
        );
        assert_probe(&case_catalog, 1);

        let mut trailing = ordinary.clone();
        trailing.extend_from_slice(b"trailing");
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&trailing),
            None
        );
        assert_probe(&trailing, 1);

        let mut malformed = ordinary.clone();
        malformed.pop();
        assert_probe(&malformed, 1);

        let mut zip64 = ordinary.clone();
        let eocd = zip64
            .windows(4)
            .rposition(|signature| signature == b"PK\x05\x06")
            .unwrap();
        zip64[eocd + 8..eocd + 12].copy_from_slice(&[0xff, 0xff, 0xff, 0xff]);
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&zip64),
            None
        );
        assert_probe(&zip64, 1);

        let mut invalid_utf8 = ordinary.clone();
        let central = central_record(&invalid_utf8, b"content.xml");
        invalid_utf8[central + 46] = 0xff;
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&invalid_utf8),
            None
        );
        assert_probe(&invalid_utf8, 1);

        for flag in [0x0008_u16, 0x0001_u16] {
            let mut flagged = ordinary.clone();
            let central = central_record(&flagged, b"content.xml");
            flagged[central + 8..central + 10].copy_from_slice(&flag.to_le_bytes());
            assert_eq!(
                litchi_odf_common::detect::packaged_has_ooxml_catalog(&flagged),
                None
            );
            assert_probe(&flagged, 1);
        }

        let mut prefixed = b"prefix".to_vec();
        prefixed.extend_from_slice(&ordinary);
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(&prefixed),
            None
        );
        assert_probe(&prefixed, 0);
    }

    #[test]
    #[cfg(all(feature = "odt", feature = "docx"))]
    fn valid_ooxml_odf_polyglot_keeps_ooxml_precedence() {
        let mut output = std::io::Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(litchi_odf_common::constants::ODF_TEXT.as_bytes())
            .unwrap();
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("word/document.xml", options).unwrap();
        writer
            .write_all(br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#)
            .unwrap();
        writer.finish().unwrap();

        super::super::reset_opc_probe_count();
        assert_eq!(
            detect_file_format_from_bytes(output.get_ref()),
            Some(FileFormat::Docx)
        );
        assert_eq!(super::super::opc_probe_count(), 1);
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

    #[cfg(feature = "rtf")]
    #[test]
    fn detects_compressed_rtf_bytes_and_reader_without_moving_reader_cursor() {
        let source = br#"{\rtf1\ansi Compressed RTF\par}"#;
        for use_lzfu in [true, false] {
            let compressed = litchi_rtf::transport::compress(source, use_lzfu).unwrap();

            assert_eq!(
                detect_file_format_from_bytes(&compressed),
                Some(FileFormat::Rtf)
            );

            let mut reader = std::io::Cursor::new(compressed);
            reader.set_position(3);
            assert_eq!(
                detect_format_from_reader(&mut reader),
                Some(FileFormat::Rtf)
            );
            assert_eq!(reader.position(), 3);
        }

        // A ZIP local header can carry the LZFu marker at offset eight. The
        // container signature must retain precedence in both detector APIs,
        // including an rtf-only build where ZIP owners are feature-elided.
        let mut zip_polyglot = b"PK\x03\x04\x00\x00\x00\x00".to_vec();
        zip_polyglot.extend_from_slice(b"LZFu");
        assert_eq!(detect_file_format_from_bytes(&zip_polyglot), None);

        let mut zip_reader = std::io::Cursor::new(zip_polyglot);
        zip_reader.set_position(2);
        assert_eq!(detect_format_from_reader(&mut zip_reader), None);
        assert_eq!(zip_reader.position(), 2);

        // Apply the same regression to the eight-byte OLE2 signature. An
        // invalid OLE candidate must not fall through to compressed RTF.
        let mut ole_polyglot = utils::OLE2_SIGNATURE.to_vec();
        ole_polyglot.extend_from_slice(b"LZFu");
        assert_eq!(detect_file_format_from_bytes(&ole_polyglot), None);

        let mut ole_reader = std::io::Cursor::new(ole_polyglot);
        ole_reader.set_position(1);
        assert_eq!(detect_format_from_reader(&mut ole_reader), None);
        assert_eq!(ole_reader.position(), 1);
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
