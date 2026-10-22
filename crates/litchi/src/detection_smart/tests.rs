//! Detection tests that need the facade's iWork tail.

#[cfg(all(feature = "odt", feature = "pages"))]
#[test]
fn rejected_odf_probe_preserves_the_pages_source_for_lower_precedence_detection() {
    let bytes = include_bytes!("../../../../test-data/iwork/pages/basic.pages").to_vec();
    let pointer = bytes.as_ptr();
    let capacity = bytes.capacity();

    let detected = super::detect_format_smart(bytes, &super::SourceLimits::default())
        .expect("Pages fixture should be detected");
    let super::DetectedFormat::Pages(retained) = detected else {
        panic!("ODF rejection must continue to Pages detection");
    };

    assert_eq!(retained.as_ptr(), pointer);
    assert_eq!(retained.capacity(), capacity);
}

/// iWork moved out of `litchi_ops::detect` into a facade tail. These tests pin
/// that it is still the last ZIP candidate: it classifies a package nothing
/// else claims, and loses to OOXML and ODF when a package carries both.
#[cfg(feature = "pages")]
mod iwork_is_classified_last {
    #[cfg(any(feature = "docx", feature = "odt"))]
    use std::io::Read;
    use std::io::{Cursor, Seek, SeekFrom, Write};

    use litchi_core::detection::FileFormat;

    use super::super::{
        DetectedFormat, detect_file_format, detect_file_format_from_bytes,
        detect_format_from_reader, detect_format_smart,
    };

    const PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

    /// Rewrite the Pages fixture with `leading` entries before its own and
    /// `trailing` entries after them. Every entry is stored.
    #[cfg(any(feature = "docx", feature = "odt"))]
    fn repack(leading: &[(&str, &[u8])], trailing: &[(&str, &[u8])]) -> Vec<u8> {
        let mut source = zip::ZipArchive::new(Cursor::new(PAGES)).expect("Pages fixture is a ZIP");
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, data) in leading {
            writer.start_file(*name, options).unwrap();
            writer.write_all(data).unwrap();
        }
        for index in 0..source.len() {
            let mut entry = source.by_index(index).unwrap();
            let name = entry.name().to_owned();
            let mut data = Vec::new();
            entry.read_to_end(&mut data).unwrap();
            writer.start_file(name, options).unwrap();
            writer.write_all(&data).unwrap();
        }
        for (name, data) in trailing {
            writer.start_file(*name, options).unwrap();
            writer.write_all(data).unwrap();
        }
        writer.finish().unwrap();
        output.into_inner()
    }

    fn iwork_leaf_says_pages(bytes: &[u8]) -> bool {
        matches!(
            litchi_iwa_detect::bytes(bytes, litchi_iwa_detect::Limits::default()),
            Ok(Some(litchi_iwa_detect::Format::Pages))
        )
    }

    /// The format each facade entry point reports, with the reader's cursor
    /// checked for restoration.
    fn facade_formats(bytes: &[u8]) -> Vec<Option<FileFormat>> {
        let limits = crate::detection_smart::SourceLimits::default();
        let mut formats = vec![detect_file_format_from_bytes(bytes, &limits)];

        let mut reader = Cursor::new(bytes.to_vec());
        reader.seek(SeekFrom::Start(3)).unwrap();
        formats.push(detect_format_from_reader(&mut reader, &limits));
        assert_eq!(
            reader.stream_position().unwrap(),
            3,
            "cursor must be restored"
        );

        let mut file = tempfile::NamedTempFile::new().unwrap();
        file.write_all(bytes).unwrap();
        file.flush().unwrap();
        formats.push(detect_file_format(file.path(), &limits));
        #[cfg(all(
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
            any(unix, windows)
        ))]
        formats.push(
            super::super::detect_and_prepare(file.path(), &limits)
                .map(|detected| detected.format()),
        );

        formats
            .push(detect_format_smart(bytes.to_vec(), &limits).map(|detected| detected.format()));
        formats
    }

    #[test]
    fn a_package_nothing_else_claims_is_classified_by_the_iwork_tail() {
        assert!(iwork_leaf_says_pages(PAGES));
        for format in facade_formats(PAGES) {
            assert_eq!(format, Some(FileFormat::Pages));
        }
        assert!(matches!(
            detect_format_smart(
                PAGES.to_vec(),
                &crate::detection_smart::SourceLimits::default(),
            ),
            Some(DetectedFormat::Pages(_))
        ));

        // The ops coordinator itself is iWork-free: without the facade tail
        // the same package is unrecognized.
        assert_eq!(
            litchi_ops::detect::detect_file_format_from_bytes(
                PAGES,
                &litchi_ops::detect::SourceLimits::default(),
            ),
            None
        );
        assert_eq!(
            litchi_ops::detect::detect_format_from_reader(
                &mut Cursor::new(PAGES),
                &litchi_ops::detect::SourceLimits::default(),
            ),
            None
        );
        assert!(
            litchi_ops::detect::detect_format_smart(
                PAGES.to_vec(),
                &litchi_ops::detect::SourceLimits::default(),
            )
            .is_none()
        );
    }

    #[cfg(feature = "docx")]
    #[test]
    fn an_ooxml_package_with_iwork_markers_stays_ooxml() {
        let polyglot = repack(
            &[],
            &[
                (
                    "[Content_Types].xml",
                    br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
                ),
                (
                    "_rels/.rels",
                    br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
                ),
                (
                    "word/document.xml",
                    br#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#,
                ),
            ],
        );
        assert!(
            iwork_leaf_says_pages(&polyglot),
            "the polyglot must still look like Pages to the iWork leaf"
        );
        for format in facade_formats(&polyglot) {
            assert_eq!(format, Some(FileFormat::Docx));
        }
    }

    #[cfg(feature = "odt")]
    #[test]
    fn an_odf_package_with_iwork_markers_stays_odf() {
        let polyglot = repack(
            &[("mimetype", b"application/vnd.oasis.opendocument.text")],
            &[(
                "content.xml",
                br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>odf</t:p></o:text></o:body></o:document-content>"#,
            )],
        );
        assert!(
            iwork_leaf_says_pages(&polyglot),
            "the polyglot must still look like Pages to the iWork leaf"
        );
        for format in facade_formats(&polyglot) {
            assert_eq!(format, Some(FileFormat::Odt));
        }
    }
}
