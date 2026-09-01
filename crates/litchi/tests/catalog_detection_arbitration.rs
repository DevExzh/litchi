//! Integration regressions for package-catalog format arbitration.

#[cfg(any(feature = "pptx", feature = "xlsx"))]
const PPTX_PRESENTATION_MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

#[cfg(any(feature = "pptx", feature = "xlsx"))]
const PPTX_PACKAGE_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#;

#[cfg(any(feature = "pptx", feature = "xlsx"))]
const PPTX_PRESENTATION_XML: &[u8] = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#;

#[cfg(any(feature = "pptx", feature = "xlsx"))]
fn zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
    use std::io::{Cursor, Write};

    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for &(name, data) in entries {
        writer
            .start_file(name, options)
            .expect("ZIP member name is valid");
        writer.write_all(data).expect("ZIP member is writable");
    }
    writer.finish().expect("ZIP archive is writable");
    output.into_inner()
}

#[cfg(any(feature = "pptx", feature = "xlsx"))]
fn pptx_catalog(content_type: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="{content_type}"/></Types>"#
    );
    zip_bytes(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", PPTX_PACKAGE_RELATIONSHIPS),
        ("ppt/presentation.xml", PPTX_PRESENTATION_XML),
    ])
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
fn odf_package(mimetype: &str, content: &[u8]) -> Vec<u8> {
    let mut writer = litchi::odf_common::core::PackageWriter::new();
    writer.set_mimetype(mimetype).expect("ODF MIME is valid");
    writer
        .add_file("content.xml", content)
        .expect("ODF content is writable");
    writer.finish_to_bytes().expect("ODF package is writable")
}

#[cfg(all(feature = "docx", feature = "odt", any(unix, windows)))]
mod document_limit_arbitration {
    use super::odf_package;
    use litchi::Document;
    use litchi::common::{Error, Resource};

    const ODT_MIME: &str = "application/vnd.oasis.opendocument.text";
    const ODT_CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">bounded</text:p></office:text></office:body></office:document-content>"#;
    const DOCX_CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    const DOCX_ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    const DOCX_DOCUMENT: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>OOXML wins</w:t></w:r></w:p></w:body></w:document>"#;

    fn odt_docx_polyglot(catalog_name: &str) -> Vec<u8> {
        let mut writer = litchi::odf_common::core::PackageWriter::new();
        writer.set_mimetype(ODT_MIME).expect("ODT MIME is valid");
        for (name, data) in [
            ("content.xml", ODT_CONTENT),
            (catalog_name, DOCX_CONTENT_TYPES),
            ("_rels/.rels", DOCX_ROOT_RELATIONSHIPS),
            ("word/document.xml", DOCX_DOCUMENT),
        ] {
            writer
                .add_file(name, data)
                .expect("polyglot member is writable");
        }
        writer
            .finish_to_bytes()
            .expect("polyglot package is writable")
    }

    fn assert_input_limit(error: Error, observed: u64, maximum: u64) {
        let Error::ResourceLimit(limit) = error else {
            panic!("expected an input resource limit, got {error}");
        };
        assert_eq!(limit.resource, Resource::InputBytes);
        assert_eq!(limit.observed, observed);
        assert_eq!(limit.limit, maximum);
    }

    #[test]
    fn ordinary_odt_uses_native_policy_but_polyglot_honors_docx_input_limit() {
        let ordinary = odf_package(ODT_MIME, ODT_CONTENT);
        let native_limits = litchi::docx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("limits are consistent");
        assert!(Document::from_bytes_with_limits(ordinary.clone(), native_limits).is_ok());
        let ordinary_path = tempfile::Builder::new()
            .suffix(".odt")
            .tempfile()
            .expect("temporary ordinary ODT path");
        std::fs::write(ordinary_path.path(), &ordinary).expect("write ordinary ODT fixture");
        assert!(Document::open_with_limits(ordinary_path.path(), native_limits).is_ok());

        let renamed_path = tempfile::Builder::new()
            .suffix(".DOCX")
            .tempfile()
            .expect("temporary DOCX-suffixed ODT path");
        std::fs::write(renamed_path.path(), &ordinary).expect("write renamed ODT fixture");
        let renamed_error = match Document::open_with_limits(renamed_path.path(), native_limits) {
            Ok(_) => panic!("DOCX-suffixed ODT bypassed the caller's input limit"),
            Err(error) => error,
        };
        assert_input_limit(
            renamed_error,
            u64::try_from(ordinary.len()).expect("fixture length fits in u64"),
            1,
        );

        for catalog_name in ["[Content_Types].xml", "[content_types].xml"] {
            let bytes = odt_docx_polyglot(catalog_name);
            let observed = u64::try_from(bytes.len()).expect("fixture length fits in u64");
            let maximum = observed - 1;
            let limits = litchi::docx::ReadLimits::builder()
                .max_input_bytes(maximum)
                .expect("positive input limit")
                .build()
                .expect("limits are consistent");

            let bytes_error = match Document::from_bytes_with_limits(bytes.clone(), limits) {
                Ok(_) => panic!("ODT/DOCX bytes bypassed the caller's DOCX input limit"),
                Err(error) => error,
            };
            assert_input_limit(bytes_error, observed, maximum);

            let temporary = tempfile::Builder::new()
                .suffix(".odt")
                .tempfile()
                .expect("temporary ODT/DOCX path");
            std::fs::write(temporary.path(), bytes).expect("write ODT/DOCX fixture");
            let path_error = match Document::open_with_limits(temporary.path(), limits) {
                Ok(_) => panic!("ODT/DOCX path bypassed the caller's DOCX input limit"),
                Err(error) => error,
            };
            assert_input_limit(path_error, observed, maximum);
        }
    }
}

#[cfg(feature = "pptx")]
mod content_type_matching {
    use super::{PPTX_PRESENTATION_MAIN_CONTENT_TYPE, pptx_catalog};
    use litchi::common::FileFormat;
    use litchi::common::detection::detect_file_format_from_bytes;

    #[test]
    fn accepts_canonical_content_type_casing_but_rejects_prefix_and_suffix_near_matches() {
        let uppercase = PPTX_PRESENTATION_MAIN_CONTENT_TYPE.to_ascii_uppercase();
        for content_type in [PPTX_PRESENTATION_MAIN_CONTENT_TYPE.to_owned(), uppercase] {
            assert_eq!(
                detect_file_format_from_bytes(&pptx_catalog(&content_type)),
                Some(FileFormat::Pptx),
                "canonical MIME type casing should identify PPTX: {content_type}"
            );
        }

        for content_type in [
            format!("x-{PPTX_PRESENTATION_MAIN_CONTENT_TYPE}"),
            format!("{PPTX_PRESENTATION_MAIN_CONTENT_TYPE}-suffix"),
        ] {
            assert_eq!(
                detect_file_format_from_bytes(&pptx_catalog(&content_type)),
                None,
                "near-match MIME type must not identify PPTX: {content_type}"
            );
        }
    }
}

#[cfg(feature = "ods")]
mod ordinary_ods_detection {
    use super::odf_package;
    use litchi::common::FileFormat;
    use litchi::common::detection::{detect_file_format_from_bytes, detect_format_from_reader};
    use litchi::detection_smart::{DetectedFormat, detect_format_smart};
    use std::io::Cursor;

    const MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet/></office:body></office:document-content>"#;

    #[test]
    fn ordinary_ods_remains_detectable_without_an_opc_catalog() {
        let bytes = odf_package(MIME, CONTENT);
        assert_eq!(
            litchi::odf_common::detect::packaged_has_ooxml_catalog(&bytes),
            Some(false)
        );
        assert_eq!(detect_file_format_from_bytes(&bytes), Some(FileFormat::Ods));

        let mut reader = Cursor::new(bytes.clone());
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Ods)
        );
        assert!(matches!(
            detect_format_smart(bytes),
            Some(DetectedFormat::Ods(_))
        ));
    }
}

#[cfg(feature = "odp")]
mod ordinary_odp_detection {
    use super::odf_package;
    use litchi::common::FileFormat;
    use litchi::common::detection::{detect_file_format_from_bytes, detect_format_from_reader};
    use litchi::detection_smart::{DetectedFormat, detect_format_smart};
    use std::io::Cursor;

    const MIME: &str = "application/vnd.oasis.opendocument.presentation";
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page draw:name="one"/></office:presentation></office:body></office:document-content>"#;

    #[test]
    fn ordinary_odp_remains_detectable_without_an_opc_catalog() {
        let bytes = odf_package(MIME, CONTENT);
        assert_eq!(
            litchi::odf_common::detect::packaged_has_ooxml_catalog(&bytes),
            Some(false)
        );
        assert_eq!(detect_file_format_from_bytes(&bytes), Some(FileFormat::Odp));

        let mut reader = Cursor::new(bytes.clone());
        assert_eq!(
            detect_format_from_reader(&mut reader),
            Some(FileFormat::Odp)
        );
        assert!(matches!(
            detect_format_smart(bytes),
            Some(DetectedFormat::Odp(_))
        ));
    }
}

#[cfg(all(feature = "pptx", feature = "odp"))]
mod ooxml_odf_arbitration {
    use super::{
        PPTX_PACKAGE_RELATIONSHIPS, PPTX_PRESENTATION_MAIN_CONTENT_TYPE, PPTX_PRESENTATION_XML,
        odf_package, zip_bytes,
    };
    use litchi::Presentation;
    use litchi::common::detection::detect_file_format_from_bytes;
    use litchi::common::{Error, FileFormat};
    use litchi::detection_smart::{
        DetectedFormat, detect_format_smart, detect_format_smart_with_limits,
    };

    const ODP_MIME: &str = "application/vnd.oasis.opendocument.presentation";
    const ODP_CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page draw:name="one"/></office:presentation></office:body></office:document-content>"#;
    const POLYGLOT_ODP_CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page draw:name="one"/><draw:page draw:name="two"/></office:presentation></office:body></office:document-content>"#;
    const ODP_MANIFEST: &[u8] = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.presentation"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#;
    const PPTX_PRESENTATION_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#;
    const PPTX_SLIDE: &[u8] = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;

    fn pptx_odf_polyglot() -> Vec<u8> {
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="{PPTX_PRESENTATION_MAIN_CONTENT_TYPE}"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#
        );
        zip_bytes(&[
            ("mimetype", ODP_MIME.as_bytes()),
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", PPTX_PACKAGE_RELATIONSHIPS),
            ("ppt/presentation.xml", PPTX_PRESENTATION_XML),
            (
                "ppt/_rels/presentation.xml.rels",
                PPTX_PRESENTATION_RELATIONSHIPS,
            ),
            ("ppt/slides/slide1.xml", PPTX_SLIDE),
            ("content.xml", POLYGLOT_ODP_CONTENT),
            ("META-INF/manifest.xml", ODP_MANIFEST),
        ])
    }

    #[test]
    fn ooxml_precedence_survives_an_odf_mimetype_marker() {
        let bytes = pptx_odf_polyglot();
        assert_eq!(
            detect_file_format_from_bytes(&bytes),
            Some(FileFormat::Pptx)
        );
        assert!(matches!(
            detect_format_smart(bytes.clone()),
            Some(DetectedFormat::Pptx(_))
        ));

        let presentation = Presentation::from_bytes(bytes).expect("valid PPTX owner wins");
        assert_eq!(
            presentation.slide_count().expect("slide count is readable"),
            1
        );
    }

    #[test]
    fn ordinary_odp_bytes_use_native_policy_before_pptx_limits() {
        let bytes = odf_package(ODP_MIME, ODP_CONTENT);
        let limits = litchi::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive PPTX input limit")
            .build()
            .expect("limits are consistent");

        let presentation = Presentation::from_bytes_with_limits(bytes, limits)
            .expect("ordinary ODP bytes use the ODP policy");
        assert_eq!(
            presentation.slide_count().expect("slide count is readable"),
            1
        );
    }

    #[test]
    fn tight_limits_do_not_fall_back_to_the_odp_marker() {
        let bytes = pptx_odf_polyglot();
        let input_limits = litchi::pptx::ReadLimits::builder()
            .max_input_bytes(u64::try_from(bytes.len() - 1).expect("fixture length fits in u64"))
            .expect("positive input limit")
            .build()
            .expect("limits are consistent");
        let catalog_limits = litchi::pptx::ReadLimits::builder()
            .max_content_types_bytes(1)
            .expect("positive content-types limit")
            .build()
            .expect("limits are consistent");

        assert!(detect_format_smart_with_limits(bytes.clone(), input_limits).is_none());
        assert!(detect_format_smart_with_limits(bytes.clone(), catalog_limits).is_none());
        let error = match Presentation::from_bytes_with_limits(bytes, catalog_limits) {
            Ok(_) => panic!("tight limits were replaced by an ODP fallback"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::ResourceLimit(ref limit)
                if limit.resource == litchi::common::Resource::InputBytes && limit.limit == 1
        ));
    }

    #[test]
    fn renamed_odp_with_tiny_pptx_limit_uses_the_odp_policy() {
        let bytes = odf_package(ODP_MIME, ODP_CONTENT);
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary renamed ODP path");
        std::fs::write(temporary.path(), &bytes).expect("write renamed ODP");

        let presentation = Presentation::open(temporary.path()).expect("renamed ODP opens");
        assert_eq!(
            presentation.slide_count().expect("slide count is readable"),
            1
        );

        let limits = litchi::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("limits are consistent");
        let limited = Presentation::open_with_limits(temporary.path(), limits)
            .expect("renamed ordinary ODP uses the neutral ODP budget");
        assert_eq!(limited.slide_count().expect("slide count is readable"), 1);
    }
}

#[cfg(feature = "xlsx")]
mod wrong_family_workbook {
    use super::{PPTX_PRESENTATION_MAIN_CONTENT_TYPE, pptx_catalog};
    use litchi::Workbook;
    use litchi::common::Error;

    #[test]
    fn wrong_family_ooxml_path_returns_not_office_file() {
        let temporary = tempfile::Builder::new()
            .suffix(".xlsx")
            .tempfile()
            .expect("temporary wrong-family workbook path");
        std::fs::write(
            temporary.path(),
            pptx_catalog(PPTX_PRESENTATION_MAIN_CONTENT_TYPE),
        )
        .expect("write wrong-family OOXML fixture");

        let error = match Workbook::open(temporary.path()) {
            Ok(_) => panic!("PPTX content was accepted by the XLSX facade"),
            Err(error) => error,
        };
        assert!(
            error
                .downcast_ref::<Error>()
                .is_some_and(|error| matches!(error, Error::NotOfficeFile))
        );
    }
}
