//! Markdown tests: the facade's in-source Markdown tests for its `Document`,
//! ported onto `File` with the same fixtures and outputs, and the `File`
//! dispatch: workbooks, locked and signed files, and the family views.
//!
//! Refusals that the facade reported as a string-flattened
//! `Error::Unsupported` are asserted through the typed code here.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic by design"
)]

#[allow(unused_imports, reason = "used by the suites of the enabled formats")]
use litchi_markdown::{MarkdownOptions, ToMarkdown};

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
#[allow(unused_imports, reason = "used by the suites of the enabled formats")]
use crate::tests::test_data;
#[allow(unused_imports, reason = "used by the suites of the enabled formats")]
use crate::{Code, File};

#[cfg(feature = "doc")]
#[test]
fn real_doc_inline_images_are_refused_in_plain_output() {
    let file = crate::open(test_data("ole/doc/testPictures.doc")).unwrap();
    for options in [
        MarkdownOptions::new(),
        MarkdownOptions::new().with_styles(false),
    ] {
        let error = file
            .to_markdown_with_options(&options)
            .expect_err("DOC inline images must be refused");
        assert_eq!(error.code(), Code::UnsupportedElement, "{error:?}");
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Doc));
    }
}

#[cfg(feature = "docx")]
mod docx {
    use super::*;
    use crate::format::docx::DocxOwner;
    use crate::tests::{content_type, minimal_ooxml};
    use crate::{OpenOptions, SourceMode};

    fn minimal_docx(document_xml: &[u8]) -> Vec<u8> {
        minimal_ooxml("word/document.xml", content_type::DOCX, document_xml, None)
    }

    /// A DOCX package with a styles part, which the package-aware
    /// conversion reads, as the facade's Markdown tests build it.
    fn styled_docx(document_xml: &str) -> Vec<u8> {
        use std::io::Write as _;

        const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#;
        const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        const DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
        const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;

        let mut archive = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, contents) in [
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", document_xml),
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/styles.xml", STYLES),
        ] {
            archive.start_file(name, options).unwrap();
            archive.write_all(contents.as_bytes()).unwrap();
        }
        archive.finish().unwrap().into_inner()
    }

    const ALT_CHUNK: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdChunk"/><w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>"#;
    const UNKNOWN_BLOCK: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:customBodyBlock/><w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>"#;

    /// The three DOCX routes the facade tested: the eager oracle, the owned
    /// source-backed byte route and, where positional sources exist, the
    /// path route. The temporary file of the path route lives as long as
    /// the returned guard.
    fn routes(bytes: Vec<u8>) -> (Vec<(&'static str, File)>, Option<impl Sized>) {
        let mut files = vec![
            (
                "eager",
                crate::open_bytes_with(
                    bytes.clone(),
                    &OpenOptions::new().source_mode(SourceMode::Eager),
                )
                .unwrap(),
            ),
            ("owned", crate::open_bytes(bytes.clone()).unwrap()),
        ];
        #[cfg(any(unix, windows))]
        let guard = {
            let temporary = crate::tests::temporary_with(&bytes, ".docx");
            files.push(("source", crate::open(temporary.path()).unwrap()));
            Some(temporary)
        };
        #[cfg(not(any(unix, windows)))]
        let guard: Option<()> = None;
        (files, guard)
    }

    #[test]
    fn active_alt_chunk_blocks_are_refused_on_every_route() {
        let (files, _guard) = routes(minimal_docx(ALT_CHUNK));
        for (route, file) in files {
            let error = file
                .to_markdown()
                .expect_err("Markdown must not silently drop altChunk blocks");
            assert_eq!(error.code(), Code::UnsupportedElement, "{route}: {error:?}");
            assert!(error.message().contains("altChunk"), "{route}: {error:?}");
            assert_eq!(error.format(), Some(litchi_core::FileFormat::Docx));
        }
    }

    #[test]
    fn active_unknown_blocks_are_refused_on_every_route() {
        let (files, _guard) = routes(minimal_docx(UNKNOWN_BLOCK));
        for (route, file) in files {
            let error = file
                .to_markdown()
                .expect_err("Markdown must not silently drop unknown body blocks");
            assert_eq!(error.code(), Code::UnsupportedElement, "{route}: {error:?}");
            assert!(
                error.message().contains("unmodeled DOCX body blocks"),
                "{route}: {error:?}"
            );
        }
    }

    #[test]
    fn every_route_renders_the_same_markdown() {
        let document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#;
        let (files, _guard) = routes(styled_docx(document));
        for (route, file) in files {
            assert_eq!(
                file.to_markdown().unwrap(),
                "# Title\n\n**bold** plain\n\n| A | B |\n|----------|----------|\n\n\n",
                "{route}"
            );
            assert_eq!(
                file.document().unwrap().to_markdown().unwrap(),
                file.to_markdown().unwrap(),
                "{route}: the view and the file render alike"
            );
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn a_source_changed_after_the_cached_package_is_reported() {
        let temporary =
            crate::tests::temporary_copy(&test_data("ooxml/docx/FancyFoot.docx"), ".docx");
        let file = crate::open(temporary.path()).unwrap();
        let File::Docx(docx) = &file else {
            panic!("expected DOCX, got {file:?}");
        };
        let DocxOwner::Source(package) = docx.owner() else {
            panic!("a filesystem DOCX keeps the source owner");
        };
        let expected = package.source_version().unwrap();

        // Fill the handle's Markdown cache, then change the file under it.
        docx.markdown_package()
            .expect("source Markdown package should populate the cache");
        crate::tests::mutate(temporary.path());

        let error = file
            .to_markdown()
            .expect_err("Markdown must reject a stale cached package");
        crate::tests::assert_source_changed(&error, expected);
        // A clone shares the cache, and the same refusal.
        let error = file
            .clone()
            .to_markdown()
            .expect_err("a clone shares the stale cache");
        crate::tests::assert_source_changed(&error, expected);
    }

    #[test]
    fn the_markdown_package_is_built_once_per_handle() {
        let file = crate::open_bytes(styled_docx(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>cached</w:t></w:r></w:p></w:body></w:document>"#,
        ))
        .unwrap();
        let File::Docx(docx) = &file else {
            panic!("expected DOCX, got {file:?}");
        };
        assert!(docx.is_source_backed());
        let first: *const litchi_docx::Package = docx.markdown_package().unwrap();
        let clone = file.clone();
        let File::Docx(cloned) = &clone else {
            unreachable!("a clone keeps its variant");
        };
        let second: *const litchi_docx::Package = cloned.markdown_package().unwrap();
        assert_eq!(first, second, "clones share one materialized package");
        assert_eq!(file.to_markdown().unwrap(), "cached\n\n");
    }

    #[cfg(feature = "yaml")]
    #[test]
    fn front_matter_renders_the_core_properties() {
        let file = crate::open(test_data("ooxml/docx/documentProperties.docx")).unwrap();
        let markdown = file
            .to_markdown_with_options(&MarkdownOptions::new().with_metadata(true))
            .unwrap();
        assert!(markdown.starts_with("---\n"), "{markdown}");
        assert!(markdown.ends_with("Hello World\\!\n\n"), "{markdown}");
    }

    #[cfg(not(feature = "yaml"))]
    #[test]
    fn front_matter_needs_the_yaml_feature() {
        let file = crate::open(test_data("ooxml/docx/documentProperties.docx")).unwrap();
        let error = file
            .to_markdown_with_options(&MarkdownOptions::new().with_metadata(true))
            .expect_err("front matter needs the `yaml` feature");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
        assert_eq!(error.suggestion(), Some("enable the `yaml` cargo feature"));
    }
}

#[cfg(all(feature = "odt", any(unix, windows)))]
#[test]
fn a_source_changed_odt_is_reported_after_markdown() {
    use crate::format::odt::OdtOwner;

    let mut builder = litchi_odt::Builder::new();
    builder.add_paragraph("Source-backed ODT").unwrap();
    let temporary =
        crate::tests::temporary_with(&builder.build().unwrap().to_bytes().unwrap(), ".odt");
    let file = crate::open(temporary.path()).unwrap();
    let File::Odt(odt) = &file else {
        panic!("expected ODT, got {file:?}");
    };
    let OdtOwner::Source(document) = odt.owner() else {
        panic!("a filesystem ODT keeps the source owner");
    };
    let expected = document.source_version().unwrap();

    assert_eq!(file.to_markdown().unwrap(), "Source\\-backed ODT\n\n");
    crate::tests::mutate(temporary.path());
    let error = file
        .to_markdown()
        .expect_err("Markdown must reject a stale ODT source");
    crate::tests::assert_source_changed(&error, expected);
}

#[cfg(feature = "pptx")]
#[test]
fn presentations_render_one_section_per_slide() {
    let file = crate::open(test_data(
        "libreoffice-core/oox/qa/unit/data/shape-text-alignment.pptx",
    ))
    .unwrap();
    let markdown = file.to_markdown().unwrap();
    assert!(markdown.starts_with("# Slide 1"), "{markdown}");
    assert_eq!(
        file.presentation().unwrap().to_markdown().unwrap(),
        markdown,
        "the view and the file render alike"
    );
    #[cfg(not(feature = "yaml"))]
    {
        let error = file
            .to_markdown_with_options(&MarkdownOptions::new().with_metadata(true))
            .expect_err("front matter needs the `yaml` feature");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Pptx));
    }
}

#[cfg(feature = "xlsx")]
#[test]
fn workbooks_are_refused_as_an_unsupported_capability() {
    let file = crate::open(test_data(
        "libreoffice-core/chart2/qa/extras/data/xlsx/boxWhisker.xlsx",
    ))
    .unwrap();
    let error = file
        .to_markdown()
        .expect_err("workbooks have no Markdown output");
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
    assert_eq!(error.format(), Some(litchi_core::FileFormat::Xlsx));
}

#[cfg(feature = "docx")]
#[test]
fn a_signed_file_renders_the_file_it_wraps() {
    let file = crate::open(test_data("poi/test-data/xmldsign/hello-world-signed.docx")).unwrap();
    let File::Signed(signed) = &file else {
        panic!("expected a signed file, got {file:?}");
    };
    let markdown = file.to_markdown().unwrap();
    assert!(!markdown.is_empty());
    assert_eq!(markdown, signed.file().to_markdown().unwrap());
}

#[cfg(any(
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "encryption"
))]
#[test]
fn a_locked_file_is_refused_as_protected() {
    use std::io::Cursor;

    // An encrypted OOXML container: `EncryptionInfo` and an
    // `EncryptedPackage`, without any real key material.
    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["EncryptionInfo"], &[4, 0, 4, 0, 0x40, 0, 0, 0])
        .unwrap();
    writer
        .create_stream(
            &["EncryptedPackage"],
            &[8, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 4, 5, 6, 7, 8],
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let file = crate::open_bytes(output.into_inner()).unwrap();
    assert!(matches!(file, File::Locked(_)), "{file:?}");
    let error = file
        .to_markdown()
        .expect_err("an encrypted file has no Markdown");
    assert_eq!(error.code(), Code::DocumentProtected, "{error:?}");
}
