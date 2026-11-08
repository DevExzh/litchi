//! The facade's `document` in-source tests (`doc.rs`, `table.rs`,
//! `paragraph.rs`, `run.rs`), ported onto `File` and `DocumentView`.
//!
//! Each test keeps its facade semantics: the route a fixture takes, the
//! owner it ends in, and the result or typed refusal it reports. Facade
//! errors that were string-flattened are asserted through the typed code and
//! leaf instead. The Markdown tests move with `ToMarkdown` (step S7).

#[cfg(any(feature = "docx", feature = "odt"))]
use std::io::Cursor;
#[cfg(any(feature = "docx", feature = "odt"))]
use std::io::Write as _;

use super::test_data;
#[cfg(feature = "docx")]
use super::{content_type, minimal_ooxml};
use crate::compat::word::{CellMerge, Paragraph};
#[cfg(any(feature = "docx", feature = "odt"))]
use crate::compat::word::{DocumentElement, Table};
use crate::{Code, File};

fn open(relative: &str) -> File {
    crate::open(test_data(relative)).unwrap_or_else(|error| panic!("open {relative}: {error}"))
}

fn paragraph_signature(paragraph: &Paragraph) -> (String, Vec<String>) {
    (
        paragraph.text().expect("paragraph text"),
        paragraph
            .runs()
            .expect("paragraph runs")
            .into_iter()
            .map(|run| run.text().expect("run text"))
            .collect(),
    )
}

#[cfg(any(feature = "docx", feature = "odt"))]
fn table_text(table: &Table) -> String {
    let mut text = String::new();
    for row in table.rows().expect("table rows") {
        for cell in row.cells().expect("table cells") {
            text.push_str(&cell.text().expect("cell text"));
            text.push('|');
        }
    }
    text
}

#[cfg(any(feature = "docx", feature = "odt"))]
fn element_texts(file: &File) -> Vec<String> {
    file.document()
        .unwrap()
        .elements()
        .unwrap()
        .into_iter()
        .map(|element| match element {
            DocumentElement::Paragraph(paragraph) => format!("p:{}", paragraph.text().unwrap()),
            DocumentElement::Table(table) => format!("t:{}", table_text(&table)),
        })
        .collect()
}

#[cfg(any(feature = "docx", feature = "odt"))]
fn paragraph_texts(file: &File) -> Vec<String> {
    file.document()
        .unwrap()
        .paragraphs()
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

#[cfg(any(feature = "docx", feature = "odt"))]
fn table_texts(file: &File) -> Vec<String> {
    file.document()
        .unwrap()
        .tables()
        .unwrap()
        .iter()
        .map(table_text)
        .collect()
}

// ---- DOCX ---------------------------------------------------------------

#[cfg(feature = "docx")]
mod docx {
    use super::*;
    use crate::format::docx::{Docx, DocxOwner};
    use crate::{OpenOptions, SourceMode};

    pub(super) fn minimal_docx(document_xml: &[u8]) -> Vec<u8> {
        minimal_ooxml("word/document.xml", content_type::DOCX, document_xml, None)
    }

    pub(super) const SECTION_XML: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rHeader1"/><w:footerReference w:type="default" r:id="rFooter1"/><w:headerReference w:type="first" r:id="rHeader2"/><w:footerReference w:type="first" r:id="rFooter2"/><w:headerReference w:type="even" r:id="rHeader3"/><w:footerReference w:type="even" r:id="rFooter3"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="first" r:id="rHeader2"/><w:footerReference w:type="even" r:id="rFooter2"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#;

    pub(super) fn docx(file: &File) -> &Docx {
        match file {
            File::Docx(docx) => docx,
            other => panic!("expected DOCX, got {other:?}"),
        }
    }

    fn cold_loads(file: &File) -> u64 {
        match docx(file).owner() {
            DocxOwner::Source(package) => package.cache_diagnostics().cold_loads,
            DocxOwner::Eager(_) => panic!("expected the source-backed owner"),
        }
    }

    fn source_version(file: &File) -> litchi_core::SourceVersion {
        match docx(file).owner() {
            DocxOwner::Source(package) => package.source_version().expect("source version"),
            DocxOwner::Eager(_) => panic!("expected the source-backed owner"),
        }
    }

    /// The facade's eager oracle: `from_detected(DetectedFormat::Docx)`.
    fn eager(bytes: Vec<u8>) -> File {
        let file =
            crate::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
                .expect("open eager DOCX oracle");
        assert!(!docx(&file).is_source_backed());
        file
    }

    #[test]
    fn test_document_open_docx() {
        let file = open("ooxml/docx/FancyFoot.docx");
        assert!(matches!(file, File::Docx(_)));
    }

    #[test]
    fn owned_docx_bytes_use_source_owner_and_match_independent_eager_oracle() {
        let bytes = std::fs::read(test_data("ooxml/docx/FancyFoot.docx")).unwrap();
        let source = crate::open_bytes(bytes.clone()).expect("open source-backed DOCX bytes");
        let eager = eager(bytes);
        assert!(docx(&source).is_source_backed());
        assert_eq!(cold_loads(&source), 0);

        let text = source.text().expect("source-backed text");
        assert_eq!(text, eager.text().expect("eager text"));
        let after_first = cold_loads(&source);
        assert_eq!(after_first, 1);
        assert_eq!(source.text().unwrap(), text);
        assert_eq!(cold_loads(&source), after_first);
        assert_eq!(
            source.document().unwrap().paragraph_count().unwrap(),
            eager.document().unwrap().paragraph_count().unwrap()
        );
        assert_eq!(
            source.metadata().unwrap().unwrap().title,
            eager.metadata().unwrap().unwrap().title
        );
    }

    #[test]
    fn owned_docx_bytes_defer_malformed_main_xml_until_query() {
        let bytes = minimal_docx(b"<w:document>");
        let source = crate::open_bytes(bytes.clone()).expect("catalog-only source open");
        assert_eq!(cold_loads(&source), 0);
        assert!(
            source.text().is_err(),
            "malformed main XML must fail on the first semantic query"
        );
        assert!(
            crate::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
                .is_err(),
            "the independent eager oracle must reject malformed main XML during open"
        );
    }

    #[test]
    fn owned_docx_bytes_honor_limits_and_docx_precedence() {
        let bytes = minimal_docx(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>bounded</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_bytes_with(bytes.clone(), &OpenOptions::new().ooxml_limits(limits))
            .expect_err("an explicit input limit applies before source-backed publication");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                observed: actual,
                maximum: 1,
            }) if *actual == bytes.len() as u64
        ));

        #[cfg(feature = "odt")]
        {
            let polyglot = minimal_ooxml(
                "word/document.xml",
                content_type::DOCX,
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>OOXML wins</w:t></w:r></w:p></w:body></w:document>"#,
                Some("application/vnd.oasis.opendocument.text"),
            );
            let file = crate::open_bytes(polyglot).expect("DOCX precedence");
            assert!(docx(&file).is_source_backed());
            assert_eq!(file.text().unwrap(), "OOXML wins");
        }
    }

    #[test]
    fn owned_docx_wrong_family_is_terminal() {
        let bytes = minimal_ooxml(
            "ppt/presentation.xml",
            content_type::PPTX,
            br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            None,
        );
        // Detection names the family whether or not this build has the
        // PowerPoint owner; the document view refuses it either way.
        let error = match crate::open_bytes(bytes) {
            Ok(file) => file.document().expect_err("a PPTX is not a document"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), Some(litchi_core::FileFormat::Pptx));
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_uses_source_owner_and_matches_eager_projection() {
        let path = test_data("ooxml/docx/FancyFoot.docx");
        let source = crate::open(&path).expect("open source-backed DOCX");
        let eager = eager(std::fs::read(&path).unwrap());
        assert!(docx(&source).is_source_backed());
        let before = cold_loads(&source);
        assert_eq!(before, 0);
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
        assert!(cold_loads(&source) > before);
        assert_eq!(
            source.document().unwrap().paragraph_count().unwrap(),
            eager.document().unwrap().paragraph_count().unwrap()
        );
        assert_eq!(paragraph_texts(&source), paragraph_texts(&eager));
        assert_eq!(table_texts(&source), table_texts(&eager));
        assert_eq!(element_texts(&source), element_texts(&eager));
        let source_metadata = source.metadata().unwrap().unwrap();
        let eager_metadata = eager.metadata().unwrap().unwrap();
        assert_eq!(source_metadata.has_data(), eager_metadata.has_data());
        assert_eq!(source_metadata.title, eager_metadata.title);
    }

    #[test]
    fn paragraph_text_matches_eager_and_source_and_returns_none_out_of_range() {
        let bytes = minimal_docx(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let source = crate::open_bytes(bytes.clone()).expect("open source-backed DOCX");
        let eager = eager(bytes);
        assert_eq!(cold_loads(&source), 0);
        let (source_view, eager_view) = (source.document().unwrap(), eager.document().unwrap());
        for index in 0..2 {
            assert_eq!(
                source_view.paragraph_text(index).unwrap(),
                eager_view.paragraph_text(index).unwrap()
            );
        }
        assert_eq!(cold_loads(&source), 1);
        assert_eq!(source_view.paragraph_text(usize::MAX).unwrap(), None);
        assert_eq!(eager_view.paragraph_text(usize::MAX).unwrap(), None);
    }

    #[test]
    fn paragraph_matches_bulk_text_and_runs_for_eager_and_source() {
        let bytes = minimal_docx(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>middle </w:t></w:r><w:r><w:t>value</w:t></w:r></w:p><w:p><w:r><w:t>last</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let source = crate::open_bytes(bytes.clone()).expect("open source-backed DOCX");
        let eager = eager(bytes);
        let (source_view, eager_view) = (source.document().unwrap(), eager.document().unwrap());
        let source_bulk = source_view.paragraphs().unwrap();
        let eager_bulk = eager_view.paragraphs().unwrap();
        assert_eq!(source_bulk.len(), 3);
        assert_eq!(eager_bulk.len(), 3);
        for index in [0, 1] {
            let source_selected = source_view.paragraph(index).unwrap().unwrap();
            let eager_selected = eager_view.paragraph(index).unwrap().unwrap();
            assert_eq!(
                paragraph_signature(&source_selected),
                paragraph_signature(&source_bulk[index])
            );
            assert_eq!(
                paragraph_signature(&eager_selected),
                paragraph_signature(&eager_bulk[index])
            );
            assert_eq!(
                paragraph_signature(&source_selected),
                paragraph_signature(&eager_selected)
            );
        }
        assert!(source_view.paragraph(usize::MAX).unwrap().is_none());
        assert!(eager_view.paragraph(usize::MAX).unwrap().is_none());
    }

    #[test]
    fn docx_section_inventory_matches_eager_and_source_ownership() {
        use litchi_docx::section::Ownership;

        let bytes = minimal_docx(SECTION_XML);
        let source = crate::open_bytes(bytes.clone()).expect("open source-backed DOCX");
        let eager = eager(bytes);
        let source_inventory = docx(&source)
            .section_inventory(&litchi_docx::section::Limits::default())
            .unwrap();
        let eager_inventory = docx(&eager)
            .section_inventory(&litchi_docx::section::Limits::default())
            .unwrap();
        assert_eq!(source_inventory.sections(), eager_inventory.sections());
        assert_eq!(source_inventory.paragraph_count(), 2);
        assert_eq!(source_inventory.sections().len(), 2);
        let first = source_inventory.section(0).expect("first section");
        assert_eq!(
            first.ownership(),
            Ownership::Paragraph(litchi_core::Position::new(0))
        );
        assert_eq!(first.paragraphs().start().get(), 0);
        assert_eq!(first.paragraphs().end().get(), 1);
        assert!(first.page_size().is_some());
        assert!(first.margins().is_some());
        assert_eq!(first.headers().len(), 3);
        assert_eq!(first.footers().len(), 3);
        let last = source_inventory.section(1).expect("body-final section");
        assert_eq!(last.ownership(), Ownership::BodyFinal);
        assert_eq!(last.paragraphs().start().get(), 1);
        assert_eq!(last.paragraphs().end().get(), 2);
        assert!(last.margins().is_some());
        assert_eq!(last.headers().len(), 1);
        assert_eq!(last.footers().len(), 1);
        assert!(source_inventory.section(usize::MAX).is_none());
    }

    #[test]
    fn docx_header_footer_inventory_keeps_inert_relationship_ids() {
        let bytes = std::fs::read(test_data("ooxml/docx/headerFooter.docx")).unwrap();
        let file = eager(bytes);
        let inventory = docx(&file)
            .section_inventory(&litchi_docx::section::Limits::default())
            .unwrap();
        let mut ids = inventory
            .sections()
            .iter()
            .flat_map(|section| {
                section
                    .headers()
                    .iter()
                    .chain(section.footers().iter())
                    .map(|reference| reference.relationship_id.clone())
            })
            .collect::<Vec<_>>();
        ids.sort();
        assert_eq!(
            ids,
            (4..=9).map(|id| format!("rId{id}")).collect::<Vec<_>>()
        );
    }

    /// A budget-managed source-backed DOCX, built by the leaf as the facade
    /// test built it: `open` never builds managed owners (record 0008 A1),
    /// so the handle adopts one directly.
    fn managed(
        xml: &[u8],
        memory: u64,
        scope: &'static str,
    ) -> (litchi_core::CancellationSource, litchi_core::Budget, File) {
        let bytes = minimal_docx(xml);
        let budget = litchi_core::Budget::root(
            scope,
            litchi_core::Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = litchi_core::CancellationSource::pair();
        let execution_limits = litchi_core::ExecutionLimits::new(
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroU64::new(memory).unwrap(),
            0,
        )
        .unwrap();
        let context =
            litchi_core::ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let package = litchi_docx::source_backed::Package::from_read_at_with_options(
            std::sync::Arc::new(litchi_core::OwnedSource::new(bytes)),
            litchi_docx::OpenOptions::default()
                .source_limits(litchi_docx::SourceLimits::default())
                .execution_context(context),
        )
        .expect("open managed DOCX source");
        let file = File::Docx(Docx::from_source(package, &OpenOptions::new()).unwrap());
        (cancellation_source, budget, file)
    }

    #[test]
    fn managed_docx_section_inventory_owns_result_after_facade_drop() {
        let memory = (minimal_docx(SECTION_XML).len() as u64).saturating_mul(4);
        let (_cancellation, budget, file) =
            managed(SECTION_XML, memory, "facade-managed-docx-section-inventory");
        let inventory = docx(&file)
            .section_inventory(&litchi_docx::section::Limits::default())
            .unwrap();
        drop(file);
        assert_eq!(inventory.sections().len(), 2);
        assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
    }

    const MANAGED_XML: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>managed</w:t></w:r></w:p><w:p><w:r><w:t>selected</w:t></w:r></w:p></w:body></w:document>"#;

    /// On a managed source-backed DOCX the light `paragraph_text` projection
    /// is admitted while the rich paragraph views are refused, because those
    /// views cannot retain the managed payload reservation; a budget sized to
    /// the package alone refuses even the projection, with a typed memory
    /// refusal naming the scope.
    #[test]
    fn managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal() {
        const AMPLE_MEMORY: u64 = 1 << 20;
        let (_cancellation, budget, file) = managed(
            MANAGED_XML,
            AMPLE_MEMORY,
            "facade-managed-docx-paragraph-text",
        );
        let view = file.document().unwrap();
        assert_eq!(view.paragraph_text(1).unwrap().as_deref(), Some("selected"));
        assert!(budget.used(litchi_core::Resource::Memory) > 0);
        let paragraphs = view.paragraphs().expect_err("rich paragraphs are refused");
        assert!(
            paragraphs.message().contains("document paragraphs"),
            "{paragraphs:?}"
        );
        let paragraph = view.paragraph(1).expect_err("a rich paragraph is refused");
        assert!(
            paragraph.message().contains("document paragraph"),
            "{paragraph:?}"
        );
        drop(file);
        assert_eq!(budget.used(litchi_core::Resource::Memory), 0);

        let package_bytes = minimal_docx(MANAGED_XML).len() as u64;
        let (_tight_cancellation, tight_budget, tight) = managed(
            MANAGED_XML,
            package_bytes,
            "facade-managed-docx-paragraph-text",
        );
        let refusal = tight
            .document()
            .unwrap()
            .paragraph_text(1)
            .expect_err("a package-sized budget cannot admit the parser workspace");
        assert!(
            refusal.message().contains("Memory budget exceeded")
                && refusal
                    .message()
                    .contains("facade-managed-docx-paragraph-text"),
            "expected a typed memory-budget refusal naming the scope, got: {refusal:?}"
        );
        drop(tight);
        assert_eq!(tight_budget.used(litchi_core::Resource::Memory), 0);
    }

    #[test]
    fn docx_section_inventory_maps_deferred_malformed_and_tight_limits() {
        let malformed = crate::open_bytes(minimal_docx(b"<w:document>"))
            .expect("malformed main XML stays deferred at source open");
        assert_eq!(
            docx(&malformed)
                .section_inventory(&litchi_docx::section::Limits::default())
                .unwrap_err()
                .code(),
            Code::CorruptFile
        );

        let file = crate::open_bytes(minimal_docx(SECTION_XML)).unwrap();
        let limits = litchi_docx::section::Limits {
            max_input_bytes: 1,
            ..Default::default()
        };
        let error = docx(&file).section_inventory(&limits).unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
        assert!(error.message().contains("limit"), "{error:?}");
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_section_inventory_reports_source_mutation() {
        let temporary = super::super::temporary_copy(&test_data("ooxml/docx/FancyFoot.docx"), "");
        let file = crate::open(temporary.path()).unwrap();
        let expected = source_version(&file);
        super::super::mutate(temporary.path());
        super::super::assert_source_changed(
            &docx(&file)
                .section_inventory(&litchi_docx::section::Limits::default())
                .unwrap_err(),
            expected,
        );
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_reports_source_mutation_on_deferred_queries() {
        let temporary = super::super::temporary_copy(&test_data("ooxml/docx/FancyFoot.docx"), "");
        let file = crate::open(temporary.path()).unwrap();
        let expected = source_version(&file);
        file.text().expect("initial source query");
        super::super::mutate(temporary.path());
        let view = file.document().unwrap();
        for error in [
            file.text().unwrap_err(),
            file.metadata().unwrap_err(),
            view.paragraph_text(0).unwrap_err(),
            view.paragraph(0).unwrap_err(),
        ] {
            super::super::assert_source_changed(&error, expected);
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_defers_malformed_main_xml_until_query() {
        let bytes = minimal_docx(b"<w:document>");
        let temporary = super::super::temporary_with(&bytes, ".docx");
        let source = crate::open(temporary.path()).expect("catalog-only source open");
        assert!(docx(&source).is_source_backed());
        assert!(source.text().is_err(), "malformed main XML fails on query");
        assert!(
            crate::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
                .is_err(),
            "the eager oracle validates malformed main XML immediately"
        );
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_suffix_and_input_limits_keep_detector_precedence() {
        let temporary = super::super::temporary_with(b"not a ZIP", ".docx");
        let error = crate::open(temporary.path()).expect_err("an OOXML suffix needs ZIP magic");
        assert_eq!(error.code(), Code::CorruptFile);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::ZipError(_))
        ));

        let oversized = super::super::temporary_with(&[b'x'; 4096], ".docx");
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        // The OPC input limit applies before the suffix's ZIP validation.
        let error = crate::open_with(oversized.path(), &OpenOptions::new().ooxml_limits(limits))
            .expect_err("the input limit applies first");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                observed: 4096,
                maximum: 1,
            })
        ));
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_path_reports_typed_zip_and_part_errors() {
        let malformed = super::super::temporary_with(b"PK\x03\x04not a complete ZIP", ".docx");
        let error = crate::open(malformed.path()).expect_err("a broken ZIP is refused");
        assert_eq!(error.code(), Code::CorruptFile, "{error:?}");

        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#,
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
            .write_all(
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            )
            .unwrap();
        writer.start_file("word/styles.xml", options).unwrap();
        writer
            .write_all(br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#)
            .unwrap();
        let parts = super::super::temporary_with(&writer.finish().unwrap().into_inner(), ".docx");
        let limits = litchi_opc::SourceLimits::builder()
            .max_parts(1)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_with(parts.path(), &OpenOptions::new().ooxml_limits(limits))
            .expect_err("the part limit rejects the package");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::Parts,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_docx_preserves_other_ooxml_precedence_when_owner_disabled() {
        let bytes = minimal_ooxml(
            "ppt/presentation.xml",
            content_type::PPTX,
            br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            None,
        );
        let temporary = super::super::temporary_with(&bytes, ".docx");
        let error = match crate::open(temporary.path()) {
            Ok(file) => file
                .document()
                .expect_err("a PPTX named .docx is not a document"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), Some(litchi_core::FileFormat::Pptx));
    }

    #[cfg(all(feature = "odt", any(unix, windows)))]
    #[test]
    fn filesystem_docx_wins_over_a_lower_precedence_odf_marker() {
        let bytes = minimal_ooxml(
            "word/document.xml",
            content_type::DOCX,
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>OOXML wins</w:t></w:r></w:p></w:body></w:document>"#,
            Some("application/vnd.oasis.opendocument.text"),
        );
        let temporary = super::super::temporary_with(&bytes, ".odt");
        let file = crate::open(temporary.path()).expect("OOXML precedence");
        assert!(docx(&file).is_source_backed());
        assert_eq!(file.text().unwrap(), "OOXML wins");
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn extensionless_docx_path_uses_the_source_owner() {
        let bytes = minimal_docx(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>extensionless</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let temporary = super::super::temporary_with(&bytes, "");
        let file = crate::open(temporary.path()).expect("open extensionless DOCX");
        assert!(docx(&file).is_source_backed());
        assert_eq!(file.text().unwrap(), "extensionless");
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn extensionless_zip_without_content_types_uses_byte_fallback() {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("plain.txt", options).unwrap();
        writer.write_all(b"not an office package").unwrap();
        let temporary = super::super::temporary_with(&writer.finish().unwrap().into_inner(), "");
        let error = crate::open(temporary.path()).expect_err("not an Office file");
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn extensionless_non_ooxml_path_uses_the_neutral_fallback_ceiling() {
        let temporary = super::super::temporary_with(&[b'x'; 4096], "");
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_with(temporary.path(), &OpenOptions::new().ooxml_limits(limits))
            .expect_err("a non-OOXML input is still classified");
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }

    #[cfg(all(feature = "odt", any(unix, windows)))]
    #[test]
    fn docx_polyglot_wins_before_missing_or_malformed_odf_manifest() {
        let valid = minimal_ooxml(
            "word/document.xml",
            content_type::DOCX,
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>polyglot</w:t></w:r></w:p></w:body></w:document>"#,
            None,
        );
        let valid = odt::add_member(
            &valid,
            "mimetype",
            b"application/vnd.oasis.opendocument.text",
        );
        let malformed = odt::add_member(&valid, "META-INF/manifest.xml", b"<manifest:broken>");

        let from_bytes = crate::open_bytes(malformed.clone()).expect("DOCX byte precedence");
        assert!(docx(&from_bytes).is_source_backed());
        assert_eq!(from_bytes.text().unwrap(), "polyglot");

        let temporary = super::super::temporary_with(&malformed, ".odt");
        let from_path = crate::open(temporary.path()).expect("DOCX path precedence");
        assert!(docx(&from_path).is_source_backed());
        assert_eq!(from_path.text().unwrap(), "polyglot");
    }

    #[cfg(all(feature = "odt", any(unix, windows)))]
    #[test]
    fn docx_polyglot_without_odf_manifest_wins_for_bytes_and_path() {
        let bytes = odt::add_member(
            &minimal_ooxml(
                "word/document.xml",
                content_type::DOCX,
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>manifest absent</w:t></w:r></w:p></w:body></w:document>"#,
                None,
            ),
            "mimetype",
            b"application/vnd.oasis.opendocument.text",
        );
        let from_bytes = crate::open_bytes(bytes.clone()).expect("DOCX byte precedence");
        assert!(docx(&from_bytes).is_source_backed());
        assert_eq!(from_bytes.text().unwrap(), "manifest absent");

        let temporary = super::super::temporary_with(&bytes, ".odt");
        let from_path = crate::open(temporary.path()).expect("DOCX path precedence");
        assert!(docx(&from_path).is_source_backed());
        assert_eq!(from_path.text().unwrap(), "manifest absent");
    }

    #[cfg(feature = "odt")]
    #[test]
    fn owned_odt_bytes_do_not_use_docx_limits() {
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let file =
            crate::open_bytes_with(odt::minimal_odt(), &OpenOptions::new().ooxml_limits(limits))
                .expect("ODT uses the ODT policy");
        assert!(matches!(file, File::Odt(_)));
        assert_eq!(file.text().unwrap(), "Source-backed ODT");
        let view = file.document().unwrap();
        assert_eq!(
            view.paragraph_text(0).unwrap().as_deref(),
            Some("Source-backed ODT")
        );
        assert_eq!(view.paragraph_text(1).unwrap(), None);
    }

    #[cfg(feature = "odt")]
    #[test]
    fn owned_docx_bytes_do_not_use_odf_limits() {
        let bytes = minimal_docx(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>OPC precedence</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let polyglot = odt::add_member(
            &bytes,
            "mimetype",
            b"application/vnd.oasis.opendocument.text",
        );
        let limits = litchi_odf_common::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .max_members(1)
            .unwrap()
            .build()
            .unwrap();
        let options = OpenOptions::new()
            .odf_options(litchi_odf_common::OpenOptions::new().source_limits(limits));
        for input in [bytes, polyglot] {
            let file = crate::open_bytes_with(input, &options).expect("DOCX uses OPC limits");
            assert!(matches!(file, File::Docx(_)));
            assert_eq!(file.text().unwrap(), "OPC precedence");
        }
    }

    #[cfg(feature = "odt")]
    #[test]
    fn owned_odt_bytes_do_not_hide_malformed_ooxml_catalog() {
        let bytes = odt::add_member(
            &odt::minimal_odt(),
            "[Content_Types].xml",
            b"<Types><broken>",
        );
        let error = crate::open_bytes(bytes).expect_err("a malformed catalog is not hidden by ODT");
        assert_eq!(error.code(), Code::CorruptFile);
        assert!(
            error
                .message()
                .contains("Invalid [Content_Types].xml manifest:")
                && error.message().contains("root must be Types"),
            "{error:?}"
        );
    }

    #[cfg(all(feature = "odt", any(unix, windows)))]
    #[test]
    fn filesystem_odt_does_not_hide_malformed_ooxml_catalog() {
        let bytes = odt::add_member(
            &odt::minimal_odt(),
            "[Content_Types].xml",
            b"<Types><broken>",
        );
        let temporary = super::super::temporary_with(&bytes, ".odt");
        let error = crate::open(temporary.path()).expect_err("a malformed catalog is not hidden");
        assert_eq!(error.code(), Code::CorruptFile);
        assert!(
            error
                .message()
                .contains("Invalid [Content_Types].xml manifest:")
                && error.message().contains("root must be Types"),
            "{error:?}"
        );
    }

    #[cfg(all(feature = "odt", any(unix, windows)))]
    #[test]
    fn filesystem_odt_does_not_use_docx_limits() {
        let temporary = super::super::temporary_with(&odt::minimal_odt(), ".odt");
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let file = crate::open_with(temporary.path(), &OpenOptions::new().ooxml_limits(limits))
            .expect("ODT opening uses the ODT source policy");
        assert!(matches!(file, File::Odt(_)));
        assert_eq!(file.text().unwrap(), "Source-backed ODT");
    }

    #[test]
    fn test_document_from_bytes_docx() {
        let bytes = std::fs::read(test_data("ooxml/docx/FancyFoot.docx")).unwrap();
        assert!(matches!(crate::open_bytes(bytes).unwrap(), File::Docx(_)));
    }

    #[test]
    fn owned_docx_facade_survives_moves_and_repeated_reads() {
        fn move_file(file: File) -> File {
            file
        }
        let file = move_file(open("ooxml/docx/FancyFoot.docx"));
        let text = file.text().unwrap();
        assert!(!text.is_empty());
        assert_eq!(file.text().unwrap(), text);
        let view = file.document().unwrap();
        assert_eq!(
            view.paragraph_count().unwrap(),
            view.paragraphs().unwrap().len()
        );
        assert!(!view.elements().unwrap().is_empty());
        view.tables().unwrap();
        file.metadata().unwrap();
        // Clones share the opened state.
        let clone = file.clone();
        drop(file);
        assert_eq!(clone.text().unwrap(), text);
    }

    #[test]
    fn test_document_text_docx() {
        assert!(!open("ooxml/docx/FancyFoot.docx").text().unwrap().is_empty());
    }

    #[test]
    fn test_document_paragraph_count_docx() {
        let file = open("ooxml/docx/FancyFoot.docx");
        assert!(file.document().unwrap().paragraph_count().unwrap() > 0);
    }

    #[test]
    fn test_document_paragraphs_docx() {
        let file = open("ooxml/docx/FancyFoot.docx");
        let paragraphs = file.document().unwrap().paragraphs().unwrap();
        assert!(!paragraphs.is_empty());
        for paragraph in paragraphs {
            paragraph.text().unwrap();
        }
    }

    #[test]
    fn test_document_tables_docx() {
        let file = open("ooxml/docx/table_footnotes.docx");
        if let Some(table) = file.document().unwrap().tables().unwrap().first() {
            assert!(table.row_count().unwrap() > 0);
        }
    }

    #[test]
    fn test_document_elements_docx() {
        let file = open("ooxml/docx/FancyFoot.docx");
        let elements = file.document().unwrap().elements().unwrap();
        assert!(!elements.is_empty());
        for element in elements {
            assert!(element.is_paragraph() || element.is_table());
        }
    }

    #[test]
    fn test_document_metadata_docx() {
        let metadata = open("ooxml/docx/documentProperties.docx")
            .metadata()
            .unwrap()
            .expect("DOCX metadata is always reported");
        let _ = (metadata.title, metadata.author);
    }

    #[test]
    fn docx_text_reads_every_fixture_family() {
        for (fixture, nonempty) in [
            ("ooxml/docx/ComplexNumberedLists.docx", true),
            ("ooxml/docx/footnotes.docx", true),
            ("ooxml/docx/endnotes.docx", true),
            ("ooxml/docx/Headers.docx", false),
            ("ooxml/docx/headerFooter.docx", false),
            ("ooxml/docx/comment.docx", false),
            ("ooxml/docx/drawing.docx", true),
        ] {
            let file = open(fixture);
            let text = file
                .text()
                .unwrap_or_else(|error| panic!("{fixture}: {error}"));
            if nonempty {
                assert!(!text.is_empty(), "{fixture}");
            }
        }
        let lists = open("ooxml/docx/ComplexNumberedLists.docx");
        assert!(!lists.document().unwrap().paragraphs().unwrap().is_empty());
    }

    // ---- table.rs, paragraph.rs and run.rs tests (DOCX) -------------------

    #[test]
    fn docx_tables_rows_and_cells() {
        let file = open("ooxml/docx/table_footnotes.docx");
        let tables = file.document().unwrap().tables().unwrap();
        for table in &tables {
            assert!(table.row_count().unwrap() > 0);
            let rows = table.rows().unwrap();
            assert!(!rows.is_empty());
            assert!(table.row_at(0).unwrap().is_some());
            assert!(table.row_at(9999).unwrap().is_none());
            for row in &rows {
                assert!(row.cell_count().unwrap() > 0);
                let cells = row.cells().unwrap();
                assert!(!cells.is_empty());
                assert!(row.cell_at(0).unwrap().is_some());
                assert!(row.cell_at(9999).unwrap().is_none());
                for cell in &cells {
                    cell.text().unwrap();
                    assert!(cell.grid_span().unwrap() >= 1);
                }
            }
        }
    }

    #[test]
    fn test_table_document_with_tables() {
        for fixture in [
            "ooxml/docx/table_footnotes.docx",
            "ooxml/docx/table-alignment.docx",
        ] {
            let file = open(fixture);
            for table in file.document().unwrap().tables().unwrap() {
                assert!(table.row_count().unwrap() > 0, "{fixture}");
            }
        }
        // The orphaned core-properties relationship is read past and recorded,
        // so the table semantics and the metadata stay readable.
        let file = open("ooxml/docx/table-indent.docx");
        assert!(!file.document().unwrap().tables().unwrap().is_empty());
        file.metadata()
            .expect("the orphaned core part is read past");
    }

    /// DOCX absorbs covered columns into `w:gridSpan`: merged cells own their
    /// range and are never continuations.
    #[test]
    fn docx_grid_span_owners_are_never_continuations() {
        let file = open("ooxml/docx/drawing.docx");
        let mut widest = 1;
        for table in file.document().unwrap().tables().unwrap() {
            for row in &table.rows().unwrap() {
                for (index, cell) in row.cells().unwrap().iter().enumerate() {
                    let span = cell.grid_span().unwrap();
                    widest = widest.max(span);
                    assert_ne!(cell.horizontal_merge().unwrap(), CellMerge::Continuation);
                    assert_eq!(row.grid_span_at(index).unwrap(), Some(span));
                    assert!(cell.row_span().unwrap().is_none());
                }
            }
        }
        assert!(widest >= 3, "the fixture declares w:gridSpan=\"3\"");
    }

    /// The format-neutral vertical merge agrees with the DOCX-specific one.
    #[test]
    fn docx_vertical_merge_matches_the_format_specific_accessor() {
        let file = open("ooxml/docx/drawing.docx");
        for table in file.document().unwrap().tables().unwrap() {
            for row in &table.rows().unwrap() {
                for cell in &row.cells().unwrap() {
                    let expected = match cell.v_merge().unwrap() {
                        None => CellMerge::None,
                        Some(litchi_docx::VMergeState::Restart) => CellMerge::Start,
                        Some(litchi_docx::VMergeState::Continue) => CellMerge::Continuation,
                    };
                    assert_eq!(cell.vertical_merge().unwrap(), expected);
                }
            }
        }
    }

    #[test]
    fn docx_paragraphs_and_runs_read() {
        let file = open("ooxml/docx/FancyFoot.docx");
        for paragraph in file.document().unwrap().paragraphs().unwrap() {
            paragraph.text().unwrap();
            for run in paragraph.runs().unwrap() {
                run.text().unwrap();
                run.bold().unwrap();
                run.italic().unwrap();
                run.underline().unwrap();
                run.strikethrough().unwrap();
                run.vertical_position().unwrap();
            }
        }
    }

    /// A direct `w:u` is reported, `w:val="none"` as an explicit `false`,
    /// and a run without one inherits (`None`).
    #[test]
    fn docx_runs_report_direct_underline() {
        const XML: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>a</w:t></w:r><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r></w:p></w:body></w:document>"#;
        let file = crate::open_bytes(minimal_docx(XML)).unwrap();
        let paragraphs = file.document().unwrap().paragraphs().unwrap();
        let underline: Vec<_> = paragraphs[0]
            .runs()
            .unwrap()
            .iter()
            .map(|run| run.underline().unwrap())
            .collect();
        assert_eq!(underline, [Some(true), Some(false), None]);
    }
}

// ---- DOC ----------------------------------------------------------------

#[cfg(feature = "doc")]
mod doc {
    use super::*;
    use crate::{DocOptions, OpenOptions};

    #[test]
    fn paragraph_text_matches_doc_projection_and_bounds() {
        let file = open("ole/doc/FancyFoot.doc");
        let view = file.document().unwrap();
        let paragraphs = view.paragraphs().unwrap();
        assert!(!paragraphs.is_empty());
        for index in [0, paragraphs.len() - 1] {
            assert_eq!(
                view.paragraph_text(index).unwrap(),
                Some(paragraphs[index].text().unwrap())
            );
        }
        assert_eq!(view.paragraph_text(paragraphs.len()).unwrap(), None);
    }

    #[test]
    fn paragraph_matches_doc_projection_and_bounds() {
        let file = open("ole/doc/FancyFoot.doc");
        let view = file.document().unwrap();
        let paragraphs = view.paragraphs().unwrap();
        for index in [0, paragraphs.len() - 1] {
            let selected = view.paragraph(index).unwrap().expect("paragraph exists");
            assert_eq!(
                paragraph_signature(&selected),
                paragraph_signature(&paragraphs[index])
            );
        }
        assert!(view.paragraph(paragraphs.len()).unwrap().is_none());
    }

    #[test]
    fn test_document_open_and_bytes_doc() {
        assert!(matches!(open("ole/doc/FancyFoot.doc"), File::Doc(_)));
        let bytes = std::fs::read(test_data("ole/doc/FancyFoot.doc")).unwrap();
        assert!(matches!(crate::open_bytes(bytes).unwrap(), File::Doc(_)));
    }

    #[test]
    fn doc_leniency_is_available_through_the_unified_facade() {
        let path = test_data("ole/doc/duplicate-style-names.doc");
        assert!(crate::open(&path).is_err());
        let options = OpenOptions::new()
            .doc(DocOptions::new().leniency(litchi_doc::Leniency::TolerateStylesheetDefects));
        let from_path =
            crate::open_with(&path, &options).expect("lenient DOC options reach the open");
        let text = from_path.text().unwrap();
        assert!(!text.is_empty());
        let from_bytes = crate::open_bytes_with(std::fs::read(&path).unwrap(), &options)
            .expect("lenient DOC options reach the byte open");
        assert_eq!(from_bytes.text().unwrap(), text);
    }

    #[test]
    fn doc_admission_residues_open_by_default() {
        let file = open("ole/doc/watermark.doc");
        assert!(!file.text().unwrap().is_empty());
        let file = open("poi/test-data/document/test.doc");
        assert!(!file.text().unwrap().is_empty());
    }

    #[test]
    fn test_document_text_paragraphs_and_metadata_doc() {
        let file = open("ole/doc/FancyFoot.doc");
        assert!(!file.text().unwrap().is_empty());
        let view = file.document().unwrap();
        let paragraphs = view.paragraphs().unwrap();
        assert!(!paragraphs.is_empty());
        for paragraph in paragraphs {
            paragraph.text().unwrap();
            for run in paragraph.runs().unwrap() {
                run.text().unwrap();
                run.bold().unwrap();
                run.italic().unwrap();
                run.strikethrough().unwrap();
            }
        }
        let metadata = file
            .metadata()
            .unwrap()
            .expect("DOC metadata is always reported");
        let _ = (metadata.title, metadata.author);
        assert!(
            open("ole/doc/Lists.doc")
                .document()
                .unwrap()
                .paragraph_count()
                .unwrap()
                > 0
        );
    }

    #[test]
    fn test_document_open_nonexistent_file() {
        let error = crate::open(test_data("nonexistent_file.docx")).unwrap_err();
        assert_eq!(error.code(), Code::FileNotFound);
    }

    #[test]
    fn test_document_from_bytes_invalid_data() {
        let error = crate::open_bytes(b"This is not a valid document file".to_vec()).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }

    /// Binary DOC records `TC80` merge roles per cell; the span is only
    /// recoverable with row context.
    #[test]
    fn doc_merge_roles_resolve_through_row_context() {
        let file = open("ole/doc/table-merged-cells.doc");
        let tables = file.document().unwrap().tables().unwrap();
        assert!(!tables.is_empty());
        let (mut horizontal_starts, mut vertical_participants) = (0, 0);
        for table in &tables {
            for row in &table.rows().unwrap() {
                let cells = row.cells().unwrap();
                for (index, cell) in cells.iter().enumerate() {
                    match cell.horizontal_merge().unwrap() {
                        CellMerge::Start => {
                            horizontal_starts += 1;
                            let span = row.grid_span_at(index).unwrap().unwrap();
                            assert!(span >= 1 && span <= cells.len());
                        },
                        CellMerge::Continuation => {
                            assert_eq!(row.grid_span_at(index).unwrap(), None);
                        },
                        _ => assert_eq!(row.grid_span_at(index).unwrap(), Some(1)),
                    }
                    if cell.vertical_merge().unwrap().is_merged() {
                        vertical_participants += 1;
                    }
                }
            }
        }
        assert_eq!(horizontal_starts, 0);
        assert!(vertical_participants > 0);
    }
}

// ---- RTF ----------------------------------------------------------------

#[cfg(feature = "rtf")]
mod rtf {
    use super::*;

    #[test]
    fn paragraph_text_matches_rtf_projection_and_bounds() {
        let file = open("rtf/testUnicode.rtf");
        let view = file.document().unwrap();
        let paragraphs = view.paragraphs().unwrap();
        assert!(!paragraphs.is_empty());
        for index in [0, paragraphs.len() - 1] {
            assert_eq!(
                view.paragraph_text(index).unwrap(),
                Some(paragraphs[index].text().unwrap())
            );
        }
        assert_eq!(view.paragraph_text(paragraphs.len()).unwrap(), None);
    }

    #[test]
    fn paragraph_matches_rtf_projection_and_bounds() {
        let file = open("rtf/testUnicode.rtf");
        let view = file.document().unwrap();
        let paragraphs = view.paragraphs().unwrap();
        for index in [0, paragraphs.len() - 1] {
            let selected = view.paragraph(index).unwrap().expect("paragraph exists");
            assert_eq!(
                paragraph_signature(&selected),
                paragraph_signature(&paragraphs[index])
            );
        }
        assert!(view.paragraph(paragraphs.len()).unwrap().is_none());
    }

    #[test]
    fn test_document_open_and_bytes_rtf() {
        assert!(matches!(open("rtf/testUnicode.rtf"), File::Rtf(_)));
        let bytes = std::fs::read(test_data("rtf/testUnicode.rtf")).unwrap();
        assert!(matches!(crate::open_bytes(bytes).unwrap(), File::Rtf(_)));
    }

    #[test]
    fn unified_owned_rtf_route_accepts_all_supported_transport_forms() {
        let plain = br#"{\rtf1\ansi Plain owned route\par Second paragraph}"#.to_vec();
        let mut cp1252 = br#"{\rtf1\ansi\ansicpg1252 caf"#.to_vec();
        cp1252.extend_from_slice(&[0xe9, b'}']);
        let lzfu = litchi_rtf::transport::compress(
            &plain,
            true,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .unwrap();
        let mela = litchi_rtf::transport::compress(
            &plain,
            false,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .unwrap();
        for source in [plain, cp1252, lzfu, mela] {
            let file = crate::open_bytes(source).expect("owned RTF route");
            assert!(!file.text().unwrap().is_empty());
        }
    }

    #[test]
    fn unified_rtf_accepts_literal_cp1252_transport_bytes() {
        let mut source = br#"{\rtf1\ansi\ansicpg1252 caf"#.to_vec();
        source.extend_from_slice(&[0xe9, b'}']);
        assert_eq!(crate::open_bytes(source).unwrap().text().unwrap(), "café");
    }

    #[test]
    fn unified_rtf_compressed_transport_matches_plain_source() {
        let source = br#"{\rtf1\ansi Plain source\par Second paragraph}"#.to_vec();
        let native = litchi_rtf::Package::from_bytes(&source).unwrap();
        let plain = crate::open_bytes(source.clone()).unwrap();
        assert_eq!(plain.text().unwrap(), native.text());
        assert_eq!(
            plain.document().unwrap().paragraph_count().unwrap(),
            native.paragraph_count()
        );
        for use_lzfu in [true, false] {
            let compressed = litchi_rtf::transport::compress(
                &source,
                use_lzfu,
                litchi_rtf::transport::CompressionLimits::default(),
            )
            .unwrap();
            let file = crate::open_bytes(compressed).unwrap();
            assert_eq!(file.text().unwrap(), native.text());
            assert_eq!(
                file.document().unwrap().paragraph_count().unwrap(),
                native.paragraph_count()
            );
        }
    }

    #[test]
    fn unified_rtf_malformed_compressed_frame_is_a_typed_error() {
        let mut compressed = litchi_rtf::transport::compress(
            br#"{\rtf1\ansi malformed compressed frame}"#,
            true,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .unwrap();
        compressed.truncate(15);
        let error = crate::open_bytes(compressed).expect_err("a truncated frame fails");
        assert_eq!(error.code(), Code::CorruptFile);
        assert!(
            matches!(error.leaf(), crate::error::Source::Rtf(_)),
            "{error:?}"
        );
    }

    #[test]
    fn unified_rtf_metadata_preserves_info_without_inventing_timezones() {
        let source = concat!(
            r"{\rtf1\ansi{\*\generator Litchi 1.0;}{\info",
            r"{\title Unified title}{\subject Subject}{\author Ada}",
            r"{\operator Grace}{\keywords one,two}{\comment fallback comment}",
            r"{\doccomm Primary description}{\manager Lin}{\company ACME}",
            r"{\category Draft}{\hlinkbase https://example.test/base/}",
            r"{\creatim\yr2026\mo7\dy15\hr12\min34\sec56}",
            r"{\revtim\yr2026\mo7\dy16\hr9\min8}",
            r"{\printim\yr0\mo0\dy0\hr0\min0}",
            r"{\buptim\yr2026\mo7\dy17}",
            r"\version7\vern191\edmins42\nofpages3\nofwords9\nofchars44",
            r"\nofcharsws50\id77}Body}",
        );
        let metadata = crate::open_bytes(source.as_bytes().to_vec())
            .unwrap()
            .metadata()
            .unwrap()
            .expect("RTF metadata is always reported");
        assert_eq!(metadata.title.as_deref(), Some("Unified title"));
        assert_eq!(metadata.subject.as_deref(), Some("Subject"));
        assert_eq!(metadata.author.as_deref(), Some("Ada"));
        assert_eq!(metadata.last_modified_by.as_deref(), Some("Grace"));
        assert_eq!(metadata.description.as_deref(), Some("Primary description"));
        assert_eq!(metadata.manager.as_deref(), Some("Lin"));
        assert_eq!(metadata.company.as_deref(), Some("ACME"));
        assert_eq!(metadata.category.as_deref(), Some("Draft"));
        assert_eq!(
            metadata.hyperlink_base.as_deref(),
            Some("https://example.test/base/")
        );
        assert_eq!(metadata.revision.as_deref(), Some("7"));
        assert_eq!(metadata.version.as_deref(), Some("191"));
        assert_eq!(metadata.editing_time_minutes, Some(42));
        assert_eq!(metadata.page_count, Some(3));
        assert_eq!(metadata.word_count, Some(9));
        assert_eq!(metadata.character_count, Some(44));
        assert_eq!(metadata.character_count_with_spaces, Some(50));
        assert_eq!(metadata.identifier.as_deref(), Some("77"));
        assert_eq!(metadata.application.as_deref(), Some("Litchi 1.0"));
        assert_eq!(
            metadata
                .created
                .and_then(litchi_core::metadata::Timestamp::unzoned),
            chrono::NaiveDate::from_ymd_opt(2026, 7, 15)
                .unwrap()
                .and_hms_opt(12, 34, 56)
        );
        assert_eq!(
            metadata
                .modified
                .and_then(litchi_core::metadata::Timestamp::unzoned),
            chrono::NaiveDate::from_ymd_opt(2026, 7, 16)
                .unwrap()
                .and_hms_opt(9, 8, 0)
        );
        assert_eq!(
            metadata
                .last_backup
                .and_then(litchi_core::metadata::Timestamp::unzoned),
            chrono::NaiveDate::from_ymd_opt(2026, 7, 17)
                .unwrap()
                .and_hms_opt(0, 0, 0)
        );
        assert_eq!(
            metadata
                .created
                .and_then(litchi_core::metadata::Timestamp::utc),
            None
        );
        assert_eq!(
            metadata
                .modified
                .and_then(litchi_core::metadata::Timestamp::utc),
            None
        );
        assert_eq!(metadata.last_printed, None);
        assert!(metadata.has_data());
    }

    #[test]
    fn rtf_text_paragraphs_runs_and_tables_read() {
        for fixture in [
            "rtf/testUnicode.rtf",
            "rtf/testStyles.rtf",
            "rtf/testHex.rtf",
        ] {
            let file = open(fixture);
            file.text().unwrap();
        }
        let file = open("rtf/testUnicode.rtf");
        let paragraphs = file.document().unwrap().paragraphs().unwrap();
        assert!(!paragraphs.is_empty());
        for paragraph in paragraphs {
            for run in paragraph.runs().unwrap() {
                run.text().unwrap();
                run.bold().unwrap();
                run.italic().unwrap();
                run.underline().unwrap();
                run.strikethrough().unwrap();
            }
        }
        let file = open("rtf/chtoutline.rtf");
        file.text().unwrap();
        for table in file.document().unwrap().tables().unwrap() {
            assert!(table.row_count().unwrap() > 0);
            for row in &table.rows().unwrap() {
                let cells = row.cells().unwrap();
                assert!(!cells.is_empty());
                for cell in &cells {
                    cell.text().unwrap();
                }
            }
        }
        // `hyperlink.rtf` may use constructs the reader refuses; reading it
        // must not panic either way.
        if let Ok(file) = crate::open(test_data("rtf/hyperlink.rtf")) {
            file.text().unwrap();
        }
    }

    /// The `\ul` group of `tdf108947.rtf` is an underlined run; RTF states
    /// the property on every run, as it states bold and italic.
    #[test]
    fn rtf_runs_report_underline() {
        let file = open("libreoffice-core/sw/qa/extras/rtfimport/data/tdf108947.rtf");
        let paragraphs = file.document().unwrap().paragraphs().unwrap();
        let runs = paragraphs[0].runs().unwrap();
        assert_eq!(runs[0].text().unwrap(), "Sublet Agreement");
        assert_eq!(runs[0].underline().unwrap(), Some(true));
        let file = open("rtf/testStyles.rtf");
        let paragraphs = file.document().unwrap().paragraphs().unwrap();
        assert_eq!(
            paragraphs[0].runs().unwrap()[0].underline().unwrap(),
            Some(false)
        );
    }

    /// A row of two cells joined by `\clmgf` (range owner) and `\clmrg`
    /// (covered) spans two columns; an unversioned header is refused.
    #[test]
    fn rtf_horizontal_merge_roles_resolve_to_a_two_column_span() {
        let bytes = std::fs::read(test_data("rtf/table-cell-horizontal-merge.rtf")).unwrap();
        assert!(bytes.starts_with(b"{\\rtf1"));
        let mut unversioned = bytes.clone();
        unversioned.remove(b"{\\rtf1".len() - 1);
        assert!(crate::open_bytes(unversioned).is_err());

        let file = crate::open_bytes(bytes).unwrap();
        let tables = file.document().unwrap().tables().unwrap();
        let row = tables[0].row_at(0).unwrap().expect("one row");
        let cells = row.cells().unwrap();
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].horizontal_merge().unwrap(), CellMerge::Start);
        assert_eq!(
            cells[1].horizontal_merge().unwrap(),
            CellMerge::Continuation
        );
        assert_eq!(row.grid_span_at(0).unwrap(), Some(2));
        assert_eq!(row.grid_span_at(1).unwrap(), None);
        assert_eq!(row.grid_span_at(99).unwrap(), None);
    }

    #[test]
    fn rtf_vertical_merge_roles_are_reported() {
        let file = open("rtf/table-cell-vertical-merge.rtf");
        let mut starts = 0;
        for table in file.document().unwrap().tables().unwrap() {
            for row in &table.rows().unwrap() {
                for cell in &row.cells().unwrap() {
                    if cell.vertical_merge().unwrap() == CellMerge::Start {
                        starts += 1;
                    }
                }
            }
        }
        assert!(starts > 0);
    }
}

// ---- ODT ----------------------------------------------------------------

#[cfg(feature = "odt")]
mod odt {
    use super::*;
    use std::io::Read as _;

    const FLAT_ODT: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.3"
  office:mimetype="application/vnd.oasis.opendocument.text">
  <office:body><office:text><text:h>Title</text:h><text:p>Hello flat text</text:p></office:text></office:body>
</office:document>"#;

    #[test]
    fn owned_odt_source_limits_remain_typed() {
        use litchi_odf_common::{LimitKind, SourceLimitKind, SourceLimits};

        let bytes = minimal_odt();
        let length = u64::try_from(bytes.len()).unwrap();
        let members =
            u64::try_from(zip::ZipArchive::new(Cursor::new(&bytes)).unwrap().len()).unwrap();
        for (kind, observed) in [
            (SourceLimitKind::InputBytes, length),
            (SourceLimitKind::Members, members),
        ] {
            for maximum in [observed, observed - 1] {
                let limits = match kind {
                    SourceLimitKind::InputBytes => SourceLimits::builder().max_input_bytes(maximum),
                    SourceLimitKind::Members => {
                        SourceLimits::builder().max_members(usize::try_from(maximum).unwrap())
                    },
                    _ => unreachable!(),
                }
                .unwrap()
                .build()
                .unwrap();
                let options = crate::OpenOptions::new()
                    .odf_options(litchi_odf_common::OpenOptions::new().source_limits(limits));
                let result = crate::open_bytes_with(bytes.clone(), &options);
                if maximum == observed {
                    assert_eq!(result.unwrap().text().unwrap(), "Source-backed ODT");
                } else {
                    let error = result.expect_err("the selected source ceiling is enforced");
                    assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
                    assert!(
                        matches!(
                            error.leaf(),
                            crate::error::Source::Odf(litchi_odf_common::Error::Limit {
                                kind: actual_kind,
                                observed: actual_observed,
                                maximum: actual_maximum,
                            }) if *actual_kind == LimitKind::Source(kind)
                                && *actual_observed == observed && *actual_maximum == maximum
                        ),
                        "{error:?}"
                    );
                }
            }
        }
    }

    #[test]
    fn flat_odt_source_limits_remain_typed() {
        let length = u64::try_from(FLAT_ODT.len()).unwrap();
        for maximum in [length, length - 1] {
            let limits = litchi_odf_common::SourceLimits::builder()
                .max_input_bytes(maximum)
                .unwrap()
                .build()
                .unwrap();
            let options = crate::OpenOptions::new()
                .odf_options(litchi_odf_common::OpenOptions::new().source_limits(limits));
            let result = crate::open_bytes_with(FLAT_ODT.to_vec(), &options);
            if maximum == length {
                assert_eq!(result.unwrap().to_bytes().unwrap(), FLAT_ODT);
            } else {
                let error = result.expect_err("the selected flat input ceiling is enforced");
                assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
                assert!(
                    matches!(error.leaf(),
                        crate::error::Source::Odf(litchi_odf_common::Error::Limit {
                            kind: litchi_odf_common::LimitKind::Source(
                                litchi_odf_common::SourceLimitKind::InputBytes),
                            observed, maximum: actual_maximum,
                        }) if *observed == length && *actual_maximum == maximum
                    ),
                    "{error:?}"
                );
            }
        }
    }

    pub(super) fn minimal_odt() -> Vec<u8> {
        let mut builder = litchi_odt::Builder::new();
        builder.add_paragraph("Source-backed ODT").unwrap();
        builder.build().unwrap().to_bytes().unwrap()
    }

    fn malformed_odt() -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(litchi_odf_common::constants::ODF_TEXT.as_bytes())
            .unwrap();
        writer.start_file("META-INF/manifest.xml", options).unwrap();
        writer
            .write_all(br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#)
            .unwrap();
        writer.start_file("content.xml", options).unwrap();
        writer.write_all(b"<office:document-content>").unwrap();
        writer.finish().unwrap().into_inner()
    }

    pub(super) fn add_member(bytes: &[u8], path: &str, payload: &[u8]) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).unwrap();
            let name = entry.name().expect("valid ZIP member name").into_owned();
            let mut data = Vec::new();
            entry.read_to_end(&mut data).unwrap();
            writer.start_file(name, options).unwrap();
            writer.write_all(&data).unwrap();
        }
        writer.start_file(path, options).unwrap();
        writer.write_all(payload).unwrap();
        writer.finish().unwrap().into_inner()
    }

    fn is_source(file: &File) -> bool {
        matches!(
            file,
            File::Odt(odt) if !matches!(odt.owner(), crate::format::odt::OdtOwner::Eager(_))
        )
    }

    #[cfg(any(unix, windows))]
    fn source_version(file: &File) -> litchi_core::SourceVersion {
        match file {
            File::Odt(odt) => match odt.owner() {
                crate::format::odt::OdtOwner::Source(document) => {
                    document.source_version().unwrap()
                },
                crate::format::odt::OdtOwner::Eager(_) => panic!("expected the source owner"),
            },
            other => panic!("expected ODT, got {other:?}"),
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_source_matches_eager_projection() {
        let path = test_data("odf/odt/table-cell-column-span.odt");
        let source = crate::open(&path).unwrap();
        let eager = crate::open_bytes(std::fs::read(&path).unwrap()).unwrap();
        assert!(is_source(&source));
        assert!(!is_source(&eager));
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
        assert_eq!(
            source.document().unwrap().paragraph_count().unwrap(),
            eager.document().unwrap().paragraph_count().unwrap()
        );
        assert_eq!(paragraph_texts(&source), paragraph_texts(&eager));
        assert_eq!(table_texts(&source), table_texts(&eager));
        assert_eq!(element_texts(&source), element_texts(&eager));
        let source_metadata = source.metadata().unwrap().unwrap();
        let eager_metadata = eager.metadata().unwrap().unwrap();
        assert_eq!(source_metadata.has_data(), eager_metadata.has_data());
        assert_eq!(source_metadata.title, eager_metadata.title);
        assert_eq!(source_metadata.author, eager_metadata.author);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_paragraph_matches_eager_source_and_bulk_projection() {
        let path = test_data("odf/odt/table-cell-column-span.odt");
        let source = crate::open(&path).unwrap();
        let eager = crate::open_bytes(std::fs::read(&path).unwrap()).unwrap();
        let (source_view, eager_view) = (source.document().unwrap(), eager.document().unwrap());
        let source_bulk = source_view.paragraphs().unwrap();
        let eager_bulk = eager_view.paragraphs().unwrap();
        assert!(!source_bulk.is_empty());
        assert_eq!(source_bulk.len(), eager_bulk.len());
        let mut indexes = vec![0, source_bulk.len() / 2, source_bulk.len() - 1];
        indexes.dedup();
        for index in indexes {
            let source_selected = source_view.paragraph(index).unwrap().unwrap();
            let eager_selected = eager_view.paragraph(index).unwrap().unwrap();
            assert_eq!(
                paragraph_signature(&source_selected),
                paragraph_signature(&source_bulk[index])
            );
            assert_eq!(
                paragraph_signature(&eager_selected),
                paragraph_signature(&eager_bulk[index])
            );
            assert_eq!(
                paragraph_signature(&source_selected),
                paragraph_signature(&eager_selected)
            );
        }
        assert!(source_view.paragraph(source_bulk.len()).unwrap().is_none());
        assert!(eager_view.paragraph(eager_bulk.len()).unwrap().is_none());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_paragraph_reports_source_mutation() {
        let temporary =
            super::super::temporary_copy(&test_data("odf/odt/table-cell-column-span.odt"), "");
        let file = crate::open(temporary.path()).unwrap();
        let expected = source_version(&file);
        assert!(file.document().unwrap().paragraph(0).unwrap().is_some());
        super::super::mutate(temporary.path());
        super::super::assert_source_changed(
            &file.document().unwrap().paragraph(0).unwrap_err(),
            expected,
        );
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_keeps_unselected_media_cold() {
        let bytes = add_member(&minimal_odt(), "Pictures/deferred.bin", b"not read yet");
        let temporary = super::super::temporary_with(&bytes, "");
        let file = crate::open(temporary.path()).unwrap();
        let media = match &file {
            File::Odt(odt) => match odt.owner() {
                crate::format::odt::OdtOwner::Source(document) => document.media_files().unwrap(),
                crate::format::odt::OdtOwner::Eager(_) => panic!("expected the source owner"),
            },
            other => panic!("expected ODT, got {other:?}"),
        };
        assert!(media.iter().any(|path| path == "Pictures/deferred.bin"));
        let view = file.document().unwrap();
        assert_eq!(
            view.paragraph_text(0).unwrap().as_deref(),
            Some("Source-backed ODT")
        );
        assert_eq!(view.paragraph_text(1).unwrap(), None);
        assert_eq!(file.text().unwrap(), "Source-backed ODT");
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_ignores_suffix_and_accepts_extensionless_paths() {
        let bytes = minimal_odt();
        for suffix in ["", ".wrong"] {
            let temporary = super::super::temporary_with(&bytes, suffix);
            let file = crate::open(temporary.path()).expect("open ODT by package MIME");
            assert!(is_source(&file));
            assert_eq!(file.text().unwrap(), "Source-backed ODT");
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_refuses_malformed_content_with_typed_error() {
        let bytes = malformed_odt();
        let temporary = super::super::temporary_with(&bytes, ".odt");
        let error = match crate::open(temporary.path()) {
            Ok(file) => file.text().expect_err("malformed ODT is refused"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
        assert!(crate::open_bytes(bytes).is_err());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_odt_reports_source_mutation_on_deferred_queries() {
        let temporary = super::super::temporary_with(&minimal_odt(), "");
        let file = crate::open(temporary.path()).unwrap();
        let expected = source_version(&file);
        file.text().expect("initial source query");
        super::super::mutate(temporary.path());
        for error in [
            file.text().unwrap_err(),
            file.metadata().unwrap_err(),
            file.document().unwrap().paragraph_text(0).unwrap_err(),
        ] {
            super::super::assert_source_changed(&error, expected);
        }
    }

    #[cfg(all(any(feature = "ods", feature = "odp"), any(unix, windows)))]
    #[test]
    fn filesystem_odt_probe_does_not_claim_other_odf_families() {
        #[cfg(feature = "ods")]
        let (mimetype, body, expected) = (
            litchi_odf_common::constants::ODF_SPREADSHEET,
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet/></office:body></office:document-content>"#.as_slice(),
            litchi_core::FileFormat::Ods,
        );
        #[cfg(all(not(feature = "ods"), feature = "odp"))]
        let (mimetype, body, expected) = (
            litchi_odf_common::constants::ODF_PRESENTATION,
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#.as_slice(),
            litchi_core::FileFormat::Odp,
        );
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(litchi_odf_common::constants::ODF_CONTENT, body)
            .unwrap();
        let temporary = super::super::temporary_with(&writer.finish_to_bytes().unwrap(), "");
        let file = crate::open(temporary.path()).unwrap();
        assert_eq!(file.format(), Some(expected));
        let error = file
            .document()
            .expect_err("another ODF family is not a document");
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), Some(expected));
    }

    /// `filesystem_other_ooxml_wins_without_docx_feature`: a PPTX that also
    /// carries an ODT marker is a presentation, never claimed as ODT.
    #[cfg(all(feature = "pptx", any(unix, windows)))]
    #[test]
    fn filesystem_other_ooxml_wins_over_an_odt_marker() {
        use super::super::{content_type, minimal_ooxml};

        let bytes = minimal_ooxml(
            "ppt/presentation.xml",
            content_type::PPTX,
            br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            Some("application/vnd.oasis.opendocument.text"),
        );
        let temporary = super::super::temporary_with(&bytes, ".odt");
        let error = match crate::open(temporary.path()) {
            Ok(file) => file
                .document()
                .map(drop)
                .expect_err("a PPTX is not a document"),
            Err(error) => error,
        };
        assert_eq!(
            error.format().or(error.detected()),
            Some(litchi_core::FileFormat::Pptx)
        );
    }

    #[test]
    fn flat_odt_opens_through_its_read_projection() {
        // Record 0292: flat ODT reads through the leaf's packaged projection
        // and publishes its flat bytes.
        let file = crate::open_bytes(FLAT_ODT.to_vec()).expect("flat ODT opens");
        assert_eq!(file.format(), Some(litchi_core::FileFormat::Odt));
        assert_eq!(file.text().unwrap(), "Title\nHello flat text");
        assert_eq!(file.to_bytes().unwrap(), FLAT_ODT);

        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(FLAT_ODT, ".fodt");
            let file = crate::open(temporary.path()).expect("flat ODT path opens");
            assert_eq!(file.to_bytes().unwrap(), FLAT_ODT);
        }
    }

    /// ODF stores `table:number-columns-spanned`, so the span is readable
    /// from the cell alone.
    #[test]
    fn odt_column_span_is_read_from_the_stored_count() {
        let file = open("odf/odt/table-cell-column-span.odt");
        let mut spanned = Vec::new();
        for (table_index, table) in file
            .document()
            .unwrap()
            .tables()
            .unwrap()
            .iter()
            .enumerate()
        {
            for row in &table.rows().unwrap() {
                for (index, cell) in row.cells().unwrap().iter().enumerate() {
                    let span = cell.grid_span().unwrap();
                    if span > 1 {
                        spanned.push((table_index, span));
                        assert_eq!(cell.horizontal_merge().unwrap(), CellMerge::Start);
                        assert_eq!(row.grid_span_at(index).unwrap(), Some(span));
                    }
                    assert!(cell.row_span().unwrap().is_some());
                }
            }
        }
        assert!(spanned.iter().any(|(_, span)| *span == 3), "{spanned:?}");
    }
}
