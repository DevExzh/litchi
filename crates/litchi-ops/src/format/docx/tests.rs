//! `Docx` handle tests: every constructor route on real fixtures, parity with
//! the leaf output the facade wraps, the source-freshness guard, editor
//! construction, and the handle's size and auto-trait bounds.

use std::io::{Cursor, Write as _};
use std::path::PathBuf;
use std::sync::Arc;

use litchi_core::{FileFormat, OwnedSource, ReadAt};
use litchi_docx::section::Ownership;
use litchi_opc::{OpcPackage, SourceLimits};

use super::{Docx, DocxEditor, DocxOwner};
use crate::detect::{DetectedFormat, detect_format_smart};
use crate::error::{Code, Error, Source};
use crate::options::{OpenOptions, SourceMode};

const FIXTURES: &[&str] = &[
    "ooxml/docx/FancyFoot.docx",
    "ooxml/docx/documentProperties.docx",
    "ooxml/docx/headerFooter.docx",
    "ooxml/docx/table_footnotes.docx",
];

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(name)
}

fn minimal_ooxml(main_part: &str, content_type: &str, main_xml: &[u8]) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/{main_part}" ContentType="{content_type}"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{main_part}"/></Relationships>"#
    );
    writer.start_file("[Content_Types].xml", options).unwrap();
    writer.write_all(content_types.as_bytes()).unwrap();
    writer.start_file("_rels/.rels", options).unwrap();
    writer.write_all(root_relationships.as_bytes()).unwrap();
    writer.start_file(main_part, options).unwrap();
    writer.write_all(main_xml).unwrap();
    writer.finish().unwrap().into_inner()
}

const DOCX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

fn minimal_docx(document_xml: &[u8]) -> Vec<u8> {
    minimal_ooxml("word/document.xml", DOCX_MAIN, document_xml)
}

const SECTION_XML: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rHeader1"/><w:footerReference w:type="default" r:id="rFooter1"/><w:headerReference w:type="first" r:id="rHeader2"/><w:footerReference w:type="first" r:id="rFooter2"/><w:headerReference w:type="even" r:id="rHeader3"/><w:footerReference w:type="even" r:id="rFooter3"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="first" r:id="rHeader2"/><w:footerReference w:type="even" r:id="rFooter2"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#;

// ---- routes -------------------------------------------------------------

/// The path route: [`crate::open_with`], which `litchi::open` wraps.
#[cfg(any(unix, windows))]
fn open_path(path: &std::path::Path, options: &OpenOptions) -> Docx {
    match crate::open_with(path, options) {
        Ok(crate::File::Docx(docx)) => docx,
        other => panic!("{} is not opened as DOCX: {other:?}", path.display()),
    }
}

/// The source projection retained by either unsigned editor.
#[track_caller]
fn retained_source(docx: &Docx) -> litchi_docx::document::Snapshot {
    match docx.edit() {
        Ok(DocxEditor::Edit(edit)) => edit.source().clone(),
        Ok(DocxEditor::Source(edit)) => edit.source().clone(),
        other => panic!("expected a main-document edit, got {other:?}"),
    }
}

/// The byte route: [`crate::open_bytes_with`], which `litchi::open_bytes`
/// wraps.
fn open_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Result<Docx, Error> {
    match crate::open_bytes_with(bytes, options)? {
        crate::File::Docx(docx) => Ok(docx),
        other => panic!("bytes are not opened as DOCX: {other:?}"),
    }
}

/// The facade's eager route: `DetectedFormat::Docx` from smart detection.
fn open_eager(bytes: Vec<u8>, options: &OpenOptions) -> Result<Docx, Error> {
    match detect_format_smart(bytes) {
        Some(DetectedFormat::Docx(package)) => Docx::from_opc_package(package, options),
        _ => panic!("bytes are not detected as DOCX"),
    }
}

/// The facade's `Document::from_prepared` route, through the unified
/// `open_prepared_with`.
#[cfg(any(unix, windows))]
fn open_prepared(path: &std::path::Path, options: &OpenOptions) -> Result<Docx, Error> {
    let detection = crate::detect::detect_and_prepare(path).expect("a DOCX path is prepared");
    match crate::open_prepared_with(detection, options)? {
        crate::File::Docx(docx) => Ok(docx),
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

fn cold_loads(docx: &Docx) -> u64 {
    match docx.owner() {
        DocxOwner::Source(package) => package.cache_diagnostics().cold_loads,
        DocxOwner::Eager(_) => panic!("expected a source-backed owner"),
    }
}

// ---- leaf oracles (what the facade returns today) -----------------------

struct Oracle {
    text: String,
    metadata: String,
    tolerance: litchi_docx::ToleranceReport,
}

/// What `Docx::text` reads: each paragraph's text, separated by `\n`.
fn paragraph_lines(paragraphs: &[litchi_docx::Paragraph]) -> String {
    paragraphs
        .iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect::<Vec<_>>()
        .join("\n")
}

/// The eager facade output: `docx::Package` paragraphs, props, and the
/// report of the open-time text pass.
fn eager_oracle(bytes: &[u8]) -> Oracle {
    let package = litchi_docx::Package::from_opc_package(
        OpcPackage::from_bytes_with_limits(bytes, SourceLimits::default()).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    document.text().unwrap();
    let text = paragraph_lines(&document.paragraphs().unwrap());
    let metadata = package
        .props()
        .cloned()
        .map(litchi_core::Metadata::from)
        .unwrap_or_default();
    Oracle {
        text,
        metadata: format!("{metadata:?}"),
        tolerance: document.tolerance_report(),
    }
}

/// The source-backed facade output, through the leaf directly.
fn source_oracle(bytes: &[u8]) -> Oracle {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.to_vec()));
    let package = litchi_docx::source_backed::Package::from_read_at(source).unwrap();
    let text = paragraph_lines(&package.document().unwrap().paragraphs().unwrap());
    let metadata = package.metadata().unwrap();
    Oracle {
        text,
        metadata: format!("{metadata:?}"),
        tolerance: package.tolerance_report(),
    }
}

#[track_caller]
fn assert_matches(docx: &Docx, oracle: &Oracle) {
    assert_eq!(docx.text().unwrap(), oracle.text);
    let metadata = || {
        docx.metadata()
            .unwrap()
            .expect("DOCX metadata is always reported")
    };
    assert_eq!(format!("{:?}", metadata()), oracle.metadata);
    // A second call reads the sidecar and must agree.
    assert_eq!(format!("{:?}", metadata()), oracle.metadata);
    assert_eq!(docx.tolerance_report().unwrap(), oracle.tolerance);
}

// ---- tests --------------------------------------------------------------

const fn assert_handle<T: Send + Sync + Clone + std::fmt::Debug + 'static>() {}
const _: () = assert_handle::<Docx>();

#[test]
fn handle_is_one_pointer_and_clones_share_state() {
    assert_eq!(size_of::<Docx>(), size_of::<usize>());
    let docx = open_bytes(minimal_docx(SECTION_XML), &OpenOptions::new()).unwrap();
    let clone = docx.clone();
    assert!(Arc::ptr_eq(&docx.0, &clone.0));
    assert_eq!(docx.text().unwrap(), clone.text().unwrap());
    assert!(format!("{docx:?}").contains("source_backed: true"));
}

/// Oracle drift B3: the text keeps one line per paragraph, empty paragraphs
/// and table-cell paragraphs included, on both owners.
#[test]
fn text_separates_paragraphs_with_newlines() {
    const XML: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p/><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>last</w:t></w:r></w:p></w:body></w:document>"#;
    let source = open_bytes(minimal_docx(XML), &OpenOptions::new()).unwrap();
    let eager = open_eager(minimal_docx(XML), &OpenOptions::new()).unwrap();
    for docx in [source, eager] {
        assert_eq!(docx.text().unwrap(), "first\n\ncell\nlast");
    }
    let fixture = std::fs::read(fixture("ooxml/docx/Hyperlink.docx")).unwrap();
    let docx = open_bytes(fixture, &OpenOptions::new()).unwrap();
    assert_eq!(
        docx.text().unwrap(),
        "Following is a hyperlink.\nEricWhite.com"
    );
}

#[cfg(any(unix, windows))]
#[test]
fn path_route_adopts_the_source_owner_and_matches_the_facade() {
    for name in FIXTURES {
        let path = fixture(name);
        let bytes = std::fs::read(&path).unwrap();
        let docx = open_path(&path, &OpenOptions::new());
        assert!(docx.is_source_backed(), "{name}");
        assert_eq!(cold_loads(&docx), 0, "{name}: open reads no part payload");
        let source = source_oracle(&bytes);
        assert_matches(&docx, &source);
        // Source and eager facade routes agree on text and metadata.
        let eager = eager_oracle(&bytes);
        assert_eq!(source.text, eager.text, "{name}");
        assert_eq!(source.metadata, eager.metadata, "{name}");
    }
}

#[test]
fn bytes_route_adopts_the_source_owner_and_matches_the_facade() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).unwrap();
        let docx = open_bytes(bytes.clone(), &OpenOptions::new()).unwrap();
        assert!(docx.is_source_backed(), "{name}");
        assert_eq!(cold_loads(&docx), 0, "{name}: open reads no part payload");
        assert_matches(&docx, &source_oracle(&bytes));
    }
}

#[test]
fn eager_route_validates_at_open_and_matches_the_facade() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).unwrap();
        let docx = open_eager(bytes.clone(), &OpenOptions::new()).unwrap();
        assert!(!docx.is_source_backed(), "{name}");
        assert_matches(&docx, &eager_oracle(&bytes));
    }
}

#[cfg(any(unix, windows))]
#[test]
fn prepared_route_adopts_the_prepared_package_and_matches_the_facade() {
    for name in FIXTURES {
        let path = fixture(name);
        let bytes = std::fs::read(&path).unwrap();
        let docx = open_prepared(&path, &OpenOptions::new()).unwrap();
        assert!(docx.is_source_backed(), "{name}");
        assert_eq!(cold_loads(&docx), 0, "{name}: open reads no part payload");
        assert_matches(&docx, &source_oracle(&bytes));
    }
}

#[test]
fn eager_source_mode_materializes_the_owning_package() {
    let bytes = std::fs::read(fixture("ooxml/docx/FancyFoot.docx")).unwrap();
    let options = OpenOptions::new().source_mode(SourceMode::Eager);
    let docx = open_bytes(bytes.clone(), &options).unwrap();
    assert!(!docx.is_source_backed());
    assert_matches(&docx, &eager_oracle(&bytes));
}

#[test]
fn malformed_main_xml_is_deferred_on_source_and_refused_on_eager() {
    let bytes = minimal_docx(b"<w:document>");
    let docx = open_bytes(bytes.clone(), &OpenOptions::new())
        .expect("a source-backed open defers the main document");
    assert_eq!(cold_loads(&docx), 0);
    let error = docx.text().unwrap_err();
    assert_eq!(error.code(), Code::CorruptFile);
    assert_eq!(error.format(), Some(FileFormat::Docx));
    assert!(matches!(error.leaf(), Source::Docx(_)), "{error:?}");
    assert_eq!(
        docx.section_inventory().unwrap_err().code(),
        Code::CorruptFile
    );

    let error = open_eager(bytes, &OpenOptions::new()).unwrap_err();
    assert_eq!(error.code(), Code::CorruptFile);
}

#[test]
fn section_inventory_matches_on_both_owners() {
    let bytes = minimal_docx(SECTION_XML);
    let source = open_bytes(bytes.clone(), &OpenOptions::new()).unwrap();
    let eager = open_eager(bytes, &OpenOptions::new()).unwrap();
    let inventory = source.section_inventory().unwrap();
    assert_eq!(
        inventory.sections(),
        eager.section_inventory().unwrap().sections()
    );
    assert_eq!(inventory.paragraph_count(), 2);
    assert_eq!(inventory.sections().len(), 2);
    let first = inventory.section(0).unwrap();
    assert_eq!(
        first.ownership(),
        Ownership::Paragraph(litchi_core::Position::new(0))
    );
    assert_eq!(first.headers().len(), 3);
    assert_eq!(first.footers().len(), 3);
    let last = inventory.section(1).unwrap();
    assert_eq!(last.ownership(), Ownership::BodyFinal);
    assert_eq!(last.headers().len(), 1);
    assert!(inventory.section(usize::MAX).is_none());

    // The facade flattened this to `InvalidFormat`; it is now typed.
    let limits = litchi_docx::section::Limits {
        max_input_bytes: 1,
        ..Default::default()
    };
    for docx in [&source, &eager] {
        let error = docx.section_inventory_with_limits(&limits).unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
    }
}

#[cfg(any(unix, windows))]
#[test]
fn every_source_read_reports_a_changed_file() {
    let temporary = tempfile::NamedTempFile::new().unwrap();
    std::fs::copy(fixture("ooxml/docx/FancyFoot.docx"), temporary.path()).unwrap();
    let docx = open_path(temporary.path(), &OpenOptions::new());
    // Fill the metadata sidecar first: the guard must still apply to it.
    docx.metadata().unwrap();
    let DocxOwner::Source(package) = docx.owner() else {
        panic!("filesystem DOCX keeps the source owner");
    };
    let expected = package.source_version().unwrap();
    assert!(docx.owner().check_source().is_ok());

    let mut file = std::fs::OpenOptions::new()
        .append(true)
        .open(temporary.path())
        .unwrap();
    file.write_all(b"source mutation").unwrap();
    drop(file);

    let errors = [
        docx.text().unwrap_err(),
        docx.metadata().unwrap_err(),
        docx.tolerance_report().unwrap_err(),
        docx.section_inventory().unwrap_err(),
        docx.edit().unwrap_err(),
        docx.owner().check_source().unwrap_err(),
    ];
    for error in errors {
        assert_eq!(error.code(), Code::IoError, "{error:?}");
        let changed = match error.leaf() {
            Source::Docx(litchi_docx::Error::Opc(error)) | Source::Opc(error) => error,
            Source::DocxTransaction(litchi_docx::document::Error::Document(
                litchi_docx::Error::Opc(error),
            )) => error,
            other => panic!("expected a typed OPC source change, got {other:?}"),
        };
        match changed {
            litchi_opc::OpcError::SourceChanged {
                expected: was,
                actual,
            } => {
                assert_eq!(*was, expected);
                assert_ne!(*actual, expected);
            },
            other => panic!("expected SourceChanged, got {other:?}"),
        }
    }
}

#[test]
fn finish_prefers_the_source_failure_and_passes_eager_results_through() {
    let bytes = std::fs::read(fixture("ooxml/docx/FancyFoot.docx")).unwrap();
    let eager = open_eager(bytes.clone(), &OpenOptions::new()).unwrap();
    let error = Error::new(Code::InternalError, "operation failed");
    assert_eq!(
        eager.owner().finish::<()>(Err(error)).unwrap_err().code(),
        Code::InternalError
    );
    assert_eq!(eager.owner().finish(Ok(7)).unwrap(), 7);

    let source = open_bytes(bytes, &OpenOptions::new()).unwrap();
    assert_eq!(source.owner().finish(Ok(7)).unwrap(), 7);
}

#[test]
fn editor_is_built_from_either_owner_without_reindexing() {
    let bytes = minimal_docx(SECTION_XML);
    let source = open_bytes(bytes.clone(), &OpenOptions::new()).unwrap();
    let eager = open_eager(bytes, &OpenOptions::new()).unwrap();

    let source_edit = retained_source(&source);
    let eager_edit = retained_source(&eager);
    assert_eq!(source_edit.paragraph_count(), 2);
    assert_eq!(source_edit.paragraph_count(), eager_edit.paragraph_count());
    assert_eq!(source_edit.xml_bytes(), eager_edit.xml_bytes());
    // Building the editor read only the main part on the source owner.
    assert_eq!(cold_loads(&source), 1);
}

/// Record 0011 closed the leaf gap M0a recorded: both owners admit Word
/// documents whose markup compatibility cannot change content (a declared
/// namespace, a root `mc:Ignorable`) and build editors over the same bytes.
#[test]
fn editor_admits_word_markup_compatibility_on_both_owners() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).unwrap();
        let eager = open_eager(bytes.clone(), &OpenOptions::new()).unwrap();
        let source = open_bytes(bytes, &OpenOptions::new()).unwrap();
        let eager_edit = retained_source(&eager);
        let source_edit = match source.edit() {
            Ok(DocxEditor::Edit(edit)) => edit.source().clone(),
            Ok(DocxEditor::Source(edit)) => edit.source().clone(),
            other => panic!("{name}: {other:?}"),
        };
        assert!(source_edit.paragraph_count() > 0, "{name}");
        assert_eq!(
            source_edit.paragraph_count(),
            eager_edit.paragraph_count(),
            "{name}"
        );
        assert_eq!(source_edit.xml_bytes(), eager_edit.xml_bytes(), "{name}");
    }
}

/// Markup compatibility that selects content stays refused, as a typed leaf
/// error, whichever owner holds the document: here a body-level
/// `AlternateContent` whose branches hold different paragraphs.
#[test]
fn editor_refuses_alternate_content_on_both_owners() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:r><w:t>P0</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><w:p><w:r><w:t>P2</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = minimal_docx(document);
    let source = open_bytes(bytes.clone(), &OpenOptions::new()).unwrap();
    let eager = open_eager(bytes, &OpenOptions::new()).unwrap();
    // Reads apply the projection: the selected branch only.
    assert_eq!(source.text().unwrap(), eager.text().unwrap());
    assert!(!source.text().unwrap().contains("FALLBACK"));
    for docx in [&source, &eager] {
        let error = docx.edit().unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
        assert_eq!(error.format(), Some(FileFormat::Docx));
        assert!(
            matches!(
                error.leaf(),
                Source::DocxTransaction(litchi_docx::document::Error::Document(
                    litchi_docx::Error::UnsafeEdit { .. }
                ))
            ),
            "{error:?}"
        );
    }
}

#[test]
fn editor_refuses_when_the_source_is_not_retained() {
    let bytes = std::fs::read(fixture("ooxml/docx/FancyFoot.docx")).unwrap();
    let options = OpenOptions::new().retain_source(false);
    for docx in [
        open_bytes(bytes.clone(), &options).unwrap(),
        open_eager(bytes.clone(), &options).unwrap(),
    ] {
        let error = docx.edit().unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(FileFormat::Docx));
        assert!(error.suggestion().is_some());
        // Reads are unaffected.
        assert!(!docx.text().unwrap().is_empty());
    }
}

#[test]
fn transaction_errors_are_classified_and_typed() {
    use litchi_docx::document::Error as TransactionError;

    let cases = [
        (
            TransactionError::OutOfBounds {
                position: 3,
                len: 1,
            },
            Code::OutOfRange,
        ),
        (
            TransactionError::Limit {
                kind: litchi_docx::document::LimitKind::Operations,
                maximum: 1,
                observed: 2,
            },
            Code::LimitExceeded,
        ),
        (TransactionError::StaleSource, Code::IoError),
        (
            TransactionError::SemanticPrecondition,
            Code::InvalidArgument,
        ),
        (
            TransactionError::InvalidDurable("bad".into()),
            Code::InvalidInput,
        ),
        (
            TransactionError::Document(litchi_docx::Error::Xml("bad".into())),
            Code::CorruptFile,
        ),
    ];
    for (leaf, code) in cases {
        let error = Error::from(leaf);
        assert_eq!(error.code(), code, "{error:?}");
        assert_eq!(error.format(), Some(FileFormat::Docx));
        assert!(matches!(error.leaf(), Source::DocxTransaction(_)));
        assert!(std::error::Error::source(&error).is_some());
    }
}
