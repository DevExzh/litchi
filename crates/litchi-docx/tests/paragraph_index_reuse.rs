use std::io::{self, Cursor, Write};
use std::sync::Arc;
use std::sync::Mutex;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion, TextOutputOptions,
};
use litchi_docx::{Error, Package, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAIN: &str = "/word/document.xml";

const FINITE_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OBJECTS: u64 = 1_000_000;
const FINITE_DEPTH: u64 = 1024;
const FINITE_WORK: u64 = 1 << 30;
const SOURCE_DOCUMENT_SCAN_WORKSPACE_BASE: u64 = 131_072;
const SOURCE_DOCUMENT_SCAN_WORKSPACE_PER_BYTE: u64 = 32;

fn document_xml(paragraphs: &[&str]) -> Vec<u8> {
    let body = paragraphs
        .iter()
        .map(|text| format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#))
        .collect::<String>();
    format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#).into_bytes()
}

fn docx_bytes(document: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document,
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn docx_bytes_with_core(document: Vec<u8>, core: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document,
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_owned(),
            core,
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    package.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
    PackageWriter::to_bytes(&package).unwrap()
}

fn core_properties_xml(title: &str) -> Vec<u8> {
    format!(
        r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title><dc:creator>Ada</dc:creator><cp:revision>7</cp:revision></cp:coreProperties>"#
    )
    .into_bytes()
}

fn docx_bytes_with_alternate(primary: Vec<u8>, alternate: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            primary,
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/alternate.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            alternate,
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn eager_texts(document: &litchi_docx::Document<'_>) -> Vec<String> {
    document
        .paragraphs()
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

fn source_texts(document: &source_backed::Document) -> Vec<String> {
    document
        .paragraphs()
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

fn assert_eager_view(document: &litchi_docx::Document<'_>, expected: &[&str]) {
    assert_eq!(document.paragraph_count().unwrap(), expected.len());
    assert_eq!(document.text().unwrap(), expected.concat());
    assert_eq!(eager_texts(document), expected);
    for (index, expected_text) in expected.iter().enumerate() {
        assert_eq!(
            document.paragraph(index).unwrap().unwrap().text().unwrap(),
            *expected_text
        );
    }
}

fn assert_source_view(document: &source_backed::Document, expected: &[&str]) {
    assert_eq!(document.paragraph_count().unwrap(), expected.len());
    assert_eq!(document.extract_text().unwrap(), expected.concat());
    assert_eq!(source_texts(document), expected);
    for (index, expected_text) in expected.iter().enumerate() {
        assert_eq!(
            document.paragraph_text(index).unwrap().as_deref(),
            Some(*expected_text)
        );
        assert_eq!(
            document.paragraph(index).unwrap().unwrap().text().unwrap(),
            *expected_text
        );
    }
    assert!(document.paragraph(expected.len()).unwrap().is_none());
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-paragraph-index-reuse-test",
        Limits::new(
            memory,
            FINITE_INPUT_BYTES,
            FINITE_OUTPUT_BYTES,
            FINITE_OBJECTS,
            FINITE_DEPTH,
            FINITE_WORK,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn source_document_scan_workspace(xml_len: usize) -> u64 {
    (xml_len as u64)
        .saturating_mul(SOURCE_DOCUMENT_SCAN_WORKSPACE_PER_BYTE)
        .saturating_add(SOURCE_DOCUMENT_SCAN_WORKSPACE_BASE)
}

fn source_document_index_admission(xml_len: usize) -> u64 {
    (xml_len as u64 / 4 + 1)
        .saturating_mul(24)
        .saturating_add(1024)
}

struct FailingWriter;

impl Write for FailingWriter {
    fn write(&mut self, _buffer: &[u8]) -> io::Result<usize> {
        Err(io::Error::other("injected DOCX save sink failure"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn repeated_eager_and_source_views_keep_count_and_text_stable() {
    let expected = ["first", "second", "third"];
    let bytes = docx_bytes(document_xml(&expected));

    let eager = Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    for _ in 0..4 {
        assert_eager_view(&eager.document().unwrap(), &expected);
    }

    let source = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    for _ in 0..4 {
        assert_source_view(&source.document().unwrap(), &expected);
    }
}

#[test]
fn bom_and_mce_visibility_are_preserved_across_repeated_index_queries() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:body></w:document>"#
    );
    let mut marked = b"\xEF\xBB\xBF".to_vec();
    marked.extend_from_slice(xml.as_bytes());
    let bytes = docx_bytes(marked);
    let expected = ["fallback", "tail"];

    let eager = Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    for _ in 0..3 {
        assert_eager_view(&eager.document().unwrap(), &expected);
    }

    let source = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    for _ in 0..3 {
        assert_source_view(&source.document().unwrap(), &expected);
    }
}

#[test]
fn eager_mutable_save_replaces_cached_ranges_and_survives_reopen() {
    let mut package =
        Package::from_reader(Cursor::new(docx_bytes(document_xml(&["before"])))).unwrap();
    assert_eager_view(&package.document().unwrap(), &["before"]);

    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("after");
    let output = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    package.save(output.path()).unwrap();

    assert_eager_view(&package.document().unwrap(), &["before", "after"]);
    let reopened = Package::open(output.path()).unwrap();
    assert_eager_view(&reopened.document().unwrap(), &["before", "after"]);
}

#[test]
fn eager_raw_edit_replaces_cached_ranges_and_keeps_old_snapshot_stable() {
    let mut package =
        Package::from_reader(Cursor::new(docx_bytes(document_xml(&["before"])))).unwrap();
    assert_eager_view(&package.document().unwrap(), &["before"]);
    let old_snapshot = package.document_snapshot().unwrap();

    let main = PackURI::new(MAIN).unwrap();
    package
        .edit_opc(|candidate| {
            candidate
                .get_part_mut(&main)?
                .set_blob(document_xml(&["after", "new tail"]));
            Ok::<_, Error>(())
        })
        .unwrap();

    assert_eq!(old_snapshot.paragraph_count(), 1);
    assert_eq!(
        old_snapshot
            .paragraph(litchi_core::Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "before"
    );
    assert_eager_view(&package.document().unwrap(), &["after", "new tail"]);
}

#[test]
fn eager_root_relationship_retarget_replaces_cached_ranges_and_survives_save() {
    let bytes = docx_bytes_with_alternate(
        document_xml(&["primary"]),
        document_xml(&["alternate", "second alternate"]),
    );
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    assert_eager_view(&package.document().unwrap(), &["primary"]);

    package
        .edit_opc(|candidate| {
            let office_document_id = candidate
                .rels()
                .iter()
                .find(|relationship| relationship.reltype() == rt::OFFICE_DOCUMENT)
                .map(|relationship| relationship.r_id().to_owned())
                .expect("fixture must have one officeDocument relationship");
            candidate
                .relationships_mut()
                .retarget(&office_document_id, "word/alternate.xml".to_owned())?;
            Ok::<_, Error>(())
        })
        .unwrap();

    assert_eager_view(
        &package.document().unwrap(),
        &["alternate", "second alternate"],
    );
    let output = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert_eager_view(
        &reopened.document().unwrap(),
        &["alternate", "second alternate"],
    );
}

#[test]
fn eager_failed_save_keeps_warmed_generation_retryable() {
    let mut package =
        Package::from_reader(Cursor::new(docx_bytes(document_xml(&["before"])))).unwrap();
    assert_eager_view(&package.document().unwrap(), &["before"]);

    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("after");
    assert!(package.to_stream(FailingWriter).is_err());

    // The failed sink must leave the published OPC generation unchanged. The
    // staged mutable document remains retryable, while the warmed index still
    // describes the old XML rather than an unpublished candidate.
    assert_eager_view(&package.document().unwrap(), &["before"]);

    let output = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    package.save(output.path()).unwrap();
    assert_eager_view(&package.document().unwrap(), &["before", "after"]);

    let reopened = Package::open(output.path()).unwrap();
    assert_eager_view(&reopened.document().unwrap(), &["before", "after"]);
}

#[derive(Debug)]
struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn replace(&self, bytes: Vec<u8>) {
        *self.bytes.lock().unwrap() = bytes;
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        let bytes = self.bytes.lock().unwrap();
        if offset >= bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(bytes.len());
        output[..end - offset].copy_from_slice(&bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x0682,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn source_change_after_a_warm_index_refuses_new_view_but_old_view_stays_pinned() {
    let source = Arc::new(MutableSource::new(docx_bytes(document_xml(&["before"]))));
    let package = source_backed::Package::from_read_at(source.clone()).unwrap();
    let first = package.document().unwrap();
    assert_source_view(&first, &["before"]);
    let old_view = first.clone();
    drop(first);

    source.replace(docx_bytes(document_xml(&["after", "new tail"])));
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert_source_view(&old_view, &["before"]);
}

#[test]
fn cancellation_after_a_warm_index_is_checked_before_cache_reuse() {
    let bytes = docx_bytes(document_xml(&["managed", "read"]));
    let (budget, cancellation_source, context) = managed_context(1 << 20);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.paragraph_count().unwrap(), 2);
    let old_view = document.clone();
    drop(document);

    cancellation_source.cancel();
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        old_view.paragraph_count(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    drop(old_view);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_clean_index_pressure_does_not_refuse_a_following_text_sink() {
    let document = document_xml(&["managed", "read"]);
    let bytes = docx_bytes(document.clone());
    let xml_len = document.len();
    let index_admission = source_document_index_admission(xml_len);
    let parser_workspace = source_document_scan_workspace(xml_len);
    let memory_limit = (xml_len as u64)
        .saturating_add(index_admission)
        .saturating_add(parser_workspace);
    let (budget, _cancellation_source, context) = managed_context(memory_limit);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();

    let warm = package.document().unwrap();
    assert_eq!(warm.paragraph_count().unwrap(), 2);
    drop(warm);

    // This public reservation leaves exactly enough room for the following
    // source-backed text pass only if a clean memo can release its admission.
    let hold = budget.reserve(Resource::Memory, index_admission).unwrap();
    let mut output = Vec::new();
    package
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "", u64::MAX, u64::MAX),
        )
        .expect("clean paragraph-index retention must not block text output");
    assert_eq!(output, b"managed\nread");

    drop(hold);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_clean_index_pressure_does_not_block_core_metadata_read() {
    let document = document_xml(&["managed", "read"]);
    let title = "managed metadata ".repeat(16);
    let core = core_properties_xml(&title);
    // The core payload must force the OPC payload cache to give up the main
    // payload after a failed admission. The index is still the reservation
    // that makes this operation fail unless the DOCX cache trims its clean
    // generation first.
    assert!(core.len() > document.len());
    let bytes = docx_bytes_with_core(document.clone(), core);
    let xml_len = document.len();
    let index_admission = source_document_index_admission(xml_len);
    let parser_workspace = source_document_scan_workspace(xml_len);
    let memory_limit = (xml_len as u64)
        .saturating_add(index_admission)
        .saturating_add(parser_workspace);
    let (budget, _cancellation_source, context) = managed_context(memory_limit);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();

    let warm = package.document().unwrap();
    assert_eq!(warm.paragraph_count().unwrap(), 2);
    drop(warm);
    assert!(package.cache_diagnostics().retained_entries >= 1);
    assert!(budget.used(Resource::Memory) >= (xml_len as u64).saturating_add(index_admission));

    // Consume every currently free byte. The metadata payload can fit after
    // the clean paragraph index is released, while the same reservation keeps
    // the warmed index from being silently bypassed.
    let free = budget
        .limit(Resource::Memory)
        .saturating_sub(budget.used(Resource::Memory));
    let hold = budget.reserve(Resource::Memory, free).unwrap();
    let metadata = package
        .metadata()
        .expect("metadata must trim an unused paragraph index before reading core XML");
    assert_eq!(metadata.title.as_deref(), Some(title.as_str()));

    drop(hold);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}
