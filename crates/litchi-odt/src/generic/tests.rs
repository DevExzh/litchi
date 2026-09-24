//! Regression tests for the generic ODF snapshot owners.

use super::{Family, FlatDocument, Package};
use crate::constants;
use crate::core::PackageWriter;
use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    Position, Resource,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::io::{Cursor, Read};
use std::num::{NonZeroU64, NonZeroUsize};

fn replace_zip_member_raw(package: &[u8], path: &str, replacement: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(package).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    let mut replaced = false;
    for name in archive.file_names() {
        let data = if name == path {
            replaced = true;
            replacement.to_vec()
        } else {
            archive.read(name).unwrap()
        };
        writer.write_stored(name, &data).unwrap();
    }
    assert!(replaced, "test ZIP member {path} must exist");
    writer.finish_to_bytes().unwrap()
}

fn package(mimetype: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer
        .add_file(
            constants::ODF_CONTENT,
            br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#,
        )
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.png", b"PNG", "image/png")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn opens_every_document_family_and_template_losslessly() {
    for (mimetype, family, template) in [
        (constants::ODF_TEXT, Family::Text, false),
        (constants::ODF_TEXT_TEMPLATE, Family::Text, true),
        (constants::ODF_SPREADSHEET, Family::Spreadsheet, false),
        (
            constants::ODF_SPREADSHEET_TEMPLATE,
            Family::Spreadsheet,
            true,
        ),
        (constants::ODF_PRESENTATION, Family::Presentation, false),
        (
            constants::ODF_PRESENTATION_TEMPLATE,
            Family::Presentation,
            true,
        ),
        (constants::ODF_DRAWING, Family::Drawing, false),
        (constants::ODF_DRAWING_TEMPLATE, Family::Drawing, true),
        (constants::ODF_CHART, Family::Chart, false),
        (constants::ODF_CHART_TEMPLATE, Family::Chart, true),
        (constants::ODF_FORMULA, Family::Formula, false),
        (constants::ODF_FORMULA_TEMPLATE, Family::Formula, true),
        (constants::ODF_IMAGE, Family::Image, false),
        (constants::ODF_IMAGE_TEMPLATE, Family::Image, true),
        (constants::ODF_MASTER, Family::Master, false),
        (constants::ODF_MASTER_TEMPLATE, Family::Master, true),
        (constants::ODF_WEB, Family::Web, true),
        (constants::ODF_DATABASE, Family::Database, false),
    ] {
        let bytes = package(mimetype);
        let document = Package::from_bytes(bytes.clone()).unwrap();
        assert_eq!(document.family(), family);
        assert_eq!(document.is_template(), template);
        assert_eq!(document.mimetype(), mimetype);
        assert!(document.content_xml().unwrap().contains("office:body"));
        assert!(document.odf_metadata().unwrap().is_none());
        assert_eq!(document.media_files().unwrap(), ["Pictures/pixel.png"]);
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.into_bytes(), bytes);
    }
}

#[test]
fn rejects_non_odf_missing_content_and_invalid_xml_bytes() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/zip").unwrap();
    writer.add_file(constants::ODF_CONTENT, b"<x/>").unwrap();
    assert!(Package::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_DRAWING).unwrap();
    assert!(Package::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_CHART).unwrap();
    writer.add_file(constants::ODF_CONTENT, b"<x/>").unwrap();
    let package = writer.finish_to_bytes().unwrap();
    let invalid_xml = replace_zip_member_raw(&package, constants::ODF_CONTENT, &[0xff]);
    assert!(Package::from_bytes(invalid_xml).is_err());
}

#[test]
fn opens_standard_and_odfdo_compatible_flat_documents_losslessly() {
    for (mimetype, body, family, extension) in [
        (constants::ODF_TEXT, "text", Family::Text, "fodt"),
        (constants::ODF_TEXT_TEMPLATE, "text", Family::Text, "fott"),
        (
            constants::ODF_SPREADSHEET,
            "spreadsheet",
            Family::Spreadsheet,
            "fods",
        ),
        (
            constants::ODF_PRESENTATION,
            "presentation",
            Family::Presentation,
            "fodp",
        ),
        (constants::ODF_DRAWING, "drawing", Family::Drawing, "fodg"),
        (constants::ODF_CHART, "chart", Family::Chart, "fodc"),
        (constants::ODF_FORMULA, "formula", Family::Formula, "fodf"),
        (constants::ODF_IMAGE, "image", Family::Image, "fodi"),
    ] {
        let xml = format!(
            r#"<?xml version="1.0"?><!-- keep --><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{mimetype}" o:version="1.3"><o:body><o:{body}/></o:body></o:document>"#
        );
        let document = FlatDocument::from_bytes(xml.clone().into_bytes()).unwrap();
        assert_eq!(document.family(), family);
        assert_eq!(
            document.is_template(),
            mimetype == constants::ODF_TEXT_TEMPLATE
        );
        assert_eq!(document.mimetype(), mimetype);
        assert_eq!(document.extension(), extension);
        assert_eq!(document.xml(), xml);
        assert_eq!(document.to_bytes(), xml.as_bytes());
        assert_eq!(document.into_bytes(), xml.into_bytes());
    }
}

#[test]
fn flat_document_owned_and_reader_limits_share_exact_boundaries() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let exact = u64::try_from(xml.len()).unwrap();
    assert!(FlatDocument::from_bytes_with_limit(xml.as_bytes().to_vec(), exact).is_ok());
    assert!(FlatDocument::from_reader_with_limit(Cursor::new(xml.as_bytes()), exact).is_ok());

    let mut oversized = xml.as_bytes().to_vec();
    oversized.push(b'x');
    assert!(matches!(
        FlatDocument::from_bytes_with_limit(oversized.clone(), exact),
        Err(Error::ResourceLimit(_))
    ));
    assert!(matches!(
        FlatDocument::from_reader_with_limit(Cursor::new(oversized), exact),
        Err(Error::ResourceLimit(_))
    ));
    assert!(matches!(
        FlatDocument::from_bytes_with_limit(xml.into_bytes(), 0),
        Err(Error::InvalidFormat(_))
    ));
    assert!(matches!(
        FlatDocument::from_bytes_with_limit(
            b"<o:document/>".to_vec(),
            super::flat::HARD_MAX_FLAT_DOCUMENT_BYTES + 1,
        ),
        Err(Error::InvalidFormat(_))
    ));
}

#[test]
fn rejects_flat_mimetype_body_mismatch_and_incomplete_xml() {
    for xml in [
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:spreadsheet/></o:body></o:document>"#,
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body>"#,
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body></o:document><o:document/>"#,
    ] {
        assert!(
            FlatDocument::from_bytes(xml.as_bytes().to_vec()).is_err(),
            "accepted invalid flat document {xml}"
        );
    }
}

#[test]
fn flat_document_exposes_namespace_aware_metadata() {
    let xml = br#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:d="http://purl.org/dc/elements/1.1/"
        o:mimetype="application/vnd.oasis.opendocument.text">
        <o:meta><d:title>A &amp; B</d:title></o:meta>
        <o:body><o:text/></o:body>
    </o:document>"#;
    let document = FlatDocument::from_bytes(xml.to_vec()).unwrap();
    assert_eq!(
        document.odf_metadata().unwrap().title.as_deref(),
        Some("A & B")
    );
    assert_eq!(document.metadata().unwrap().title.as_deref(), Some("A & B"));
}

#[test]
fn flat_document_accepts_entity_escaped_office_namespace_uri() {
    let xml = br#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:&#x31;.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body></o:document>"#;
    let document = FlatDocument::from_bytes(xml.to_vec()).expect("escaped office URI is semantic");
    assert_eq!(document.family(), Family::Text);
    assert_eq!(document.as_bytes(), xml);
}

fn flat_test_context(memory: u64) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "flat generic test",
        BudgetLimits::new(memory, 1 << 30, 1 << 30, 1_000_000, 4_096, 1_000_000_000),
    );
    let (_source, cancellation) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution);
    (budget, context)
}

struct CountingReader {
    bytes: Vec<u8>,
    reads: usize,
}

impl Read for CountingReader {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        self.reads = self.reads.saturating_add(1);
        let amount = output.len().min(self.bytes.len());
        output[..amount].copy_from_slice(&self.bytes[..amount]);
        self.bytes.drain(..amount);
        Ok(amount)
    }
}

#[test]
fn flat_reader_admits_input_incrementally_before_full_stream_consumption() {
    let input_limit = 8 * 1024_u64;
    let budget = Budget::root(
        "flat incremental reader test",
        BudgetLimits::new(
            1 << 30,
            input_limit,
            1 << 30,
            1_000_000,
            4_096,
            1_000_000_000,
        ),
    );
    let (_source, cancellation) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(1 << 30).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution);
    let mut reader = CountingReader {
        bytes: vec![b'x'; 64 * 1024],
        reads: 0,
    };
    let result = FlatDocument::from_reader_with_execution_context(&mut reader, 1 << 20, context);
    assert!(matches!(
        result,
        Err(Error::ResourceLimit(limit)) if limit.resource == Resource::InputBytes
    ));
    assert_eq!(reader.reads, 2);
    assert_eq!(budget.used(Resource::InputBytes), input_limit);
    assert!(
        !reader.bytes.is_empty(),
        "reader input must not be consumed past the first bounded refusal"
    );
}

#[test]
fn flat_reader_charges_old_and_new_allocations_during_large_growth() {
    let input = vec![b'x'; 72 * 1024];
    let maximum = u64::try_from(input.len()).unwrap();
    let (probe_budget, probe_context) = flat_test_context(u64::MAX);
    let (probe_bytes, probe_memory) =
        super::flat::read_flat_input(Cursor::new(input.clone()), maximum, &probe_context)
            .expect("probe input read");
    assert_eq!(probe_bytes.len(), input.len());
    assert!(probe_bytes.capacity() > 64 * 1024);
    let retained_capacity = probe_budget.used(Resource::Memory);
    assert_eq!(
        retained_capacity,
        u64::try_from(probe_bytes.capacity()).unwrap()
    );
    drop(probe_memory);
    assert_eq!(probe_budget.used(Resource::Memory), 0);

    // The first growth keeps the 64 KiB initial allocation live while the
    // larger destination is allocated. One byte below that measured peak
    // must refuse before the replacement allocation can be admitted.
    let initial_capacity = u64::try_from(64 * 1024).unwrap();
    let exact_peak = initial_capacity
        .checked_add(retained_capacity)
        .expect("test peak fits in u64");
    let (exact_budget, exact_context) = flat_test_context(exact_peak);
    let (exact_bytes, exact_memory) =
        super::flat::read_flat_input(Cursor::new(input.clone()), maximum, &exact_context)
            .expect("exact old-plus-new peak must admit");
    assert_eq!(exact_bytes.len(), input.len());
    drop(exact_memory);
    assert_eq!(exact_budget.used(Resource::Memory), 0);

    let (_under_budget, under_context) = flat_test_context(exact_peak - 1);
    assert!(matches!(
        super::flat::read_flat_input(Cursor::new(input), maximum, &under_context),
        Err(Error::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
}

#[test]
fn flat_snapshot_requires_and_retains_exact_source_memory() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:h="http://www.w3.org/1999/xhtml" xmlns:dc="http://purl.org/dc/elements/1.1/" o:mimetype="{}"><o:body><o:text><t:p>Alpha</t:p></o:text></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let source_len = u64::try_from(xml.len()).unwrap();
    let bytes = xml.into_bytes();
    let source_capacity = u64::try_from(bytes.capacity()).unwrap();
    let maximum = source_len + 256;
    let (_zero_budget, zero_context) = flat_test_context(0);
    assert!(matches!(
        FlatDocument::from_bytes_with_execution_context(
            bytes.clone(),
            maximum,
            zero_context,
        ),
        Err(Error::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
    let (budget, context) = flat_test_context(source_capacity);
    let document =
        FlatDocument::from_bytes_with_execution_context(bytes, maximum, context).unwrap();
    assert_eq!(budget.used(Resource::Memory), source_capacity);
    drop(document);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn flat_mutation_refusal_keeps_source_reservation_and_bytes() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:h="http://www.w3.org/1999/xhtml" xmlns:dc="http://purl.org/dc/elements/1.1/" o:mimetype="{}"><o:body><o:text><t:p>Alpha</t:p></o:text></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let source_len = u64::try_from(xml.len()).unwrap();
    let bytes = xml.into_bytes();
    let source_capacity = u64::try_from(bytes.capacity()).unwrap();
    let maximum = source_len + 256;
    let (budget, context) = flat_test_context(source_capacity);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(bytes, maximum, context).unwrap();
    let before = document.to_bytes();
    let error = document
        .set_paragraph_rdfa(
            Position::new(0),
            &crate::RdfaAttributes {
                about: Some("#changed".to_string()),
                property: Some("dc:title".to_string()),
                ..crate::RdfaAttributes::default()
            },
        )
        .expect_err("source-only memory budget must refuse before candidate allocation");
    assert!(matches!(error, Error::ResourceLimit(limit) if limit.resource == Resource::Memory));
    assert_eq!(document.as_bytes(), before.as_slice());
    assert_eq!(budget.used(Resource::Memory), source_capacity);
}

#[test]
fn flat_mutations_retain_only_the_new_candidate_and_ignore_large_output_cap() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:h="http://www.w3.org/1999/xhtml" xmlns:dc="http://purl.org/dc/elements/1.1/" o:mimetype="{}"><o:body><o:text><t:p>Alpha</t:p></o:text></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let maximum = super::flat::HARD_MAX_FLAT_DOCUMENT_BYTES;
    let (budget, context) = flat_test_context(1 << 20);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(xml.into_bytes(), maximum, context)
            .unwrap();
    document
        .set_paragraph_rdfa(
            Position::new(0),
            &crate::RdfaAttributes {
                about: Some("#changed".to_string()),
                property: Some("dc:title".to_string()),
                ..crate::RdfaAttributes::default()
            },
        )
        .expect("exact aggregate reservation must admit the edit");
    assert!(document.xml().contains("xhtml:about"));
    assert_eq!(
        budget.used(Resource::Memory),
        u64::try_from(document.xml.capacity()).unwrap()
    );
}

#[test]
fn flat_mutation_noop_keeps_source_usage_unchanged() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:mimetype="{}"><o:body><o:text><t:p>Alpha</t:p></o:text></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let source_len = u64::try_from(xml.len()).unwrap();
    let bytes = xml.into_bytes();
    let source_capacity = u64::try_from(bytes.capacity()).unwrap();
    let (budget, context) = flat_test_context(1 << 20);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(bytes, source_len + 1, context).unwrap();
    let before = document.to_bytes();
    document
        .set_paragraph_rdfa(Position::new(0), &crate::RdfaAttributes::default())
        .expect("equivalent RDFa must be an atomic no-op");
    assert_eq!(document.to_bytes(), before);
    assert_eq!(budget.used(Resource::Memory), source_capacity);
}

#[test]
fn flat_variable_declarations_expand_replace_and_remove_atomically() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let mut document = FlatDocument::from_bytes(xml.into_bytes()).unwrap();
    let scope = crate::variable_declaration::Scope::Body(crate::variable_declaration::Body::Text);
    let first = crate::variable_declaration::Group {
        kind: crate::variable_declaration::Kind::Simple,
        part: crate::variable_declaration::Part::Flat,
        scope: scope.clone(),
        declarations: vec![crate::variable_declaration::Declaration::Simple {
            name: "counter".to_string(),
            value_type: crate::variable_declaration::ValueType::Float,
        }],
    };
    assert!(
        document
            .set_variable_declaration_group(&first)
            .unwrap()
            .is_none()
    );
    assert!(document.xml().contains("<o:text><text:variable-decls"));
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(crate::variable_declaration::Kind::Simple, "counter")
            .is_some()
    );

    let second = crate::variable_declaration::Group {
        declarations: vec![crate::variable_declaration::Declaration::Simple {
            name: "replacement".to_string(),
            value_type: crate::variable_declaration::ValueType::String,
        }],
        ..first.clone()
    };
    assert_eq!(
        document.set_variable_declaration_group(&second).unwrap(),
        Some(first.clone()),
    );
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(crate::variable_declaration::Kind::Simple, "replacement")
            .is_some()
    );
    assert_eq!(
        document
            .remove_variable_declaration_group(&scope, crate::variable_declaration::Kind::Simple)
            .unwrap(),
        Some(second),
    );
    assert!(document.variable_declarations().unwrap().groups.is_empty());
}

#[cfg(windows)]
#[test]
fn flat_document_save_replaces_a_windows_destination_atomically() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let destination = directory.path().join("template.fott");
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
        constants::ODF_TEXT_TEMPLATE,
    );
    let document = FlatDocument::from_bytes(xml.into_bytes()).expect("template opens");

    document.save(&destination).expect("Windows atomic save");

    assert_eq!(
        std::fs::read(destination).expect("saved template"),
        document.as_bytes()
    );
}
