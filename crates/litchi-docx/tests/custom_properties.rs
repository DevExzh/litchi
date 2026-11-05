#![allow(
    clippy::redundant_closure_call,
    reason = "raw OPC test edits stage one closure over the edit's graph"
)]
use std::io::{self, Cursor, Write};

use litchi_docx::Package;
use litchi_ooxml_common::custom::{Props as CustomProps, Value};
use litchi_opc::OpcPackage;
use litchi_opc::PackURI;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const HEADER_FIXTURE: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx"
);

fn custom_properties_xml(bytes: &[u8]) -> Vec<u8> {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("open DOCX OPC package");
    let uri = PackURI::new("/docProps/custom.xml").expect("static custom-properties URI");
    package
        .get_part(&uri)
        .expect("real fixture has custom properties")
        .blob()
        .to_vec()
}

/// Opens and drops a document edit. The alt-chunk-header fixture (an
/// `w:altChunk` after the body-final section properties) was refused until
/// record 0293 admitted that tail; either way the package keeps its custom
/// properties.
fn abandoned_document_edit(package: &mut Package) {
    assert!(
        package.edit().is_ok(),
        "the alt-chunk-header fixture admits document editing"
    );
}

/// Stage a custom-properties change through the metadata edit family and
/// publish it as a new package.
fn edit_custom_props(
    package: &Package,
    change: impl FnOnce(&mut CustomProps),
) -> litchi_docx::metadata::Result<Package> {
    let mut edit = package.edit_metadata();
    let mut custom = edit.custom_props().clone();
    change(&mut custom);
    edit.set_custom_props(custom);
    Ok(edit.commit()?.into_package())
}

fn remove_custom_props(package: &Package) -> litchi_docx::metadata::Result<Package> {
    let mut edit = package.edit_metadata();
    edit.remove_custom_props();
    Ok(edit.commit()?.into_package())
}

struct FailingSink {
    remaining: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::other("injected sink failure"));
        }
        let accepted = bytes.len().min(self.remaining);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn custom_props_round_trip_and_clear_remove_the_package_graph() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("custom-props.docx");
    let package = edit_custom_props(&Package::new().expect("new DOCX package"), |custom| {
        custom
            .insert("Project", "Litchi")
            .expect("valid custom property");
        custom
            .insert("Version", Value::I32(38))
            .expect("valid custom property");
    })
    .expect("publish custom properties");
    package.save(&path).expect("save custom properties");

    let reopened = Package::open(&path).expect("reopen custom properties");
    assert_eq!(
        reopened.custom_props().get("project"),
        Some(&Value::Text("Litchi".to_string()))
    );
    assert_eq!(
        reopened.custom_props().get("VERSION"),
        Some(&Value::I32(38))
    );
    let reopened = remove_custom_props(&reopened).expect("clear custom properties");
    reopened.save(&path).expect("save cleared properties");

    let graph = OpcPackage::open(&path, litchi_opc::SourceLimits::default())
        .expect("open cleared OPC graph");
    assert!(
        graph
            .rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::CUSTOM_PROPERTIES)
    );
    assert!(
        graph
            .iter_parts()
            .all(|part| part.content_type() != ct::OFC_CUSTOM_PROPERTIES)
    );
    assert!(
        Package::open(&path)
            .expect("reopen cleared package")
            .custom_props()
            .is_empty()
    );
}

#[test]
fn malformed_custom_props_are_not_treated_as_absent() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("malformed-custom-props.docx");
    let package = Package::new().expect("new DOCX package");
    package.save(&path).expect("save base package");

    let mut opc = OpcPackage::open(&path, litchi_opc::SourceLimits::default())
        .expect("open base OPC package");
    let part = PackURI::new("/docProps/custom.xml").expect("static package URI");
    opc.add_part(Box::new(BlobPart::new(
        part,
        ct::OFC_CUSTOM_PROPERTIES.to_string(),
        b"<not-properties/>".to_vec(),
    )));
    opc.relate_to("docProps/custom.xml", rt::CUSTOM_PROPERTIES);

    assert!(Package::from_opc_package(opc).is_err());
}

#[test]
fn word_reserved_custom_properties_are_host_scoped_and_transactional() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let valid_path = directory.path().join("word-reserved-props.docx");
    let valid = edit_custom_props(&Package::new().expect("new DOCX package"), |custom| {
        custom
            .insert(
                "ClassificationContentMarkingHeaderFontProps",
                "#ffFF00,23,Calibri",
            )
            .expect("valid Word header font properties");
    })
    .expect("publish Word reserved properties");
    valid
        .save(&valid_path)
        .expect("save Word reserved properties");
    assert_eq!(
        Package::open(&valid_path)
            .expect("reopen Word reserved properties")
            .custom_props()
            .get("ClassificationContentMarkingHeaderFontProps"),
        Some(&Value::Text("#ffFF00,23,Calibri".to_owned()))
    );

    let invalid_path = directory.path().join("non-word-reserved-props.docx");
    let invalid = Package::new().expect("new DOCX package");
    // The host-scoped property is refused when the edit is published, and the
    // source package keeps no custom-properties part.
    assert!(
        edit_custom_props(&invalid, |custom| {
            custom
                .insert(
                    "ClassificationContentMarkingHeaderLocations",
                    r"Office\:Suite:10",
                )
                .expect("valid PowerPoint location syntax");
        })
        .is_err()
    );
    assert!(invalid.save(&invalid_path).is_ok());
    assert!(
        invalid
            .opc_package()
            .iter_parts()
            .all(|part| part.content_type() != ct::OFC_CUSTOM_PROPERTIES)
    );
}

#[test]
fn clean_real_fixture_preserves_custom_properties_for_stream_and_path_save() {
    let expected = custom_properties_xml(HEADER_FIXTURE);

    let streamed = Package::from_reader(Cursor::new(HEADER_FIXTURE))
        .expect("open alt-chunk-header fixture for stream save");
    let mut stream_output = Vec::new();
    streamed
        .write_to(&mut stream_output)
        .expect("stream clean alt-chunk-header fixture");
    assert_eq!(custom_properties_xml(&stream_output), expected);

    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("clean-alt-chunk-header.docx");
    let saved = Package::from_reader(Cursor::new(HEADER_FIXTURE))
        .expect("open alt-chunk-header fixture for path save");
    saved
        .save(&path)
        .expect("save clean alt-chunk-header fixture");
    let path_output = std::fs::read(path).expect("read path-saved fixture");
    assert_eq!(custom_properties_xml(&path_output), expected);
}

#[test]
fn abandoned_document_edit_preserves_custom_properties_for_stream_and_path_save() {
    let expected = custom_properties_xml(HEADER_FIXTURE);

    let mut streamed = Package::from_reader(Cursor::new(HEADER_FIXTURE))
        .expect("open alt-chunk-header fixture for refused stream save");
    abandoned_document_edit(&mut streamed);
    let mut stream_output = Vec::new();
    streamed
        .write_to(&mut stream_output)
        .expect("stream after refused document edit");
    assert_eq!(custom_properties_xml(&stream_output), expected);

    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("refused-alt-chunk-header.docx");
    let mut saved = Package::from_reader(Cursor::new(HEADER_FIXTURE))
        .expect("open alt-chunk-header fixture for refused path save");
    abandoned_document_edit(&mut saved);
    saved
        .save(&path)
        .expect("path save after refused document edit");
    let path_output = std::fs::read(path).expect("read refused path-saved fixture");
    assert_eq!(custom_properties_xml(&path_output), expected);
}

#[test]
fn raw_opc_custom_property_replacement_remains_byte_exact_after_save() {
    let custom_uri = PackURI::new("/docProps/custom.xml").expect("static custom-properties URI");
    let mut replacement = custom_properties_xml(HEADER_FIXTURE);
    replacement.push(b'\n');

    let mut package = Package::from_reader(Cursor::new(HEADER_FIXTURE))
        .expect("open alt-chunk-header fixture for raw replacement");
    package = {
        let mut edit = package.edit_opc();
        (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
            {
                candidate
                    .get_part_mut(&custom_uri)?
                    .set_blob(replacement.clone());
                Ok(())
            }
        })(edit.package_mut())
        .expect("replace custom-properties bytes through raw OPC API");
        edit.commit()
            .expect("replace custom-properties bytes through raw OPC API")
            .into_package()
    };
    assert_eq!(
        package
            .opc_package()
            .get_part(&custom_uri)
            .expect("replaced custom-properties part")
            .blob(),
        replacement.as_slice()
    );

    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("save raw OPC custom-properties replacement");
    assert_eq!(custom_properties_xml(&output), replacement);
}

#[test]
fn raw_opc_empty_custom_properties_part_is_preserved_until_explicit_clear() {
    let custom_uri = PackURI::new("/docProps/custom.xml").expect("static custom-properties URI");
    let replacement = br#"<?xml version="1.0"?><op:Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns:op="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" />
"#
    .to_vec();

    let mut package = Package::from_reader(Cursor::new(HEADER_FIXTURE))
        .expect("open alt-chunk-header fixture for empty replacement");
    package = {
        let mut edit = package.edit_opc();
        (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
            {
                candidate
                    .get_part_mut(&custom_uri)?
                    .set_blob(replacement.clone());
                Ok(())
            }
        })(edit.package_mut())
        .expect("replace custom-properties bytes with an empty root");
        edit.commit()
            .expect("replace custom-properties bytes with an empty root")
            .into_package()
    };
    assert!(package.custom_props().is_empty());

    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("save empty raw OPC custom-properties replacement");
    assert_eq!(custom_properties_xml(&output), replacement);

    let cleared = Package::from_reader(Cursor::new(output)).expect("reopen empty properties");
    let cleared = edit_custom_props(&cleared, CustomProps::clear).expect("clear custom properties");
    let mut cleared_output = Vec::new();
    cleared
        .write_to(&mut cleared_output)
        .expect("save explicitly cleared custom properties");
    // Clearing properties that are already empty is an exact no-op (record
    // 0355): the empty part is kept byte for byte. Removing the part itself
    // is a raw graph change.
    assert_eq!(custom_properties_xml(&cleared_output), replacement);

    let removed = remove_custom_props(&cleared).expect("explicitly remove custom graph");
    let mut removed_output = Vec::new();
    removed
        .write_to(&mut removed_output)
        .expect("save explicitly removed custom properties");
    let removed_graph =
        OpcPackage::from_bytes(&removed_output, litchi_opc::SourceLimits::default())
            .expect("open explicitly removed OPC graph");
    assert!(
        removed_graph
            .rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::CUSTOM_PROPERTIES)
    );
    assert!(
        removed_graph
            .iter_parts()
            .all(|part| part.content_type() != ct::OFC_CUSTOM_PROPERTIES)
    );
}

#[test]
fn exact_source_stream_failure_keeps_typed_cause_and_accepted_count() {
    let authored = Package::new().unwrap();
    let mut original = Vec::new();
    authored.write_to(&mut original).unwrap();
    for accepted in [0, 17] {
        let package = Package::from_bytes(&original).unwrap();
        let error = package
            .write_to(FailingSink {
                remaining: accepted,
            })
            .unwrap_err();
        let litchi_docx::Error::Opc(error) = error else {
            panic!("publication lost its OPC error type");
        };
        let cause = if accepted == 0 {
            error
        } else {
            let litchi_opc::OpcError::IncompleteOutput { written, source } = error else {
                panic!("partial output lost its accepted-byte count");
            };
            assert_eq!(written, accepted as u64);
            *source
        };
        assert!(
            matches!(cause, litchi_opc::OpcError::IoError(ref error) if error.kind() == io::ErrorKind::Other)
        );
        let mut retry = Vec::new();
        package.write_to(&mut retry).unwrap();
        assert_eq!(retry, original);
    }
}

#[test]
fn canonical_bytes_publish_pending_edits_and_reopen_exactly() {
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("canonical byte output");
    let package = edit_custom_props(&builder.build().unwrap(), |custom| {
        custom.insert("ByteMarker", "retained").unwrap();
    })
    .unwrap();
    let bytes = package.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.document().unwrap().text().unwrap(),
        "canonical byte output"
    );
    assert_eq!(
        reopened.custom_props().get("ByteMarker"),
        Some(&Value::Text("retained".into()))
    );
    assert_eq!(reopened.to_bytes().unwrap(), bytes);
    assert_eq!(reopened.to_plain_bytes().unwrap(), bytes);
}
