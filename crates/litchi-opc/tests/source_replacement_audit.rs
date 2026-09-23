#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused publication assertions intentionally panic on fixture errors"
)]

//! Change 0747: source-backed publication audits a replaced Part's original
//! bytes and its replacement with one `verify_source_replacement` call. The
//! verdict, the refusal identity and the emitted archive are those of the two
//! `verify_source` audits it replaces, through every public door.

use litchi_opc::{OpcError, PackURI, SourceBackedPackage, SourceTopologyPlan};
use soapberry_zip::office::StreamingArchiveWriter;
use xml_minifier::audit::{Limits, verify_source};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT: &str = "/word/document.xml";
const STYLES: &[u8] = b"<styles>\n  <style id=\"a\" />\n</styles>";

/// A formatted producer document with `items` records; `item` rewrites one
/// record's markup.
fn document(items: usize, item: impl Fn(usize) -> Option<String>) -> Vec<u8> {
    let mut document = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<document>\n");
    for index in 0..items {
        let record =
            item(index).unwrap_or_else(|| format!("<item n=\"{index}\">value {index}</item>"));
        document.push_str("  ");
        document.push_str(&record);
        document.push('\n');
    }
    document.push_str("</document>\n");
    document.into_bytes()
}

fn plain() -> Vec<u8> {
    document(200, |_| None)
}

fn edited(index: usize, record: &str) -> Vec<u8> {
    document(200, |at| (at == index).then(|| record.to_string()))
}

fn archive(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    writer.write_stored("word/document.xml", document).unwrap();
    writer.write_stored("word/styles.xml", STYLES).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn uri() -> PackURI {
    PackURI::new(DOCUMENT).unwrap()
}

#[derive(Clone, Copy, Debug)]
enum Door {
    Topology,
    SingleOverlay,
    MultipleOverlays,
}

const DOORS: [Door; 3] = [Door::Topology, Door::SingleOverlay, Door::MultipleOverlays];

fn publish(door: Door, original: &[u8], replacement: Vec<u8>) -> (Result<(), OpcError>, Vec<u8>) {
    let package = SourceBackedPackage::from_vec(archive(original)).expect("fixture opens");
    let mut output = Vec::new();
    let result = match door {
        Door::Topology => {
            let mut plan = SourceTopologyPlan::new();
            plan.try_replace_part(uri(), replacement).unwrap();
            package.write_topology_to_stream(&mut output, plan)
        },
        Door::SingleOverlay => {
            package.write_part_overlay_to_stream(&mut output, &uri(), replacement)
        },
        Door::MultipleOverlays => {
            package.write_part_overlays_to_stream(&mut output, vec![(uri(), replacement)])
        },
    };
    (result, output)
}

fn assert_refused(
    door: Door,
    result: Result<(), OpcError>,
    output: &[u8],
    expected: &xml_minifier::audit::Error,
) {
    match result {
        Err(OpcError::XmlPublication { part, source }) => {
            assert_eq!(part, DOCUMENT, "{door:?}");
            assert_eq!(&source, expected, "{door:?}");
        },
        other => panic!("{door:?}: expected an XML publication refusal, got {other:?}"),
    }
    assert!(
        output.is_empty(),
        "{door:?}: a refusal emits no archive byte"
    );
}

#[test]
fn a_local_edit_publishes_the_replacement_and_leaves_every_other_member_exact() {
    let original = plain();
    let replacement = edited(57, "<item n=\"57\">value 57, edited</item>");
    for door in DOORS {
        let (result, output) = publish(door, &original, replacement.clone());
        result.unwrap_or_else(|error| panic!("{door:?}: the local edit publishes: {error:?}"));
        let before = SourceBackedPackage::from_vec(archive(&original)).unwrap();
        let after = SourceBackedPackage::from_vec(output).expect("published package reopens");
        for part in before.iter_parts() {
            let name = part.partname().clone();
            let published = after.part(&name).unwrap().data().unwrap();
            let expected = if name == uri() {
                replacement.clone()
            } else {
                part.data().unwrap().as_bytes().to_vec()
            };
            assert_eq!(
                published.as_bytes(),
                expected.as_slice(),
                "{door:?}: {name}"
            );
        }
    }
}

#[test]
fn a_defect_inside_the_edited_element_is_refused_with_the_complete_audit_error() {
    let original = plain();
    for record in [
        "<item n=\"57\" xml:space=\"bogus\">value</item>",
        "<item n=\"57\">value</itemx>",
        "<item n=\"57\"><!DOCTYPE x>value</item>",
        "<item n=\"57\"a=\"1\">value</item>",
        "<item n=\"57\">value &unclosed</item>",
    ] {
        let replacement = edited(57, record);
        let expected =
            verify_source(&replacement, Limits::default()).expect_err("the fixture is invalid");
        for door in DOORS {
            let (result, output) = publish(door, &original, replacement.clone());
            assert_refused(door, result, &output, &expected);
        }
    }
    // Invalid UTF-8 in the edited value keeps its physical offset.
    let mut replacement = edited(57, "<item n=\"57\">value X</item>");
    let at = replacement
        .windows(7)
        .position(|window| window == b"value X")
        .unwrap()
        + 6;
    replacement[at] = 0xFF;
    let expected = verify_source(&replacement, Limits::default()).unwrap_err();
    assert_eq!(
        expected,
        xml_minifier::audit::Error::Encoding { valid_up_to: at }
    );
    for door in DOORS {
        let (result, output) = publish(door, &original, replacement.clone());
        assert_refused(door, result, &output, &expected);
    }
}

#[test]
fn a_substituted_defect_outside_the_edit_is_still_audited_and_refused() {
    // The intended edit is one record; a second, distant record was also
    // substituted and carries a defect.
    let original = plain();
    let replacement = document(200, |index| match index {
        57 => Some("<item n=\"57\">value 57, edited</item>".to_string()),
        150 => Some("<item n=\"150\" xml:space=\"x\">value 150</item>".to_string()),
        _ => None,
    });
    let expected = verify_source(&replacement, Limits::default()).unwrap_err();
    for door in DOORS {
        let (result, output) = publish(door, &original, replacement.clone());
        assert_refused(door, result, &output, &expected);
    }
}

#[test]
fn a_refused_original_is_reported_before_its_replacement() {
    // Both sides are invalid, and the replacement's first defect comes
    // earlier than the one it copies from the original. The original's error
    // wins, as it did when the two sides were audited by two calls.
    let original = document(200, |index| {
        (index == 150).then(|| "<item n=\"150\" xml:space=\"q\">value 150</item>".to_string())
    });
    let replacement = document(200, |index| match index {
        57 => Some("<item n=\"57\"><!DOCTYPE x></item>".to_string()),
        150 => Some("<item n=\"150\" xml:space=\"q\">value 150</item>".to_string()),
        _ => None,
    });
    let expected = verify_source(&original, Limits::default()).unwrap_err();
    assert_ne!(
        Some(&expected),
        verify_source(&replacement, Limits::default())
            .err()
            .as_ref()
    );
    for door in DOORS {
        let (result, output) = publish(door, &original, replacement.clone());
        assert_refused(door, result, &output, &expected);
    }
}
