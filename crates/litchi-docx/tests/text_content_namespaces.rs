//! Typed text controls must retain expanded XML names and source ownership.
use litchi_docx::paragraph::{Run, RunBreak, RunBreakClear, RunBreakType, RunContent};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

fn fixture() -> Vec<u8> {
    let xml = br#"<a:document xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:f="urn:foreign"><a:body><a:p><a:r><b:t>A&#10;B</b:t><b:br b:type="column" f:clear="all" b:clear="left"/><b:footnoteReference f:id="bad" b:id="7"/><b:endnoteReference b:id="9"/><f:br/></a:r></a:p></a:body></a:document>"#;
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            xml.to_vec(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn check(run: Run) {
    assert_eq!(run.text().unwrap(), "A\nB\n");
    let contents = run.contents().unwrap();
    assert_eq!(contents.len(), 5);
    assert!(matches!(&contents[0], RunContent::Text(text) if text=="A\nB"));
    assert!(matches!(
        contents[1],
        RunContent::Break(RunBreak {
            break_type: RunBreakType::Column,
            clear: RunBreakClear::Left
        })
    ));
    assert!(matches!(contents[2], RunContent::FootnoteReference(7)));
    assert!(matches!(contents[3], RunContent::EndnoteReference(9)));
    assert!(matches!(&contents[4], RunContent::Unknown(raw) if raw.xml_bytes()==b"<f:br/>"));
    assert_eq!(
        run.breaks().unwrap().as_slice(),
        &[RunBreak {
            break_type: RunBreakType::Column,
            clear: RunBreakClear::Left
        }]
    );
}

#[test]
fn eager_and_source_runs_keep_namespace_context_after_owner_drop() {
    let bytes = fixture();
    let run = {
        let package = litchi_docx::Package::from_bytes(&bytes).unwrap();
        let paragraph = package.document().unwrap().paragraphs().unwrap().remove(0);
        assert_eq!(paragraph.text().unwrap(), "A\nB\n");
        paragraph.runs().unwrap().remove(0)
    };
    check(run);
    let run = {
        let package = litchi_docx::source_backed::Package::from_bytes(&bytes).unwrap();
        let paragraph = package.document().unwrap().paragraph(0).unwrap().unwrap();
        assert_eq!(paragraph.text().unwrap(), "A\nB\n");
        paragraph.runs().unwrap().remove(0)
    };
    check(run);
}

#[test]
fn strict_default_namespaces_and_local_rebinding_preserve_break_types() {
    let run = Run::new(br#"<r xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:x="urn:foreign"><br w:type="page"/><x:br/><x:br xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main" x:type="column"/><w:br xmlns:w="urn:foreign"/></r>"#.to_vec());
    let breaks = run.breaks().unwrap();
    assert_eq!(breaks.len(), 2);
    assert_eq!(breaks[0].break_type, RunBreakType::Page);
    assert_eq!(breaks[1].break_type, RunBreakType::Column);
    let contents = run.contents().unwrap();
    assert_eq!(contents.len(), 4);
    assert!(matches!(contents[0], RunContent::Break(_)));
    assert!(matches!(contents[1], RunContent::Unknown(_)));
    assert!(matches!(contents[2], RunContent::Break(_)));
    assert!(matches!(contents[3], RunContent::Unknown(_)));
}

#[test]
fn resolved_default_root_does_not_enable_unqualified_fragment_fallback() {
    let run = Run::new(br#"<r xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><t xmlns="">foreign</t><br xmlns=""/></r>"#.to_vec());
    assert!(
        run.contents()
            .unwrap()
            .iter()
            .all(|content| matches!(content, RunContent::Unknown(_)))
    );
    assert!(run.breaks().unwrap().is_empty());
    assert_eq!(run.text().unwrap(), "");
    let run = Run::new(br#"<r xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><footnoteReference id="1"/></r>"#.to_vec());
    assert!(run.contents().is_err());
}

#[test]
fn flattened_text_respects_inherited_aliases_and_local_shadowing() {
    let paragraph = litchi_docx::Paragraph::new(br#"<a:p xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:f="urn:foreign"><a:r><b:t>one</b:t><b:t xmlns:b="urn:foreign">hidden</b:t><b:t>two</b:t><f:t>foreign</f:t><b:tab/><b:br/><b:cr/><b:noBreakHyphen/><b:softHyphen/></a:r></a:p>"#.to_vec());
    let expected = "onetwo\t\n\n\u{2011}\u{ad}";
    assert_eq!(paragraph.text().unwrap(), expected);
    let run = paragraph.runs().unwrap().remove(0);
    drop(paragraph);
    assert_eq!(run.text().unwrap(), expected);
}

#[test]
fn empty_elements_obey_the_run_content_depth_limit() {
    let xml = format!(
        "<w:r>{}<w:tab/>{}</w:r>",
        "<x>".repeat(127),
        "</x>".repeat(127)
    );
    let run = Run::new(xml.into_bytes());
    assert!(run.contents().is_err());
    assert!(run.breaks().is_err());
}
