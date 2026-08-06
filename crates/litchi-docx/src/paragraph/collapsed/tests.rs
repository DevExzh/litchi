use super::*;
use crate::Paragraph;
use crate::writer::MutableDocument;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W12: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn paragraph_xml(body: &str) -> String {
    format!(
        r#"<w:p xmlns:w="{W}" xmlns:x="urn:opaque" xmlns:w12="{W12}" xmlns:mc="{MC}">{body}</w:p>"#
    )
}

#[test]
fn reads_closed_on_off_domain_and_preserves_opaque_ppr_content() {
    let xml = paragraph_xml(
        r#"<w:pPr><w:pStyle w:val="Heading1"/><w12:collapsed w12:val="on"/><x:opaque x:value="keep"/></w:pPr><w:r><w:t>kept</w:t></w:r>"#,
    );
    let snapshot = Snapshot::from_xml(xml.as_bytes().to_vec()).unwrap();
    assert_eq!(snapshot.collapsed(), Some(Collapsed::Enabled));
    assert_eq!(
        Paragraph::new(xml.as_bytes().to_vec()).collapsed().unwrap(),
        Some(Collapsed::Enabled)
    );
    assert_eq!(snapshot.xml_bytes(), xml.as_bytes());

    let disabled = Snapshot::from_xml(
        paragraph_xml(r#"<w:pPr><w12:collapsed w12:val="0"/></w:pPr>"#).into_bytes(),
    )
    .unwrap();
    assert_eq!(disabled.collapsed(), Some(Collapsed::Disabled));
}

#[test]
fn transaction_commit_and_inverse_are_atomic_and_loss_preserving() {
    let source = Snapshot::from_xml(
        paragraph_xml(
            r#"<w:pPr><w:pStyle w:val="Heading1"/><x:before/><x:after/></w:pPr><w:r><w:t>text</w:t></w:r>"#,
        )
        .into_bytes(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_collapsed(Some(Collapsed::Enabled)).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.collapsed(), None);
    let changed = commit.snapshot();
    assert_eq!(changed.collapsed(), Some(Collapsed::Enabled));
    let changed_xml = std::str::from_utf8(changed.xml_bytes()).unwrap();
    assert!(changed_xml.contains("x:before"));
    assert!(changed_xml.contains("x:after"));
    assert_eq!(commit.patch().before(), None);
    assert_eq!(commit.patch().after(), Some(Collapsed::Enabled));

    let restored = commit.patch().inverse().apply(changed).unwrap();
    assert_eq!(restored.collapsed(), None);
    assert!(
        std::str::from_utf8(restored.xml_bytes())
            .unwrap()
            .contains("w:pStyle")
    );
}

#[test]
fn invalid_inputs_and_patch_preconditions_leave_sources_unchanged() {
    assert!(
        Snapshot::from_xml(
            paragraph_xml(r#"<w:pPr><w12:collapsed w12:val="maybe"/></w:pPr>"#).into_bytes()
        )
        .is_err()
    );
    assert!(
        Snapshot::from_xml(
            paragraph_xml(r#"<w:pPr><w12:collapsed/><w12:collapsed/></w:pPr>"#).into_bytes()
        )
        .is_err()
    );
    assert!(
        Snapshot::from_xml(paragraph_xml(r#"<w:pPr><x:collapsed/></w:pPr>"#).into_bytes()).is_ok()
    );

    let source =
        Snapshot::from_xml(paragraph_xml(r#"<w:r><w:t>unchanged</w:t></w:r>"#).into_bytes())
            .unwrap();
    let mut edit = source.edit();
    edit.set_collapsed(Some(Collapsed::Disabled)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().apply(&source).is_ok());
    assert!(commit.patch().apply(commit.snapshot()).is_err());
}

#[test]
fn direct_and_generated_paragraph_facades_emit_ignorable_extension() {
    let mut paragraph = Paragraph::new(paragraph_xml(r#"<w:r><w:t>one</w:t></w:r>"#).into_bytes());
    paragraph.set_collapsed(Some(Collapsed::Disabled)).unwrap();
    let output = String::from_utf8(paragraph.xml_bytes().to_vec()).unwrap();
    assert!(output.contains("w12:collapsed"));
    assert!(output.contains("mc:Ignorable=\"w12\""));
    assert_eq!(paragraph.collapsed().unwrap(), Some(Collapsed::Disabled));

    let mut document = MutableDocument::new();
    let _ = document.add_paragraph_with_text("generated").collapse();
    let generated = document.to_xml().unwrap();
    assert!(generated.contains("w12:collapsed"));
    assert_eq!(
        document.paragraph_collapsed(0).unwrap(),
        Some(Collapsed::Enabled)
    );
}
