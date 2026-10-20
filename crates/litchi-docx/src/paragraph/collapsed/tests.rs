use super::*;
use crate::Paragraph;
use crate::builder::Document as DraftDocument;

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
fn direct_and_generated_paragraph_facades_emit_ignorable_extension() {
    let mut paragraph = Paragraph::new(paragraph_xml(r#"<w:r><w:t>one</w:t></w:r>"#).into_bytes());
    paragraph.set_collapsed(Some(Collapsed::Disabled)).unwrap();
    let output = String::from_utf8(paragraph.xml_bytes().to_vec()).unwrap();
    assert!(output.contains("w12:collapsed"));
    assert!(output.contains("mc:Ignorable=\"w12\""));
    assert_eq!(paragraph.collapsed().unwrap(), Some(Collapsed::Disabled));

    let mut document = DraftDocument::new();
    let _ = document.add_paragraph_with_text("generated").collapse();
    let generated = document.to_xml().unwrap();
    assert!(generated.contains("w12:collapsed"));
    assert_eq!(
        document.paragraph_collapsed(0).unwrap(),
        Some(Collapsed::Enabled)
    );
}
