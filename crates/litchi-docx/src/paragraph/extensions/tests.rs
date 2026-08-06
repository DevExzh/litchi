use super::*;
use crate::paragraph::Paragraph;
use crate::table::Row;
use crate::writer::MutableDocument;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = WORD_2010_NAMESPACE;
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[test]
fn parses_paragraph_ids_and_spell_state_without_touching_content() {
    let xml = format!(
        r#"<w:p xmlns:w="{W}" xmlns:w14="{W14}" xmlns:mc="{MC}" mc:Ignorable="w14" w14:paraId="0000002a" w14:textId="7fffffff" w14:noSpellErr="1"><w:r r:id="keep"><x:unknown xmlns:x="urn:test" x:value="keep"/></w:r></w:p>"#
    )
    .into_bytes();
    let paragraph = Paragraph::new(xml.clone());
    let extensions = paragraph.extensions().unwrap();
    assert_eq!(extensions.ids().para_id(), Id::new(0x2a));
    assert_eq!(extensions.ids().text_id(), Id::new(0x7fff_ffff));
    assert_eq!(extensions.no_spell_err(), Some(true));
    assert_eq!(paragraph.xml_bytes(), xml.as_slice());
}

#[test]
fn parses_row_ids_and_ignores_nested_paragraph_attributes() {
    let xml = format!(
        r#"<w:tr xmlns:w="{W}" xmlns:w14="{W14}" w14:paraId="00000001"><w:tc><w:p w14:paraId="00000002"/></w:tc></w:tr>"#
    );
    let row = Row::new(xml.into_bytes());
    let ids = row.extension_ids().unwrap();
    assert_eq!(ids.para_id(), Id::new(1));
    assert_eq!(ids.text_id(), None);
}

#[test]
fn rejects_invalid_lexical_values_and_dependency_states() {
    for value in [
        "", "0000000", "00000000", "80000000", "ffffffff", "0000000g",
    ] {
        assert!(Id::parse(value).is_err(), "accepted invalid ID {value:?}");
    }

    let mut ids = Ids::new();
    let text = Id::new(2).unwrap();
    assert!(ids.set_text_id(Some(text)).is_err());
    assert_eq!(ids, Ids::new());

    let invalid = format!(r#"<w:p xmlns:w="{W}" xmlns:w14="{W14}" w14:textId="00000002"/>"#);
    assert!(Paragraph::new(invalid.into_bytes()).extensions().is_err());

    let invalid_spell = format!(r#"<w:p xmlns:w="{W}" xmlns:w14="{W14}" w14:noSpellErr="on"/>"#);
    assert!(
        Paragraph::new(invalid_spell.into_bytes())
            .extensions()
            .is_err()
    );
}

#[test]
fn generated_paragraph_and_row_attributes_are_checked_and_ignorable() {
    let para_id = Id::new(0x2a).unwrap();
    let text_id = Id::new(0x2b).unwrap();
    let mut document = MutableDocument::new();
    let paragraph = document.add_paragraph_with_text("typed");
    paragraph
        .set_para_id(Some(para_id))
        .unwrap()
        .set_text_id(Some(text_id))
        .unwrap()
        .set_no_spell_err(Some(false));
    let table = document.add_table(1, 1);
    table
        .row(0)
        .unwrap()
        .set_para_id(Some(para_id))
        .unwrap()
        .set_text_id(Some(text_id))
        .unwrap();

    let xml = document.to_xml().unwrap();
    assert!(xml.contains(
        r#"mc:Ignorable="w14" w14:paraId="0000002a" w14:textId="0000002b" w14:noSpellErr="0""#
    ));
    assert!(xml.contains(r#"mc:Ignorable="w14" w14:paraId="0000002a" w14:textId="0000002b""#));
    assert!(xml.contains(&format!(r#"xmlns:w14="{W14}""#)));
}

#[test]
fn invalid_replacement_is_atomic_and_unknown_body_xml_survives() {
    let source = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:test"><w:body><w:p><w:r><w:hyperlink r:id="rId9"><w:t>kept</w:t></w:hyperlink></w:r><x:opaque x:value="keep"/></w:p></w:body></w:document>"#
    );
    let mut document = MutableDocument::from_xml(&source).unwrap();
    let paragraph = document.add_paragraph();
    let before = paragraph.extensions().to_owned();
    assert!(paragraph.set_text_id(Some(Id::new(3).unwrap())).is_err());
    assert_eq!(*paragraph.extensions(), before);

    let output = document.to_xml().unwrap();
    assert!(output.contains(r#"r:id="rId9""#));
    assert!(output.contains(r#"x:opaque x:value="keep""#));
}
