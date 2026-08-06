use super::*;

use crate::paragraph::Run;
use crate::writer::MutableDocument;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const SX: &str = NAMESPACE;

fn run_xml(body: &str) -> String {
    format!(
        r#"<w:r xmlns:w="{W}" xmlns:sx="{SX}" xmlns:mc="{MC}" xmlns:x="urn:future">{body}</w:r>"#
    )
}

#[test]
fn reads_repeated_symbols_and_preserves_absent_attributes() {
    let xml = run_xml(
        r#"<w:rPr><w:b/></w:rPr><sx:symEx/><sx:symEx sx:font="Segoe UI Symbol" sx:char="0001F600"/><w:t>tail</w:t>"#,
    );
    let run = Run::new(xml.into_bytes());
    let symbols = run.symbols().unwrap();

    assert_eq!(symbols.len(), 2);
    assert_eq!(symbols.first().unwrap().font(), None);
    assert_eq!(symbols.first().unwrap().character(), None);
    assert_eq!(symbols.get(1).unwrap().font(), Some("Segoe UI Symbol"));
    assert_eq!(symbols.get(1).unwrap().char_code(), Some(0x1F600));
    assert_eq!(run.sym_ex().unwrap(), symbols.first().cloned());
}

#[test]
fn snapshots_are_failure_atomic_and_patchable_without_losing_unknown_run_xml() {
    let source_xml = run_xml(
        r#"<w:rPr><w:i/></w:rPr><x:future x:value="keep"/><sx:symEx sx:font="Old" sx:char="00000041"/><w:t>tail</w:t>"#,
    );
    let source = Snapshot::from_xml(source_xml.as_bytes().to_vec()).unwrap();
    assert_eq!(source.symbol().unwrap().font(), Some("Old"));

    let mut edit = source.edit();
    let replacement = Symbol::new("New & Font", 0x1F642).unwrap();
    edit.set_symbol(Some(replacement.clone())).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.xml_bytes(), source_xml.as_bytes());
    assert_eq!(commit.snapshot().symbol(), Some(&replacement));
    let changed = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(changed.contains(r#"x:future x:value="keep""#));
    assert!(changed.contains("<w:rPr><w:i/></w:rPr>"));
    assert!(changed.contains("<w:t>tail</w:t>"));
    assert!(changed.contains("New &amp; Font"));

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.symbol().unwrap().font(), Some("Old"));
    assert_eq!(commit.patch().before().len(), 1);
    assert_eq!(commit.patch().after().first(), Some(&replacement));
    assert!(commit.patch().apply(commit.snapshot()).is_err());
}

#[test]
fn raw_run_edits_keep_no_symbol_absent_from_explicit_empty_distinct() {
    let mut absent = Run::new(run_xml(r#"<w:t>plain</w:t>"#).into_bytes());
    assert_eq!(absent.symbol().unwrap(), None);

    absent.set_symbol(Some(Symbol::empty())).unwrap();
    assert_eq!(absent.symbol().unwrap(), Some(Symbol::empty()));
    let authored = String::from_utf8(absent.xml_bytes_for_test().to_vec()).unwrap();
    assert!(authored.contains("symEx"));
    assert!(authored.contains("mc:Ignorable=\"w15sym\""));

    absent.clear_symbol().unwrap();
    assert_eq!(absent.symbol().unwrap(), None);
    assert!(
        String::from_utf8(absent.xml_bytes_for_test().to_vec())
            .unwrap()
            .contains("<w:t>plain</w:t>")
    );
}

#[test]
fn invalid_symbol_domains_are_rejected_before_editing() {
    assert!(Symbol::new("font", 0xD800).is_err());
    assert!(Symbol::new("font", 0x11_0000).is_err());
    assert!(Symbol::from_parts(Some("f".repeat(MAX_FONT_CHARS + 1)), Some(0x41)).is_err());

    let wrong_attribute_namespace = run_xml(
        r#"<sx:symEx w:char="00000041" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
    );
    assert!(
        Run::new(wrong_attribute_namespace.into_bytes())
            .symbols()
            .is_err()
    );

    let invalid_char = run_xml(r#"<sx:symEx sx:char="0000D800"/>"#);
    assert!(Run::new(invalid_char.into_bytes()).symbols().is_err());
}

#[test]
fn mutable_writer_authors_and_reopens_sym_ex() {
    let mut document = MutableDocument::new();
    let run = document.add_paragraph().add_run();
    run.set_symbol_character("Segoe UI Symbol", 0x1F600)
        .unwrap();
    assert_eq!(run.symbol().unwrap().character(), Some(0x1F600));

    let xml = document.to_xml().unwrap();
    assert!(xml.contains("w15sym:symEx"));
    assert!(xml.contains("w15sym:char=\"0001F600\""));
    assert!(xml.contains("mc:Ignorable=\"w15sym\""));

    let start = xml.find("<w:r>").unwrap();
    let end = xml[start..].find("</w:r>").unwrap() + start + "</w:r>".len();
    let reopened = Run::new(xml.as_bytes()[start..end].to_vec());
    assert_eq!(
        reopened.symbol().unwrap().unwrap().font(),
        Some("Segoe UI Symbol")
    );
}

// `Run` deliberately keeps its source bytes crate-private; this narrow test
// helper avoids weakening that facade solely for byte-preservation assertions.
trait RunBytes {
    fn xml_bytes_for_test(&self) -> &[u8];
}

impl RunBytes for Run {
    fn xml_bytes_for_test(&self) -> &[u8] {
        self.xml_bytes()
    }
}
