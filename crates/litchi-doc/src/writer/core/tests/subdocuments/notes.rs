use super::super::support::*;

#[test]
fn test_footnotes() {
    let mut writer = Writer::new();
    let entry = FootnoteEntry::new(0u32, "This is a footnote", 1u16);
    writer.add_footnote(entry);
    assert_eq!(writer.footnotes.len(), 1);
    assert_eq!(writer.footnotes[0].text, "This is a footnote");
}

#[test]
fn test_endnotes() {
    let mut writer = Writer::new();
    let entry = FootnoteEntry::new(0u32, "This is an endnote", 1u16);
    writer.add_endnote(entry);
    assert_eq!(writer.endnotes.len(), 1);
    assert_eq!(writer.endnotes[0].text, "This is an endnote");
}
