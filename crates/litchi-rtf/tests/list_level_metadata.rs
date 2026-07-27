//! Round-trip tests for list-level metadata flags: `\lvltentative`,
//! `\levellegal`, `\levelnorestart`, `\levelold`, `\levelprev`,
//! `\levelprevspace`, and `\leveltemplateid`.

use litchi_rtf::{RtfDocument, RtfWriter};

const SOURCE: &str = concat!(
    r"{\rtf1\ansi{\*\listtable{\list\listtemplateid42{\listlevel\levelnfc0\leveljc0\levelfollow0",
    r"\levelstartat1\lvltentative\levellegal1\levelnorestart1\levelold1\levelprev1\levelprevspace1",
    r"\leveltemplateid1234{\leveltext\'02\'00.;}{\levelnumbers\'01;}\f0}{\listname L;}\listid7}}",
    r"\pard Body\par}"
);

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn list_level_metadata_reaches_the_model() {
    let document = RtfDocument::parse(SOURCE).unwrap();
    let level = &document.list_table().lists()[0].levels[0];
    assert!(level.tentative);
    assert!(level.legal_format);
    assert!(level.no_restart);
    assert!(level.legacy);
    assert!(level.include_previous);
    assert!(level.include_previous_space);
    assert_eq!(level.template_id, Some(1234));
}

#[test]
fn list_level_metadata_round_trips_through_the_writer() {
    let document = RtfDocument::parse(SOURCE).unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output).unwrap();
    for marker in [
        r"\lvltentative",
        r"\levellegal",
        r"\levelnorestart",
        r"\levelold",
        r"\levelprev",
        r"\levelprevspace",
        r"\leveltemplateid1234",
    ] {
        assert!(serialized.contains(marker), "missing {marker} in {serialized}");
    }

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    let level = &reparsed.list_table().lists()[0].levels[0];
    assert!(level.tentative);
    assert!(level.legal_format);
    assert!(level.no_restart);
    assert!(level.legacy);
    assert!(level.include_previous);
    assert!(level.include_previous_space);
    assert_eq!(level.template_id, Some(1234));
}

#[test]
fn list_levels_default_to_no_metadata() {
    let source = r"{\rtf1\ansi{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}\pard X\par}";
    let document = RtfDocument::parse(source).unwrap();
    let level = &document.list_table().lists()[0].levels[0];
    assert!(!level.tentative);
    assert!(!level.legal_format);
    assert!(!level.no_restart);
    assert!(!level.legacy);
    assert!(!level.include_previous);
    assert!(!level.include_previous_space);
    assert_eq!(level.template_id, None);
}
