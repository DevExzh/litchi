use litchi_rtf::{ProtectionUser, ProtectionUserTable, RtfDocument, RtfWriter};
use std::borrow::Cow;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_spec_grammar_as_ordered_inert_metadata_and_round_trips() {
    let document = RtfDocument::parse(
        r#"{\rtf1\ansi{\*\protusertbl{DOMAIN\'5cuserone}{\u30740?\u21457?\u92?\u20320?}}Body}"#,
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    let users = document.protection_user_table().unwrap().users();
    assert_eq!(users.len(), 2);
    assert_eq!(users[0].name, r#"DOMAIN\userone"#);
    assert_eq!(users[1].name, "研发\\你");

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.protection_user_table(),
        document.protection_user_table()
    );
}

#[test]
fn typed_document_api_mutates_validated_inert_table() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let mut table = ProtectionUserTable::new(vec![
        ProtectionUser::new(Cow::Borrowed("DOMAIN\\alice")).unwrap(),
    ])
    .unwrap();
    table
        .push(ProtectionUser::new(Cow::Borrowed("bob@example.test")).unwrap())
        .unwrap();
    document.set_protection_user_table(table).unwrap();

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(
        reparsed.protection_user_table(),
        document.protection_user_table()
    );
    assert_eq!(reparsed.text(), "Body");

    document.clear_protection_user_table();
    assert!(document.protection_user_table().is_none());
}

#[test]
fn rejects_malformed_active_or_oversized_tables() {
    let cases = [
        r#"{\rtf1{\protusertbl{DOMAIN\'5calice}}Body}"#,
        r#"{\rtf1{\*\protusertbl}Body}"#,
        r#"{\rtf1{\*\protusertbl{}}Body}"#,
        r#"{\rtf1{\*\protusertbl alice}Body}"#,
        r#"{\rtf1{\*\protusertbl{{nested}}}Body}"#,
        r#"{\rtf1{\*\protusertbl{alice\b bob}}Body}"#,
        r#"{\rtf1{\*\protusertbl{\bin2 xx}}Body}"#,
        r#"{\rtf1\protusertbl Body}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }

    let oversized = format!(r#"{{\rtf1{{\*\protusertbl{{{}}}}}}}"#, "x".repeat(65_537));
    assert!(RtfDocument::parse(&oversized).is_err());
    assert!(ProtectionUser::new(Cow::Borrowed("\n")).is_err());
    assert!(ProtectionUserTable::new(Vec::new()).is_err());
}
