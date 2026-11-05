#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{ProtectionUser, ProtectionUserTable, write::Writer};
use std::borrow::Cow;

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_spec_grammar_as_ordered_inert_metadata_and_round_trips() {
    let document = Tree::parse(
        r"{\rtf1\ansi{\*\protusertbl{DOMAIN\'5cuserone}{\u30740?\u21457?\u92?\u20320?}}Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    let users = document.protection_user_table().unwrap().users();
    assert_eq!(users.len(), 2);
    assert_eq!(users[0].name, r"DOMAIN\userone");
    assert_eq!(users[1].name, "研发\\你");

    let output = write(&document);
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.protection_user_table(),
        document.protection_user_table()
    );
}

#[test]
fn typed_document_api_mutates_validated_inert_table() {
    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    let mut table = ProtectionUserTable::new(vec![
        ProtectionUser::new(Cow::Borrowed("DOMAIN\\alice")).unwrap(),
    ])
    .unwrap();
    table
        .push(ProtectionUser::new(Cow::Borrowed("bob@example.test")).unwrap())
        .unwrap();
    document.set_protection_user_table(table).unwrap();

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
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
        r"{\rtf1{\protusertbl{DOMAIN\'5calice}}Body}",
        r"{\rtf1{\*\protusertbl}Body}",
        r"{\rtf1{\*\protusertbl{}}Body}",
        r"{\rtf1{\*\protusertbl alice}Body}",
        r"{\rtf1{\*\protusertbl{{nested}}}Body}",
        r"{\rtf1{\*\protusertbl{alice\b bob}}Body}",
        r"{\rtf1{\*\protusertbl{\bin2 xx}}Body}",
        r"{\rtf1\protusertbl Body}",
    ];
    for rtf in cases {
        assert!(
            Tree::parse(rtf, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {rtf}"
        );
    }

    let oversized = format!(r"{{\rtf1{{\*\protusertbl{{{}}}}}}}", "x".repeat(65_537));
    assert!(Tree::parse(&oversized, litchi_rtf::read::OpenOptions::default()).is_err());
    assert!(ProtectionUser::new(Cow::Borrowed("\n")).is_err());
    assert!(ProtectionUserTable::new(Vec::new()).is_err());
}
