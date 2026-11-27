//! Stored native mail-merge data and legacy form caches remain inert.
use litchi_docx::{
    Package,
    field::{LegacyFormKind, MergeCounterKind},
    mail_merge::{
        Conformance, DataType, Destination, MainDocumentType, RelationshipId, Source, Target,
    },
};
use std::collections::BTreeMap;

const SOURCE: &str = "file:////Users/ryker/CodeProjects/litchi/test-data/microsoft-office/word/word-mailmerge-source.docx";
const QUERY: &str = "SELECT * FROM /Users/ryker/CodeProjects/litchi/test-data/microsoft-office/word/word-mailmerge-source.docx";

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}

fn members(package: &Package) -> BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        package.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

#[test]
fn native_mail_merge_settings_and_field_caches_are_inert() {
    let package = native("word-mailmerge-fields.docx");
    let settings = package.mail_merge_settings().unwrap().unwrap();
    assert_eq!(settings.main_document_type(), MainDocumentType::FormLetters);
    assert!(settings.link_to_query());
    assert_eq!(settings.data_type(), Some(DataType::TextFile));
    assert_eq!(settings.query(), Some(QUERY));
    assert_eq!(settings.data_source_relationship_id(), Some("rId1"));
    assert_eq!(settings.header_source_relationship_id(), None);
    assert_eq!(settings.destination(), Destination::NewDocument);
    assert_eq!(settings.active_record(), 1);
    assert_eq!(settings.check_errors(), 2);
    assert!(settings.odso().is_none());
    assert_eq!(
        package
            .mail_merge_target(&RelationshipId::new("rId1").unwrap())
            .unwrap(),
        Target::External(SOURCE.into())
    );
    let document = package.document().unwrap();
    assert!(document.mail_merge_recipients().unwrap().is_none());
    let fields = document.fields().unwrap();
    assert_eq!(fields.len(), 7);
    assert!(fields.iter().all(|f| !f.is_dirty() && !f.is_locked()));
    let merge = document
        .merge_fields()
        .unwrap()
        .iter()
        .map(|f| f.merge_field().unwrap().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(
        merge
            .iter()
            .map(|f| (f.field_name(), f.cached_result()))
            .collect::<Vec<_>>(),
        [
            ("FirstName", Some("«FirstName»")),
            ("LastName", Some("«LastName»")),
            ("Amount", Some("«Amount»")),
            ("FirstName", None)
        ]
    );
    let condition = fields
        .iter()
        .find(|f| f.instruction().trim_start().starts_with("IF "))
        .unwrap();
    assert_eq!(
        condition.instruction().trim(),
        "IF Alex = \"Alex\" \"MATCH_009\" \"OTHER_009\""
    );
    assert_eq!(condition.result(), Some("MATCH_009"));
    let counters = document.mail_merge_counters().unwrap();
    assert_eq!(
        counters
            .iter()
            .map(|f| (f.kind(), f.cached_result()))
            .collect::<Vec<_>>(),
        [
            (MergeCounterKind::Record, Some("«Merge Record #»")),
            (MergeCounterKind::Sequence, Some("«Merge Sequence #»")),
        ]
    );
}

#[test]
fn native_mail_merge_source_table_has_two_exact_records() {
    let package = native("word-mailmerge-source.docx");
    assert!(package.mail_merge_settings().unwrap().is_none());
    let document = package.document().unwrap();
    assert!(document.fields().unwrap().is_empty());
    let tables = document.tables().unwrap();
    assert_eq!(tables.len(), 1);
    let rows = tables[0]
        .rows()
        .unwrap()
        .iter()
        .map(|row| {
            row.cells()
                .unwrap()
                .iter()
                .map(|cell| cell.text().unwrap())
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    assert_eq!(
        rows,
        [
            vec![
                "Title",
                "FirstName",
                "LastName",
                "JobTitle",
                "Company",
                "Address1",
                "Address2",
                "City",
                "State",
                "PostalCode",
                "Country",
                "HomePhone",
                "WorkPhone",
                "Amount"
            ],
            vec![
                "",
                "Alex",
                "Fixture",
                "",
                "NATIVE_MERGE_009",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "42.50"
            ],
            vec![
                "",
                "Blair",
                "Sample",
                "",
                "NATIVE_MERGE_009",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "84.00"
            ],
        ]
    );
}

#[test]
fn native_mail_merge_update_clear_and_replay_preserve_body_and_other_parts() {
    let package = native("word-mailmerge-fields.docx");
    let original = members(&package);
    let mut settings = package.mail_merge_settings().unwrap().unwrap();
    settings.set_active_record(2).set_check_errors(1);
    let mut edit = package.edit_mail_merge();
    edit.update(
        settings,
        Some(Source::External(SOURCE.into())),
        None,
        None,
        Conformance::Transitional,
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("changed.docx");
    commit.package().save(&path).unwrap();
    let reopened = Package::open(&path).unwrap();
    let changed = reopened.mail_merge_settings().unwrap().unwrap();
    assert_eq!(changed.active_record(), 2);
    assert_eq!(changed.check_errors(), 1);
    assert_eq!(changed.data_source_relationship_id(), Some("rId1"));
    assert_eq!(changed.query(), Some(QUERY));
    assert_eq!(
        reopened
            .mail_merge_target(
                &RelationshipId::new(changed.data_source_relationship_id().unwrap()).unwrap()
            )
            .unwrap(),
        Target::External(SOURCE.into())
    );
    let updated_members = members(&reopened);
    assert_eq!(
        updated_members.keys().collect::<Vec<_>>(),
        original.keys().collect::<Vec<_>>()
    );
    for (name, bytes) in &original {
        if name != "word/settings.xml" {
            assert!(updated_members.get(name) == Some(bytes), "changed {name}");
        }
    }
    let restored = commit
        .package()
        .apply_mail_merge(&commit.patch().inverse())
        .unwrap();
    let restored_members = members(restored.package());
    assert_eq!(
        restored_members.keys().collect::<Vec<_>>(),
        original.keys().collect::<Vec<_>>()
    );
    for (name, bytes) in &original {
        assert!(
            restored_members.get(name) == Some(bytes),
            "inverse changed {name}"
        );
    }
    let mut clear = reopened.edit_mail_merge();
    assert!(clear.clear().unwrap());
    let clear_commit = clear.commit().unwrap();
    let recovered = clear_commit
        .package()
        .apply_mail_merge(&clear_commit.patch().inverse())
        .unwrap();
    assert!(members(recovered.package()) == updated_members);
    let cleared = Package::from_bytes(&clear_commit.package().to_bytes().unwrap()).unwrap();
    assert!(cleared.mail_merge_settings().unwrap().is_none());
    let cleared_members = members(&cleared);
    for (name, bytes) in &original {
        if name != "word/settings.xml" && name != "word/_rels/settings.xml.rels" {
            assert!(
                cleared_members.get(name) == Some(bytes),
                "clear changed {name}"
            );
        }
    }
    // OPC may retain an empty relationship member after removing its edges.
    assert!(
        cleared
            .opc_package()
            .get_part(&litchi_opc::PackURI::new("/word/settings.xml").unwrap())
            .unwrap()
            .rels()
            .is_empty()
    );
    assert!(members(&package) == original);
}

#[test]
fn native_mail_merge_source_reuse_requires_matching_target_and_dialect() {
    for (target, conformance, prefix) in [
        (
            "https://example.invalid/never-fetch.csv",
            Conformance::Transitional,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
        ),
        (
            SOURCE,
            Conformance::Strict,
            "http://purl.oclc.org/ooxml/officeDocument/relationships/",
        ),
    ] {
        let package = native("word-mailmerge-fields.docx");
        let mut edit = package.edit_mail_merge();
        edit.update(
            package.mail_merge_settings().unwrap().unwrap(),
            Some(Source::External(target.into())),
            None,
            None,
            conformance,
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
        let settings = reopened.mail_merge_settings().unwrap().unwrap();
        let id = settings.data_source_relationship_id().unwrap();
        assert_eq!(
            reopened
                .mail_merge_target(&RelationshipId::new(id).unwrap())
                .unwrap(),
            Target::External(target.into())
        );
        let part = reopened
            .opc_package()
            .get_part(&litchi_opc::PackURI::new("/word/settings.xml").unwrap())
            .unwrap();
        assert_eq!(part.rels().len(), 1);
        assert_eq!(
            part.rels().get(id).unwrap().reltype(),
            format!("{prefix}mailMergeSource")
        );
        assert_eq!(
            package
                .mail_merge_target(&RelationshipId::new("rId1").unwrap())
                .unwrap(),
            Target::External(SOURCE.into())
        );
    }
}

#[test]
fn native_legacy_form_instructions_caches_and_bookmarks_are_not_evaluated() {
    for name in ["word-legacy-form.docx", "word-legacy-form-unprotected.docx"] {
        let package = native(name);
        let document = package.document().unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(
            fields
                .iter()
                .map(|f| (f.kind(), f.cached_result()))
                .collect::<Vec<_>>(),
            [
                (LegacyFormKind::Text, Some("FORM_VALUE_010")),
                (LegacyFormKind::CheckBox, None),
                (LegacyFormKind::DropDown, None)
            ]
        );
        assert!(
            fields
                .iter()
                .all(|f| !f.is_dirty() && !f.is_locked() && f.opaque_instructions().is_empty())
        );
        let bookmarks = document.bookmarks().unwrap();
        assert_eq!(
            bookmarks
                .iter()
                .map(|b| (b.id(), b.name()))
                .collect::<Vec<_>>(),
            [
                (0, "FixtureText010"),
                (1, "FixtureCheck010"),
                (2, "FixtureDropdown010")
            ]
        );
    }
}
