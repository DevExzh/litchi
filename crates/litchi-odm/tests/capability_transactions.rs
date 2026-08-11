#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::{HistoryLimits, Position};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, Profile},
};
use litchi_odm::{
    Master,
    style::Origin,
    subdocument::Target,
    transaction::{Conflict, MergeError},
};
use std::io::{Cursor, Read as _};

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const CONTENT: &str = concat!(
    r#"<?xml version="1.0"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    r#"<office:automatic-styles><style:style style:name="AutoSection" style:family="section"/></office:automatic-styles>"#,
    r#"<office:body><office:text text:global="true"><text:section text:name="Book" xml:id="book">"#,
    r#"<text:section text:name="Chapter" text:style-name="AutoSection" text:protected="true">"#,
    r#"<text:section-source xlink:type="simple" xlink:href="Chapters/a.odt" xlink:show="embed" "#,
    r#"text:section-name="Body" text:filter-name="writer8"/><text:p>Cached</text:p>"#,
    r#"</text:section></text:section></office:text></office:body></office:document-content>"#,
);
const META: &str = concat!(
    r#"<?xml version="1.0"?><office:document-meta "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:dc="http://purl.org/dc/elements/1.1/" "#,
    r#"xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">"#,
    r#"<office:meta><dc:title>Before</dc:title><dc:creator>Editor</dc:creator>"#,
    r#"<meta:user-defined meta:name="opaque">keep</meta:user-defined></office:meta>"#,
    r#"</office:document-meta>"#,
);
const STYLES: &str = concat!(
    r#"<?xml version="1.0"?><office:document-styles "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">"#,
    r#"<office:styles><style:style style:name="Sect1" style:family="section" "#,
    r#"style:parent-style-name="Standard"/></office:styles></office:document-styles>"#,
);

fn package(content: &str, signed: bool) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    writer
        .add_file_with_media_type(
            "Chapters/a.odt",
            b"inert",
            "application/vnd.oasis.opendocument.text",
        )
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/cover.png", b"image", "image/png")
        .unwrap();
    if signed {
        writer
            .add_file("META-INF/documentsignatures.xml", b"<document-signatures/>")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn projects_section_tree_style_catalog_metadata_and_resource_graph() {
    let master = Master::from_bytes(package(CONTENT, false)).unwrap();
    let tree = master.section_tree();
    assert_eq!(tree.roots(), &[Position::new(0)]);
    assert_eq!(tree.sections().len(), 2);
    let chapter = tree.get(Position::new(1)).unwrap();
    assert_eq!(chapter.name(), "Chapter");
    assert_eq!(chapter.parent(), Some(Position::new(0)));
    assert_eq!(chapter.style_name(), Some("AutoSection"));
    assert_eq!(chapter.protected(), Some(true));
    assert_eq!(chapter.reference(), Some(Position::new(0)));

    assert!(
        master
            .styles()
            .iter()
            .any(|style| { style.name() == "AutoSection" && style.origin() == Origin::Content })
    );
    assert!(master.styles().iter().any(|style| {
        style.name() == "Sect1"
            && style.family() == Some("section")
            && style.parent() == Some("Standard")
            && style.origin() == Origin::Styles
    }));
    assert_eq!(master.metadata().unwrap().author.as_deref(), Some("Editor"));

    let chapter_resource = master
        .resources()
        .resources()
        .iter()
        .find(|resource| resource.path() == "Chapters/a.odt")
        .unwrap();
    assert_eq!(chapter_resource.references(), &[Position::new(0)]);
    assert_eq!(master.resources().missing(), &[]);
}

#[test]
fn unified_title_and_link_commit_is_atomic_compact_and_reversible() {
    let source = Master::from_bytes(package(CONTENT, false)).unwrap();
    let mut edit = source.edit();
    edit.set_title("After & exact")
        .unwrap()
        .set_link("Chapter", "Chapters/revised & exact.odt")
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot();
    assert_eq!(changed.title(), Some("After & exact"));
    assert_eq!(
        changed.subdocuments()[0].href(),
        "Chapters/revised & exact.odt"
    );
    assert!(matches!(
        changed.subdocuments()[0].target(),
        Target::Package(_)
    ));
    assert!(changed.content_xml().contains("revised &amp; exact.odt"));
    compact_xml::validate(changed.content_xml().as_bytes()).unwrap();
    assert_eq!(commit.patch().changes().links().len(), 1);
    assert!(commit.patch().changes().title().is_some());
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        source.as_bytes()
    );

    let reopened = Master::from_bytes(changed.as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.title(), Some("After & exact"));
    assert_eq!(reopened.section_tree().sections().len(), 2);
}

#[test]
fn durable_patch_merge_and_bounded_history_preserve_lineage() {
    let source = Master::from_bytes(package(CONTENT, false)).unwrap();

    let mut title_edit = source.edit();
    title_edit.set_title("Merged title").unwrap();
    let title_commit = title_edit.commit().unwrap();
    let mut link_edit = source.edit();
    link_edit
        .set_link(Position::new(0), "Chapters/merged.odt")
        .unwrap();
    let link_commit = link_edit.commit().unwrap();
    let merged = title_commit.patch().merge(link_commit.patch()).unwrap();
    let merged_snapshot = merged.apply(&source).unwrap();
    assert_eq!(merged_snapshot.title(), Some("Merged title"));
    assert_eq!(
        merged_snapshot.subdocuments()[0].href(),
        "Chapters/merged.odt"
    );

    let durable = merged.durable().unwrap();
    let json = durable.to_deterministic_json().unwrap();
    assert_eq!(json, durable.to_deterministic_json().unwrap());
    let decoded = litchi_odm::transaction::DurablePatch::from_deterministic_json(&json).unwrap();
    let applied = decoded.apply(&source).unwrap();
    assert_eq!(applied.as_bytes(), merged_snapshot.as_bytes());
    assert_eq!(
        decoded.inverse().apply(&applied).unwrap().as_bytes(),
        source.as_bytes()
    );
    assert!(decoded.apply(&applied).is_err());

    let mut history = source.history(HistoryLimits::new(4, u64::MAX));
    history.record(&title_commit).unwrap();
    assert_eq!(history.current().title(), Some("Merged title"));
    assert!(history.undo());
    assert_eq!(history.current().title(), Some("Before"));
    assert!(history.redo());
    assert_eq!(history.current().title(), Some("Merged title"));
}

#[test]
fn merge_reports_divergent_title_conflict() {
    let source = Master::from_bytes(package(CONTENT, false)).unwrap();
    let mut left_edit = source.edit();
    left_edit.set_title("Left").unwrap();
    let left_commit = left_edit.commit().unwrap();
    let mut right_edit = source.edit();
    right_edit.set_title("Right").unwrap();
    let right_commit = right_edit.commit().unwrap();
    assert!(matches!(
        left_commit.patch().merge(right_commit.patch()),
        Err(MergeError::Conflicts(conflicts)) if conflicts.len() == 1
    ));
}

#[test]
fn changed_signed_and_encrypted_transactions_are_refused_but_noops_are_exact() {
    let signed = Master::from_bytes(package(CONTENT, true)).unwrap();
    assert_eq!(
        signed.security().changed_write_disposition(),
        litchi_odm::security::ChangedWriteDisposition::RefusedSigned
    );
    assert_eq!(
        signed.security().write_capabilities().signatures(),
        litchi_odm::security::CryptographicWriteCapability::ExactPreservationOnly
    );
    assert_eq!(
        signed.security().write_capabilities().encryption(),
        litchi_odm::security::CryptographicWriteCapability::Unavailable
    );
    assert_eq!(
        signed.edit().commit().unwrap().snapshot().as_bytes(),
        signed.as_bytes()
    );
    let mut signed_edit = signed.edit();
    signed_edit.set_title("invalidates signature").unwrap();
    assert!(signed_edit.commit().is_err());

    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption("secret", Profile::compatible())
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    let encrypted_bytes = writer.finish_to_bytes().unwrap();
    let encrypted = Master::from_bytes_with_password(encrypted_bytes, "secret").unwrap();
    assert_eq!(
        encrypted.security().changed_write_disposition(),
        litchi_odm::security::ChangedWriteDisposition::RefusedEncrypted
    );
    assert_eq!(
        encrypted.security().write_capabilities().encryption(),
        litchi_odm::security::CryptographicWriteCapability::ExactPreservationOnly
    );
    assert_eq!(
        encrypted.edit().commit().unwrap().snapshot().as_bytes(),
        encrypted.as_bytes()
    );
    let mut encrypted_edit = encrypted.edit();
    encrypted_edit
        .set_link(Position::new(0), "changed.odt")
        .unwrap();
    assert!(encrypted_edit.commit().is_err());
}

#[test]
fn malformed_section_and_style_semantics_are_rejected() {
    let invalid_bool = CONTENT.replace("text:protected=\"true\"", "text:protected=\"maybe\"");
    assert!(Master::from_bytes(package(&invalid_bool, false)).is_err());
    let invalid_style = STYLES.replace(" style:name=\"Sect1\"", "");
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer
        .add_file("styles.xml", invalid_style.as_bytes())
        .unwrap();
    assert!(Master::from_bytes(writer.finish_to_bytes().unwrap()).is_err());
}

#[test]
fn original_libreoffice_odm_ingests_edits_and_reopens_without_repacking() {
    let original = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/odfexport/data/tdf121119.odm"
    );
    let master = Master::from_bytes(original.to_vec()).unwrap();
    assert_eq!(master.as_bytes(), original);
    assert_eq!(master.section_tree().sections().len(), 2);
    assert!(
        master
            .section_tree()
            .sections()
            .iter()
            .all(|section| section.protected() == Some(true))
    );
    assert!(master.styles().iter().any(|style| style.name() == "Sect1"));
    assert_eq!(master.subdocuments().len(), 2);
    assert!(
        master
            .subdocuments()
            .iter()
            .all(|reference| reference.target().is_external())
    );
    let mut edit = master.edit();
    edit.set_title("Edited genuine ODM")
        .unwrap()
        .set_link(Position::new(0), "../edited-DUMMY2.odt")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().title(), Some("Edited genuine ODM"));
    assert_eq!(
        commit.snapshot().subdocuments()[0].href(),
        "../edited-DUMMY2.odt"
    );
    let mut changed_archive =
        zip::ZipArchive::new(Cursor::new(commit.snapshot().as_bytes())).unwrap();
    let mut mimetype = String::new();
    {
        let mut first = changed_archive.by_index(0).unwrap();
        assert_eq!(first.name(), "mimetype");
        assert_eq!(first.compression(), zip::CompressionMethod::Stored);
        first.read_to_string(&mut mimetype).unwrap();
    }
    assert_eq!(mimetype, MIME);
    assert!(Master::from_bytes(commit.snapshot().as_bytes().to_vec()).is_ok());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        original
    );
}

#[test]
fn three_way_planning_detects_parent_remove_against_child_edit() {
    let source = Master::from_bytes(package(CONTENT, false)).unwrap();
    let mut remove_parent_edit = source.edit();
    remove_parent_edit.remove_section(Position::new(0)).unwrap();
    let remove_parent_commit = remove_parent_edit.commit().unwrap();
    let mut rename_child_edit = source.edit();
    rename_child_edit
        .rename_section(Position::new(1), "Renamed child")
        .unwrap();
    let rename_child_commit = rename_child_edit.commit().unwrap();
    let plan = remove_parent_commit
        .patch()
        .plan_three_way(rename_child_commit.patch())
        .unwrap();
    assert!(plan.conflicts().conflicts().iter().any(
        |conflict| matches!(conflict, Conflict::Section(position) if *position == Position::new(1))
    ));
    assert!(!plan.can_commit());
}
