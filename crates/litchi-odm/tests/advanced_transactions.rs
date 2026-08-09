#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Position;
use litchi_odf_common::core::PackageWriter;
use litchi_odm::{
    Master,
    style::Origin,
    transaction::{MergeError, ResourceSpec, SectionSpec, SecurityPolicy, StyleSpec},
};

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const CONTENT: &str = concat!(
    r#"<?xml version="1.0"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#,
    r#"<style:style style:name="AutoSection" style:family="section"/></office:automatic-styles>"#,
    r#"<office:body><office:text><text:section text:name="Chapter" text:style-name="AutoSection">"#,
    r#"<text:section-source xlink:href="Chapters/a.odt"/><text:p>cached</text:p>"#,
    r#"</text:section><text:section text:name="Local"><text:p>local</text:p></text:section>"#,
    r#"</office:text></office:body></office:document-content>"#,
);
const STYLES: &str = concat!(
    r#"<?xml version="1.0"?><office:document-styles "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">"#,
    r#"<office:styles><style:style style:name="Standard" style:family="paragraph"/>"#,
    r#"</office:styles></office:document-styles>"#,
);
const META: &str = concat!(
    r#"<?xml version="1.0"?><office:document-meta "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:dc="http://purl.org/dc/elements/1.1/" "#,
    r#"xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">"#,
    r#"<office:meta><dc:title>Before</dc:title><dc:creator>Author</dc:creator>"#,
    r#"</office:meta></office:document-meta>"#,
);

fn source() -> Master {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer
        .add_file_with_media_type(
            "Chapters/a.odt",
            b"chapter",
            "application/vnd.oasis.opendocument.text",
        )
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/cover.png", b"cover", "image/png")
        .unwrap();
    Master::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
}

#[test]
fn one_transaction_updates_tree_styles_resources_metadata_and_dependencies() {
    let source = source();
    let mut metadata = source.metadata().unwrap().clone();
    metadata.title = Some("After".to_string());
    metadata.author = Some("Revised author".to_string());
    metadata.subject = Some("Book".to_string());

    let mut edit = source.edit();
    edit.set_metadata(metadata)
        .unwrap()
        .rename_section(Position::new(0), "Renamed Chapter")
        .unwrap()
        .rename_style(Origin::Content, "AutoSection", "RenamedSectionStyle")
        .unwrap()
        .add_style(StyleSpec::new("TailStyle", "section").unwrap())
        .unwrap()
        .add_section(
            SectionSpec::new("Tail")
                .unwrap()
                .with_style("TailStyle")
                .unwrap(),
        )
        .unwrap()
        .put_resource(ResourceSpec::new("Pictures/new.png", "image/png", b"new".to_vec()).unwrap())
        .unwrap()
        .remove_resource("Pictures/cover.png")
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot();
    assert_eq!(changed.title(), Some("After"));
    assert_eq!(
        changed.metadata().unwrap().author.as_deref(),
        Some("Revised author")
    );
    assert!(
        changed
            .section_tree()
            .sections()
            .iter()
            .any(|section| section.name() == "Renamed Chapter"
                && section.style_name() == Some("RenamedSectionStyle"))
    );
    assert!(
        changed
            .section_tree()
            .sections()
            .iter()
            .any(|section| section.name() == "Tail" && section.style_name() == Some("TailStyle"))
    );
    assert!(
        changed
            .styles()
            .iter()
            .any(|style| style.name() == "RenamedSectionStyle")
    );
    assert!(
        changed
            .resources()
            .resources()
            .iter()
            .any(|resource| resource.path() == "Pictures/new.png")
    );
    assert!(
        !changed
            .resources()
            .resources()
            .iter()
            .any(|resource| resource.path() == "Pictures/cover.png")
    );
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        source.as_bytes()
    );
    assert!(Master::from_bytes(changed.as_bytes().to_vec()).is_ok());
}

#[test]
fn section_and_style_create_delete_are_reversible() {
    let source = source();
    let mut add = source.edit();
    add.add_style(StyleSpec::new("Disposable", "section").unwrap())
        .unwrap()
        .add_section(
            SectionSpec::new("Disposable Section")
                .unwrap()
                .with_style("Disposable")
                .unwrap(),
        )
        .unwrap();
    let added = add.commit().unwrap().into_snapshot();
    let section_position = Position::new(added.section_tree().sections().len() - 1);
    let mut remove = added.edit();
    remove.remove_section(section_position).unwrap();
    let without_section = remove.commit().unwrap().into_snapshot();
    let mut remove_style = without_section.edit();
    remove_style
        .remove_style(Origin::Styles, "Disposable")
        .unwrap();
    let commit = remove_style.commit().unwrap();
    assert!(
        !commit
            .snapshot()
            .styles()
            .iter()
            .any(|style| style.name() == "Disposable")
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        without_section.as_bytes()
    );
}

#[test]
fn dependency_and_security_policy_refusals_are_typed_errors() {
    let source = source();
    let mut linked_remove = source.edit();
    linked_remove.remove_resource("Chapters/a.odt").unwrap();
    assert!(linked_remove.commit().is_err());

    let mut style_remove = source.edit();
    style_remove
        .remove_style(Origin::Content, "AutoSection")
        .unwrap();
    assert!(style_remove.commit().is_err());

    let mut strict = source.edit_with_policy(SecurityPolicy::strict());
    assert!(
        strict
            .set_link(Position::new(0), "https://example.invalid/chapter.odt")
            .is_err()
    );

    let mut resolved = source.edit();
    resolved
        .put_resource(
            ResourceSpec::new(
                "Chapters/replacement.odt",
                "application/vnd.oasis.opendocument.text",
                b"replacement".to_vec(),
            )
            .unwrap(),
        )
        .unwrap()
        .set_link(Position::new(0), "Chapters/replacement.odt")
        .unwrap()
        .remove_resource("Chapters/a.odt")
        .unwrap();
    assert!(resolved.commit().is_ok());
}

#[test]
fn transfer_durable_stale_and_three_way_plan_share_the_strong_boundary() {
    let source = source();
    let mut transfer = source.edit();
    transfer
        .transfer_resource(&source, "Pictures/cover.png", "Pictures/copied.png")
        .unwrap();
    let transfer_commit = transfer.commit().unwrap();
    let durable = transfer_commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded = litchi_odm::transaction::DurablePatch::from_deterministic_json(&wire).unwrap();
    assert_eq!(
        decoded.apply(&source).unwrap().as_bytes(),
        transfer_commit.snapshot().as_bytes()
    );
    assert!(decoded.apply(transfer_commit.snapshot()).is_err());

    let mut metadata = source.metadata().unwrap().clone();
    metadata.description = Some("planned".to_string());
    let mut left = source.edit();
    left.set_metadata(metadata).unwrap();
    let left = left.commit().unwrap();
    let plan = left
        .patch()
        .plan_three_way(transfer_commit.patch())
        .unwrap();
    assert!(plan.can_commit());
    let merged = plan.commit().unwrap().apply(&source).unwrap();
    assert_eq!(
        merged.metadata().unwrap().description.as_deref(),
        Some("planned")
    );
    assert!(
        merged
            .resources()
            .resources()
            .iter()
            .any(|resource| resource.path() == "Pictures/copied.png")
    );

    let mut right = source.edit();
    right.rename_section(Position::new(0), "Right").unwrap();
    let right = right.commit().unwrap();
    let mut conflicting = source.edit();
    conflicting
        .rename_section(Position::new(0), "Conflicting")
        .unwrap();
    let conflicting = conflicting.commit().unwrap();
    let plan = right.patch().plan_three_way(conflicting.patch()).unwrap();
    assert!(!plan.can_commit());
    assert!(matches!(plan.commit(), Err(MergeError::Conflicts(_))));
}
