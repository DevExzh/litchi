#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::{HistoryLimits, Position};
use litchi_odf_common::core::PackageWriter;
use litchi_odm::{
    Master,
    style::Origin,
    transaction::{
        ActiveContentPolicy, Conflict, MergeError, ResourceSpec, SectionSpec, SecurityPolicy,
        StyleSpec, SubdocumentSpec,
    },
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

fn master_from_parts(content: &str, styles: &str, chapter: &[u8]) -> Master {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer
        .add_file_with_media_type(
            "Chapters/a.odt",
            chapter,
            "application/vnd.oasis.opendocument.text",
        )
        .unwrap();
    Master::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
}

fn master_without_styles(content: &str, chapter: &[u8]) -> Master {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer
        .add_file_with_media_type(
            "Chapters/a.odt",
            chapter,
            "application/vnd.oasis.opendocument.text",
        )
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
    let mut metadata_edit = source.edit();
    metadata_edit.set_metadata(metadata).unwrap();
    let metadata_commit = metadata_edit.commit().unwrap();
    let metadata_plan = metadata_commit
        .patch()
        .plan_three_way(transfer_commit.patch())
        .unwrap();
    assert!(metadata_plan.can_commit());
    let merged = metadata_plan.commit().unwrap().apply(&source).unwrap();
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

    let mut right_edit = source.edit();
    right_edit
        .rename_section(Position::new(0), "Right")
        .unwrap();
    let right_commit = right_edit.commit().unwrap();
    let mut conflicting_edit = source.edit();
    conflicting_edit
        .rename_section(Position::new(0), "Conflicting")
        .unwrap();
    let conflicting_commit = conflicting_edit.commit().unwrap();
    let conflict_plan = right_commit
        .patch()
        .plan_three_way(conflicting_commit.patch())
        .unwrap();
    assert!(!conflict_plan.can_commit());
    assert!(matches!(
        conflict_plan.commit(),
        Err(MergeError::Conflicts(_))
    ));
}

#[test]
fn linked_section_transfer_and_subtree_cleanup_close_package_dependencies() {
    let source = source();
    let mut destination_edit = source.edit();
    destination_edit.set_title("Destination master").unwrap();
    let destination = destination_edit.commit().unwrap().into_snapshot();
    let mut transfer = destination.edit_with_policy(SecurityPolicy::strict());
    transfer
        .transfer_linked_section(
            &source,
            Position::new(0),
            "Imported chapter",
            "Chapters/imported.odt",
        )
        .unwrap();
    let commit = transfer.commit().unwrap();
    let changed = commit.snapshot();
    let imported = changed
        .section_tree()
        .sections()
        .iter()
        .find(|section| section.name() == "Imported chapter")
        .unwrap();
    let reference = &changed.subdocuments()[imported.reference().unwrap().get()];
    assert_eq!(reference.href(), "Chapters/imported.odt");
    assert_eq!(reference.source_section(), None);
    assert_eq!(changed.resources().missing(), &[]);
    assert_eq!(
        changed
            .resources()
            .resources()
            .iter()
            .find(|resource| resource.path() == "Chapters/imported.odt")
            .unwrap()
            .references(),
        &[imported.reference().unwrap()]
    );

    let imported_position = changed
        .section_tree()
        .sections()
        .iter()
        .position(|section| section.name() == "Imported chapter")
        .map(Position::new)
        .unwrap();
    let mut cleanup = changed.edit();
    cleanup
        .remove_section_with_orphaned_resources(imported_position)
        .unwrap();
    let cleaned = cleanup.commit().unwrap().into_snapshot();
    assert!(
        cleaned
            .section_tree()
            .sections()
            .iter()
            .all(|section| section.name() != "Imported chapter")
    );
    assert!(
        cleaned
            .resources()
            .resources()
            .iter()
            .all(|resource| resource.path() != "Chapters/imported.odt")
    );
    assert!(Master::from_bytes(cleaned.as_bytes().to_vec()).is_ok());
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        destination.as_bytes()
    );
}

#[test]
fn final_security_graph_atomic_style_cleanup_and_named_plans_are_checked() {
    let source = source();

    let mut missing_edit = source.edit();
    missing_edit
        .set_link(Position::new(0), "Chapters/missing.odt")
        .unwrap();
    let missing_snapshot = missing_edit.commit().unwrap().into_snapshot();
    assert!(!missing_snapshot.resources().missing().is_empty());
    let mut strict_change = missing_snapshot.edit_with_policy(SecurityPolicy::strict());
    strict_change.set_title("blocked by final graph").unwrap();
    assert!(strict_change.commit().is_err());

    let mut atomic_cleanup = source.edit();
    atomic_cleanup
        .remove_section_with_orphaned_resources(Position::new(0))
        .unwrap()
        .remove_style(Origin::Content, "AutoSection")
        .unwrap();
    let cleaned = atomic_cleanup.commit().unwrap().into_snapshot();
    assert!(
        cleaned
            .styles()
            .iter()
            .all(|style| style.name() != "AutoSection")
    );
    assert!(
        cleaned
            .resources()
            .resources()
            .iter()
            .all(|resource| resource.path() != "Chapters/a.odt")
    );

    let mut author_metadata = source.metadata().unwrap().clone();
    author_metadata.author = Some("Left author".to_string());
    let mut author_edit = source.edit();
    author_edit.set_metadata(author_metadata).unwrap();
    let author_commit = author_edit.commit().unwrap();
    let mut description_metadata = source.metadata().unwrap().clone();
    description_metadata.description = Some("Right description".to_string());
    let mut description_edit = source.edit();
    description_edit.set_metadata(description_metadata).unwrap();
    let description_commit = description_edit.commit().unwrap();
    let metadata_plan = author_commit
        .patch()
        .plan_three_way(description_commit.patch())
        .unwrap();
    assert!(metadata_plan.can_commit());
    let metadata_merged = metadata_plan.commit().unwrap().apply(&source).unwrap();
    assert_eq!(
        metadata_merged.metadata().unwrap().author.as_deref(),
        Some("Left author")
    );
    assert_eq!(
        metadata_merged.metadata().unwrap().description.as_deref(),
        Some("Right description")
    );
    let mut title_edit = source.edit();
    title_edit.set_title("Planned title").unwrap();
    let title_commit = title_edit.commit().unwrap();
    for plan in [
        title_commit
            .patch()
            .plan_three_way(author_commit.patch())
            .unwrap(),
        author_commit
            .patch()
            .plan_three_way(title_commit.patch())
            .unwrap(),
    ] {
        let merged = plan.commit().unwrap().apply(&source).unwrap();
        assert_eq!(merged.title(), Some("Planned title"));
        assert_eq!(
            merged.metadata().unwrap().author.as_deref(),
            Some("Left author")
        );
    }

    let mut left_edit = source.edit();
    left_edit
        .add_section(SectionSpec::new("Same destination").unwrap())
        .unwrap();
    let left_commit = left_edit.commit().unwrap();
    let identical_plan = left_commit
        .patch()
        .plan_three_way(left_commit.patch())
        .unwrap();
    assert!(identical_plan.can_commit());
    let identical = identical_plan.commit().unwrap().apply(&source).unwrap();
    assert_eq!(
        identical
            .section_tree()
            .sections()
            .iter()
            .filter(|section| section.name() == "Same destination")
            .count(),
        1
    );
    let mut right_edit = source.edit();
    right_edit
        .add_section(
            SectionSpec::new("Same destination")
                .unwrap()
                .with_subdocument(
                    SubdocumentSpec::new("Chapters/other.odt")
                        .unwrap()
                        .with_source_section("Body")
                        .unwrap(),
                ),
        )
        .unwrap();
    let right_commit = right_edit.commit().unwrap();
    let named_conflict_plan = left_commit
        .patch()
        .plan_three_way(right_commit.patch())
        .unwrap();
    assert!(named_conflict_plan.conflicts().conflicts().iter().any(
        |conflict| matches!(conflict, Conflict::SectionName(name) if name == "Same destination")
    ));
    assert!(!named_conflict_plan.can_commit());
}

#[test]
fn common_master_structure_local_references_and_active_write_policy_are_explicit() {
    use litchi_odm::{
        security::ActiveKind,
        structure::{IndexKind, Kind},
    };

    let content = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink"><office:scripts>"#,
        r#"<script:event-listener script:event-name="on-load" script:macro-name="Main.Run" xlink:href="vnd.sun.star.script:Main.Run"/>"#,
        r#"</office:scripts><office:body><office:text><text:p text:style-name="Standard">front</text:p>"#,
        r#"<text:table-of-content text:name="Contents" xml:id="toc1">"#,
        r#"<text:table-of-content-source/><text:index-body><text:p>cached</text:p>"#,
        r#"</text:index-body></text:table-of-content>"#,
        r#"<text:section text:name="Target"><text:p>target</text:p></text:section>"#,
        r#"<text:section text:name="LocalSource"><text:section-source text:section-name="Target"/>"#,
        r#"</text:section><text:section text:name="Dde"><office:dde-source/></text:section>"#,
        r#"<text:section text:name="Unresolved"><text:section-source text:section-name="Missing"/>"#,
        r#"</text:section>"#,
        r#"<text:list><text:list-item><text:p>one</text:p></text:list-item></text:list>"#,
        r#"<table:table><table:table-column/><table:table-row><table:table-cell>"#,
        r#"<text:p>cell</text:p></table:table-cell></table:table-row></table:table>"#,
        r#"</office:text></office:body></office:document-content>"#,
    );
    let master = master_from_parts(content, STYLES, b"chapter");
    assert_eq!(
        master.structure().items(),
        &[
            Kind::Paragraph,
            Kind::GeneratedIndex(IndexKind::TableOfContents),
            Kind::Section(Position::new(0)),
            Kind::Section(Position::new(1)),
            Kind::Section(Position::new(2)),
            Kind::Section(Position::new(3)),
            Kind::List,
            Kind::Table,
        ]
    );
    let local = &master.section_tree().local_references()[0];
    assert_eq!(local.owner(), Position::new(1));
    assert_eq!(local.target_name(), "Target");
    assert_eq!(local.target(), Some(Position::new(0)));
    let unresolved = master
        .section_tree()
        .unresolved_local_references()
        .next()
        .unwrap();
    assert_eq!(unresolved.owner(), Position::new(3));
    assert_eq!(unresolved.target_name(), "Missing");
    assert_eq!(
        master
            .section_tree()
            .get(Position::new(1))
            .unwrap()
            .local_reference(),
        Some(Position::new(0))
    );
    assert_eq!(
        master
            .section_tree()
            .get(Position::new(1))
            .unwrap()
            .source(),
        Some(litchi_odm::section::Source::Local(Position::new(0)))
    );
    assert!(
        master
            .section_tree()
            .get(Position::new(2))
            .unwrap()
            .has_dde_source()
    );
    assert_eq!(
        master
            .section_tree()
            .get(Position::new(2))
            .unwrap()
            .source(),
        Some(litchi_odm::section::Source::Dde)
    );
    assert!(
        master
            .security()
            .active_content()
            .iter()
            .any(|item| item.kind() == ActiveKind::Dde)
    );
    let generated = &master.structure().generated_indexes()[0];
    assert_eq!(generated.item(), Position::new(1));
    assert_eq!(generated.name(), Some("Contents"));
    assert_eq!(generated.xml_id(), Some("toc1"));
    let listener = master
        .security()
        .active_content()
        .iter()
        .find(|item| item.kind() == ActiveKind::EventListener)
        .unwrap();
    assert_eq!(listener.trigger(), Some("on-load"));
    assert_eq!(listener.target(), Some("Main.Run"));
    assert_eq!(listener.link(), Some("vnd.sun.star.script:Main.Run"));
    assert_eq!(
        master.security().changed_write_disposition(),
        litchi_odm::security::ChangedWriteDisposition::RequiresInertActiveContentOptIn
    );

    let policy = SecurityPolicy::strict().with_active_content(ActiveContentPolicy::PreserveInert);
    let mut index_edit = master.edit_with_policy(policy);
    index_edit
        .rename_generated_index(Position::new(1), "Renamed Contents")
        .unwrap();
    let index_commit = index_edit.commit().unwrap();
    let renamed = index_commit.snapshot();
    assert_eq!(
        renamed.structure().generated_indexes()[0].name(),
        Some("Renamed Contents")
    );
    assert_eq!(index_commit.patch().changes().generated_indexes().len(), 1);
    assert_eq!(
        index_commit
            .patch()
            .inverse()
            .apply(renamed)
            .unwrap()
            .as_bytes(),
        master.as_bytes()
    );
    let durable = index_commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&master).unwrap().as_bytes(),
        renamed.as_bytes()
    );
    let mut history = master.history(HistoryLimits::new(2, u64::MAX));
    history.record(&index_commit).unwrap();
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), master.as_bytes());
    assert!(history.redo());
    assert_eq!(history.current().as_bytes(), renamed.as_bytes());
    let mut title_edit = master.edit_with_policy(policy);
    title_edit.set_title("Index merge").unwrap();
    let title_commit = title_edit.commit().unwrap();
    let merged = index_commit
        .patch()
        .merge(title_commit.patch())
        .unwrap()
        .apply(&master)
        .unwrap();
    assert_eq!(merged.title(), Some("Index merge"));
    assert_eq!(
        merged.structure().generated_indexes()[0].name(),
        Some("Renamed Contents")
    );
    let mut competing_edit = master.edit_with_policy(policy);
    competing_edit
        .rename_generated_index(Position::new(1), "Competing Contents")
        .unwrap();
    let competing = competing_edit.commit().unwrap();
    let conflict_plan = index_commit
        .patch()
        .plan_three_way(competing.patch())
        .unwrap();
    assert!(conflict_plan.conflicts().conflicts().iter().any(
        |candidate| matches!(candidate, Conflict::GeneratedIndex(item) if *item == Position::new(1))
    ));

    let mut body_edit = master.edit_with_policy(policy);
    body_edit
        .remove_body_item(Position::new(0))
        .unwrap()
        .remove_style(Origin::Styles, "Standard")
        .unwrap();
    let body_commit = body_edit.commit().unwrap();
    let body_changed = body_commit.snapshot();
    assert_eq!(
        body_changed.structure().items()[0],
        Kind::GeneratedIndex(IndexKind::TableOfContents)
    );
    assert!(!body_changed.content_xml().contains(">front</text:p>"));
    assert!(
        !body_changed
            .styles()
            .iter()
            .any(|style| style.name() == "Standard")
    );
    assert_eq!(body_commit.patch().changes().body_items().len(), 1);
    assert_eq!(
        body_commit
            .patch()
            .inverse()
            .apply(body_changed)
            .unwrap()
            .as_bytes(),
        master.as_bytes()
    );
    assert_eq!(
        body_commit
            .patch()
            .durable()
            .unwrap()
            .apply(&master)
            .unwrap()
            .as_bytes(),
        body_changed.as_bytes()
    );
    let mut body_history = master.history(HistoryLimits::new(2, u64::MAX));
    body_history.record(&body_commit).unwrap();
    assert!(body_history.undo());
    assert_eq!(body_history.current().as_bytes(), master.as_bytes());
    assert!(body_history.redo());
    assert_eq!(body_history.current().as_bytes(), body_changed.as_bytes());
    let merged_body = body_commit
        .patch()
        .merge(index_commit.patch())
        .unwrap()
        .apply(&master)
        .unwrap();
    assert_eq!(
        merged_body.structure().generated_indexes()[0].name(),
        Some("Renamed Contents")
    );
    assert_eq!(
        merged_body.structure().items()[0],
        Kind::GeneratedIndex(IndexKind::TableOfContents)
    );
    let mut remove_index_edit = master.edit_with_policy(policy);
    remove_index_edit
        .remove_body_item(Position::new(1))
        .unwrap();
    let remove_index_commit = remove_index_edit.commit().unwrap();
    let index_remove_conflict = remove_index_commit
        .patch()
        .plan_three_way(index_commit.patch())
        .unwrap();
    assert!(index_remove_conflict.conflicts().conflicts().iter().any(
        |candidate| matches!(candidate, Conflict::BodyItem(item) if *item == Position::new(1))
    ));

    let mut container_edit = master.edit_with_policy(policy);
    container_edit
        .remove_body_item(Position::new(6))
        .unwrap()
        .remove_body_item(Position::new(7))
        .unwrap();
    let container_commit = container_edit.commit().unwrap();
    assert!(
        !container_commit
            .snapshot()
            .structure()
            .items()
            .iter()
            .any(|kind| matches!(kind, Kind::List | Kind::Table))
    );
    assert_eq!(
        container_commit
            .patch()
            .inverse()
            .apply(container_commit.snapshot())
            .unwrap()
            .as_bytes(),
        master.as_bytes()
    );

    let no_op = master.edit().commit().unwrap();
    assert_eq!(no_op.snapshot().as_bytes(), master.as_bytes());
    let mut refused = master.edit();
    refused.set_title("refused").unwrap();
    assert!(refused.commit().is_err());
    let mut preserved = master.edit_with_policy(policy);
    preserved.set_title("preserved inertly").unwrap();
    let changed = preserved.commit().unwrap().into_snapshot();
    assert!(changed.content_xml().contains("office:dde-source"));
    assert!(Master::from_bytes(changed.as_bytes().to_vec()).is_ok());
}

#[test]
fn collision_safe_transfer_renames_complete_style_and_resource_closures_atomically() {
    let source_content = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
        r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#,
        r##"<style:style style:name="BaseSection" style:family="section"><style:section-properties fo:background-color="#ffffff"/></style:style>"##,
        r#"<style:style style:name="ImportedSection" style:family="section" style:parent-style-name="BaseSection"><style:section-properties fo:margin-left="1cm"/></style:style>"#,
        r#"</office:automatic-styles><office:body><office:text><text:section text:name="Source" text:style-name="ImportedSection"><text:section-source xlink:href="Chapters/a.odt"/></text:section></office:text></office:body></office:document-content>"#,
    );
    let destination_content = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
        r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>"#,
        r##"<style:style style:name="BaseSection" style:family="section"><style:section-properties fo:background-color="#ff0000"/></style:style>"##,
        r#"<style:style style:name="ImportedSection" style:family="section" style:parent-style-name="BaseSection"><style:section-properties fo:margin-left="2cm"/></style:style>"#,
        r#"</office:automatic-styles><office:body><office:text><text:p>destination</text:p></office:text></office:body></office:document-content>"#,
    );
    let source = master_from_parts(source_content, STYLES, b"source chapter");
    let source_before = source.as_bytes().to_vec();
    let destination = master_from_parts(destination_content, STYLES, b"occupied chapter");
    let mut edit = destination.edit();
    edit.transfer_linked_section(&source, Position::new(0), "Imported", "Chapters/a.odt")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.as_bytes(), source_before);
    let changed = commit.snapshot();
    let imported = changed
        .section_tree()
        .sections()
        .iter()
        .find(|section| section.name() == "Imported")
        .unwrap();
    assert_eq!(imported.style_name(), Some("ImportedSection__import1"));
    let reference = &changed.subdocuments()[imported.reference().unwrap().get()];
    assert_eq!(reference.href(), "Chapters/a__import1.odt");
    assert!(
        changed
            .styles()
            .iter()
            .any(|style| { style.name() == "BaseSection__import1" && style.parent().is_none() })
    );
    assert!(changed.styles().iter().any(|style| {
        style.name() == "ImportedSection__import1" && style.parent() == Some("BaseSection__import1")
    }));
    assert!(changed.content_xml().contains("fo:margin-left=\"1cm\""));
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        destination.as_bytes()
    );

    let mut history = destination.history(HistoryLimits::new(4, u64::MAX));
    history.record(&commit).unwrap();
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), destination.as_bytes());
    assert!(history.redo());
    assert_eq!(history.current().as_bytes(), changed.as_bytes());

    let mut title_edit = destination.edit();
    title_edit.set_title("parallel metadata").unwrap();
    let title_commit = title_edit.commit().unwrap();
    let merged = commit
        .patch()
        .plan_three_way(title_commit.patch())
        .unwrap()
        .commit()
        .unwrap()
        .apply(&destination)
        .unwrap();
    assert_eq!(merged.title(), Some("parallel metadata"));
    assert!(
        merged
            .section_tree()
            .sections()
            .iter()
            .any(|section| section.name() == "Imported")
    );
    assert!(Master::from_bytes(merged.as_bytes().to_vec()).is_ok());
    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&destination).unwrap().as_bytes(),
        changed.as_bytes()
    );
    assert_eq!(
        changed.security().changed_write_disposition(),
        litchi_odm::security::ChangedWriteDisposition::Allowed
    );
}

#[test]
fn styles_owned_transfer_creates_a_compact_styles_part_when_absent() {
    let source_content = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text>"#,
        r#"<text:section text:name="Source" text:style-name="NamedSection">"#,
        r#"<text:section-source xlink:href="Chapters/a.odt"/></text:section>"#,
        r#"</office:text></office:body></office:document-content>"#,
    );
    let source_styles = concat!(
        r#"<?xml version="1.0"?><office:document-styles "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
        r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">"#,
        r#"<office:styles><style:style style:name="NamedBase" style:family="section">"#,
        r##"<style:section-properties fo:background-color="#eeeeee"/></style:style>"##,
        r#"<style:style style:name="NamedSection" style:family="section" style:parent-style-name="NamedBase">"#,
        r#"<style:section-properties fo:margin-right="1cm"/></style:style>"#,
        r#"</office:styles></office:document-styles>"#,
    );
    let destination_content = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
        r#"<office:body><office:text><text:p>destination</text:p></office:text></office:body>"#,
        r#"</office:document-content>"#,
    );
    let source = master_from_parts(source_content, source_styles, b"source chapter");
    let source_before = source.as_bytes().to_vec();
    let destination = master_without_styles(destination_content, b"occupied chapter");
    assert!(destination.styles_xml().is_none());

    let mut edit = destination.edit();
    edit.transfer_linked_section(
        &source,
        Position::new(0),
        "Imported named style",
        "Chapters/imported.odt",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.as_bytes(), source_before);
    let changed = commit.snapshot();
    let styles = changed.styles_xml().unwrap();
    assert!(!styles.contains('\n'));
    assert!(styles.contains("style:name=\"NamedBase\""));
    assert!(styles.contains("style:name=\"NamedSection\""));
    assert!(styles.contains("fo:margin-right=\"1cm\""));
    assert!(
        changed
            .styles()
            .iter()
            .any(|style| { style.name() == "NamedSection" && style.parent() == Some("NamedBase") })
    );
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        destination.as_bytes()
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&destination).unwrap().as_bytes(),
        changed.as_bytes()
    );
    let mut history = destination.history(HistoryLimits::new(3, u64::MAX));
    history.record(&commit).unwrap();
    assert!(history.undo());
    assert!(history.current().styles_xml().is_none());
    assert!(history.redo());
    assert!(history.current().styles_xml().is_some());
    let mut title_edit = destination.edit();
    title_edit.set_title("parallel styles owner").unwrap();
    let title_commit = title_edit.commit().unwrap();
    let merged = commit
        .patch()
        .merge(title_commit.patch())
        .unwrap()
        .apply(&destination)
        .unwrap();
    assert_eq!(merged.title(), Some("parallel styles owner"));
    assert!(merged.styles_xml().is_some());
    assert!(Master::from_bytes(changed.as_bytes().to_vec()).is_ok());
}
