use std::sync::Arc;

use litchi_docx::builder::Document as MutableDocument;
use litchi_docx::revision::conflict::{Id, Kind, Limits, Metadata, Scope, Snapshot};
use litchi_docx::{ConflictKind, Package};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const HEADER_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const OPAQUE_REL: &str = "urn:litchi:test:opaque";
const OPAQUE_BYTES: &[u8] = b"\x00opaque conflict fixture\xff\x10";

fn main_xml() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" xmlns:x="urn:litchi:opaque" mc:Ignorable="x q14" x:root="keep">
  <w:body>
    <!-- preserve-before -->
    <w:p x:sentinel="inline">
      <q14:conflictIns w:id="1" w:author="Alice &amp; Co" w:date="2026-08-08T01:02:03Z"><w:r><w:t>added &amp; kept</w:t></w:r></q14:conflictIns>
      <q14:conflictDel w:id="2" w:author="Bob" w:date="2026-08-08T02:03:04+08:00"><w:r><w:delText>removed</w:delText></w:r></q14:conflictDel>
      <x:conflictIns w:id="900" w:author="opaque"><w:r><w:t>not a W14 conflict</w:t></w:r></x:conflictIns>
    </w:p>
    <w:tbl><w:tr><w:trPr>
      <q14:conflictIns w:id="5" w:author="Property Insert"/>
      <q14:conflictDel w:id="6" w:author="Property Delete"/>
    </w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>
    <w:p x:sentinel="ranges">
      <q14:customXmlConflictInsRangeStart w:id="3" w:author="Carol" w:date="2026-08-08T03:04:05Z"/>
      <w:customXml w:uri="urn:litchi:test" w:element="inserted"><w:r><w:t>range insert payload</w:t></w:r></w:customXml>
      <q14:customXmlConflictInsRangeEnd w:id="3"/>
      <q14:customXmlConflictDelRangeStart w:id="4" w:author="Dana"/>
      <w:customXml w:uri="urn:litchi:test" w:element="deleted"><w:r><w:t>range delete payload</w:t></w:r></w:customXml>
      <q14:customXmlConflictDelRangeEnd w:id="4"/>
    </w:p>
    <mc:AlternateContent>
      <mc:Choice Requires="q14"><w:p><w:r><w:t>active branch</w:t></w:r></w:p></mc:Choice>
      <mc:Fallback><q14:customXmlConflictDelRangeEnd w:id="777"/></mc:Fallback>
    </mc:AlternateContent>
    <!-- preserve-after -->
    <w:sectPr/>
  </w:body>
</w:document>"#
    )
    .into_bytes()
}

fn header_xml() -> Vec<u8> {
    format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:p><w:r><w:t>unrelated header story</w:t></w:r></w:p></w:hdr>"#
    )
    .into_bytes()
}

fn package_with_conflicts() -> Package {
    let mut document = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main_xml(),
    );
    document.rels_mut().add_relationship(
        HEADER_REL.to_owned(),
        "header1.xml".to_owned(),
        "rIdHeader".to_owned(),
        false,
    );
    document.rels_mut().add_relationship(
        OPAQUE_REL.to_owned(),
        "media/opaque.bin".to_owned(),
        "rIdOpaque".to_owned(),
        false,
    );

    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(document));
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/header1.xml").unwrap(),
        ct::WML_HEADER.to_owned(),
        header_xml(),
    )));
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/media/opaque.bin").unwrap(),
        "application/x-litchi-opaque".to_owned(),
        OPAQUE_BYTES.to_vec(),
    )));
    opc.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "word/document.xml".to_owned(),
        "rId1".to_owned(),
        false,
    );
    Package::from_opc_package(opc).unwrap()
}

fn slice(source: &[u8], span: litchi_docx::revision::conflict::Span) -> &[u8] {
    &source[span.start()..span.end()]
}

fn minimal_document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn package_with_header_stories(main: Vec<u8>, headers: Vec<Vec<u8>>) -> Package {
    let mut document = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main,
    );
    let mut opc = OpcPackage::new();
    for (index, header) in headers.into_iter().enumerate() {
        let name = format!("header{}.xml", index + 1);
        document.rels_mut().add_relationship(
            HEADER_REL.to_owned(),
            name.clone(),
            format!("rIdHeader{}", index + 1),
            false,
        );
        opc.add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/word/{name}")).unwrap(),
            ct::WML_HEADER.to_owned(),
            header,
        )));
    }
    opc.add_part(Box::new(document));
    opc.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "word/document.xml".to_owned(),
        "rId1".to_owned(),
        false,
    );
    Package::from_opc_package(opc).unwrap()
}

#[test]
fn hand_authored_package_exposes_typed_main_and_story_inventories() {
    let package = package_with_conflicts();
    let snapshot = package.conflicts().unwrap();
    let inventory = snapshot.inventory();

    assert_eq!(inventory.conflicts.len(), 4);
    assert_eq!(inventory.ranges.len(), 2);
    assert_eq!(
        inventory
            .conflicts
            .iter()
            .map(|value| (value.kind, value.scope, value.metadata.id.get()))
            .collect::<Vec<_>>(),
        vec![
            (Kind::Insert, Scope::Inline, 1),
            (Kind::Delete, Scope::Inline, 2),
            (Kind::Insert, Scope::Property, 5),
            (Kind::Delete, Scope::Property, 6),
        ]
    );
    assert_eq!(inventory.conflicts[0].metadata.author, "Alice & Co");
    assert_eq!(
        inventory.conflicts[0].metadata.date.as_deref(),
        Some("2026-08-08T01:02:03Z")
    );
    assert_eq!(inventory.conflicts[1].metadata.author, "Bob");
    assert_eq!(
        slice(
            snapshot.source(),
            inventory.conflicts[0].text_extent().unwrap()
        ),
        b"added &amp; kept"
    );
    assert_eq!(
        slice(
            snapshot.source(),
            inventory.conflicts[1].text_extent().unwrap()
        ),
        b"removed"
    );
    assert_eq!(
        inventory
            .ranges
            .iter()
            .map(|value| (
                value.kind,
                value.metadata.id.get(),
                value.metadata.author.as_str()
            ))
            .collect::<Vec<_>>(),
        vec![(Kind::Insert, 3, "Carol"), (Kind::Delete, 4, "Dana")]
    );
    assert!(
        snapshot
            .source()
            .windows(b"<x:conflictIns".len())
            .any(|bytes| bytes == b"<x:conflictIns")
    );
    assert!(
        !inventory
            .conflicts
            .iter()
            .any(|value| value.metadata.id.get() == 900)
    );
    assert!(
        !inventory
            .ranges
            .iter()
            .any(|value| value.metadata.id.get() == 777)
    );

    let stories = package.conflict_stories().unwrap();
    assert_eq!(stories.len(), 2);
    assert_eq!(stories[0].part().as_str(), "/word/document.xml");
    assert_eq!(stories[0].snapshot().inventory(), inventory);
    assert_eq!(stories[1].part().as_str(), "/word/header1.xml");
    assert!(stories[1].snapshot().inventory().conflicts.is_empty());
    assert!(stories[1].snapshot().inventory().ranges.is_empty());
}

#[test]
fn explicit_limits_are_exact_and_independent() {
    let source = main_xml();
    let exact = Limits {
        max_source_bytes: source.len(),
        max_conflicts: 4,
        max_ranges: 2,
        ..Limits::default()
    };
    let snapshot = Snapshot::from_xml(source.clone(), exact).unwrap();
    assert_eq!(snapshot.inventory().conflicts.len(), 4);
    assert_eq!(snapshot.inventory().ranges.len(), 2);

    assert!(
        Snapshot::from_xml(
            source.clone(),
            Limits {
                max_source_bytes: source.len() - 1,
                ..exact
            }
        )
        .is_err()
    );
    assert!(
        Snapshot::from_xml(
            source.clone(),
            Limits {
                max_conflicts: 3,
                ..exact
            }
        )
        .is_err()
    );
    assert!(
        Snapshot::from_xml(
            source,
            Limits {
                max_ranges: 1,
                ..exact
            }
        )
        .is_err()
    );
    assert!(
        Limits {
            max_events: usize::MAX,
            ..Limits::default()
        }
        .validate()
        .is_err()
    );
}

#[test]
fn malformed_range_topologies_are_refused_but_inactive_orphans_are_ignored() {
    let orphan = minimal_document(r#"<w:p><q14:customXmlConflictInsRangeEnd w:id="7"/></w:p>"#);
    assert!(Snapshot::from_xml(orphan, Default::default()).is_err());

    let duplicate = minimal_document(
        r#"<w:p><q14:customXmlConflictInsRangeStart w:id="7" w:author="a"/><q14:customXmlConflictInsRangeStart w:id="7" w:author="b"/><q14:customXmlConflictInsRangeEnd w:id="7"/></w:p>"#,
    );
    assert!(Snapshot::from_xml(duplicate, Default::default()).is_err());

    let wrong_kind = minimal_document(
        r#"<w:p><q14:customXmlConflictInsRangeStart w:id="7" w:author="a"/><q14:customXmlConflictDelRangeEnd w:id="7"/></w:p>"#,
    );
    assert!(Snapshot::from_xml(wrong_kind, Default::default()).is_err());

    let inactive = minimal_document(
        r#"<mc:AlternateContent><mc:Choice Requires="q14"><w:p><w:r><w:t>active</w:t></w:r></w:p></mc:Choice><mc:Fallback><q14:customXmlConflictDelRangeEnd w:id="99"/></mc:Fallback></mc:AlternateContent>"#,
    );
    let snapshot = Snapshot::from_xml(inactive, Default::default()).unwrap();
    assert!(snapshot.inventory().ranges.is_empty());
}

fn with_main_story(package: &Package, main: Vec<u8>) -> Package {
    let main_uri = PackURI::new("/word/document.xml").unwrap();
    let mut edit = package.edit_opc();
    edit.package_mut()
        .get_part_mut(&main_uri)
        .unwrap()
        .set_blob(main);
    edit.commit().unwrap().into_package()
}

fn contains(haystack: &[u8], needle: &[u8]) -> bool {
    haystack.windows(needle.len()).any(|bytes| bytes == needle)
}

#[test]
fn source_transaction_is_noop_lossless_reversible_stale_checked_and_retryable() {
    let package = package_with_conflicts();
    let source = package.conflicts().unwrap();
    let noop = package.edit_conflicts().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(
        noop.package().conflicts().unwrap().source(),
        source.source()
    );

    let mut edit = package.edit_conflicts();
    assert!(edit.set_conflict_author(0, "x".repeat(256)).is_err());
    assert!(!edit.clone().commit().unwrap().diagnostics().changed());
    edit.set_conflict_author(0, "Alice Updated").unwrap();
    edit.set_conflict_date(0, Some("2026-08-09T10:11:12Z".to_owned()))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let committed = commit.package().conflicts().unwrap();
    assert_eq!(
        committed.inventory().conflicts[0].metadata.author,
        "Alice Updated"
    );
    assert!(contains(committed.source(), b"<!-- preserve-before -->"));
    assert!(contains(committed.source(), b"<x:conflictIns"));
    // The source package is never changed by an edit.
    assert_eq!(package.conflicts().unwrap().source(), source.source());

    let inverse = commit.patch().inverse();
    let changed = package.apply_conflicts(commit.patch()).unwrap();
    assert_eq!(
        changed.package().conflicts().unwrap().source(),
        committed.source()
    );
    // Replaying onto the changed package fails its preconditions.
    assert!(changed.package().apply_conflicts(commit.patch()).is_err());
    let restored = changed.package().apply_conflicts(&inverse).unwrap();
    assert_eq!(
        restored.package().conflicts().unwrap().source(),
        source.source()
    );

    let mut stale_bytes = source.source().to_vec();
    stale_bytes.extend_from_slice(b" ");
    let stale = with_main_story(&package, stale_bytes);
    let mut retry_edit = package.edit_conflicts();
    retry_edit.set_range_author(0, "Range Updated").unwrap();
    let retry = retry_edit.commit().unwrap();
    assert!(stale.apply_conflicts(retry.patch()).is_err());
    assert_eq!(
        package
            .apply_conflicts(retry.patch())
            .unwrap()
            .package()
            .conflicts()
            .unwrap()
            .inventory()
            .ranges[0]
            .metadata
            .author,
        "Range Updated"
    );

    let mut remove = package.edit_conflicts();
    remove.remove_range(0).unwrap();
    let removed = remove.commit().unwrap();
    let removed_snapshot = removed.package().conflicts().unwrap();
    assert_eq!(removed_snapshot.inventory().ranges.len(), 1);
    assert!(contains(removed_snapshot.source(), b"range insert payload"));
    assert_eq!(
        removed
            .package()
            .apply_conflicts(&removed.patch().inverse())
            .unwrap()
            .package()
            .conflicts()
            .unwrap()
            .source(),
        source.source()
    );
}

#[test]
fn package_publication_preserves_unrelated_parts_relationships_and_noop_arcs() {
    let package = package_with_conflicts();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let opaque_uri = PackURI::new("/word/media/opaque.bin").unwrap();
    let main_before = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .blob_arc();
    let opaque_before = package
        .opc_package()
        .get_part(&opaque_uri)
        .unwrap()
        .blob_arc();

    let noop = package.edit_conflicts().commit().unwrap();
    let after_noop = package
        .apply_conflicts(noop.patch())
        .unwrap()
        .into_package();
    let main_after_noop = after_noop
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .blob_arc();
    assert!(Arc::ptr_eq(&main_before, &main_after_noop));

    let mut edit = package.edit_conflicts();
    edit.set_range_date(1, Some("2026-08-10T00:00:00Z".to_owned()))
        .unwrap();
    let commit = edit.commit().unwrap();
    let package = package
        .apply_conflicts(commit.patch())
        .unwrap()
        .into_package();
    let published = package.conflicts().unwrap();
    assert_eq!(
        published.inventory().ranges[1].metadata.date.as_deref(),
        Some("2026-08-10T00:00:00Z")
    );

    let opaque_after = package
        .opc_package()
        .get_part(&opaque_uri)
        .unwrap()
        .blob_arc();
    assert!(Arc::ptr_eq(&opaque_before, &opaque_after));
    assert_eq!(opaque_after.as_slice(), OPAQUE_BYTES);
    let main = package.opc_package().get_part(&document_uri).unwrap();
    let relationship = main
        .rels()
        .iter()
        .find(|relationship| relationship.r_id() == "rIdOpaque")
        .unwrap();
    assert_eq!(relationship.reltype(), OPAQUE_REL);
    assert_eq!(relationship.target_ref(), "media/opaque.bin");
    assert!(!relationship.is_external());

    assert!(package.apply_conflicts(commit.patch()).is_err());
    assert_eq!(
        package.conflicts().unwrap().source(),
        published.source(),
        "stale publication changed the already-published story"
    );
}

#[test]
fn package_target_limit_lineage_keeps_unrelated_story_ceiling() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="51" w:author="before"><w:r><w:t>small</w:t></w:r></q14:conflictIns></w:p>"#,
    );
    let header_text = "h".repeat(8 * 1024);
    let header = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:p><w:r><w:t>{header_text}</w:t></w:r></w:p></w:hdr>"#
    )
    .into_bytes();
    assert!(header.len() > main.len());

    let mut document = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main.clone(),
    );
    document.rels_mut().add_relationship(
        HEADER_REL.to_owned(),
        "header1.xml".to_owned(),
        "rIdHeader".to_owned(),
        false,
    );
    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(document));
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/header1.xml").unwrap(),
        ct::WML_HEADER.to_owned(),
        header.clone(),
    )));
    opc.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "word/document.xml".to_owned(),
        "rId1".to_owned(),
        false,
    );
    let package = Package::from_opc_package(opc).unwrap();
    let limits = Limits {
        max_source_bytes: header.len(),
        max_output_bytes: header.len(),
        max_total_story_bytes: main.len() + header.len() + 1_024,
        ..Limits::default()
    };

    let source = package
        .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
        .conflicts()
        .unwrap();
    assert_eq!(source.limits(), limits);
    let mut edit = package.edit_conflicts();
    edit.set_conflict_author(0, "after with a longer author")
        .unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();

    let changed = package
        .apply_conflicts(commit.patch())
        .unwrap()
        .into_package();
    assert_eq!(
        changed
            .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
            .conflicts()
            .unwrap()
            .inventory()
            .conflicts[0]
            .metadata
            .author,
        "after with a longer author"
    );
    let restored = changed
        .apply_conflicts(&inverse)
        .unwrap()
        .into_package()
        .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
        .conflicts()
        .unwrap();
    assert_eq!(restored.source(), main);
    assert_eq!(restored.inventory().conflicts[0].metadata.author, "before");
}

#[test]
fn package_story_fanout_stops_at_the_configured_cap() {
    let main = minimal_document("<w:p/>");
    let headers = (0..3).map(|_| header_xml()).collect::<Vec<_>>();
    let package = package_with_header_stories(main, headers);
    let limits = Limits {
        max_stories: 2,
        ..Limits::default()
    };

    assert!(
        package
            .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
            .conflicts()
            .is_err()
    );
}

#[test]
fn package_aggregate_conflict_budget_is_applied_before_next_story_parse() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="61" w:author="main"><w:r><w:t>one</w:t></w:r></q14:conflictIns></w:p>"#,
    );
    let header = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:p><q14:conflictIns w:id="62" w:author="header"><w:r><w:t>two</w:t></w:r></q14:conflictIns></w:p></w:hdr>"#
    )
    .into_bytes();
    let package = package_with_header_stories(main, vec![header]);
    let limits = Limits {
        max_conflicts: 2,
        max_total_conflicts: 1,
        ..Limits::default()
    };

    assert!(
        package
            .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
            .conflict_stories()
            .is_err()
    );
}

#[test]
fn package_exact_aggregate_quotas_allow_a_later_empty_story() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="71" w:author="a"><w:r><w:t>one</w:t></w:r></q14:conflictIns><q14:customXmlConflictInsRangeStart w:id="72" w:author="b"/><w:r><w:t>outside</w:t></w:r><q14:customXmlConflictInsRangeEnd w:id="72"/></w:p>"#,
    );
    let package = package_with_header_stories(main, vec![header_xml()]);
    let limits = Limits {
        max_conflicts: 1,
        max_ranges: 1,
        max_metadata_bytes: 2,
        max_text_segments: 1,
        max_total_conflicts: 1,
        max_total_ranges: 1,
        max_total_metadata_bytes: 2,
        max_total_text_segments: 1,
        ..Limits::default()
    };

    let stories = package
        .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
        .conflict_stories()
        .unwrap();
    assert_eq!(stories.len(), 2);
    assert_eq!(stories[0].snapshot().inventory().conflicts.len(), 1);
    assert_eq!(stories[0].snapshot().inventory().ranges.len(), 1);
    assert!(stories[1].snapshot().inventory().conflicts.is_empty());
    assert!(stories[1].snapshot().inventory().ranges.is_empty());
    assert_eq!(stories[0].snapshot().limits(), limits);
    assert_eq!(stories[1].snapshot().limits(), limits);
}

#[test]
fn package_aggregate_text_segment_budget_precedes_later_story_parse() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="81" w:author="main"><w:r><w:t>one</w:t></w:r></q14:conflictIns></w:p>"#,
    );
    let header = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:p><q14:conflictIns w:id="82" w:author="header"><w:r><w:t>two</w:t></w:r></q14:conflictIns></w:p></w:hdr>"#
    )
    .into_bytes();
    let package = package_with_header_stories(main, vec![header]);
    let limits = Limits {
        max_conflicts: 2,
        max_total_conflicts: 2,
        max_text_segments: 2,
        max_total_text_segments: 1,
        ..Limits::default()
    };

    assert!(
        package
            .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
            .conflict_stories()
            .is_err()
    );
}

#[test]
fn later_story_patch_retains_the_original_package_limits() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="91" w:author="main"><w:r><w:t>one</w:t></w:r></q14:conflictIns></w:p>"#,
    );
    let header = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:p><q14:conflictIns w:id="92" w:author="header"><w:r><w:t>two</w:t></w:r></q14:conflictIns></w:p></w:hdr>"#
    )
    .into_bytes();
    let package = package_with_header_stories(main, vec![header]);
    let limits = Limits {
        max_conflicts: 4,
        max_total_conflicts: 2,
        max_text_segments: 4,
        max_total_text_segments: 2,
        ..Limits::default()
    };
    let stories = package
        .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
        .conflict_stories()
        .unwrap();
    let header = stories[1].snapshot();
    assert_eq!(header.limits(), limits);

    let mut edit = package.edit_story_conflicts(stories[1].part());
    edit.set_conflict_author(0, "updated header").unwrap();
    let commit = edit.commit().unwrap();
    // The published package still satisfies the original package limits.
    let republished = commit
        .package()
        .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
        .conflict_stories()
        .unwrap();
    assert_eq!(republished[1].part(), stories[1].part());
    let published = republished[1].snapshot();
    assert_eq!(published.limits(), limits);
    assert_eq!(
        published.inventory().conflicts[0].metadata.author,
        "updated header"
    );
}

#[test]
fn strict_word_namespace_and_public_writer_round_trip() {
    let strict = format!(
        r#"<w:document xmlns:w="{W_STRICT}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:body><w:p><q14:conflictIns w:id="31" w:author="strict"><w:r><w:t>strict text</w:t></w:r></q14:conflictIns></w:p></w:body></w:document>"#
    );
    let parsed = Snapshot::from_xml(strict.into_bytes(), Default::default()).unwrap();
    assert_eq!(parsed.inventory().conflicts.len(), 1);
    assert_eq!(parsed.inventory().conflicts[0].scope, Scope::Inline);

    let insert = Metadata::new(
        Id::new(40).unwrap(),
        "Writer Insert".to_owned(),
        Some("2026-08-08T04:05:06Z".to_owned()),
    )
    .unwrap();
    let delete = Metadata::new(Id::new(41).unwrap(), "Writer Delete".to_owned(), None).unwrap();
    let range = Metadata::new(Id::new(42).unwrap(), "Writer Range".to_owned(), None).unwrap();
    let mut document = MutableDocument::new();
    let paragraph = document.add_paragraph();
    paragraph
        .add_conflict(ConflictKind::Insert, insert)
        .unwrap()
        .add_run_with_text("writer added");
    paragraph
        .add_conflict(ConflictKind::Delete, delete)
        .unwrap()
        .add_run_with_text("writer removed");
    paragraph
        .add_custom_xml_conflict_range(ConflictKind::Delete, range)
        .unwrap()
        .add_run_with_text("writer range payload");

    let xml = document.to_xml().unwrap();
    assert!(xml.contains("mc:Ignorable=\"w14\""));
    assert!(xml.contains("<w14:conflictIns"));
    assert!(xml.contains("<w14:conflictDel"));
    assert!(xml.contains("<w:delText"));
    assert!(xml.contains(">writer removed</w:delText>"));
    assert!(xml.contains("<w14:customXmlConflictDelRangeStart"));
    assert!(xml.contains("<w14:customXmlConflictDelRangeEnd"));
    assert!(xml.contains("<w:t"));
    assert!(xml.contains(">writer range payload</w:t>"));
    assert!(!xml.contains("<w:delText>writer range payload</w:delText>"));

    let reopened = Snapshot::from_xml(xml.into_bytes(), Default::default()).unwrap();
    assert_eq!(reopened.inventory().conflicts.len(), 2);
    assert_eq!(reopened.inventory().ranges.len(), 1);
    assert_eq!(reopened.inventory().ranges[0].kind, Kind::Delete);
}

#[test]
fn conflict_removal_is_payload_exact_leaf_only_and_metadata_coalesced() {
    let package = package_with_conflicts();
    let source = package.conflicts().unwrap();
    let after_bytes = |commit: &litchi_docx::revision::conflict::Commit| {
        commit.package().conflicts().unwrap().source().to_vec()
    };

    let conflict = &source.inventory().conflicts[0];
    let mut expected_inline = Vec::new();
    expected_inline.extend_from_slice(&source.source()[..conflict.start_tag.start()]);
    expected_inline
        .extend_from_slice(&source.source()[conflict.content.start()..conflict.content.end()]);
    expected_inline.extend_from_slice(&source.source()[conflict.span.end()..]);

    let mut inline = package.edit_conflicts();
    inline.remove_conflict(0).unwrap();
    let inline = inline.commit().unwrap();
    assert_eq!(after_bytes(&inline), expected_inline);
    assert_eq!(
        inline
            .package()
            .conflicts()
            .unwrap()
            .inventory()
            .conflicts
            .len(),
        3
    );
    assert!(
        after_bytes(&inline)
            .windows(b"<w:r><w:t>added &amp; kept</w:t></w:r>".len())
            .any(|window| window == b"<w:r><w:t>added &amp; kept</w:t></w:r>")
    );

    let property = &source.inventory().conflicts[2];
    let mut expected_property = Vec::new();
    expected_property.extend_from_slice(&source.source()[..property.span.start()]);
    expected_property.extend_from_slice(&source.source()[property.span.end()..]);

    let mut leaf = package.edit_conflicts();
    leaf.remove_conflict(2).unwrap();
    let leaf = leaf.commit().unwrap();
    assert_eq!(after_bytes(&leaf), expected_property);
    assert_eq!(
        leaf.package()
            .conflicts()
            .unwrap()
            .inventory()
            .conflicts
            .iter()
            .map(|conflict| conflict.metadata.id.get())
            .collect::<Vec<_>>(),
        vec![1, 2, 6]
    );

    let mut coalesced = package.edit_conflicts();
    coalesced
        .set_conflict_author(0, "temporary".to_owned())
        .unwrap();
    coalesced
        .set_conflict_date(0, Some("2026-08-11T00:00:00Z".to_owned()))
        .unwrap();
    coalesced.remove_conflict(0).unwrap();
    assert!(
        coalesced
            .set_conflict_author(0, "later".to_owned())
            .is_err()
    );
    assert!(coalesced.set_conflict_date(0, None).is_err());
    let coalesced = coalesced.commit().unwrap();
    assert_eq!(after_bytes(&coalesced), after_bytes(&inline));
}

#[test]
fn header_story_patch_publishes_only_to_its_story_and_refuses_stale_topology() {
    let main = minimal_document("<w:p/>");
    let header = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:q14="{W14}" xmlns:mc="{MC}" mc:Ignorable="q14"><w:p><q14:conflictIns w:id="71" w:author="header before"><w:r><w:t>header payload</w:t></w:r></q14:conflictIns></w:p></w:hdr>"#
    )
    .into_bytes();
    let header_uri = PackURI::new("/word/header1.xml").unwrap();
    let header_story = |package: &Package| {
        package
            .conflict_stories()
            .unwrap()
            .into_iter()
            .find(|story| story.part() == &header_uri)
            .unwrap()
            .snapshot()
            .clone()
    };
    let package = package_with_header_stories(main.clone(), vec![header.clone(), header_xml()]);
    let mut transaction = package.edit_story_conflicts(&header_uri);
    transaction
        .set_conflict_author(0, "header after".to_owned())
        .unwrap();
    let commit = transaction.commit().unwrap();
    let published = package
        .apply_conflicts(commit.patch())
        .unwrap()
        .into_package();
    assert_eq!(
        header_story(&published).inventory().conflicts[0]
            .metadata
            .author,
        "header after"
    );
    assert_eq!(published.conflicts().unwrap().source(), main.as_slice());

    let mut stale_package = package_with_header_stories(main, vec![header, header_xml()]);
    let mut stale_transaction = stale_package.edit_story_conflicts(&header_uri);
    stale_transaction
        .set_conflict_author(0, "must not publish".to_owned())
        .unwrap();
    let stale_commit = stale_transaction.commit().unwrap();
    let main_uri = PackURI::new("/word/document.xml").unwrap();
    stale_package = {
        let mut edit = stale_package.edit_opc();
        edit.package_mut()
            .get_part_mut(&main_uri)
            .unwrap()
            .rels_mut()
            .remove("rIdHeader2");
        edit.commit().unwrap().into_package()
    };
    assert!(stale_package.apply_conflicts(stale_commit.patch()).is_err());
    // The refused replay left the (now orphan-bearing) graph untouched.
    let header_part = stale_package
        .opc_package()
        .get_part(&header_uri)
        .unwrap()
        .blob()
        .to_vec();
    assert!(
        String::from_utf8(header_part)
            .unwrap()
            .contains("header before")
    );
}

#[test]
fn conflict_publication_preserves_document_edit_composition() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="81" w:author="before"><w:r><w:t>payload</w:t></w:r></q14:conflictIns></w:p>"#,
    );
    let package = package_with_header_stories(main, Vec::new());
    let mut transaction = package.edit_conflicts();
    transaction
        .set_conflict_author(0, "after".to_owned())
        .unwrap();
    let commit = transaction.commit().unwrap();

    let package = package
        .apply_conflicts(commit.patch())
        .unwrap()
        .into_package();

    assert!(package.edit().is_ok());
}

#[test]
fn main_conflicts_ignore_a_malformed_unrelated_header_and_aggregate_quota() {
    let main = minimal_document(
        r#"<w:p><q14:conflictIns w:id="82" w:author="main"><w:r><w:t>payload</w:t></w:r></q14:conflictIns></w:p>"#,
    );
    let main_len = main.len();
    let package = package_with_header_stories(main, vec![b"<w:hdr".to_vec()]);
    let limits = Limits {
        max_source_bytes: main_len,
        max_total_story_bytes: main_len,
        ..Limits::default()
    };

    assert_eq!(
        package
            .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
            .conflicts()
            .unwrap()
            .inventory()
            .conflicts
            .len(),
        1
    );
    assert!(package.conflict_stories().is_err());
    assert!(
        package
            .with_semantic_limits(litchi_docx::SemanticLimits::default().conflicts(limits))
            .conflict_stories()
            .is_err()
    );
}

/// Record 0355a: the edit's output budget is exact through `EditOptions`;
/// one byte below refuses, and the refused edit can be retried with a
/// smaller change.
#[test]
fn edit_output_limit_is_exact_one_below_retryable_and_reversible() {
    use litchi_docx::revision::conflict::EditOptions;
    let changed_author = "output-limit author";
    let package = package_with_conflicts();
    let main = PackURI::new("/word/document.xml").unwrap();
    let story = |package: &Package| {
        package
            .opc_package()
            .get_part(&main)
            .unwrap()
            .blob()
            .to_vec()
    };
    let probe = {
        let mut edit = package.edit_conflicts();
        edit.set_conflict_author(0, changed_author).unwrap();
        edit.commit().unwrap()
    };
    let exact_len = story(probe.package()).len();
    let options = |maximum| {
        EditOptions::default().limits(Limits {
            max_output_bytes: maximum,
            ..Limits::default()
        })
    };

    let mut exact = package.edit_conflicts_with_options(options(exact_len));
    exact.set_conflict_author(0, changed_author).unwrap();
    let exact = exact.commit().unwrap();
    assert_eq!(story(exact.package()), story(probe.package()));
    let restored = exact
        .package()
        .apply_conflicts(&exact.patch().inverse())
        .unwrap();
    assert_eq!(story(restored.package()), story(&package));

    let mut below = package.edit_conflicts_with_options(options(exact_len - 1));
    below.set_conflict_author(0, changed_author).unwrap();
    assert!(below.clone().commit().is_err());
    below.set_conflict_author(0, "A").unwrap();
    let retried = below.commit().unwrap();
    assert!(story(retried.package()).len() < exact_len);
    let restored = retried
        .package()
        .apply_conflicts(&retried.patch().inverse())
        .unwrap();
    assert_eq!(story(restored.package()), story(&package));
}
