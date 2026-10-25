use super::package::{put_presence, put_threading, remove_presence, remove_threading};
use super::{
    AuthorPresenceSnapshot, CommentThreadingSnapshot, ParentComment, PresenceInfo, ThreadingInfo,
    load_presence, load_presence_snapshot, load_threading,
};
use crate::Error;
use crate::comments::{Author, Comment, Comments, Conformance, List, store_presentation_comments};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn author_source() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cmAuthor id="7" name="A" initials="A" lastIdx="2" clrIdx="2"><p:extLst><!-- before --><p:ext uri="{uri}"><?before?><p15:presenceInfo userId="Ada &amp; One" providerId="AD"/><?after?></p:ext><!-- after --></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
    )
    .into_bytes()
}

#[test]
fn replay_presence_validator_applies_recipient_extension_limits() {
    let mut limits = crate::comments::Limits::DEFAULT;
    limits.max_string_bytes = 3;
    assert!(super::validate_presence_replay(&author_source(), limits).is_err());
}

#[test]
fn replay_presence_validator_enforces_active_aggregate_at_exact_boundary() {
    let source = author_source();
    let budget = super::codec::selected_string_budget(
        &source,
        super::codec::Kind::Presence,
        crate::comments::Limits::DEFAULT,
    )
    .unwrap();
    let mut exact = crate::comments::Limits::DEFAULT;
    exact.max_string_bytes = budget;
    assert!(super::validate_presence_replay(&source, exact).is_ok());

    let mut one_under = exact;
    one_under.max_string_bytes = budget.saturating_sub(1);
    assert!(super::validate_presence_replay(&source, one_under).is_err());

    let active_name = String::from_utf8(source)
        .unwrap()
        .replace(r#"name="A""#, r#"name="AA""#);
    assert!(super::validate_presence_replay(active_name.as_bytes(), exact).is_err());
}

#[test]
fn replay_presence_validator_accepts_partial_multi_author_extensions() {
    let mut source = author_source();
    let closing = b"</p:cmAuthorLst>";
    let position = source
        .windows(closing.len())
        .position(|window| window == closing)
        .unwrap();
    source.splice(
        position..position,
        br#"<p:cmAuthor id="8" name="B" initials="B" lastIdx="0" clrIdx="0"/>"#
            .iter()
            .copied(),
    );
    assert!(super::validate_presence_replay(&source, crate::comments::Limits::DEFAULT,).is_ok());
}

#[test]
fn replay_presence_validator_ignores_inactive_typed_uri_bytes() {
    // Keep both discarded values above the selected budget.  Their only
    // reason for being accepted is that MCE removes their branch before the
    // recipient-owned semantic aggregate is charged.
    let inactive_user = "inactive-user".repeat(32);
    let inactive_provider = "inactive-provider".repeat(32);
    let source = format!(
        r#"<p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"><p:extLst><mc:AlternateContent><mc:Choice Requires="p15"><p:ext uri="urn:other-extension"/></mc:Choice><mc:Fallback><p:ext uri="{uri}"><p15:presenceInfo userId="{inactive_user}" providerId="{inactive_provider}"><p15:unexpected/></p15:presenceInfo></p:ext></mc:Fallback></mc:AlternateContent></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
        inactive_user = inactive_user,
        inactive_provider = inactive_provider,
    );
    // Independent fixture calculation: root xmlns:p/p15/mc (7+58, 9+56,
    // 8+59), cmAuthor's five attributes (2+1, 4+1, 8+1, 7+1, 6+1), and
    // the selected p:ext URI (3+19) total 251 decoded name/value bytes.
    // Do not derive this boundary from the production preflight helper.
    let selected_aggregate = (7 + PML.len())
        + (9 + P15.len())
        + (8 + MCE.len())
        + (2 + 1)
        + (4 + 1)
        + (8 + 1)
        + (7 + 1)
        + (6 + 1)
        + (3 + "urn:other-extension".len());
    assert_eq!(selected_aggregate, 251);

    let mut exact = crate::comments::Limits::DEFAULT;
    exact.max_string_bytes = selected_aggregate;
    assert!(super::validate_presence_replay(source.as_bytes(), exact).is_ok());

    let mut one_under = exact;
    one_under.max_string_bytes = selected_aggregate - 1;
    assert!(super::validate_presence_replay(source.as_bytes(), one_under).is_err());
}

#[test]
fn replay_presence_validator_keeps_physical_attribute_cap_independent() {
    let inactive_user = "x".repeat(129);
    let source = format!(
        r#"<p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"><p:extLst><mc:AlternateContent><mc:Choice Requires="p15"><p:ext uri="urn:other-extension"/></mc:Choice><mc:Fallback><p:ext uri="{uri}"><p15:presenceInfo userId="{inactive_user}" providerId="provider"/></p:ext></mc:Fallback></mc:AlternateContent></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
        inactive_user = inactive_user,
    );
    let mut limits = crate::comments::Limits::DEFAULT;
    limits.max_string_bytes = 128;
    limits.max_attribute_bytes = 128;
    let error = super::validate_presence_replay(source.as_bytes(), limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlAttributeBytes),
            observed,
            maximum: 128,
        } if observed > 128
    ));
}

#[test]
fn replay_threading_validator_checks_typed_extension_children() {
    let source = String::from_utf8(comment_source())
        .unwrap()
        .replace("p15:parentCm", "p15:unexpected");
    assert!(
        super::validate_threading_replay(source.as_bytes(), crate::comments::Limits::DEFAULT,)
            .is_err()
    );
}

#[test]
fn replay_threading_validator_enforces_active_text_at_exact_boundary() {
    let source = comment_source();
    let budget = super::codec::selected_string_budget(
        &source,
        super::codec::Kind::Threading,
        crate::comments::Limits::DEFAULT,
    )
    .unwrap();
    let mut exact = crate::comments::Limits::DEFAULT;
    exact.max_string_bytes = budget;
    assert!(super::validate_threading_replay(&source, exact).is_ok());

    let mut one_under = exact;
    one_under.max_string_bytes = budget.saturating_sub(1);
    assert!(super::validate_threading_replay(&source, one_under).is_err());

    let active_text = String::from_utf8(source)
        .unwrap()
        .replace(">hello</p:text>", ">helloo</p:text>");
    assert!(super::validate_threading_replay(active_text.as_bytes(), exact).is_err());
}

#[test]
fn replay_threading_validator_applies_string_budget_to_date_time() {
    let source = comment_source_with_date_time("2026-10-25T00:00:00Z");
    let budget = super::codec::selected_string_budget(
        source.as_bytes(),
        super::codec::Kind::Threading,
        crate::comments::Limits::DEFAULT,
    )
    .unwrap();
    let mut exact = crate::comments::Limits::DEFAULT;
    exact.max_string_bytes = budget;
    assert!(super::validate_threading_replay(source.as_bytes(), exact).is_ok());

    let over = comment_source_with_date_time("2026-10-25T00:00:00.0Z");
    assert!(super::validate_threading_replay(over.as_bytes(), exact).is_err());
}

fn comment_source() -> Vec<u8> {
    format!(
        r#"<p:cmLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cm authorId="7" idx="2"><p:pos x="1" y="2"/><p:text>hello</p:text><p:extLst><!-- before --><p:ext uri="{uri}"><p15:threadingInfo timeZoneBias="-60"><p15:parentCm authorId="7" idx="1"/></p15:threadingInfo></p:ext><!-- after --></p:extLst></p:cm></p:cmLst>"#,
        uri = super::THREADING_EXTENSION_URI,
    )
    .into_bytes()
}

fn comment_source_with_date_time(date_time: &str) -> String {
    format!(
        r#"<p:cmLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cm authorId="7" dt="{date_time}" idx="2"><p:pos x="1" y="2"/><p:text>hello</p:text><p:extLst><p:ext uri="{uri}"><p15:threadingInfo timeZoneBias="-60"><p15:parentCm authorId="7" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm></p:cmLst>"#,
        uri = super::THREADING_EXTENSION_URI,
    )
}

#[test]
fn replay_threading_validator_ignores_inactive_ordinary_values_after_mce() {
    let inactive = "x".repeat(1024);
    let source = format!(
        r#"<p:cmLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><mc:AlternateContent><mc:Choice Requires="p15"><p:cm authorId="7" idx="2"><p:pos x="1" y="2"/><p:text>hello</p:text><p:extLst><p:ext uri="{uri}"><p15:threadingInfo timeZoneBias="-60"><p15:parentCm authorId="7" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm></mc:Choice><mc:Fallback><p:cm authorId="7" dt="{inactive}" idx="0"><p:pos x="1" y="2"/><p:text>{inactive}</p:text></p:cm></mc:Fallback></mc:AlternateContent></p:cmLst>"#,
        uri = super::THREADING_EXTENSION_URI,
        inactive = inactive,
    );
    // Independent fixture calculation: root xmlns:p/p15/mc (7+58, 9+56,
    // 8+59), the active comment/position attributes (8+1, 3+1, 1+1, 1+1),
    // text "hello" (5), the threading URI (3+38), timeZoneBias (12+3), and
    // parentCm's two attributes (8+1, 3+1) total 288 bytes.  The inactive
    // fallback has much larger dt/text values but is excluded by MCE.
    // Do not derive this boundary from the production preflight helper.
    let selected_aggregate = (7 + PML.len())
        + (9 + P15.len())
        + (8 + MCE.len())
        + (8 + 1)
        + (3 + 1)
        + (1 + 1)
        + (1 + 1)
        + 5
        + (3 + super::THREADING_EXTENSION_URI.len())
        + (12 + 3)
        + (8 + 1)
        + (3 + 1);
    assert_eq!(selected_aggregate, 288);

    let mut exact = crate::comments::Limits::DEFAULT;
    exact.max_string_bytes = selected_aggregate;
    assert!(super::validate_threading_replay(source.as_bytes(), exact).is_ok());

    let mut one_under = exact;
    one_under.max_string_bytes = selected_aggregate - 1;
    assert!(super::validate_threading_replay(source.as_bytes(), one_under).is_err());
}

fn package_with_sources() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "ppt/presentation.xml".into(),
        "rId1".into(),
        false,
    );
    let mut presentation = BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml".into(),
        format!(r#"<p:presentation xmlns:p="{PML}"/>"#).into_bytes(),
    );
    presentation.rels_mut().add_relationship(
        crate::comments::SLIDE_REL.into(),
        "slides/slide1.xml".into(),
        "rIdSlide".into(),
        false,
    );
    package.add_part(Box::new(presentation));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/slides/slide1.xml").unwrap(),
        crate::comments::SLIDE_CONTENT_TYPE.into(),
        format!(r#"<p:sld xmlns:p="{PML}"/>"#).into_bytes(),
    )));
    let value = Comments {
        author_relationship_id: "rIdAuthors".into(),
        author_part_name: "/ppt/commentAuthors.xml".into(),
        authors: vec![Author {
            id: 7,
            name: "A".into(),
            initials: "A".into(),
            last_index: 2,
            color_index: 2,
        }],
        slides: vec![List {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rIdComments".into(),
            part_name: "/ppt/comments/comment1.xml".into(),
            comments: vec![Comment {
                author_id: 7,
                date_time: None,
                index: 2,
                x: 1,
                y: 2,
                text: "hello".into(),
            }],
        }],
    };
    store_presentation_comments(&mut package, &value, Conformance::Transitional).unwrap();
    package
        .get_part_mut(&PackURI::new("/ppt/commentAuthors.xml").unwrap())
        .unwrap()
        .set_blob(author_source());
    package
        .get_part_mut(&PackURI::new("/ppt/comments/comment1.xml").unwrap())
        .unwrap()
        .set_blob(comment_source());
    package
}

#[test]
fn presence_reads_required_fields_and_preserves_source_on_edit_and_inverse() {
    let source = author_source();
    let snapshot =
        AuthorPresenceSnapshot::from_xml(&source, 7, crate::comments::Limits::DEFAULT).unwrap();
    assert_eq!(
        snapshot.value(),
        Some(&PresenceInfo::new("Ada & One", "AD"))
    );
    let mut edit = snapshot.edit();
    edit.set_presence(PresenceInfo::new("Ada <Two>", "None"))
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml();
    assert!(
        changed
            .windows(b"<!-- before -->".len())
            .any(|w| w == b"<!-- before -->")
    );
    assert!(
        changed
            .windows(b"<?before?>".len())
            .any(|w| w == b"<?before?>")
    );
    assert!(
        changed
            .windows(b"Ada &lt;Two&gt;".len())
            .any(|w| w == b"Ada &lt;Two&gt;")
    );
    let mut reverted = changed.to_vec();
    commit.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn threading_reads_parent_and_updates_optional_fields_without_normalizing_source() {
    let source = comment_source();
    let snapshot =
        CommentThreadingSnapshot::from_xml(&source, 7, 2, crate::comments::Limits::DEFAULT)
            .unwrap();
    assert_eq!(
        snapshot.value(),
        Some(&ThreadingInfo::new(
            Some(-60),
            Some(ParentComment::new(Some(7), Some(1))),
        ))
    );
    let mut edit = snapshot.edit();
    edit.set_threading(ThreadingInfo::new(
        None,
        Some(ParentComment::new(Some(8), None)),
    ))
    .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml();
    assert!(
        !changed
            .windows(b"timeZoneBias".len())
            .any(|w| w == b"timeZoneBias")
    );
    assert!(
        changed
            .windows(b"authorId=\"8\"".len())
            .any(|w| w == b"authorId=\"8\"")
    );
    assert!(
        changed
            .windows(b"<!-- after -->".len())
            .any(|w| w == b"<!-- after -->")
    );
    let mut reverted = changed.to_vec();
    commit.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn threading_self_closing_joint_update_coalesces_attribute_and_child_edits() {
    let source = format!(
        r#"<p:cmLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cm authorId="7" idx="2"><p:pos x="1" y="2"/><p:text>hello</p:text><p:extLst><!-- before --><?before?><p:ext uri="{uri}"><?inside?><p15:threadingInfo timeZoneBias="-60"/><?after?></p:ext><!-- after --></p:extLst></p:cm></p:cmLst>"#,
        uri = super::THREADING_EXTENSION_URI,
    )
    .into_bytes();
    let snapshot =
        CommentThreadingSnapshot::from_xml(&source, 7, 2, crate::comments::Limits::DEFAULT)
            .unwrap();
    assert_eq!(snapshot.value(), Some(&ThreadingInfo::new(Some(-60), None)));

    let mut edit = snapshot.edit();
    edit.set_threading(ThreadingInfo::new(
        Some(-30),
        Some(ParentComment::new(Some(8), Some(1))),
    ))
    .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml();
    assert!(
        changed
            .windows(b"timeZoneBias=\"-30\"".len())
            .any(|window| { window == b"timeZoneBias=\"-30\"" })
    );
    assert!(
        changed
            .windows(b"<p15:parentCm authorId=\"8\" idx=\"1\"/>".len())
            .any(|window| window == b"<p15:parentCm authorId=\"8\" idx=\"1\"/>")
    );
    for marker in [
        b"<!-- before -->".as_slice(),
        b"<?before?>",
        b"<?inside?>",
        b"<?after?>",
        b"<!-- after -->",
    ] {
        assert!(changed.windows(marker.len()).any(|window| window == marker));
    }

    let mut reverted = changed.to_vec();
    commit.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn threading_self_closing_joint_update_reuses_payload_namespace_alias() {
    let source = format!(
        r#"<p:cmLst xmlns:p="{PML}" xmlns:x="{P15}" xmlns:mc="{MCE}"><p:cm authorId="7" idx="2"><p:pos x="1" y="2"/><p:text>hello</p:text><p:extLst><?before?><p:ext uri="{uri}"><x:threadingInfo timeZoneBias="-60"/></p:ext><?after?></p:extLst></p:cm></p:cmLst>"#,
        uri = super::THREADING_EXTENSION_URI,
    )
    .into_bytes();
    let snapshot =
        CommentThreadingSnapshot::from_xml(&source, 7, 2, crate::comments::Limits::DEFAULT)
            .unwrap();
    let mut edit = snapshot.edit();
    edit.set_threading(ThreadingInfo::new(
        Some(-30),
        Some(ParentComment::new(Some(8), Some(1))),
    ))
    .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml();
    assert!(
        changed
            .windows(b"<x:parentCm authorId=\"8\" idx=\"1\"/>".len())
            .any(|window| window == b"<x:parentCm authorId=\"8\" idx=\"1\"/>")
    );
    assert!(
        !changed
            .windows(b"p15:parentCm".len())
            .any(|window| window == b"p15:parentCm")
    );
    assert_eq!(
        CommentThreadingSnapshot::from_xml(changed, 7, 2, crate::comments::Limits::DEFAULT)
            .unwrap()
            .value(),
        Some(&ThreadingInfo::new(
            Some(-30),
            Some(ParentComment::new(Some(8), Some(1))),
        ))
    );
    let mut reverted = commit.snapshot().source_xml().to_vec();
    commit.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn threading_new_parent_uses_a_bound_generated_namespace() {
    let source = format!(
        r#"<p:cmLst xmlns:p="{PML}" xmlns:x="{P15}"><p:cm authorId="7" idx="2"><p:pos x="1" y="2"/><p:text>hello</p:text><p:extLst/></p:cm></p:cmLst>"#
    )
    .into_bytes();
    let snapshot =
        CommentThreadingSnapshot::from_xml(&source, 7, 2, crate::comments::Limits::DEFAULT)
            .unwrap();
    assert_eq!(snapshot.value(), None);
    let mut edit = snapshot.edit();
    edit.set_threading(ThreadingInfo::new(
        Some(-30),
        Some(ParentComment::new(Some(8), Some(1))),
    ))
    .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml();
    assert!(
        changed
            .windows(b"<p15:threadingInfo xmlns:p15=\"".len())
            .any(|window| window == b"<p15:threadingInfo xmlns:p15=\"")
    );
    assert!(
        changed
            .windows(b"<p15:parentCm authorId=\"8\" idx=\"1\"/>".len())
            .any(|window| window == b"<p15:parentCm authorId=\"8\" idx=\"1\"/>")
    );
    assert_eq!(
        CommentThreadingSnapshot::from_xml(changed, 7, 2, crate::comments::Limits::DEFAULT)
            .unwrap()
            .value(),
        Some(&ThreadingInfo::new(
            Some(-30),
            Some(ParentComment::new(Some(8), Some(1))),
        ))
    );
    let mut reverted = commit.snapshot().source_xml().to_vec();
    commit.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn exact_noops_keep_source_and_insert_remove_round_trip() {
    let source = author_source();
    let snapshot =
        AuthorPresenceSnapshot::from_xml(&source, 7, crate::comments::Limits::DEFAULT).unwrap();
    let mut noop = snapshot.edit();
    noop.set_presence(PresenceInfo::new("Ada & One", "AD"))
        .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.snapshot().source_xml(), source.as_slice());

    let absent = format!(
        r#"<p:cmAuthorLst xmlns:p="{PML}"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"/></p:cmAuthorLst>"#
    );
    let absent =
        AuthorPresenceSnapshot::from_xml(absent.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .unwrap();
    let mut add = absent.edit();
    add.set_presence(PresenceInfo::new("u", "p")).unwrap();
    let added = add.commit().unwrap();
    assert!(
        added
            .snapshot()
            .source_xml()
            .windows(b"presenceInfo".len())
            .any(|w| w == b"presenceInfo")
    );
    let mut remove = added.snapshot().edit();
    remove.remove().unwrap();
    let removed = remove.commit().unwrap();
    assert_eq!(removed.snapshot().source_xml(), absent.source_xml());
}

#[test]
fn required_fields_duplicate_payloads_and_wrong_namespaces_fail_closed() {
    let missing = format!(
        r#"<p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"><p:extLst><p:ext uri="{uri}"><p15:presenceInfo providerId="AD"/></p:ext></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
    );
    assert!(
        AuthorPresenceSnapshot::from_xml(missing.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .is_err()
    );
    let duplicate_source = author_source();
    let duplicate = String::from_utf8(duplicate_source).unwrap();
    assert_eq!(duplicate.matches("<p15:presenceInfo").count(), 1);
    let duplicate = duplicate.replace(
        "</p:extLst>",
        &format!(
            "<p:ext uri=\"{}\"><p15:presenceInfo userId=\"two\" providerId=\"AD\"/></p:ext></p:extLst>",
            super::PRESENCE_EXTENSION_URI
        ),
    );
    assert!(
        AuthorPresenceSnapshot::from_xml(duplicate.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .is_err()
    );
    let wrong = author_source().to_vec();
    let wrong = String::from_utf8(wrong)
        .unwrap()
        .replace("providerId=\"AD\"", "p:providerId=\"AD\"");
    assert!(
        AuthorPresenceSnapshot::from_xml(wrong.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .is_err()
    );
    let text = String::from_utf8(author_source()).unwrap().replace(
        "<p15:presenceInfo userId=\"Ada &amp; One\" providerId=\"AD\"/>",
        "<p15:presenceInfo userId=\"Ada &amp; One\" providerId=\"AD\">text</p15:presenceInfo>",
    );
    assert!(
        AuthorPresenceSnapshot::from_xml(text.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .is_err()
    );
}

#[test]
fn active_foreign_payload_children_are_rejected_but_inactive_mce_children_are_ignored() {
    let author = String::from_utf8(author_source())
        .unwrap()
        .replace(
            &format!(r#"xmlns:mc="{MCE}">"#),
            &format!(r#"xmlns:mc="{MCE}" xmlns:f="urn:foreign">"#),
        )
        .replace(
            r#"<p15:presenceInfo userId="Ada &amp; One" providerId="AD"/>"#,
            r#"<p15:presenceInfo userId="Ada &amp; One" providerId="AD"><f:bad/></p15:presenceInfo>"#,
        );
    assert!(
        AuthorPresenceSnapshot::from_xml(author.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .is_err()
    );

    let comment = String::from_utf8(comment_source())
        .unwrap()
        .replace(
            &format!(r#"xmlns:mc="{MCE}">"#),
            &format!(r#"xmlns:mc="{MCE}" xmlns:f="urn:foreign">"#),
        )
        .replace(
            r#"<p15:threadingInfo timeZoneBias="-60"><p15:parentCm authorId="7" idx="1"/></p15:threadingInfo>"#,
            r#"<p15:threadingInfo timeZoneBias="-60"><f:bad/><p15:parentCm authorId="7" idx="1"/></p15:threadingInfo>"#,
    );
    assert!(
        CommentThreadingSnapshot::from_xml(
            comment.as_bytes(),
            7,
            2,
            crate::comments::Limits::DEFAULT
        )
        .is_err()
    );

    let parent_child = String::from_utf8(comment_source())
        .unwrap()
        .replace(
            &format!(r#"xmlns:mc="{MCE}">"#),
            &format!(r#"xmlns:mc="{MCE}" xmlns:f="urn:foreign">"#),
        )
        .replace(
            r#"<p15:parentCm authorId="7" idx="1"/>"#,
            r#"<p15:parentCm authorId="7" idx="1"><f:bad/></p15:parentCm>"#,
        );
    assert!(
        CommentThreadingSnapshot::from_xml(
            parent_child.as_bytes(),
            7,
            2,
            crate::comments::Limits::DEFAULT
        )
        .is_err()
    );

    let inactive = format!(
        r#"<p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}" xmlns:f="urn:foreign"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"><p:extLst><mc:AlternateContent><mc:Choice Requires="p15"><p:ext uri="{uri}"><p15:presenceInfo userId="choice" providerId="provider"/></p:ext></mc:Choice><mc:Fallback><p:ext uri="{uri}"><p15:presenceInfo userId="fallback" providerId="provider"><f:bad/></p15:presenceInfo></p:ext></mc:Fallback></mc:AlternateContent></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
    );
    let snapshot =
        AuthorPresenceSnapshot::from_xml(inactive.as_bytes(), 7, crate::comments::Limits::DEFAULT)
            .unwrap();
    assert_eq!(snapshot.value().unwrap().user_id(), "choice");
}

#[test]
fn mce_choice_is_typed_while_fallback_bytes_remain_opaque() {
    let source = format!(
        r#"<?root?><p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"><p:extLst><?before?><mc:AlternateContent><mc:Choice Requires="p15"><?choice?><p:ext uri="{uri}"><p15:presenceInfo userId="choice" providerId="provider"/></p:ext></mc:Choice><mc:Fallback><?fallback?><p:ext uri="{uri}"><p15:presenceInfo userId="fallback" providerId="provider"/></p:ext></mc:Fallback></mc:AlternateContent><?after?></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
    );
    let source = source.into_bytes();
    let snapshot =
        AuthorPresenceSnapshot::from_xml(&source, 7, crate::comments::Limits::DEFAULT).unwrap();
    assert_eq!(snapshot.value().unwrap().user_id(), "choice");
    assert!(super::validate_presence_replay(&source, crate::comments::Limits::DEFAULT,).is_ok());
    let mut edit = snapshot.edit();
    edit.set_presence(PresenceInfo::new("updated", "provider"))
        .unwrap();
    let changed = edit.commit().unwrap();
    let bytes = changed.snapshot().source_xml();
    assert!(
        bytes
            .windows(b"userId=\"updated\"".len())
            .any(|w| w == b"userId=\"updated\"")
    );
    assert!(
        bytes
            .windows(b"userId=\"fallback\"".len())
            .any(|w| w == b"userId=\"fallback\"")
    );
    for marker in [
        b"<?root?>".as_slice(),
        b"<?before?>",
        b"<?choice?>",
        b"<?fallback?>",
        b"<?after?>",
    ] {
        assert!(bytes.windows(marker.len()).any(|window| window == marker));
    }
    let mut reverted = bytes.to_vec();
    changed.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn supported_mce_choice_with_another_extension_leaves_fallback_presence_opaque() {
    let source = format!(
        r#"<p:cmAuthorLst xmlns:p="{PML}" xmlns:p15="{P15}" xmlns:mc="{MCE}"><p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"><p:extLst><mc:AlternateContent><mc:Choice Requires="p15"><p:ext uri="urn:other-extension"/></mc:Choice><mc:Fallback><p:ext uri="{uri}"><p15:presenceInfo userId="fallback" providerId="provider"/></p:ext></mc:Fallback></mc:AlternateContent></p:extLst></p:cmAuthor></p:cmAuthorLst>"#,
        uri = super::PRESENCE_EXTENSION_URI,
    )
    .into_bytes();
    let snapshot =
        AuthorPresenceSnapshot::from_xml(&source, 7, crate::comments::Limits::DEFAULT).unwrap();
    assert_eq!(snapshot.value(), None);

    let mut edit = snapshot.edit();
    edit.set_presence(PresenceInfo::new("active", "provider"))
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml();
    assert!(
        changed
            .windows(b"urn:other-extension".len())
            .any(|window| window == b"urn:other-extension")
    );
    assert!(
        changed
            .windows(b"userId=\"fallback\"".len())
            .any(|window| window == b"userId=\"fallback\"")
    );
    assert!(
        changed
            .windows(b"userId=\"active\"".len())
            .any(|window| window == b"userId=\"active\"")
    );

    let mut reverted = changed.to_vec();
    commit.patch().inverse().apply(&mut reverted).unwrap();
    assert_eq!(reverted, source);
}

#[test]
fn presence_values_reject_xml_10_forbidden_noncharacters_before_output() {
    let snapshot =
        AuthorPresenceSnapshot::from_xml(author_source(), 7, crate::comments::Limits::DEFAULT)
            .unwrap();
    for forbidden in ['\u{FFFE}', '\u{FFFF}'] {
        let mut edit = snapshot.edit();
        let error = edit
            .set_presence(PresenceInfo::new(format!("bad{forbidden}"), "AD"))
            .unwrap_err();
        assert!(
            matches!(error, Error::Invalid(message) if message.contains("XML 1.0-forbidden character"))
        );
        assert!(!edit.diagnostics().changed());
    }
}

#[test]
fn oversized_source_preserving_edit_is_rejected_before_output_allocation() {
    let source = author_source();
    let closing = b"</p:cmAuthorLst>";
    let insertion = source
        .windows(closing.len())
        .position(|window| window == closing)
        .unwrap();
    let overhead = b"<!---->".len() + 4096;
    let filler_len = crate::comments::MAX_PART_BYTES
        .saturating_sub(source.len())
        .saturating_sub(overhead);
    let mut padded = Vec::with_capacity(source.len() + filler_len + 7);
    padded.extend_from_slice(&source[..insertion]);
    padded.extend_from_slice(b"<!--");
    padded.extend(std::iter::repeat_n(b'x', filler_len));
    padded.extend_from_slice(b"-->");
    padded.extend_from_slice(&source[insertion..]);
    assert!(padded.len() <= crate::comments::MAX_PART_BYTES);

    let snapshot =
        AuthorPresenceSnapshot::from_xml(&padded, 7, crate::comments::Limits::DEFAULT).unwrap();
    let mut edit = snapshot.edit();
    edit.set_presence(PresenceInfo::new(
        "x".repeat(crate::comments::MAX_STRING_BYTES),
        "AD",
    ))
    .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Limit { .. })));
}

#[test]
fn escaped_replacement_size_is_rejected_before_serialization() {
    let source = author_source();
    let closing = b"</p:cmAuthorLst>";
    let insertion = source
        .windows(closing.len())
        .position(|window| window == closing)
        .unwrap();
    let filler_len = crate::comments::MAX_PART_BYTES
        .saturating_sub(source.len())
        .saturating_sub(4096);
    let mut padded = Vec::with_capacity(source.len() + filler_len + 7);
    padded.extend_from_slice(&source[..insertion]);
    padded.extend_from_slice(b"<!--");
    padded.extend(std::iter::repeat_n(b'x', filler_len));
    padded.extend_from_slice(b"-->");
    padded.extend_from_slice(&source[insertion..]);
    let snapshot =
        AuthorPresenceSnapshot::from_xml(&padded, 7, crate::comments::Limits::DEFAULT).unwrap();
    let mut edit = snapshot.edit();
    edit.set_presence(PresenceInfo::new(
        "&".repeat(crate::comments::MAX_STRING_BYTES),
        "AD",
    ))
    .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Limit { .. })));
}

#[test]
fn package_graph_loads_and_publishes_both_extension_owners_atomically() {
    let mut package = package_with_sources();
    assert_eq!(
        load_presence(&package, 7).unwrap(),
        Some(PresenceInfo::new("Ada & One", "AD"))
    );
    assert_eq!(
        load_threading(&package, "/ppt/slides/slide1.xml", 7, 2)
            .unwrap()
            .unwrap()
            .time_zone_bias(),
        Some(-60)
    );
    put_presence(
        &mut package,
        7,
        PresenceInfo::new("new-user", "new-provider"),
    )
    .unwrap();
    put_threading(
        &mut package,
        "/ppt/slides/slide1.xml",
        7,
        2,
        ThreadingInfo::new(Some(0), None),
    )
    .unwrap();
    assert_eq!(
        load_presence(&package, 7).unwrap(),
        Some(PresenceInfo::new("new-user", "new-provider"))
    );
    assert_eq!(
        load_threading(&package, "/ppt/slides/slide1.xml", 7, 2)
            .unwrap()
            .unwrap()
            .time_zone_bias(),
        Some(0)
    );
    remove_presence(&mut package, 7).unwrap();
    remove_threading(&mut package, "/ppt/slides/slide1.xml", 7, 2).unwrap();
    let author_after = package
        .get_part(&PackURI::new("/ppt/commentAuthors.xml").unwrap())
        .unwrap()
        .blob();
    assert!(
        author_after
            .windows(b"<!-- before -->".len())
            .any(|w| w == b"<!-- before -->")
    );
    assert!(
        author_after
            .windows(b"<!-- after -->".len())
            .any(|w| w == b"<!-- after -->")
    );
    assert!(
        !author_after
            .windows(b"presenceInfo".len())
            .any(|w| w == b"presenceInfo")
    );
    let comment_after = package
        .get_part(&PackURI::new("/ppt/comments/comment1.xml").unwrap())
        .unwrap()
        .blob();
    assert!(
        comment_after
            .windows(b"<!-- before -->".len())
            .any(|w| w == b"<!-- before -->")
    );
    assert!(
        comment_after
            .windows(b"<!-- after -->".len())
            .any(|w| w == b"<!-- after -->")
    );
    assert!(
        !comment_after
            .windows(b"threadingInfo".len())
            .any(|w| w == b"threadingInfo")
    );
}

#[test]
fn package_graph_rejects_query_and_fragment_on_recognized_relationships() {
    let mut package = package_with_sources();
    let slide = package
        .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap();
    slide.rels_mut().remove("rIdComments");
    slide.rels_mut().add_relationship(
        crate::comments::COMMENTS_REL.into(),
        "../comments/comment1.xml?mode=typed#owner".into(),
        "rIdComments".into(),
        false,
    );
    assert!(load_presence_snapshot(&package, 7).is_err());
}
