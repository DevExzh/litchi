#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{
    Anchor, BlackWhiteMode, Limits, Payload, Relationship, Snapshot, Target, TargetMode,
    apply_commit, apply_patch, load_slide, load_snapshot,
};
use crate::Error;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

#[test]
fn loads_opaque_payload_and_relationship_metadata_losslessly() {
    let payload = b"<?xml version=\"1.0\"?><opaque:root xmlns:opaque=\"urn:opaque\"><opaque:value>raw &amp; bytes</opaque:value></opaque:root>";
    let (package, slide_name) = package_with_internal_payload(payload);
    let slide = package.get_part(&slide_name).expect("slide");
    let mut limits = Limits::default();
    let parts = load_slide(&package, 3, slide, &mut limits).expect("content part");

    assert_eq!(parts.len(), 1);
    let content_part = &parts[0];
    assert_eq!(content_part.slide_index(), 3);
    assert_eq!(content_part.index(), 0);
    assert_eq!(content_part.relationship_id(), "rIdOpaque");
    assert_eq!(
        content_part.anchor().xml(),
        b"<p:contentPart r:id=\"rIdOpaque\"><p14:nvContentPartPr/></p:contentPart>"
    );
    assert_eq!(content_part.relationship().id(), "rIdOpaque");
    assert_eq!(
        content_part.relationship().relationship_type(),
        rt::CUSTOM_XML
    );
    assert_eq!(
        content_part.relationship().target_ref(),
        "../custom/opaque.xml"
    );

    let payload_view = content_part.payload().expect("internal payload");
    assert_eq!(payload_view.part_name().as_str(), "/ppt/custom/opaque.xml");
    assert_eq!(payload_view.content_type(), "application/xml");
    assert_eq!(payload_view.bytes(), payload);
    assert_eq!(payload_view.relationships().len(), 1);
    assert_eq!(payload_view.relationships()[0].id(), "rIdPayloadLink");
    assert_eq!(
        payload_view.relationships()[0].target_ref(),
        "https://example.invalid/opaque"
    );
}

#[test]
fn discovery_is_a_source_preserving_noop() {
    let (package, slide_name) = package_with_internal_payload(b"opaque");
    let before = package
        .get_part(&slide_name)
        .expect("slide")
        .blob()
        .to_vec();
    let slide = package.get_part(&slide_name).expect("slide");
    let _ = load_slide(&package, 0, slide, &mut Limits::default()).expect("content part");
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before);
}

#[test]
fn retains_external_target_mode_without_following_it() {
    let (package, slide_name) = package_with_external_payload();
    let slide = package.get_part(&slide_name).expect("slide");
    let parts = load_slide(&package, 0, slide, &mut Limits::default()).expect("content part");
    let relationship = parts[0].relationship();

    assert_eq!(relationship.target_mode(), TargetMode::External);
    assert_eq!(relationship.target_ref(), "https://example.invalid/content");
    assert!(relationship.payload().is_none());
    assert_eq!(
        relationship.target().external_ref(),
        Some("https://example.invalid/content")
    );
}

#[test]
fn rejects_missing_anchor_relationship() {
    let package = package_with_anchor(b"<p:contentPart r:id=\"rIdMissing\"/>", None, None);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Relationship(message) if message.contains("rIdMissing")));
}

#[test]
fn rejects_anchor_without_relationship_id() {
    let package = package_with_anchor(b"<p:contentPart/>", None, None);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Invalid(message) if message.contains("missing r:id")));
}

#[test]
fn enforces_content_part_count_limit() {
    let package = package_with_anchor(
        b"<p:contentPart r:id=\"rId1\"/><p:contentPart r:id=\"rId2\"/>",
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rId1", false)),
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rId2", false)),
    );
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let mut limits = Limits::new(1, 1024, 4096, 8).unwrap();
    let error = load_slide(&package, 0, slide, &mut limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: "content-part count",
            ..
        }
    ));
}

#[test]
fn exact_noop_snapshot_commit_preserves_the_complete_source_graph() {
    let mut package = package_with_internal_payload(b"opaque").0;
    let source = snapshot(&package);
    let slide_before = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let payload_before = package
        .get_part(&PackURI::new("/ppt/custom/opaque.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let commit = source.edit().commit().unwrap();

    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    apply_commit(&mut package, commit).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .blob(),
        slide_before.as_slice()
    );
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/custom/opaque.xml").unwrap())
            .unwrap()
            .blob(),
        payload_before.as_slice()
    );
}

#[test]
fn relationship_and_payload_edits_preserve_unknown_anchor_markup() {
    let mut package = package_with_internal_payload(b"opaque").0;
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_relationship_type(0, "urn:vendor:content-part")
        .unwrap();
    edit.replace_payload(0, "application/vendor+xml", b"edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    let source_xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(
        source_xml.contains("<p14:nvContentPartPr/>") || source_xml.contains("p14:nvContentPartPr")
    );

    apply_patch(&mut package, commit.patch()).unwrap();
    let current = snapshot(&package);
    assert_eq!(
        current.parts()[0].relationship().relationship_type(),
        "urn:vendor:content-part"
    );
    assert_eq!(
        current.parts()[0].payload().unwrap().content_type(),
        "application/vendor+xml"
    );
    assert_eq!(current.parts()[0].payload().unwrap().bytes(), b"edited");
    assert!(
        current
            .source_xml()
            .windows(b"p14:nvContentPartPr".len())
            .any(|window| { window == b"p14:nvContentPartPr" })
    );
}

#[test]
fn relationship_id_edit_rewrites_only_the_anchor_attribute() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_relationship_id(0, "rIdRenamed").unwrap();
    let commit = edit.commit().unwrap();
    let source_xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(source_xml.contains(r#"r:id="rIdRenamed""#));
    assert!(!source_xml.contains(r#"r:id="rIdOpaque""#));
    assert!(source_xml.contains("p14:nvContentPartPr"));

    apply_patch(&mut package, commit.patch()).unwrap();
    let current = snapshot(&package);
    assert_eq!(current.parts()[0].relationship_id(), "rIdRenamed");
    let slide = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap();
    assert!(slide.rels().get("rIdOpaque").is_none());
    assert!(slide.rels().get("rIdRenamed").is_some());

    apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(snapshot(&package).source_xml(), source.source_xml());
}

#[test]
fn add_remove_and_inverse_manage_relationships_and_payload_parts() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let payload_name = PackURI::new("/ppt/custom/added.xml").unwrap();
    let payload = Payload::new(
        payload_name.clone(),
        "application/custom+xml",
        b"added".to_vec(),
    );
    let anchor = Anchor::new(
        "rIdAdded",
        br#"<p:contentPart r:id="rIdAdded"><p14:nvContentPartPr/></p:contentPart>"#,
    );
    let relationship = Relationship::new(
        "rIdAdded",
        litchi_opc::constants::relationship_type::CUSTOM_XML,
        "../custom/added.xml",
        TargetMode::Internal,
        Target::internal(payload),
    );
    let mut edit = source.edit();
    edit.push(anchor, relationship).unwrap();
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(snapshot(&package).parts().len(), 2);
    assert_eq!(package.get_part(&payload_name).unwrap().blob(), b"added");

    let inverse = commit.patch().inverse();
    apply_patch(&mut package, &inverse).unwrap();
    let restored = snapshot(&package);
    assert_eq!(restored.source_xml(), source.source_xml());
    assert_eq!(restored.parts(), source.parts());
    assert!(package.get_part(&payload_name).is_err());
}

#[test]
fn removal_collects_orphan_payload_and_inverse_restores_it() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.remove(0).unwrap();
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    assert!(snapshot(&package).parts().is_empty());
    assert!(
        package
            .get_part(&PackURI::new("/ppt/custom/opaque.xml").unwrap())
            .is_err()
    );

    apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(snapshot(&package).source_xml(), source.source_xml());
}

#[test]
fn stale_source_rejection_is_atomic() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_relationship_type(0, "urn:changed").unwrap();
    let patch = edit.commit().unwrap().into_patch();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(b"stale".to_vec());
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), b"stale");
}

#[test]
fn invalid_edits_do_not_mutate_the_staged_snapshot() {
    let (package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    assert!(edit.set_relationship_id(0, "bad\u{0000}").is_err());
    assert_eq!(edit.parts(), source.parts());
}

#[test]
fn reads_optional_p14_bw_mode_and_round_trips_all_schema_tokens() {
    let modes = [
        ("clr", BlackWhiteMode::Color),
        ("auto", BlackWhiteMode::Auto),
        ("gray", BlackWhiteMode::Gray),
        ("ltGray", BlackWhiteMode::LightGray),
        ("invGray", BlackWhiteMode::InverseGray),
        ("grayWhite", BlackWhiteMode::GrayWhite),
        ("blackGray", BlackWhiteMode::BlackGray),
        ("blackWhite", BlackWhiteMode::BlackWhite),
        ("black", BlackWhiteMode::Black),
        ("white", BlackWhiteMode::White),
        ("hidden", BlackWhiteMode::Hidden),
    ];
    for (token, mode) in modes {
        assert_eq!(mode.as_str(), token);
        assert_eq!(BlackWhiteMode::try_from(token).unwrap(), mode);
    }
    assert!(BlackWhiteMode::try_from("AUTO").is_err());

    let (package, slide_name) = package_with_internal_payload(b"opaque");
    let slide = package.get_part(&slide_name).unwrap();
    let parts = load_slide(&package, 0, slide, &mut Limits::default()).unwrap();
    assert_eq!(parts[0].black_white_mode(), None);

    let package = package_with_anchor(
        br#"<p:contentPart p14:bwMode="auto" r:id="rIdOpaque"><p14:nvContentPartPr/></p:contentPart>"#,
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
        None,
    );
    let slide = package.get_part(&slide_name).unwrap();
    let parts = load_slide(&package, 0, slide, &mut Limits::default()).unwrap();
    assert_eq!(parts[0].black_white_mode(), Some(BlackWhiteMode::Auto));
    assert_eq!(
        parts[0].anchor().black_white_mode(),
        Some(BlackWhiteMode::Auto)
    );
}

#[test]
fn p14_bw_mode_commits_and_reopens_every_schema_token() {
    let modes = [
        ("clr", BlackWhiteMode::Color),
        ("auto", BlackWhiteMode::Auto),
        ("gray", BlackWhiteMode::Gray),
        ("ltGray", BlackWhiteMode::LightGray),
        ("invGray", BlackWhiteMode::InverseGray),
        ("grayWhite", BlackWhiteMode::GrayWhite),
        ("blackGray", BlackWhiteMode::BlackGray),
        ("blackWhite", BlackWhiteMode::BlackWhite),
        ("black", BlackWhiteMode::Black),
        ("white", BlackWhiteMode::White),
        ("hidden", BlackWhiteMode::Hidden),
    ];

    for (token, mode) in modes {
        let mut package = package_with_anchor(
            br#"<p:contentPart r:id="rIdOpaque"/>"#,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            None,
        );
        let source = snapshot(&package);
        let mut edit = source.edit();
        assert!(edit.set_black_white_mode(0, Some(mode)).unwrap());
        let commit = edit.commit().unwrap();
        let written = commit.snapshot().source_xml();
        assert!(
            written
                .windows(format!(r#"p14:bwMode="{token}""#).len())
                .any(|window| window == format!(r#"p14:bwMode="{token}""#).as_bytes())
        );

        apply_patch(&mut package, commit.patch()).unwrap();
        let reopened = snapshot(&package);
        assert_eq!(reopened.parts()[0].black_white_mode(), Some(mode));
        assert!(
            reopened
                .source_xml()
                .windows(format!(r#"p14:bwMode="{token}""#).len())
                .any(|window| window == format!(r#"p14:bwMode="{token}""#).as_bytes())
        );
    }
}

#[test]
fn p14_bw_mode_edits_preserve_opaque_anchor_bytes_and_inverse() {
    let anchor = br#"<p:contentPart xmlns:vendor="urn:vendor" vendor:flag="keep" p14:bwMode="gray" r:id="rIdOpaque"><p14:nvContentPartPr/><vendor:opaque/></p:contentPart>"#;
    let (mut package, _) = (
        package_with_anchor(
            anchor,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            None,
        ),
        PackURI::new("/ppt/slides/slide1.xml").unwrap(),
    );
    let source = snapshot(&package);
    assert_eq!(
        source.parts()[0].black_white_mode(),
        Some(BlackWhiteMode::Gray)
    );

    let mut noop = source.edit();
    assert!(
        !noop
            .set_black_white_mode(0, Some(BlackWhiteMode::Gray))
            .unwrap()
    );
    assert!(!noop.commit().unwrap().is_changed());

    let mut edit = source.edit();
    assert!(
        edit.set_black_white_mode(0, Some(BlackWhiteMode::BlackWhite))
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    let changed_xml = commit.snapshot().source_xml();
    assert!(
        changed_xml
            .windows(b"p14:bwMode=\"blackWhite\"".len())
            .any(|window| { window == b"p14:bwMode=\"blackWhite\"" })
    );
    assert!(
        changed_xml
            .windows(b"vendor:flag=\"keep\"".len())
            .any(|window| window == b"vendor:flag=\"keep\"")
    );
    assert!(
        changed_xml
            .windows(b"vendor:opaque".len())
            .any(|window| window == b"vendor:opaque")
    );

    apply_patch(&mut package, commit.patch()).unwrap();
    let changed = snapshot(&package);
    assert_eq!(
        changed.parts()[0].black_white_mode(),
        Some(BlackWhiteMode::BlackWhite)
    );
    apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(snapshot(&package).source_xml(), source.source_xml());
}

#[test]
fn p14_bw_mode_insert_remove_handles_inherited_and_local_namespaces() {
    let (mut package, _) = (
        package_with_anchor(
            br#"<p:contentPart r:id="rIdOpaque"/>"#,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            None,
        ),
        PackURI::new("/ppt/slides/slide1.xml").unwrap(),
    );
    let source = snapshot(&package);
    let mut edit = source.edit();
    assert!(
        edit.set_black_white_mode(0, Some(BlackWhiteMode::Auto))
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    let xml = commit.snapshot().source_xml();
    assert!(
        xml.windows(b"p14:bwMode=\"auto\"".len())
            .any(|window| window == b"p14:bwMode=\"auto\"")
    );
    assert!(
        xml.windows(format!("xmlns:p14=\"{P14}\"").len())
            .any(|window| window == format!("xmlns:p14=\"{P14}\"").as_bytes())
    );
    apply_patch(&mut package, commit.patch()).unwrap();

    let current = snapshot(&package);
    let mut remove = current.edit();
    assert!(remove.set_bw_mode(0, None).unwrap());
    let removed = remove.commit().unwrap();
    let removed_xml = removed.snapshot().source_xml();
    assert!(
        !removed_xml
            .windows(b"p14:bwMode".len())
            .any(|window| window == b"p14:bwMode")
    );
    assert!(
        removed_xml
            .windows(format!("xmlns:p14=\"{P14}\"").len())
            .any(|window| window == format!("xmlns:p14=\"{P14}\"").as_bytes())
    );
    apply_patch(&mut package, removed.patch()).unwrap();
    assert_eq!(snapshot(&package).parts()[0].black_white_mode(), None);

    let mut exact = source.edit();
    exact
        .set_black_white_mode(0, Some(BlackWhiteMode::Auto))
        .unwrap();
    exact.set_black_white_mode(0, None).unwrap();
    let exact = exact.commit().unwrap();
    assert!(!exact.is_changed());
    assert_eq!(exact.snapshot().source_xml(), source.source_xml());
}

#[test]
fn p14_bw_mode_accepts_inherited_custom_prefix_and_rejects_bad_values() {
    let anchor = br#"<p:contentPart x:bwMode="gray" r:id="rIdOpaque"/>"#;
    let (mut package, _) = (
        package_with_anchor(
            anchor,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            None,
        ),
        PackURI::new("/ppt/slides/slide1.xml").unwrap(),
    );
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).unwrap();
    let mut limits = Limits::default();
    let parts = load_slide(&package, 0, slide, &mut limits).unwrap();
    assert_eq!(parts[0].black_white_mode(), None);

    let slide = package.get_part_mut(&slide_name).unwrap();
    let anchor_text = std::str::from_utf8(anchor).unwrap();
    let xml = format!(
        "<p:sld xmlns:p=\"{PML}\" xmlns:r=\"{REL}\" xmlns:x=\"{P14}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{anchor_text}</p:spTree></p:cSld></p:sld>"
    );
    slide.set_blob(xml.into_bytes());
    let source = snapshot(&package);
    assert_eq!(
        source.parts()[0].black_white_mode(),
        Some(BlackWhiteMode::Gray)
    );
    let mut edit = source.edit();
    edit.set_black_white_mode(0, Some(BlackWhiteMode::White))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit
            .snapshot()
            .source_xml()
            .windows(b"x:bwMode=\"white\"".len())
            .any(|window| window == b"x:bwMode=\"white\"")
    );

    for value in ["AUTO", "unknown"] {
        let package = package_with_anchor(
            format!(r#"<p:contentPart p14:bwMode="{value}" r:id="rIdOpaque"/>"#).as_bytes(),
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            None,
        );
        let slide = package.get_part(&slide_name).unwrap();
        let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
        assert!(matches!(error, Error::Invalid(message) if message.contains("bwMode")));
    }

    let package = package_with_anchor(
        br#"<p:contentPart xmlns:x="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:bwMode="gray" x:bwMode="white" r:id="rIdOpaque"/>"#,
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
        None,
    );
    let slide = package.get_part(&slide_name).unwrap();
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Invalid(message) if message.contains("duplicate")));
}

#[test]
fn p14_bw_mode_ignores_unqualified_foreign_and_media_lookalikes() {
    for anchor in [
        br#"<p:contentPart bwMode="gray" r:id="rIdOpaque"/>"#.as_slice(),
        br#"<p:contentPart xmlns:x="urn:foreign" x:bwMode="gray" r:id="rIdOpaque"/>"#.as_slice(),
        br#"<p:contentPart r:id="rIdOpaque"><p14:media p14:bwMode="gray"/></p:contentPart>"#
            .as_slice(),
    ] {
        let package = package_with_anchor(
            anchor,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            None,
        );
        let current = snapshot(&package);
        assert_eq!(current.parts()[0].black_white_mode(), None);
    }
}

#[test]
fn p14_bw_mode_edits_only_the_active_mce_branch() {
    let anchor = br#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:Choice Requires="x"><p:contentPart p14:bwMode="black" r:id="rIdInactive"/></mc:Choice><mc:Fallback><p:contentPart p14:bwMode="gray" r:id="rIdOpaque"><p14:nvContentPartPr/></p:contentPart></mc:Fallback></mc:AlternateContent>"#;
    let (mut package, _) = (
        package_with_anchor(
            anchor,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false)),
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdInactive", false)),
        ),
        PackURI::new("/ppt/slides/slide1.xml").unwrap(),
    );
    let source = snapshot(&package);
    assert_eq!(source.parts().len(), 1);
    assert_eq!(source.parts()[0].relationship_id(), "rIdOpaque");
    assert_eq!(
        source.parts()[0].black_white_mode(),
        Some(BlackWhiteMode::Gray)
    );

    let mut edit = source.edit();
    edit.set_black_white_mode(0, Some(BlackWhiteMode::White))
        .unwrap();
    let commit = edit.commit().unwrap();
    let xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(xml.contains(r#"r:id="rIdInactive""#));
    assert!(xml.contains(r#"p14:bwMode="black""#));
    assert!(xml.contains(r#"r:id="rIdOpaque""#));
    assert!(xml.contains(r#"p14:bwMode="white""#));
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(
        snapshot(&package).parts()[0].black_white_mode(),
        Some(BlackWhiteMode::White)
    );
}

#[test]
fn p14_bw_mode_selects_a_canonical_mce_choice() {
    let anchor = br#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14"><p:contentPart p14:bwMode="black" r:id="rIdChoice"/></mc:Choice><mc:Fallback><p:contentPart p14:bwMode="gray" r:id="rIdFallback"/></mc:Fallback></mc:AlternateContent>"#;
    let (mut package, _) = (
        package_with_anchor(
            anchor,
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdChoice", false)),
            Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdFallback", false)),
        ),
        PackURI::new("/ppt/slides/slide1.xml").unwrap(),
    );
    let source = snapshot(&package);
    assert_eq!(source.parts().len(), 1);
    assert_eq!(source.parts()[0].relationship_id(), "rIdChoice");
    assert_eq!(
        source.parts()[0].black_white_mode(),
        Some(BlackWhiteMode::Black)
    );

    let mut edit = source.edit();
    edit.set_black_white_mode(0, Some(BlackWhiteMode::White))
        .unwrap();
    let commit = edit.commit().unwrap();
    let xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(xml.contains(r#"r:id="rIdChoice""#));
    assert!(xml.contains(r#"p14:bwMode="white""#));
    assert!(xml.contains(r#"r:id="rIdFallback""#));
    assert!(xml.contains(r#"p14:bwMode="gray""#));

    apply_patch(&mut package, commit.patch()).unwrap();
    let reopened = snapshot(&package);
    assert_eq!(reopened.parts()[0].relationship_id(), "rIdChoice");
    assert_eq!(
        reopened.parts()[0].black_white_mode(),
        Some(BlackWhiteMode::White)
    );
}

#[test]
fn signed_content_part_noop_is_preserved_and_mutation_requires_explicit_unsign() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
            .to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rIdSignature".to_owned(),
        false,
    );
    assert!(package.is_signed());

    let source = snapshot(&package);
    let noop = source.edit().commit().unwrap();
    assert!(!noop.is_changed());
    apply_commit(&mut package, noop).unwrap();
    assert!(package.is_signed());

    let mut edit = source.edit();
    edit.set_black_white_mode(0, Some(BlackWhiteMode::White))
        .unwrap();
    let commit = edit.commit().unwrap();
    let before = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let error = apply_patch(&mut package, commit.patch()).unwrap_err();
    assert!(matches!(
        error,
        Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(package.is_signed());
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .blob(),
        before.as_slice()
    );

    package.unsign();
    assert!(!package.is_signed());
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(
        snapshot(&package).parts()[0].black_white_mode(),
        Some(BlackWhiteMode::White)
    );
}

fn package_with_internal_payload(payload: &[u8]) -> (OpcPackage, PackURI) {
    let relationship = Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false));
    let mut package = package_with_anchor(
        b"<p:contentPart r:id=\"rIdOpaque\"><p14:nvContentPartPr/></p:contentPart>",
        relationship,
        None,
    );
    let payload_name = PackURI::new("/ppt/custom/opaque.xml").unwrap();
    package
        .get_part_mut(&payload_name)
        .expect("payload")
        .set_blob(payload.to_vec());
    (package, PackURI::new("/ppt/slides/slide1.xml").unwrap())
}

fn package_with_external_payload() -> (OpcPackage, PackURI) {
    let package = package_with_anchor(
        b"<p:contentPart r:id=\"rIdExternal\"/>",
        Some((
            rt::CUSTOM_XML,
            "https://example.invalid/content",
            "rIdExternal",
            true,
        )),
        None,
    );
    (package, PackURI::new("/ppt/slides/slide1.xml").unwrap())
}

fn package_with_anchor(
    anchors: &[u8],
    relationship: Option<(&str, &str, &str, bool)>,
    relationship_2: Option<(&str, &str, &str, bool)>,
) -> OpcPackage {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let xml = format!(
        "<p:sld xmlns:p=\"{PML}\" xmlns:r=\"{REL}\" xmlns:p14=\"{P14}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>",
        String::from_utf8(anchors.to_vec()).unwrap()
    );
    let mut slide = BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), xml.into_bytes());
    if let Some((relationship_type, target, id, external)) = relationship {
        slide.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
    }
    if let Some((relationship_type, target, id, external)) = relationship_2 {
        slide.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
    }
    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    let payload_name = PackURI::new("/ppt/custom/opaque.xml").unwrap();
    let mut payload = BlobPart::new(
        payload_name,
        "application/xml".to_owned(),
        b"opaque payload".to_vec(),
    );
    payload.rels_mut().add_relationship(
        rt::HYPERLINK.to_owned(),
        "https://example.invalid/opaque".to_owned(),
        "rIdPayloadLink".to_owned(),
        true,
    );
    package.add_part(Box::new(payload));
    package
}

fn snapshot(package: &OpcPackage) -> Snapshot {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    load_snapshot(package, 0, slide, &mut Limits::default()).expect("snapshot")
}
