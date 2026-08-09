#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI, XmlPart};

use super::model::escaped_len;
use super::package::process_owner_ooxml;
use super::*;

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

#[test]
fn strict_round_trip_preserves_inert_values_and_extensions() {
    let xml = br#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:x" x:origin="fixture"><p:tag name="PATH" val="C:\Docs\file"/><p:tag name="XML" val="&lt;root command=&quot;none&quot;/&gt;"></p:tag></p:tagLst>"#;
    let value = parse(xml).unwrap();
    assert_eq!(
        value.get("xml").unwrap().value(),
        "<root command=\"none\"/>"
    );
    assert_eq!(value.namespaces()[0].qualified_name(), "xmlns:x");
    assert_eq!(value.attrs()[0].qualified_name(), "x:origin");

    let strict = value.xml(Conformance::Strict).unwrap();
    assert!(std::str::from_utf8(&strict).unwrap().contains(STRICT_TEXT));
    assert_eq!(parse(&strict).unwrap(), value);
}

#[test]
fn mce_fallback_is_selected() {
    let xml = br#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:tag/></mc:Choice><mc:Fallback><p:tag name="fallback" val="1"/></mc:Fallback></mc:AlternateContent></p:tagLst>"#;
    assert_eq!(parse(xml).unwrap().get("FALLBACK").unwrap().value(), "1");
}

#[test]
fn mce_p14_choice_matches_powerpoint_capabilities() {
    let xml = format!(
        r#"<p:tagLst xmlns:p="{PML_TEXT}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="{P14}"><mc:AlternateContent><mc:Choice Requires="p14"><p:tag name="choice" val="selected"/></mc:Choice><mc:Fallback><p:tag name="fallback" val="wrong"/></mc:Fallback></mc:AlternateContent></p:tagLst>"#
    );
    let parsed = parse(xml.as_bytes()).unwrap();
    assert_eq!(parsed.get("choice").unwrap().value(), "selected");
    assert!(parsed.get("fallback").is_err());
}

#[test]
fn unicode_caseless_crud_is_checked_and_move_first() {
    let mut list = List::new();
    list.add(Tag::new("Straße", "one").unwrap()).unwrap();
    list.add(Tag::new("Owner", "Alice").unwrap()).unwrap();
    assert_eq!(list.get("STRASSE").unwrap().value(), "one");
    assert_eq!(list.get(1_usize).unwrap().name(), "Owner");
    assert!(matches!(
        list.get(2_usize),
        Err(Error::IndexOutOfBounds { index: 2, len: 2 })
    ));
    assert!(matches!(
        list.add(Tag::new("strasse", "duplicate").unwrap()),
        Err(Error::DuplicateName { matches: 1, .. })
    ));

    let old = list
        .replace("OWNER", Tag::new("Reviewer", "Bob").unwrap())
        .unwrap();
    assert_eq!(old.value(), "Alice");
    assert_eq!(list.set("reviewer", "Carol").unwrap(), "Bob");
    list.insert(1, Tag::new("Status", "Draft").unwrap())
        .unwrap();
    list.reorder(&["status", "strasse", "reviewer"]).unwrap();
    assert_eq!(list.tags()[0].name(), "Status");
    list.reorder(&[2_usize, 1, 0]).unwrap();
    assert_eq!(list.tags()[0].name(), "Reviewer");
    assert_eq!(list.remove("STATUS").unwrap().value(), "Draft");
}

#[test]
fn malformed_duplicate_names_have_typed_ambiguity_and_numeric_repair() {
    let xml = format!(
        r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="Straße" val="A"/><p:tag name="STRASSE" val="B"/></p:tagLst>"#
    );
    let mut list = parse(xml.as_bytes()).unwrap();
    assert!(matches!(
        list.get("strasse"),
        Err(Error::AmbiguousName { matches: 2, .. })
    ));
    assert!(matches!(
        write(&list, Conformance::Transitional),
        Err(Error::DuplicateName { matches: 1, .. })
    ));
    assert_eq!(list.remove(1_usize).unwrap().value(), "B");
    assert_eq!(list.get("STRASSE").unwrap().value(), "A");
    assert!(write(&list, Conformance::Transitional).is_ok());
}

#[test]
fn reorder_rejects_partial_and_duplicate_orders_without_mutating() {
    let mut list = List::new();
    list.add(Tag::new("one", "1").unwrap()).unwrap();
    list.add(Tag::new("two", "2").unwrap()).unwrap();
    let original = list.clone();
    assert!(matches!(
        list.reorder(&["one"]),
        Err(Error::OrderLength {
            expected: 2,
            actual: 1
        })
    ));
    assert_eq!(list, original);
    assert!(matches!(
        list.reorder(&["one", "ONE"]),
        Err(Error::DuplicateSelection { index: 0 })
    ));
    assert_eq!(list, original);
}

#[test]
fn malformed_markup_and_resource_limits_are_rejected() {
    for xml in [
        format!(r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag val="x"/></p:tagLst>"#),
        format!(r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="x"/></p:tagLst>"#),
        format!(
            r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="x" val="y"><p:tag name="z" val="q"/></p:tag></p:tagLst>"#
        ),
        format!(r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:other/></p:tagLst>"#),
        format!(r#"<!DOCTYPE x><p:tagLst xmlns:p="{TRANSITIONAL}"/>"#),
        format!(r#"<?bad x?><p:tagLst xmlns:p="{TRANSITIONAL}"/>"#),
    ] {
        assert!(parse(xml.as_bytes()).is_err(), "{xml}");
    }
    assert!(matches!(
        parse(&vec![b' '; MAX_PART_BYTES + 1]),
        Err(Error::Limit { .. })
    ));
    assert!(Tag::new("bad\0name", "value").is_err());

    let entity = format!(
        r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="x" val="y">&amp;</p:tag></p:tagLst>"#
    );
    assert!(matches!(
        parse(entity.as_bytes()),
        Err(Error::Invalid(message)) if message.contains("entity references")
    ));
}

#[test]
fn parsing_rejects_canonical_escape_expansion_past_the_wire_budget() {
    const ENTITY: &str = "&#9;";
    let references_per_tag = (MAX_PART_BYTES - 512) / (2 * ENTITY.len());
    assert!(references_per_tag <= MAX_TEXT_BYTES);
    let references = ENTITY.repeat(references_per_tag);
    let xml = format!(
        r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="one" val="{references}"/><p:tag name="two" val="{references}"/></p:tagLst>"#
    );
    assert!(xml.len() <= MAX_PART_BYTES);
    assert!(matches!(
        parse(xml.as_bytes()),
        Err(Error::Limit {
            resource: "encoded tag-list bytes",
            ..
        })
    ));
}

#[test]
fn discovery_uses_stable_relationship_id_order() {
    use litchi_opc::{Part, XmlPart};

    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let mut slide = XmlPart::new(
        slide_name.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        Vec::new(),
    );
    for (relationship_id, target) in [("zId", "../tags/tag2.xml"), ("aId", "../tags/tag1.xml")] {
        slide.rels_mut().add_relationship(
            TAG_REL.into(),
            target.into(),
            relationship_id.into(),
            false,
        );
    }

    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    for (part_name, name) in [
        ("/ppt/tags/tag1.xml", "first"),
        ("/ppt/tags/tag2.xml", "second"),
    ] {
        package.add_part(Box::new(XmlPart::new(
            PackURI::new(part_name).unwrap(),
            CONTENT_TYPE.into(),
            format!(
                r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="{name}" val="1"/></p:tagLst>"#
            )
            .into_bytes(),
        )));
    }

    let source = package.get_part(&slide_name).unwrap();
    let discovered = discover(source, &package).unwrap();
    assert_eq!(
        discovered.iter().map(Source::rel).collect::<Vec<_>>(),
        ["aId", "zId"]
    );
    assert_eq!(discovered[0].list().get("first").unwrap().value(), "1");
    assert_eq!(discovered[1].list().get("second").unwrap().value(), "1");
}

#[test]
fn raw_attributes_require_valid_bound_prefixes() {
    let unbound = raw::Attr::new("x:value", "1").unwrap();
    assert!(
        Tag::new("name", "value")
            .unwrap()
            .with_attr(unbound)
            .is_err()
    );

    let tag = Tag::new("name", "value")
        .unwrap()
        .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
        .unwrap()
        .with_attr(raw::Attr::new("x:value", "1").unwrap())
        .unwrap();
    let mut list = List::new();
    list.add(tag).unwrap();
    assert!(
        std::str::from_utf8(&write(&list, Conformance::Transitional).unwrap())
            .unwrap()
            .contains("x:value=\"1\"")
    );
}

#[test]
fn escaped_size_budget_is_cached_and_failed_edits_are_atomic() {
    let quotes = "\"".repeat(MAX_TEXT_BYTES);
    assert_eq!(escaped_len(&quotes).unwrap(), 6 * MAX_TEXT_BYTES);

    let mut list = List::new()
        .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
        .unwrap()
        .with_attr(raw::Attr::new("x:padding", quotes.clone()).unwrap())
        .unwrap();
    list.add(Tag::new("small", "ok").unwrap()).unwrap();
    assert_eq!(
        list.wire_len,
        write(&list, Conformance::Transitional).unwrap().len()
    );

    {
        let before = list.clone();
        let replacement = Tag::new("large", quotes.clone()).unwrap();
        assert!(matches!(
            list.replace("small", replacement),
            Err(Error::Limit {
                resource: "encoded tag-list bytes",
                ..
            })
        ));
        assert_eq!(list, before);
    }
    {
        let before = list.clone();
        assert!(matches!(
            list.set("small", quotes.clone()),
            Err(Error::Limit {
                resource: "encoded tag-list bytes",
                ..
            })
        ));
        assert_eq!(list, before);
        assert_eq!(list.get("small").unwrap().value(), "ok");
    }
    {
        let before = list.clone();
        assert!(matches!(
            list.add(Tag::new("large", quotes.clone()).unwrap()),
            Err(Error::Limit {
                resource: "encoded tag-list bytes",
                ..
            })
        ));
        assert_eq!(list, before);
    }
    {
        let before = list.clone();
        assert!(matches!(
            list.insert(0, Tag::new("large", quotes.clone()).unwrap()),
            Err(Error::Limit {
                resource: "encoded tag-list bytes",
                ..
            })
        ));
        assert_eq!(list, before);
    }
    assert!(write(&list, Conformance::Transitional).is_ok());

    let root_overflow = List::new()
        .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
        .unwrap()
        .with_attr(raw::Attr::new("x:first", quotes.clone()).unwrap())
        .unwrap()
        .with_attr(raw::Attr::new("x:second", quotes.clone()).unwrap());
    assert!(matches!(root_overflow, Err(Error::Limit { .. })));

    let tag_overflow = Tag::new("standalone", "ok")
        .unwrap()
        .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
        .unwrap()
        .with_attr(raw::Attr::new("x:first", quotes.clone()).unwrap())
        .unwrap()
        .with_attr(raw::Attr::new("x:second", quotes).unwrap());
    assert!(matches!(tag_overflow, Err(Error::Limit { .. })));
}

#[test]
fn namespace_builders_reject_invalid_prospective_bindings() {
    let tag = Tag::new("name", "value")
        .unwrap()
        .with_attr(raw::Attr::new("xml:lang", "en").unwrap())
        .unwrap();
    assert!(
        tag.with_namespace(raw::Attr::new("xmlns:xml", "https://example.invalid/not-xml").unwrap())
            .is_err()
    );
    assert!(
        List::new()
            .with_namespace(raw::Attr::new("xmlns:xmlns", "urn:invalid").unwrap())
            .is_err()
    );
    assert!(
        Tag::new("name", "value")
            .unwrap()
            .with_namespace(raw::Attr::new("xmlns:x", "").unwrap())
            .is_err()
    );
}

fn owner_part(
    part_name: &str,
    root: &str,
    content_type: &str,
    conformance: Conformance,
) -> XmlPart {
    let body = if root == "presentation" {
        String::new()
    } else {
        "<p:cSld><p:spTree/></p:cSld>".to_owned()
    };
    XmlPart::new(
        PackURI::new(part_name).unwrap(),
        content_type.into(),
        format!(
            r#"<p:{root} xmlns:p="{}" xmlns:r="{}">{body}</p:{root}>"#,
            conformance.namespace(),
            conformance.relationship_namespace(),
        )
        .into_bytes(),
    )
}

#[derive(Clone, Copy)]
enum MceBranch {
    Choice,
    Fallback,
}

#[derive(Clone, Copy)]
enum MceContainer {
    Empty,
    Missing,
}

fn mce_owner_fixture(
    root: &str,
    conformance: Conformance,
    branch: MceBranch,
    container: MceContainer,
) -> (Vec<u8>, String) {
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let active_container = match container {
        MceContainer::Empty => r#"<p:custDataLst data-active="keep"/>"#,
        MceContainer::Missing => "",
    };
    let common_slide = root != "presentation";
    let active = if common_slide {
        format!(r"<p:spTree/>{active_container}<p:extLst/>")
    } else {
        format!(r"{active_container}<p:defaultTextStyle/>")
    };
    let inactive = if common_slide {
        r#"<p:spTree/><p:custDataLst x:keep="inactive"><!--inactive-comment--><p:tags r:id="rIdInactive"/></p:custDataLst><p:extLst/>"#.to_owned()
    } else {
        r#"<p:custDataLst x:keep="inactive"><!--inactive-comment--><p:tags r:id="rIdInactive"/></p:custDataLst><p:defaultTextStyle/>"#.to_owned()
    };
    let (alternate, inactive_branch) = match branch {
        MceBranch::Choice => (
            format!(
                r#"<mc:Choice Requires="p14">{active}</mc:Choice><mc:Fallback>{inactive}</mc:Fallback>"#
            ),
            format!(r"<mc:Fallback>{inactive}</mc:Fallback>"),
        ),
        MceBranch::Fallback => (
            format!(
                r#"<mc:Choice Requires="x">{inactive}</mc:Choice><mc:Fallback>{active}</mc:Fallback>"#
            ),
            format!(r#"<mc:Choice Requires="x">{inactive}</mc:Choice>"#),
        ),
    };
    let body = if common_slide {
        format!(r"<p:cSld><mc:AlternateContent>{alternate}</mc:AlternateContent></p:cSld>")
    } else {
        format!(r"<mc:AlternateContent>{alternate}</mc:AlternateContent>")
    };
    (
            format!(
                r#"<p:{root} xmlns:p="{}" xmlns:r="{}" xmlns:mc="{MC}" xmlns:p14="{P14}" xmlns:x="urn:unsupported" mc:Ignorable="x">{body}</p:{root}>"#,
                conformance.namespace(),
                conformance.relationship_namespace(),
            )
            .into_bytes(),
            inactive_branch,
        )
}

fn package_with_slide(conformance: Conformance) -> (OpcPackage, PackURI) {
    let name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let mut package = OpcPackage::new();
    package.add_part(Box::new(owner_part(
        name.as_str(),
        "sld",
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        conformance,
    )));
    (package, name)
}

fn list(name: &str, value: &str) -> List {
    let mut list = List::new();
    list.add(Tag::new(name, value).unwrap()).unwrap();
    list
}

fn mark_signed(package: &mut OpcPackage) {
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
}

#[test]
fn anchored_crud_is_strict_profiled_and_signature_safe() {
    use std::sync::Arc;

    let (mut package, owner) = package_with_slide(Conformance::Strict);
    assert_eq!(load(&package, &owner).unwrap(), None);
    assert_eq!(
        put(&mut package, &owner, list("Owner", "Alice")).unwrap(),
        None
    );
    let created = load(&package, &owner).unwrap().unwrap();
    assert_eq!(created.conformance(), Conformance::Strict);
    assert_eq!(created.rel(), "rId1");
    assert_eq!(created.part().as_str(), "/ppt/tags/tag1.xml");
    let relationship = package
        .get_part(&owner)
        .unwrap()
        .rels()
        .get(created.rel())
        .unwrap();
    assert_eq!(relationship.reltype(), STRICT_TAG_REL);
    assert!(
        std::str::from_utf8(package.get_part(created.part()).unwrap().blob())
            .unwrap()
            .contains(STRICT_TEXT)
    );
    let owner_xml = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    assert!(owner_xml.contains("<p:custDataLst"));
    assert!(owner_xml.contains("<p:tags"));

    let owner_before = package.get_part(&owner).unwrap().blob_arc();
    let part_before = package.get_part(created.part()).unwrap().blob_arc();
    mark_signed(&mut package);
    let old = put(&mut package, &owner, list("Owner", "Alice"))
        .unwrap()
        .unwrap();
    assert_eq!(old.get("owner").unwrap().value(), "Alice");
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &owner_before,
        &package.get_part(&owner).unwrap().blob_arc()
    ));
    assert!(Arc::ptr_eq(
        &part_before,
        &package.get_part(created.part()).unwrap().blob_arc()
    ));

    let malformed = format!(
        r#"<p:tagLst xmlns:p="{STRICT_TEXT}"><p:tag name="Owner" val="one"/><p:tag name="OWNER" val="two"/></p:tagLst>"#
    );
    let malformed = parse(malformed.as_bytes()).unwrap();
    assert!(matches!(
        put(&mut package, &owner, malformed),
        Err(Error::DuplicateName { .. })
    ));
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &part_before,
        &package.get_part(created.part()).unwrap().blob_arc()
    ));

    let old = put(&mut package, &owner, list("Reviewer", "Bob"))
        .unwrap()
        .unwrap();
    assert_eq!(old.get("owner").unwrap().value(), "Alice");
    assert!(!package.is_signed());
    assert_eq!(
        load(&package, &owner)
            .unwrap()
            .unwrap()
            .list()
            .get("reviewer")
            .unwrap()
            .value(),
        "Bob"
    );

    mark_signed(&mut package);
    let removed = remove(&mut package, &owner).unwrap().unwrap();
    assert_eq!(removed.get("reviewer").unwrap().value(), "Bob");
    assert!(!package.is_signed());
    assert!(package.get_part(created.part()).is_err());
    assert_eq!(load(&package, &owner).unwrap(), None);

    let after_remove = package.get_part(&owner).unwrap().blob_arc();
    mark_signed(&mut package);
    assert_eq!(remove(&mut package, &owner).unwrap(), None);
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &after_remove,
        &package.get_part(&owner).unwrap().blob_arc()
    ));
}

#[test]
fn customer_data_and_schema_order_are_preserved() {
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let original = format!(
        r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree/><p:custDataLst keep="yes"><p:custData r:id="rIdData"/></p:custDataLst><p:controls/></p:cSld></p:sld>"#
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        original.as_bytes().to_vec(),
    )));

    assert_eq!(
        put(&mut package, &owner, list("Owner", "Alice")).unwrap(),
        None
    );
    let updated = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    let customer = updated.find("<p:custData ").unwrap();
    let tags = updated.find("<p:tags ").unwrap();
    let controls = updated.find("<p:controls").unwrap();
    assert!(customer < tags && tags < controls);

    assert!(remove(&mut package, &owner).unwrap().is_some());
    assert_eq!(
        package.get_part(&owner).unwrap().blob(),
        original.as_bytes()
    );

    let empty = format!(
        r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree/><p:custDataLst keep="yes"/><p:controls/></p:cSld></p:sld>"#
    );
    package
        .get_part_mut(&owner)
        .unwrap()
        .set_blob(empty.into_bytes());
    assert_eq!(put(&mut package, &owner, List::new()).unwrap(), None);
    assert_eq!(remove(&mut package, &owner).unwrap(), Some(List::new()));
    let restored = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    assert!(restored.contains("<p:custDataLst keep=\"yes\"></p:custDataLst>"));
    assert!(restored.contains("<p:controls"));
}

#[test]
fn malformed_owner_order_is_rejected_before_mutation() {
    use std::sync::Arc;

    let cases = [
        "<p:cSld><p:custDataLst/><p:spTree/></p:cSld>",
        "<p:cSld><p:spTree/><p:spTree/></p:cSld>",
        r#"<p:cSld><p:spTree/><p:custDataLst><p:tags r:id="rId1"/><p:custData r:id="rIdData"/></p:custDataLst></p:cSld>"#,
    ];
    for body in cases {
        let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            format!(r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}">{body}</p:sld>"#)
                .into_bytes(),
        )));
        let owner_before = package.get_part(&owner).unwrap().blob_arc();
        mark_signed(&mut package);

        assert!(matches!(load(&package, &owner), Err(Error::Invalid(_))));
        assert!(matches!(
            put(&mut package, &owner, list("Owner", "Alice")),
            Err(Error::Invalid(_))
        ));
        assert!(matches!(
            remove(&mut package, &owner),
            Err(Error::Invalid(_))
        ));
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &owner_before,
            &package.get_part(&owner).unwrap().blob_arc()
        ));
    }
}

#[test]
fn presentation_customer_data_after_later_children_is_rejected_atomically_under_mce() {
    use std::sync::Arc;

    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let owner = PackURI::new("/ppt/presentation.xml").unwrap();
    let tag_part = PackURI::new("/ppt/tags/tag1.xml").unwrap();
    let owner_xml = format!(
        r#"<p:presentation xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}" xmlns:mc="{MC}" xmlns:p14="{P14}"><mc:AlternateContent><mc:Choice Requires="p14"><p:defaultTextStyle/><p:custDataLst><p:tags r:id="rId1"/></p:custDataLst></mc:Choice><mc:Fallback><p:custDataLst><p:tags r:id="rIdInactive"/></p:custDataLst><p:defaultTextStyle/></mc:Fallback></mc:AlternateContent></p:presentation>"#
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml".into(),
        owner_xml.into_bytes(),
    )));
    package.add_part(Box::new(XmlPart::new(
        tag_part.clone(),
        CONTENT_TYPE.into(),
        write(&list("Owner", "Alice"), Conformance::Transitional).unwrap(),
    )));
    package
        .get_part_mut(&owner)
        .unwrap()
        .rels_mut()
        .add_relationship(
            TAG_REL.into(),
            tag_part.relative_ref(owner.base_uri()),
            "rId1".into(),
            false,
        );
    let owner_before = package.get_part(&owner).unwrap().blob_arc();
    let tag_before = package.get_part(&tag_part).unwrap().blob_arc();
    mark_signed(&mut package);

    for error in [
        load(&package, &owner).unwrap_err(),
        put(&mut package, &owner, list("Reviewer", "Bob")).unwrap_err(),
        remove(&mut package, &owner).unwrap_err(),
    ] {
        assert!(matches!(
            error,
            Error::Invalid(message)
                if message.contains("must precede later root children")
        ));
    }
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &owner_before,
        &package.get_part(&owner).unwrap().blob_arc()
    ));
    assert!(Arc::ptr_eq(
        &tag_before,
        &package.get_part(&tag_part).unwrap().blob_arc()
    ));
    assert!(
        package
            .get_part(&owner)
            .unwrap()
            .rels()
            .get("rId1")
            .is_some()
    );
}

#[test]
fn package_shared_target_forks_and_collects_only_orphans() {
    let (mut package, first_owner) = package_with_slide(Conformance::Transitional);
    assert_eq!(
        put(&mut package, &first_owner, list("Owner", "Alice")).unwrap(),
        None
    );
    let original = load(&package, &first_owner).unwrap().unwrap();
    let original_part = original.part().clone();
    let original_bytes = package.get_part(&original_part).unwrap().blob_arc();

    let second_owner = PackURI::new("/ppt/slides/slide2.xml").unwrap();
    let second_xml = format!(
        r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree/><p:custDataLst><p:tags r:id="rIdShared"/></p:custDataLst></p:cSld></p:sld>"#
    );
    package.add_part(Box::new(XmlPart::new(
        second_owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        second_xml.into_bytes(),
    )));
    package
        .get_part_mut(&second_owner)
        .unwrap()
        .rels_mut()
        .add_relationship(
            TAG_REL.into(),
            original_part.relative_ref(second_owner.base_uri()),
            "rIdShared".into(),
            false,
        );

    assert_eq!(
        load(&package, &second_owner)
            .unwrap()
            .unwrap()
            .list()
            .get("owner")
            .unwrap()
            .value(),
        "Alice"
    );
    let old = put(&mut package, &first_owner, list("Reviewer", "Bob"))
        .unwrap()
        .unwrap();
    assert_eq!(old.get("owner").unwrap().value(), "Alice");
    let first = load(&package, &first_owner).unwrap().unwrap();
    let second = load(&package, &second_owner).unwrap().unwrap();
    assert_ne!(first.part(), &original_part);
    assert_eq!(second.part(), &original_part);
    assert_eq!(first.list().get("reviewer").unwrap().value(), "Bob");
    assert_eq!(second.list().get("owner").unwrap().value(), "Alice");
    assert!(std::sync::Arc::ptr_eq(
        &original_bytes,
        &package.get_part(&original_part).unwrap().blob_arc()
    ));

    let fork = first.part().clone();
    assert!(remove(&mut package, &first_owner).unwrap().is_some());
    assert!(package.get_part(&fork).is_err());
    assert!(package.get_part(&original_part).is_ok());
    assert!(remove(&mut package, &second_owner).unwrap().is_some());
    assert!(package.get_part(&original_part).is_err());
}

#[test]
fn same_owner_reused_anchor_forks_relationship_and_part() {
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let original_part = PackURI::new("/ppt/tags/tag1.xml").unwrap();
    let owner_xml = format!(
        r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="Shape"/><p:cNvSpPr/><p:nvPr><p:custDataLst><p:tags r:id="rIdShared"/></p:custDataLst></p:nvPr></p:nvSpPr></p:sp></p:spTree><p:custDataLst><p:tags xmlns:x="urn:test" x:keep="yes" r:id='rIdShared'><!--keep-anchor--></p:tags></p:custDataLst></p:cSld></p:sld>"#
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        owner_xml.into_bytes(),
    )));
    package.add_part(Box::new(XmlPart::new(
        original_part.clone(),
        CONTENT_TYPE.into(),
        write(&list("Owner", "Alice"), Conformance::Transitional).unwrap(),
    )));
    package
        .get_part_mut(&owner)
        .unwrap()
        .rels_mut()
        .add_relationship(
            TAG_REL.into(),
            original_part.relative_ref(owner.base_uri()),
            "rIdShared".into(),
            false,
        );

    let old = put(&mut package, &owner, list("Reviewer", "Bob"))
        .unwrap()
        .unwrap();
    assert_eq!(old.get("owner").unwrap().value(), "Alice");
    let direct = load(&package, &owner).unwrap().unwrap();
    assert_ne!(direct.rel(), "rIdShared");
    assert_ne!(direct.part(), &original_part);
    assert_eq!(direct.list().get("reviewer").unwrap().value(), "Bob");
    assert_eq!(
        discover(package.get_part(&owner).unwrap(), &package)
            .unwrap()
            .len(),
        2
    );
    let updated = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    assert_eq!(updated.matches("rIdShared").count(), 1);
    assert!(updated.contains("x:keep=\"yes\""));
    assert!(updated.contains("<!--keep-anchor-->"));
    assert!(package.get_part(&original_part).is_ok());

    let fork = direct.part().clone();
    assert!(remove(&mut package, &owner).unwrap().is_some());
    assert_eq!(load(&package, &owner).unwrap(), None);
    assert!(package.get_part(&fork).is_err());
    assert!(package.get_part(&original_part).is_ok());
    assert!(
        package
            .get_part(&owner)
            .unwrap()
            .rels()
            .get("rIdShared")
            .is_some()
    );
    assert_eq!(
        discover(package.get_part(&owner).unwrap(), &package)
            .unwrap()
            .len(),
        1
    );
}

#[test]
fn direct_owner_mutates_active_mce_fallback_and_preserves_inactive_source() {
    use std::sync::Arc;

    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let part_name = PackURI::new("/ppt/tags/tag1.xml").unwrap();
    let inactive = r#"<mc:Choice Requires="x"><p:custDataLst x:keep="inactive"><!--inactive-comment--><p:tags r:id="rIdInactive"/></p:custDataLst></mc:Choice>"#;
    let owner_xml = format!(
        r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}" xmlns:mc="{MC}" xmlns:x="urn:future" mc:Ignorable="x"><p:cSld><p:spTree/><mc:AlternateContent>{inactive}<mc:Fallback><p:custDataLst><p:tags x:keep="active" r:id='rIdActive'><!--active-comment--></p:tags></p:custDataLst></mc:Fallback></mc:AlternateContent></p:cSld></p:sld>"#
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        owner_xml.into_bytes(),
    )));
    package.add_part(Box::new(XmlPart::new(
        part_name.clone(),
        CONTENT_TYPE.into(),
        write(&list("Branch", "fallback"), Conformance::Transitional).unwrap(),
    )));
    package
        .get_part_mut(&owner)
        .unwrap()
        .rels_mut()
        .add_relationship(
            TAG_REL.into(),
            part_name.relative_ref(owner.base_uri()),
            "rIdActive".into(),
            false,
        );

    assert_eq!(
        load(&package, &owner)
            .unwrap()
            .unwrap()
            .list()
            .get("branch")
            .unwrap()
            .value(),
        "fallback"
    );

    let before = package.get_part(&owner).unwrap().blob_arc();
    mark_signed(&mut package);
    let old = put(&mut package, &owner, list("Reviewer", "Ada"))
        .unwrap()
        .unwrap();
    assert_eq!(old.get("branch").unwrap().value(), "fallback");
    assert_eq!(
        load(&package, &owner)
            .unwrap()
            .unwrap()
            .list()
            .get("reviewer")
            .unwrap()
            .value(),
        "Ada"
    );
    assert!(!package.is_signed());
    assert!(Arc::ptr_eq(
        &before,
        &package.get_part(&owner).unwrap().blob_arc()
    ));

    mark_signed(&mut package);
    let removed = remove(&mut package, &owner).unwrap().unwrap();
    assert_eq!(removed.get("reviewer").unwrap().value(), "Ada");
    assert!(!package.is_signed());
    assert_eq!(load(&package, &owner).unwrap(), None);
    let updated = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    assert!(updated.contains(inactive));
    assert!(!updated.contains("rIdActive"));
    assert!(package.get_part(&part_name).is_err());
}

#[test]
fn direct_owner_mce_creation_covers_all_owners_branches_profiles_and_container_states() {
    let cases = [
        (
            "/ppt/presentation.xml",
            "presentation",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            Conformance::Strict,
            MceBranch::Choice,
            MceContainer::Empty,
        ),
        (
            "/ppt/presentation.xml",
            "presentation",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            Conformance::Transitional,
            MceBranch::Fallback,
            MceContainer::Missing,
        ),
        (
            "/ppt/slides/slide1.xml",
            "sld",
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            Conformance::Transitional,
            MceBranch::Choice,
            MceContainer::Missing,
        ),
        (
            "/ppt/slideLayouts/slideLayout1.xml",
            "sldLayout",
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
            Conformance::Strict,
            MceBranch::Fallback,
            MceContainer::Empty,
        ),
        (
            "/ppt/slideMasters/slideMaster1.xml",
            "sldMaster",
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
            Conformance::Transitional,
            MceBranch::Choice,
            MceContainer::Empty,
        ),
        (
            "/ppt/notesSlides/notesSlide1.xml",
            "notes",
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
            Conformance::Strict,
            MceBranch::Fallback,
            MceContainer::Missing,
        ),
        (
            "/ppt/notesMasters/notesMaster1.xml",
            "notesMaster",
            "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
            Conformance::Transitional,
            MceBranch::Choice,
            MceContainer::Missing,
        ),
        (
            "/ppt/handoutMasters/handoutMaster1.xml",
            "handoutMaster",
            "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
            Conformance::Strict,
            MceBranch::Fallback,
            MceContainer::Empty,
        ),
    ];

    for (path, root, content_type, conformance, branch, container) in cases {
        let owner = PackURI::new(path).unwrap();
        let (xml, inactive) = mce_owner_fixture(root, conformance, branch, container);
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            content_type.into(),
            xml,
        )));
        assert_eq!(load(&package, &owner).unwrap(), None, "{root}");
        assert_eq!(
            put(&mut package, &owner, list("Owner", root)).unwrap(),
            None,
            "{root}"
        );
        let source = load(&package, &owner).unwrap().unwrap();
        assert_eq!(source.conformance(), conformance, "{root}");
        assert_eq!(source.list().get("owner").unwrap().value(), root);
        let tag_part = source.part().clone();
        let raw = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        assert!(
            raw.contains(&inactive),
            "inactive branch changed for {root}"
        );
        let processed = process_owner_ooxml(raw.as_bytes()).unwrap();
        let processed = std::str::from_utf8(processed.as_ref()).unwrap();
        let tags = processed.find("<p:tags ").unwrap();
        if root == "presentation" {
            assert!(tags < processed.find("<p:defaultTextStyle").unwrap());
        } else {
            assert!(processed.find("<p:spTree").unwrap() < tags);
            assert!(tags < processed.find("<p:extLst").unwrap());
        }

        assert!(remove(&mut package, &owner).unwrap().is_some(), "{root}");
        assert_eq!(load(&package, &owner).unwrap(), None, "{root}");
        let raw = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        assert!(
            raw.contains(&inactive),
            "inactive branch changed for {root}"
        );
        assert!(
            package.get_part(&tag_part).is_err(),
            "orphan retained for {root}"
        );
        assert!(
            discover(package.get_part(&owner).unwrap(), &package)
                .unwrap()
                .is_empty()
        );
    }
}

#[test]
fn inactive_same_id_forces_active_fork_and_retains_raw_consumer_graph() {
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let original_part = PackURI::new("/ppt/tags/tag1.xml").unwrap();
    let inactive = r#"<mc:Fallback><p:spTree/><p:custDataLst x:keep="inactive"><!--inactive-same-id--><p:tags r:id="rIdShared"/></p:custDataLst></mc:Fallback>"#;
    let make_package = || {
        let owner_xml = format!(
            r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}" xmlns:mc="{MC}" xmlns:p14="{P14}" xmlns:x="urn:test"><p:cSld><mc:AlternateContent><mc:Choice Requires="p14"><p:spTree/><p:custDataLst><p:tags x:keep="active" r:id='rIdShared'><!--active-same-id--></p:tags></p:custDataLst></mc:Choice>{inactive}</mc:AlternateContent></p:cSld></p:sld>"#
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            owner_xml.into_bytes(),
        )));
        package.add_part(Box::new(XmlPart::new(
            original_part.clone(),
            CONTENT_TYPE.into(),
            write(&list("Owner", "Alice"), Conformance::Transitional).unwrap(),
        )));
        package
            .get_part_mut(&owner)
            .unwrap()
            .rels_mut()
            .add_relationship(
                TAG_REL.into(),
                original_part.relative_ref(owner.base_uri()),
                "rIdShared".into(),
                false,
            );
        package
    };

    let mut package = make_package();
    let old = put(&mut package, &owner, list("Reviewer", "Bob"))
        .unwrap()
        .unwrap();
    assert_eq!(old.get("owner").unwrap().value(), "Alice");
    let active = load(&package, &owner).unwrap().unwrap();
    assert_ne!(active.rel(), "rIdShared");
    assert_ne!(active.part(), &original_part);
    assert_eq!(active.list().get("reviewer").unwrap().value(), "Bob");
    let forked = active.part().clone();
    let raw = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    assert!(raw.contains(inactive));
    assert!(raw.contains("x:keep=\"active\""));
    assert!(raw.contains("<!--active-same-id-->"));
    assert_eq!(raw.matches("rIdShared").count(), 1);
    assert!(package.get_part(&original_part).is_ok());
    assert!(remove(&mut package, &owner).unwrap().is_some());
    assert!(package.get_part(&forked).is_err());
    assert!(package.get_part(&original_part).is_ok());
    assert!(
        package
            .get_part(&owner)
            .unwrap()
            .rels()
            .get("rIdShared")
            .is_some()
    );

    let mut package = make_package();
    assert!(remove(&mut package, &owner).unwrap().is_some());
    assert_eq!(load(&package, &owner).unwrap(), None);
    let raw = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
    assert!(raw.contains(inactive));
    assert!(
        package
            .get_part(&owner)
            .unwrap()
            .rels()
            .get("rIdShared")
            .is_some()
    );
    assert!(package.get_part(&original_part).is_ok());
}

#[test]
fn mixed_profile_preflight_is_atomic() {
    use std::sync::Arc;

    let (mut package, owner) = package_with_slide(Conformance::Strict);
    assert_eq!(
        put(&mut package, &owner, list("Owner", "Alice")).unwrap(),
        None
    );
    let part_name = load(&package, &owner).unwrap().unwrap().part().clone();
    package
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(write(&list("Owner", "Alice"), Conformance::Transitional).unwrap());
    let owner_before = package.get_part(&owner).unwrap().blob_arc();
    let part_before = package.get_part(&part_name).unwrap().blob_arc();
    mark_signed(&mut package);

    for result in [
        load(&package, &owner).map(|_| ()),
        put(&mut package, &owner, list("Reviewer", "Bob")).map(|_| ()),
        remove(&mut package, &owner).map(|_| ()),
    ] {
        assert!(matches!(
            result,
            Err(Error::Invalid(message)) if message.contains("namespace profile")
        ));
    }
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &owner_before,
        &package.get_part(&owner).unwrap().blob_arc()
    ));
    assert!(Arc::ptr_eq(
        &part_before,
        &package.get_part(&part_name).unwrap().blob_arc()
    ));
}

#[test]
fn mixed_owner_relationship_profiles_are_rejected_atomically() {
    use std::sync::Arc;

    for (relationship_namespace, relationship_type) in
        [(REL_TEXT, STRICT_TAG_REL), (STRICT_REL_TEXT, TAG_REL)]
    {
        let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let part_name = PackURI::new("/ppt/tags/tag1.xml").unwrap();
        let owner_xml = format!(
            r#"<p:sld xmlns:p="{STRICT_TEXT}" xmlns:r="{relationship_namespace}"><p:cSld><p:spTree/><p:custDataLst><p:tags r:id="rId1"/></p:custDataLst></p:cSld></p:sld>"#,
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            owner_xml.into_bytes(),
        )));
        package.add_part(Box::new(XmlPart::new(
            part_name.clone(),
            CONTENT_TYPE.into(),
            write(&list("Owner", "Alice"), Conformance::Strict).unwrap(),
        )));
        package
            .get_part_mut(&owner)
            .unwrap()
            .rels_mut()
            .add_relationship(
                relationship_type.into(),
                part_name.relative_ref(owner.base_uri()),
                "rId1".into(),
                false,
            );
        let owner_before = package.get_part(&owner).unwrap().blob_arc();
        let part_before = package.get_part(&part_name).unwrap().blob_arc();
        mark_signed(&mut package);

        assert!(matches!(load(&package, &owner), Err(Error::Invalid(_))));
        assert!(matches!(
            put(&mut package, &owner, list("Reviewer", "Bob")),
            Err(Error::Invalid(_))
        ));
        assert!(matches!(
            remove(&mut package, &owner),
            Err(Error::Invalid(_))
        ));
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &owner_before,
            &package.get_part(&owner).unwrap().blob_arc()
        ));
        assert!(Arc::ptr_eq(
            &part_before,
            &package.get_part(&part_name).unwrap().blob_arc()
        ));
    }
}

#[test]
fn creation_supports_all_presentationml_tag_owners_and_empty_lists() {
    let owners = [
        (
            "/ppt/presentation.xml",
            "presentation",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        ),
        (
            "/ppt/slides/slide1.xml",
            "sld",
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        ),
        (
            "/ppt/slideLayouts/slideLayout1.xml",
            "sldLayout",
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
        ),
        (
            "/ppt/slideMasters/slideMaster1.xml",
            "sldMaster",
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        ),
        (
            "/ppt/notesSlides/notesSlide1.xml",
            "notes",
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        ),
        (
            "/ppt/notesMasters/notesMaster1.xml",
            "notesMaster",
            "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
        ),
        (
            "/ppt/handoutMasters/handoutMaster1.xml",
            "handoutMaster",
            "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
        ),
    ];
    let mut package = OpcPackage::new();
    for (part_name, root, content_type) in owners {
        let owner = PackURI::new(part_name).unwrap();
        package.add_part(Box::new(owner_part(
            part_name,
            root,
            content_type,
            Conformance::Transitional,
        )));
        assert_eq!(put(&mut package, &owner, List::new()).unwrap(), None);
        let source = load(&package, &owner).unwrap().unwrap();
        assert!(source.list().is_empty());
        assert_eq!(source.conformance(), Conformance::Transitional);
        assert_eq!(
            discover(package.get_part(&owner).unwrap(), &package)
                .unwrap()
                .len(),
            1
        );
        assert_eq!(remove(&mut package, &owner).unwrap(), Some(List::new()));
        assert_eq!(load(&package, &owner).unwrap(), None);
    }
}

#[test]
fn part_allocation_avoids_ascii_case_and_derived_name_collisions() {
    use litchi_opc::BlobPart;

    let (mut package, owner) = package_with_slide(Conformance::Transitional);
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/PPT/TAGS/TAG1.XML").unwrap(),
        "application/octet-stream".into(),
        Vec::new(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/tags/tag2.xml/child").unwrap(),
        "application/octet-stream".into(),
        Vec::new(),
    )));

    assert_eq!(put(&mut package, &owner, List::new()).unwrap(), None);
    let source = load(&package, &owner).unwrap().unwrap();
    assert_eq!(source.part().as_str(), "/ppt/tags/tag3.xml");
}
