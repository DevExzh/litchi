//! Public package-owner integration coverage for PresentationML ink actions.
//!
//! The action namespace has no checked-in native Office fixture.  These tests
//! therefore use a small synthetic OPC slide whose owner, relationship, MIME,
//! MCE branches, and action payload are all explicit.  The fixture is kept in
//! this integration test so it cannot become an undocumented production
//! authoring profile.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration fixtures and assertions panic on failure by design"
)]

use litchi_drawingml::ink::actions::{ActionSelector, ActionType, ChildSelector};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, Part, ReadLimits, ReadResource};
use litchi_pptx::Package;
use litchi_pptx::presentation::embedded::ink_actions;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_MCE: &str = "http://purl.oclc.org/ooxml/markup-compatibility/2006";
const P14_MAIN: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const ACTION: &str = "http://schemas.microsoft.com/office/powerpoint/2014/inkAction";
const INKML: &str = "http://www.w3.org/2003/InkML";

const SLIDE_NAME: &str = "/ppt/slides/slide1.xml";
const ACTION_NAME: &str = "/ppt/custom/action.xml";
const ACTION2_NAME: &str = "/ppt/custom/action2.xml";

const ACTION_REL: &str = "rIdAction";
const ACTION_TARGET: &str = "../custom/action.xml";

/// A valid shared-profile payload with direct and grouped actions and opaque
/// InkML descendants.  The package owner must classify this only after its
/// enclosing PresentationML owner and OPC closure have been checked.
fn action_payload() -> Vec<u8> {
    format!(
        r###"<?xml version="1.0" encoding="UTF-8"?><ia:actions xmlns:ia="{ACTION}" xmlns:i="{INKML}" xmlns:v="urn:vendor" lengthUnit="cm" timeUnit="ms" xml:id="root"><i:definitions><v:future v:flag="keep"/></i:definitions><ia:action xml:id="a0" type="add" startTime="0"><ia:property name="kind" value="old"/><ia:actionData xml:id="d0" name="stroke" ref="#d1"><ia:transform matrix="1,0,0,1"/><i:trace><v:opaque>keep</v:opaque></i:trace><i:traceView/></ia:actionData><ia:actionDataGroup xml:id="dg0" name="group"><ia:actionData xml:id="d1" name="other"/></ia:actionDataGroup></ia:action><ia:actionGroup xml:id="ag0" type="transform" startTime="1"><ia:action xml:id="a1" type="remove" startTime="2"><ia:actionData/></ia:action></ia:actionGroup></ia:actions>"###
    )
    .into_bytes()
}

/// A complete MCE owner shape.  Prefixes are intentionally arbitrary and the
/// fallback is a real `p:pic`; tests that vary prefixes/default namespaces use
/// this function as the only source of owner markup.
fn action_owner_anchor(
    pml_prefix: &str,
    rel_prefix: &str,
    mce_prefix: &str,
    required: &str,
    relationship_id: &str,
) -> String {
    format!(
        r#"<{mce_prefix}:AlternateContent><{mce_prefix}:Choice Requires="{required}"><{pml_prefix}:contentPart {rel_prefix}:id="{relationship_id}"/></{mce_prefix}:Choice><{mce_prefix}:Fallback><{pml_prefix}:pic><{pml_prefix}:nvPicPr><{pml_prefix}:cNvPr id="42" name="Fallback"/><{pml_prefix}:cNvPicPr/><{pml_prefix}:nvPr/></{pml_prefix}:nvPicPr><{pml_prefix}:blipFill/><{pml_prefix}:spPr/></{pml_prefix}:pic></{mce_prefix}:Fallback></{mce_prefix}:AlternateContent>"#
    )
}

fn slide_xml(_strict: bool, relationship_ids: &[&str]) -> Vec<u8> {
    let pml_prefix = "pp";
    let rel_prefix = "rr";
    let mce_prefix = "compat";
    let action_prefix = "inkAction";
    let required = format!("p14main {action_prefix}");
    let anchors = relationship_ids
        .iter()
        .map(|relationship_id| {
            action_owner_anchor(
                pml_prefix,
                rel_prefix,
                mce_prefix,
                &required,
                relationship_id,
            )
        })
        .collect::<String>();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><{pml_prefix}:sld xmlns:{pml_prefix}="{pml_namespace}" xmlns:{rel_prefix}="{rel_namespace}" xmlns:{mce_prefix}="{mce_namespace}" xmlns:p14main="{P14_MAIN}" xmlns:{action_prefix}="{ACTION}" {mce_prefix}:Ignorable="p14main {action_prefix}"><{pml_prefix}:cSld><{pml_prefix}:spTree><{pml_prefix}:nvGrpSpPr/><{pml_prefix}:grpSpPr/>{anchors}</{pml_prefix}:spTree></{pml_prefix}:cSld></{pml_prefix}:sld>"#,
        anchors = anchors,
        pml_namespace = if _strict { STRICT_PML } else { PML },
        rel_namespace = if _strict { STRICT_REL } else { REL },
        // ECMA-376 Part 3 §7.1 fixes the markup-compatibility namespace URI
        // for both Transitional and Strict packages.  Strict changes the
        // PresentationML and relationship URIs, but not this namespace.
        mce_namespace = MCE,
    )
    .into_bytes()
}

/// Build the smallest OPC graph needed by the public owner reader.  The
/// helper deliberately uses the existing OPC part/relationship seams instead
/// of ZIP implementation types.
fn package_with_action(strict: bool) -> (OpcPackage, PackURI, PackURI) {
    package_with_actions(strict, &[ACTION_REL])
}

fn package_with_actions(strict: bool, relationship_ids: &[&str]) -> (OpcPackage, PackURI, PackURI) {
    let target_refs = vec![ACTION_TARGET; relationship_ids.len()];
    package_with_action_refs(strict, relationship_ids, &target_refs)
}

fn package_with_action_refs(
    strict: bool,
    relationship_ids: &[&str],
    target_refs: &[&str],
) -> (OpcPackage, PackURI, PackURI) {
    assert_eq!(relationship_ids.len(), target_refs.len());
    let slide_name = PackURI::new(SLIDE_NAME).unwrap();
    let action_name = PackURI::new(ACTION_NAME).unwrap();
    let mut slide = BlobPart::new(
        slide_name.clone(),
        ct::PML_SLIDE.to_owned(),
        slide_xml(strict, relationship_ids),
    );
    for (relationship_id, target_ref) in relationship_ids.iter().zip(target_refs) {
        slide.rels_mut().add_relationship(
            if strict {
                STRICT_CUSTOM_XML.to_owned()
            } else {
                rt::CUSTOM_XML.to_owned()
            },
            (*target_ref).to_owned(),
            (*relationship_id).to_owned(),
            false,
        );
    }
    let action = BlobPart::new(action_name.clone(), "text/xml".to_owned(), action_payload());
    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    package.add_part(Box::new(action));
    (package, slide_name, action_name)
}

/// Build a two-slide PresentationML owner whose second slide either shares the
/// first target or owns a distinct target.  The public presentation facade is
/// the route under test for package-wide aggregate accounting.
fn package_with_two_slide_presentation(shared_target: bool) -> OpcPackage {
    let (mut package, _slide_name, _action_name) = package_with_action(false);
    let second_slide_name = PackURI::new("/ppt/slides/slide2.xml").unwrap();
    let second_target = if shared_target {
        ACTION_TARGET
    } else {
        "../custom/action2.xml"
    };
    let mut second_slide = BlobPart::new(
        second_slide_name,
        ct::PML_SLIDE.to_owned(),
        slide_xml(false, &[ACTION_REL]),
    );
    second_slide.rels_mut().add_relationship(
        rt::CUSTOM_XML.to_owned(),
        second_target.to_owned(),
        ACTION_REL.to_owned(),
        false,
    );
    package.add_part(Box::new(second_slide));
    if !shared_target {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(ACTION2_NAME).unwrap(),
            "text/xml".to_owned(),
            action_payload(),
        )));
    }

    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRESENTATION_MAIN.to_owned(),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
    );
    presentation.rels_mut().add_relationship(
        rt::SLIDE.to_owned(),
        "slides/slide1.xml".to_owned(),
        "rIdSlide".to_owned(),
        false,
    );
    presentation.rels_mut().add_relationship(
        rt::SLIDE.to_owned(),
        "slides/slide2.xml".to_owned(),
        "rIdSlide2".to_owned(),
        false,
    );
    package.add_part(Box::new(presentation));
    package.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "ppt/presentation.xml".to_owned(),
        "rIdOfficeDocument".to_owned(),
        false,
    );
    package
}

#[test]
fn synthetic_owner_fixture_has_the_normative_closure() {
    for strict in [false, true] {
        let (package, slide_name, action_name) = package_with_action(strict);
        let slide = package.get_part(&slide_name).expect("slide part");
        let relationship = slide.rels().get(ACTION_REL).expect("action relationship");
        assert_eq!(relationship.target_partname().unwrap(), action_name);
        assert!(!relationship.is_external());
        assert_eq!(
            relationship.reltype(),
            if strict {
                STRICT_CUSTOM_XML
            } else {
                rt::CUSTOM_XML
            }
        );
        assert_eq!(
            package.get_part(&action_name).unwrap().content_type(),
            "text/xml"
        );
        let slide_xml = std::str::from_utf8(slide.blob()).unwrap();
        assert!(slide_xml.contains("Requires=\"p14main inkAction\""));
        assert!(slide_xml.contains(":contentPart rr:id=\"rIdAction\""));
        assert!(slide_xml.contains(":Fallback><pp:pic>"));
        let payload = package.get_part(&action_name).unwrap().blob();
        assert!(payload.starts_with(b"<?xml version"));
        assert!(
            payload
                .windows(ACTION.len())
                .any(|window| window == ACTION.as_bytes())
        );
    }
}

fn action_snapshot(package: &OpcPackage, slide_name: &PackURI) -> ink_actions::Snapshot {
    let slide = package.get_part(slide_name).expect("slide part");
    ink_actions::load_snapshot(package, 0, slide, &ink_actions::Limits::default())
        .expect("typed action snapshot")
}

fn action_snapshot_with_limits(
    package: &OpcPackage,
    slide_name: &PackURI,
    limits: ink_actions::Limits,
) -> Result<ink_actions::Snapshot, litchi_pptx::Error> {
    let slide = package.get_part(slide_name).expect("slide part");
    ink_actions::load_snapshot(package, 0, slide, &limits)
}

fn slide_with_default_pml_namespace(relationship_ids: &[&str]) -> Vec<u8> {
    let xml = String::from_utf8(slide_xml(false, relationship_ids)).unwrap();
    xml.replace("xmlns:pp=\"", "xmlns=\"")
        .replace("<pp:", "<")
        .replace("</pp:", "</")
        .into_bytes()
}

fn slide_with_escaped_action_namespace(relationship_ids: &[&str]) -> Vec<u8> {
    let xml = String::from_utf8(slide_xml(false, relationship_ids)).unwrap();
    xml.replace(
        ACTION,
        "http:&#x2F;&#x2F;schemas.microsoft.com&#x2F;office&#x2F;powerpoint&#x2F;2014&#x2F;inkAction",
    )
    .replace(
        P14_MAIN,
        "http:&#x2F;&#x2F;schemas.microsoft.com&#x2F;office&#x2F;powerpoint&#x2F;2010&#x2F;main",
    )
    .into_bytes()
}

#[test]
fn reads_action_profile_only_after_owner_and_opc_closure() {
    for strict in [false, true] {
        let (package, slide_name, action_name) = package_with_action(strict);
        let snapshot = action_snapshot(&package, &slide_name);
        assert_eq!(snapshot.slide_index(), 0);
        assert_eq!(
            snapshot.slide_part_name(),
            &PackURI::new(SLIDE_NAME).unwrap()
        );
        assert_eq!(snapshot.anchors().len(), 1);

        let anchor = &snapshot.anchors()[0];
        assert_eq!(anchor.semantic_ordinal(), 0);
        assert_eq!(anchor.source_ordinal(), 0);
        assert_eq!(anchor.relationship_id(), ACTION_REL);
        assert_eq!(anchor.target_part_name(), &action_name);
        assert_eq!(anchor.target_ref(), ACTION_TARGET);
        assert_eq!(anchor.target_mode(), litchi_opc::TargetMode::Internal);
        assert_eq!(anchor.content_type(), "text/xml");
        assert_eq!(
            anchor.target_bytes(),
            package.get_part(&action_name).unwrap().blob()
        );
        assert_eq!(anchor.inbound_references().len(), 1);
        assert_eq!(anchor.inbound_references()[0].relationship_id(), ACTION_REL);
        assert!(anchor.outbound_references().is_empty());
        assert_eq!(anchor.profile().length_unit().as_str(), "cm");
        assert_eq!(anchor.profile().time_unit().as_str(), "ms");
        assert_eq!(anchor.profile().actions().count(), 1);
        assert_eq!(anchor.profile().action_groups().count(), 1);
        assert!(
            anchor
                .owner_xml()
                .windows(b"compat:Fallback".len())
                .any(|window| { window == b"compat:Fallback" })
        );
        assert!(
            anchor
                .fallback_xml()
                .windows(b"pp:pic".len())
                .any(|window| { window == b"pp:pic" })
        );
        assert!(
            anchor
                .choice_xml()
                .windows(b"pp:contentPart".len())
                .any(|window| { window == b"pp:contentPart" })
        );
    }
}

#[test]
fn strict_owner_rejects_a_foreign_markup_compatibility_namespace() {
    let (mut package, slide_name, action_name) = package_with_action(true);
    let strict_source = package.get_part(&slide_name).unwrap().blob().to_vec();
    let foreign_source = String::from_utf8(strict_source.clone())
        .unwrap()
        .replace(MCE, STRICT_MCE)
        .into_bytes();
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(foreign_source.clone());
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
    let snapshot = action_snapshot(&package, &slide_name);
    assert!(
        snapshot.anchors().is_empty(),
        "foreign markup-compatibility namespace must remain opaque"
    );
    assert_eq!(snapshot.source_xml(), foreign_source.as_slice());
    assert_eq!(
        package.get_part(&slide_name).unwrap().blob(),
        foreign_source
    );
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
}

#[test]
fn strict_owner_rejects_a_transitional_presentation_namespace() {
    let (mut package, slide_name, action_name) = package_with_action(true);
    let strict_source = package.get_part(&slide_name).unwrap().blob().to_vec();
    let mismatched_source = String::from_utf8(strict_source)
        .unwrap()
        .replace(STRICT_PML, PML)
        .into_bytes();
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(mismatched_source.clone());
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
    let error = load_error(&package, &slide_name);
    assert!(
        matches!(error, litchi_pptx::Error::Invalid(_)),
        "Strict owner with Transitional PresentationML must refuse: {error:?}"
    );
    assert_eq!(
        package.get_part(&slide_name).unwrap().blob(),
        mismatched_source
    );
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
}

#[test]
fn strict_owner_rejects_a_transitional_relationship_namespace() {
    let (mut package, slide_name, action_name) = package_with_action(true);
    let strict_source = package.get_part(&slide_name).unwrap().blob().to_vec();
    let mismatched_source = String::from_utf8(strict_source)
        .unwrap()
        .replace(STRICT_REL, REL)
        .into_bytes();
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(mismatched_source.clone());
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
    let error = load_error(&package, &slide_name);
    assert!(
        matches!(error, litchi_pptx::Error::Invalid(_)),
        "Strict owner with Transitional relationships must refuse: {error:?}"
    );
    assert_eq!(
        package.get_part(&slide_name).unwrap().blob(),
        mismatched_source
    );
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
}

#[test]
fn strict_owner_rejects_a_transitional_custom_xml_relationship_type() {
    let (mut package, slide_name, action_name) = package_with_action(true);
    let strict_source = package.get_part(&slide_name).unwrap().blob().to_vec();
    let slide = package.get_part_mut(&slide_name).unwrap();
    slide.rels_mut().remove(ACTION_REL);
    slide.rels_mut().add_relationship(
        rt::CUSTOM_XML.to_owned(),
        ACTION_TARGET.to_owned(),
        ACTION_REL.to_owned(),
        false,
    );
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
    let error = load_error(&package, &slide_name);
    assert!(
        matches!(error, litchi_pptx::Error::Relationship(_)),
        "Strict owner with Transitional customXml relationship must refuse: {error:?}"
    );
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), strict_source);
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
}

#[test]
fn aliases_default_namespaces_and_escaped_uris_are_resolved_semantically() {
    let (mut package, slide_name, _action_name) = package_with_action(false);

    // The ordinary fixture already uses aliases for every relevant namespace.
    assert_eq!(action_snapshot(&package, &slide_name).anchors().len(), 1);

    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(slide_with_default_pml_namespace(&[ACTION_REL]));
    let default_namespace = action_snapshot(&package, &slide_name);
    assert_eq!(default_namespace.anchors().len(), 1);
    assert_eq!(default_namespace.anchors()[0].relationship_id(), ACTION_REL);
    assert!(
        default_namespace.anchors()[0]
            .fallback_xml()
            .windows(b"pic".len())
            .any(|window| window == b"pic")
    );

    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(slide_with_escaped_action_namespace(&[ACTION_REL]));
    let escaped_uri = action_snapshot(&package, &slide_name);
    assert_eq!(escaped_uri.anchors().len(), 1);
    assert_eq!(escaped_uri.anchors()[0].relationship_id(), ACTION_REL);
    assert_eq!(escaped_uri.anchors()[0].semantic_ordinal(), 0);
}

#[test]
fn scalar_profile_edit_changes_only_the_target_and_replays_mce_exactly() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let source = action_snapshot(&package, &slide_name);
    let selector = source.selector(0).expect("semantic action selector");
    let original_owner = source.anchors()[0].owner_xml().to_vec();
    let original_choice = source.anchors()[0].choice_xml().to_vec();
    let original_fallback = source.anchors()[0].fallback_xml().to_vec();
    let original_relationship = package
        .get_part(&slide_name)
        .unwrap()
        .rels()
        .get(ACTION_REL)
        .unwrap()
        .clone();
    let original_target = package.get_part(&action_name).unwrap().blob().to_vec();

    let mut edit = source.edit();
    assert!(
        edit.edit_profile(selector, |profile| {
            profile.set_action_type(ActionSelector::ordinal(0), ActionType::Transform)?;
            profile.set_start_time(ActionSelector::ordinal(0), " 2.50 ")?;
            profile.set_property_value(
                ChildSelector::Property {
                    action: ActionSelector::ordinal(0),
                    index: 0,
                },
                "changed",
            )?;
            Ok(())
        })
        .expect("scalar profile edit")
    );
    let commit = edit.commit().expect("action target commit");
    assert!(commit.is_changed());
    assert_eq!(commit.snapshot().source_xml(), source.source_xml());
    assert_eq!(commit.snapshot().anchors()[0].owner_xml(), original_owner);
    assert_eq!(commit.snapshot().anchors()[0].choice_xml(), original_choice);
    assert_eq!(
        commit.snapshot().anchors()[0].fallback_xml(),
        original_fallback
    );
    assert_eq!(
        commit.snapshot().anchors()[0].relationship_type(),
        original_relationship.reltype()
    );
    assert_eq!(
        commit.snapshot().anchors()[0].target_ref(),
        original_relationship.target_ref()
    );
    assert_eq!(
        commit.snapshot().anchors()[0].target_mode(),
        original_relationship.target_mode()
    );
    assert_eq!(commit.snapshot().anchors()[0].content_type(), "text/xml");
    assert_ne!(
        commit.snapshot().anchors()[0].target_bytes(),
        original_target
    );
    assert!(
        std::str::from_utf8(commit.snapshot().anchors()[0].target_bytes())
            .unwrap()
            .contains(r#"type="transform""#)
    );

    ink_actions::apply_patch(&mut package, commit.patch()).expect("publish target edit");
    let reopened = action_snapshot(&package, &slide_name);
    assert_eq!(reopened.anchors()[0].profile().actions().count(), 1);
    assert_eq!(reopened.anchors()[0].profile().action_groups().count(), 1);
    assert_eq!(
        reopened.anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .start_time(),
        "2.50"
    );
    assert_eq!(
        reopened.anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .action_type()
            .as_str(),
        "transform"
    );
    assert_eq!(reopened.anchors()[0].fallback_xml(), original_fallback);
}

#[test]
fn exact_scalar_noop_reuses_source_and_has_an_empty_forward_and_inverse_patch() {
    let (mut package, slide_name, _action_name) = package_with_action(false);
    let source = action_snapshot(&package, &slide_name);
    let selector = source.selector(0).unwrap();
    let source_pointer = source.source_xml().as_ptr();
    let mut edit = source.edit();
    assert!(
        !edit
            .edit_profile(selector, |profile| {
                profile.set_action_type(ActionSelector::ordinal(0), ActionType::Add)?;
                Ok(())
            })
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml().as_ptr(), source_pointer);
    ink_actions::apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(
        action_snapshot(&package, &slide_name).source_xml(),
        source.source_xml()
    );
    assert!(commit.patch().inverse().is_empty());
}

#[test]
fn changed_patch_and_inverse_restore_exact_target_bytes() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let source = action_snapshot(&package, &slide_name);
    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_property_value(
            ChildSelector::Property {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "forward",
        )?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let original_target = package.get_part(&action_name).unwrap().blob().to_vec();
    ink_actions::apply_patch(&mut package, commit.patch()).unwrap();
    let changed_target = package.get_part(&action_name).unwrap().blob().to_vec();
    assert_ne!(changed_target, original_target);
    ink_actions::apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        original_target
    );
    assert_eq!(
        action_snapshot(&package, &slide_name).source_xml(),
        source.source_xml()
    );
}

#[test]
fn signed_package_root_edge_allows_exact_noop_and_requires_unsign_reload_for_changes() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
    assert!(
        package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
    );

    // A no-op patch is source-checked and publishes the existing package
    // unchanged, including its package-root signing edge.
    let source = action_snapshot(&package, &slide_name);
    let selector = source.selector(0).unwrap();
    let mut noop_edit = source.edit();
    assert!(
        !noop_edit
            .edit_profile(selector, |profile| {
                profile.set_action_type(ActionSelector::ordinal(0), ActionType::Add)?;
                Ok(())
            })
            .unwrap()
    );
    let noop = noop_edit.commit().unwrap();
    assert!(!noop.is_changed());
    let signed_before = serialized_opc(&package);
    ink_actions::apply_patch(&mut package, noop.patch()).unwrap();
    assert!(package.is_signed());
    assert_eq!(serialized_opc(&package), signed_before);

    // A changed patch is refused before publication and leaves every owned
    // byte, including the signature relationship, untouched.
    let (_signed_source, signed_commit) = changed_profile_commit(&package, &slide_name);
    let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
    let before_package = serialized_opc(&package);
    let error = ink_actions::apply_patch(&mut package, signed_commit.patch()).unwrap_err();
    assert!(matches!(
        error,
        litchi_pptx::Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(package.is_signed());
    assert_eq!(serialized_opc(&package), before_package);
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );

    // Unsigning changes the package-root relationship source token.  The
    // pre-unsign patch must therefore be stale; callers reload and recreate
    // the edit before publishing under the new graph.
    package.unsign();
    assert!(!package.is_signed());
    assert!(
        !package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
    );
    let before_stale_target = package.get_part(&action_name).unwrap().blob().to_vec();
    let stale_error = ink_actions::apply_patch(&mut package, signed_commit.patch()).unwrap_err();
    assert!(matches!(stale_error, litchi_pptx::Error::StaleSource));
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_stale_target
    );

    let reloaded = action_snapshot(&package, &slide_name);
    let selector = reloaded.selector(0).unwrap();
    let mut edit = reloaded.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_property_value(
            ChildSelector::Property {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "after-unsign",
        )?;
        Ok(())
    })
    .unwrap();
    let reloaded_commit = edit.commit().unwrap();
    let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
    ink_actions::apply_patch(&mut package, reloaded_commit.patch()).unwrap();
    assert!(!package.is_signed());
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
    assert_ne!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
    let reopened = action_snapshot(&package, &slide_name);
    assert_eq!(
        reopened.anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .properties()[0]
            .value(),
        "after-unsign"
    );
}

#[test]
fn shared_target_updates_all_owners_and_retains_the_complete_inbound_graph() {
    let (mut package, slide_name, action_name) =
        package_with_actions(false, &[ACTION_REL, "rIdAction2"]);
    let source = action_snapshot(&package, &slide_name);
    assert_eq!(source.anchors().len(), 2);
    assert_eq!(source.anchors()[0].target_part_name(), &action_name);
    assert_eq!(source.anchors()[1].target_part_name(), &action_name);
    assert_eq!(source.anchors()[0].inbound_references().len(), 2);
    assert_eq!(source.anchors()[1].inbound_references().len(), 2);
    assert_eq!(
        source.anchors()[0].target_bytes(),
        source.anchors()[1].target_bytes()
    );

    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_property_value(
            ChildSelector::Property {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "shared-target-edit",
        )?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    assert_eq!(
        commit.snapshot().anchors()[0].target_bytes(),
        commit.snapshot().anchors()[1].target_bytes(),
        "one target replacement must update every owner in the snapshot"
    );
    assert_eq!(commit.snapshot().anchors()[0].inbound_references().len(), 2);

    ink_actions::apply_patch(&mut package, commit.patch()).unwrap();
    let reopened = action_snapshot(&package, &slide_name);
    assert_eq!(reopened.anchors().len(), 2);
    assert_eq!(
        reopened.anchors()[0].target_bytes(),
        reopened.anchors()[1].target_bytes()
    );
    assert_eq!(reopened.anchors()[0].inbound_references().len(), 2);
    assert_eq!(reopened.anchors()[1].inbound_references().len(), 2);
    assert_eq!(
        reopened.anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .properties()[0]
            .value(),
        "shared-target-edit"
    );
    assert_eq!(
        reopened.anchors()[1]
            .profile()
            .actions()
            .next()
            .unwrap()
            .properties()[0]
            .value(),
        "shared-target-edit"
    );
}

#[test]
fn unknown_outbound_relationship_is_a_diagnostic_preserved_by_scalar_edits() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    package
        .get_part_mut(&action_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:vendor:opaque-action-dependency".to_owned(),
            "https://example.invalid/action-dependency.xml".to_owned(),
            "rIdOpaque".to_owned(),
            true,
        );

    let source = action_snapshot(&package, &slide_name);
    let outbound = source.anchors()[0].outbound_references();
    assert_eq!(outbound.len(), 1);
    assert_eq!(outbound[0].relationship_id(), "rIdOpaque");
    assert_eq!(
        outbound[0].relationship_type(),
        "urn:vendor:opaque-action-dependency"
    );
    assert_eq!(
        outbound[0].target_ref(),
        "https://example.invalid/action-dependency.xml"
    );
    assert_eq!(outbound[0].target_mode(), litchi_opc::TargetMode::External);

    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_start_time(ActionSelector::ordinal(0), "9")?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    ink_actions::apply_patch(&mut package, commit.patch()).unwrap();

    let reopened = action_snapshot(&package, &slide_name);
    assert_eq!(reopened.anchors()[0].outbound_references(), outbound);
    assert_eq!(
        package
            .get_part(&action_name)
            .unwrap()
            .rels()
            .get("rIdOpaque")
            .unwrap()
            .target_ref(),
        "https://example.invalid/action-dependency.xml"
    );
}

#[test]
fn known_ecma_outbound_relationship_refuses_typed_action_publication_atomically() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let image_name = PackURI::new("/ppt/media/image1.png").unwrap();
    package.add_part(Box::new(BlobPart::new(
        image_name,
        "image/png".to_owned(),
        vec![0x89, b'P', b'N', b'G'],
    )));
    package
        .get_part_mut(&action_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::IMAGE.to_owned(),
            "../media/image1.png".to_owned(),
            "rIdKnownImage".to_owned(),
            false,
        );
    let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();

    let error = load_error(&package, &slide_name);
    assert!(
        matches!(error, litchi_pptx::Error::Relationship(_)),
        "known ECMA outbound edge must be a typed conformance refusal: {error:?}"
    );
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
}

#[test]
fn package_limits_accept_exact_and_one_under_caps_and_reject_one_over_atomically() {
    let payload_len = action_payload().len();
    assert!(payload_len > 1);

    let exact = ink_actions::Limits::new(1, payload_len, payload_len, 1).unwrap();
    let (package, slide_name, _action_name) = package_with_action(false);
    assert_eq!(
        action_snapshot_with_limits(&package, &slide_name, exact)
            .unwrap()
            .anchors()
            .len(),
        1
    );

    let one_under_target = ink_actions::Limits::new(1, payload_len + 1, payload_len, 1).unwrap();
    let (package, slide_name, _action_name) = package_with_action(false);
    assert_eq!(
        action_snapshot_with_limits(&package, &slide_name, one_under_target)
            .unwrap()
            .anchors()
            .len(),
        1
    );

    let one_under_aggregate = ink_actions::Limits::new(1, payload_len, payload_len + 1, 1).unwrap();
    let (package, slide_name, _action_name) = package_with_action(false);
    assert_eq!(
        action_snapshot_with_limits(&package, &slide_name, one_under_aggregate)
            .unwrap()
            .anchors()
            .len(),
        1
    );

    let one_over_target = ink_actions::Limits::new(1, payload_len - 1, payload_len - 1, 1).unwrap();
    let (package, slide_name, action_name) = package_with_action(false);
    let error = action_snapshot_with_limits(&package, &slide_name, one_over_target)
        .expect_err("target and aggregate caps must be checked independently");
    assert!(matches!(
        error,
        litchi_pptx::Error::Limit {
            resource: "ink-action target bytes",
            ..
        }
    ));
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        action_payload()
    );

    let one_over_aggregate =
        ink_actions::Limits::new(1, payload_len + 1, payload_len - 1, 1).unwrap();
    let (package, slide_name, action_name) = package_with_action(false);
    let error = action_snapshot_with_limits(&package, &slide_name, one_over_aggregate)
        .expect_err("one byte over the aggregate cap must refuse");
    assert!(matches!(
        error,
        litchi_pptx::Error::Limit {
            resource: "ink-action aggregate target bytes",
            limit: payload_limit
        } if payload_limit == payload_len - 1
    ));
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        action_payload()
    );

    let (package, slide_name, _action_name) =
        package_with_actions(false, &[ACTION_REL, "rIdAction2"]);
    let exact_shared = ink_actions::Limits::new(2, payload_len, payload_len, 2).unwrap();
    assert_eq!(
        action_snapshot_with_limits(&package, &slide_name, exact_shared)
            .unwrap()
            .anchors()
            .len(),
        2
    );

    let (package, slide_name, action_name) =
        package_with_actions(false, &[ACTION_REL, "rIdAction2"]);
    let one_over_anchor = ink_actions::Limits::new(1, payload_len, payload_len, 2).unwrap();
    let error = action_snapshot_with_limits(&package, &slide_name, one_over_anchor)
        .expect_err("one anchor over the cap must refuse");
    assert!(matches!(
        error,
        litchi_pptx::Error::Limit {
            resource: "ink-action anchor count",
            limit: 1
        }
    ));
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        action_payload()
    );

    let (package, slide_name, action_name) =
        package_with_actions(false, &[ACTION_REL, "rIdAction2"]);
    let one_over_relationship = ink_actions::Limits::new(2, payload_len, payload_len, 1).unwrap();
    let error = action_snapshot_with_limits(&package, &slide_name, one_over_relationship)
        .expect_err("one inbound edge over the cap must refuse");
    assert!(matches!(
        error,
        litchi_pptx::Error::Limit {
            resource: "ink-action relationship edges",
            limit: 1
        }
    ));
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        action_payload()
    );
}

#[test]
fn caller_limits_above_each_owner_hard_ceiling_are_rejected() {
    let hard = ink_actions::default_limits();
    let cases = [
        (
            hard.anchors.saturating_add(1),
            hard.target_bytes,
            hard.total_target_bytes,
            hard.target_relationships,
            "anchors",
        ),
        (
            hard.anchors,
            hard.target_bytes.saturating_add(1),
            hard.total_target_bytes,
            hard.target_relationships,
            "target bytes",
        ),
        (
            hard.anchors,
            hard.target_bytes,
            hard.total_target_bytes.saturating_add(1),
            hard.target_relationships,
            "aggregate target bytes",
        ),
        (
            hard.anchors,
            hard.target_bytes,
            hard.total_target_bytes,
            hard.target_relationships.saturating_add(1),
            "target relationships",
        ),
    ];
    for (anchors, target_bytes, total_target_bytes, target_relationships, resource) in cases {
        assert!(
            ink_actions::Limits::new(
                anchors,
                target_bytes,
                total_target_bytes,
                target_relationships,
            )
            .is_none(),
            "caller-supplied {resource} ceiling exceeded the hard owner bound",
        );
    }
}

#[test]
fn unrelated_owner_relationships_do_not_consume_selected_target_edge_cap() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let slide = package.get_part_mut(&slide_name).unwrap();
    slide.rels_mut().add_relationship(
        "urn:vendor:unrelated-one".to_owned(),
        "https://example.invalid/one".to_owned(),
        "rIdUnrelatedOne".to_owned(),
        true,
    );
    slide.rels_mut().add_relationship(
        "urn:vendor:unrelated-two".to_owned(),
        "https://example.invalid/two".to_owned(),
        "rIdUnrelatedTwo".to_owned(),
        true,
    );

    let payload_len = action_payload().len();
    let limits = ink_actions::Limits::new(1, payload_len, payload_len, 1).unwrap();
    let snapshot = action_snapshot_with_limits(&package, &slide_name, limits)
        .expect("selected target closure is within the edge cap");
    assert_eq!(snapshot.anchors().len(), 1);
    assert_eq!(snapshot.anchors()[0].target_part_name(), &action_name);
    assert_eq!(snapshot.anchors()[0].inbound_references().len(), 1);
    assert_eq!(snapshot.owner_relationships().len(), 3);
}

fn changed_profile_commit(
    package: &OpcPackage,
    slide_name: &PackURI,
) -> (ink_actions::Snapshot, ink_actions::Commit) {
    let source = action_snapshot(package, slide_name);
    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_property_value(
            ChildSelector::Property {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "stale-check",
        )?;
        Ok(())
    })
    .unwrap();
    (source, edit.commit().unwrap())
}

#[test]
fn patch_read_sets_reject_owner_relationship_target_and_content_type_staleness_atomically() {
    // The changed target is prepared once per case so every mutation is
    // compared with the same source-bound patch contract.
    {
        let (mut package, slide_name, _action_name) = package_with_action(false);
        let (source, commit) = changed_profile_commit(&package, &slide_name);
        let xml = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_owner = xml.clone();
        let updated = String::from_utf8(xml)
            .unwrap()
            .replace(
                r#"compat:Ignorable="p14main inkAction""#,
                r#"compat:Ignorable="p14main  inkAction""#,
            )
            .into_bytes();
        package.get_part_mut(&slide_name).unwrap().set_blob(updated);
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(error, litchi_pptx::Error::StaleSource));
        assert_ne!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(source.source_xml(), commit.snapshot().source_xml());
    }

    {
        let (mut package, slide_name, _action_name) = package_with_action(false);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let alternate_name = PackURI::new("/ppt/custom/action-alt.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            alternate_name.clone(),
            "text/xml".to_owned(),
            action_payload(),
        )));
        let slide = package.get_part_mut(&slide_name).unwrap();
        slide.rels_mut().remove(ACTION_REL);
        slide.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "../custom/action-alt.xml".to_owned(),
            ACTION_REL.to_owned(),
            false,
        );
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(error, litchi_pptx::Error::StaleSource));
    }

    {
        let (mut package, slide_name, action_name) = package_with_action(false);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        let slide = package.get_part_mut(&slide_name).unwrap();
        slide.rels_mut().remove(ACTION_REL);
        slide.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "https://example.invalid/action.xml".to_owned(),
            ACTION_REL.to_owned(),
            true,
        );
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(error, litchi_pptx::Error::Relationship(_)));
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }

    {
        let (mut package, slide_name, action_name) = package_with_action(false);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let before = package.get_part(&action_name).unwrap().blob().to_vec();
        let changed = String::from_utf8(before.clone())
            .unwrap()
            .replace(r#"value="old""#, r#"value="other""#)
            .into_bytes();
        package
            .get_part_mut(&action_name)
            .unwrap()
            .set_blob(changed);
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(error, litchi_pptx::Error::StaleSource));
        assert_ne!(package.get_part(&action_name).unwrap().blob(), before);
    }

    {
        let (mut package, slide_name, action_name) = package_with_action(false);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        package
            .get_part_mut(&action_name)
            .unwrap()
            .set_content_type("application/xml".to_owned())
            .unwrap();
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(error, litchi_pptx::Error::ContentType { .. }));
        assert_eq!(
            package.get_part(&action_name).unwrap().content_type(),
            "application/xml"
        );
    }
}

#[test]
fn selector_context_is_checked_against_the_new_snapshot() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let source = action_snapshot(&package, &slide_name);
    let selector = source.selector(0).unwrap();
    let target = package.get_part(&action_name).unwrap().blob().to_vec();
    package.get_part_mut(&action_name).unwrap().set_blob(
        String::from_utf8(target)
            .unwrap()
            .replace(r#"value="old""#, r#"value="new""#)
            .into_bytes(),
    );
    let current = action_snapshot(&package, &slide_name);
    let mut edit = current.edit();
    let error = edit
        .edit_profile(selector, |_profile| Ok(()))
        .expect_err("old selector must be stale");
    assert!(matches!(error, litchi_pptx::Error::StaleSource));
}

#[test]
fn selector_context_includes_changed_relationship_graph_diagnostics() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let source = action_snapshot(&package, &slide_name);
    let selector = source.selector(0).unwrap();
    let source_revision = source.revision();

    package
        .get_part_mut(&action_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:vendor:selector-context".to_owned(),
            "https://example.invalid/context.xml".to_owned(),
            "rIdContext".to_owned(),
            true,
        );
    let current = action_snapshot(&package, &slide_name);
    assert_ne!(
        current.revision(),
        source_revision,
        "relationship diagnostics belong to selector context"
    );
    let mut edit = current.edit();
    let error = edit
        .edit_profile(selector, |_profile| Ok(()))
        .expect_err("a selector from the pre-graph-change snapshot must be stale");
    assert!(matches!(error, litchi_pptx::Error::StaleSource));
}

fn load_error(package: &OpcPackage, slide_name: &PackURI) -> litchi_pptx::Error {
    let slide = package.get_part(slide_name).unwrap();
    ink_actions::load_snapshot(package, 0, slide, &ink_actions::Limits::default()).unwrap_err()
}

#[test]
fn malformed_inbound_relationship_is_a_typed_closure_error_and_is_not_dropped() {
    let (mut package, slide_name, action_name) = package_with_action(false);
    let inbound_name = PackURI::new("/ppt/other-owner.xml").unwrap();
    let mut inbound = BlobPart::new(
        inbound_name,
        ct::PML_SLIDE.to_owned(),
        b"<opaque/>".to_vec(),
    );
    inbound.rels_mut().add_relationship(
        rt::CUSTOM_XML.to_owned(),
        "custom/action.xml%ZZ".to_owned(),
        "rIdMalformedInbound".to_owned(),
        false,
    );
    package.add_part(Box::new(inbound));
    let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
    let before_target = package.get_part(&action_name).unwrap().blob().to_vec();

    let error = load_error(&package, &slide_name);
    assert!(
        matches!(error, litchi_pptx::Error::Relationship(_)),
        "malformed inbound edge must not disappear from the closure: {error:?}"
    );
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
    assert_eq!(
        package.get_part(&action_name).unwrap().blob(),
        before_target
    );
}

#[test]
fn unsupported_and_ambiguous_mce_shapes_remain_opaque_or_fail_closed() {
    let variants: &[(&str, fn(String) -> String, bool)] = &[
        (
            "direct-content-part",
            |xml: String| {
                xml.replace(
                    &action_owner_anchor("pp", "rr", "compat", "p14main inkAction", ACTION_REL),
                    r#"<pp:contentPart rr:id="rIdAction"/>"#,
                )
            },
            false,
        ),
        (
            "generic-content-part-choice",
            |xml: String| xml.replace("Requires=\"p14main inkAction\"", "Requires=\"p14main\""),
            false,
        ),
        (
            "unknown-required-namespace",
            |xml: String| {
                xml.replace(
                    "Requires=\"p14main inkAction\"",
                    "Requires=\"p14main inkAction future\"",
                )
                .replace(
                    "xmlns:inkAction=\"http://schemas.microsoft.com/office/powerpoint/2014/inkAction\"",
                    "xmlns:inkAction=\"http://schemas.microsoft.com/office/powerpoint/2014/inkAction\" xmlns:future=\"urn:future\"",
                )
            },
            false,
        ),
        (
            "non-picture-fallback",
            |xml: String| {
                xml.replace("<compat:Fallback><pp:pic>", "<compat:Fallback><pp:sp>")
                    .replace("</pp:pic></compat:Fallback>", "</pp:sp></compat:Fallback>")
            },
            true,
        ),
        (
            "duplicate-supported-choice",
            |xml: String| {
                let choice = r#"<compat:Choice Requires="p14main inkAction"><pp:contentPart rr:id="rIdAction"/></compat:Choice>"#;
                xml.replace(choice, &format!("{choice}{choice}"))
            },
            true,
        ),
        (
            "malformed-mce",
            |xml: String| xml.replace("</compat:AlternateContent>", ""),
            true,
        ),
    ];

    for (label, mutate, rejected) in variants {
        let (mut package, slide_name, _action_name) = package_with_action(false);
        let original = package.get_part(&slide_name).unwrap().blob().to_vec();
        let updated = mutate(String::from_utf8(original.clone()).unwrap()).into_bytes();
        package.get_part_mut(&slide_name).unwrap().set_blob(updated);
        if *rejected {
            let error = load_error(&package, &slide_name);
            assert!(
                matches!(
                    error,
                    litchi_pptx::Error::Invalid(_) | litchi_pptx::Error::Xml(_)
                ),
                "{label}: expected typed MCE rejection, got {error:?}"
            );
        } else {
            let snapshot = action_snapshot(&package, &slide_name);
            assert!(
                snapshot.anchors().is_empty(),
                "{label}: branch was promoted"
            );
            if *label == "unknown-required-namespace" {
                assert_eq!(snapshot.source_anchors().len(), 1);
                let raw = snapshot.source_anchors()[0].xml();
                assert!(raw.starts_with(b"<compat:AlternateContent>"));
                assert!(raw.ends_with(b"</compat:AlternateContent>"));
                let selector = snapshot
                    .source_selector(0)
                    .expect("opaque action-capable MCE source selector");
                let mut edit = snapshot.edit();
                let error = edit
                    .edit_profile(selector, |_profile| Ok(()))
                    .expect_err("opaque action-capable MCE must not become typed");
                assert!(matches!(error, litchi_pptx::Error::Invalid(_)));
            }
        }
        assert_ne!(package.get_part(&slide_name).unwrap().blob(), original);
    }
}

#[test]
fn malformed_anchor_and_relationship_closures_fail_without_repair() {
    type RefusalCase = (
        &'static str,
        fn(&mut OpcPackage, &PackURI),
        fn(litchi_pptx::Error) -> bool,
    );
    let variants: &[RefusalCase] = &[
        (
            "missing-rid",
            |package: &mut OpcPackage, slide_name: &PackURI| {
                let xml = String::from_utf8(package.get_part(slide_name).unwrap().blob().to_vec())
                    .unwrap()
                    .replace(r#"rr:id="rIdAction""#, "");
                package
                    .get_part_mut(slide_name)
                    .unwrap()
                    .set_blob(xml.into_bytes());
            },
            |error: litchi_pptx::Error| matches!(error, litchi_pptx::Error::Invalid(_)),
        ),
        (
            "duplicate-rid",
            |package: &mut OpcPackage, slide_name: &PackURI| {
                let xml = String::from_utf8(package.get_part(slide_name).unwrap().blob().to_vec())
                    .unwrap()
                    .replace(
                        r#"rr:id="rIdAction""#,
                        r#"rr:id="rIdAction" xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r2:id="rIdAction""#,
                    );
                package
                    .get_part_mut(slide_name)
                    .unwrap()
                    .set_blob(xml.into_bytes());
            },
            |error: litchi_pptx::Error| matches!(error, litchi_pptx::Error::Invalid(_)),
        ),
        (
            "wrong-relationship",
            |package: &mut OpcPackage, slide_name: &PackURI| {
                package
                    .get_part_mut(slide_name)
                    .unwrap()
                    .rels_mut()
                    .remove(ACTION_REL);
                package
                    .get_part_mut(slide_name)
                    .unwrap()
                    .rels_mut()
                    .add_relationship(
                        "urn:wrong".to_owned(),
                        ACTION_TARGET.to_owned(),
                        ACTION_REL.to_owned(),
                        false,
                    );
            },
            |error: litchi_pptx::Error| matches!(error, litchi_pptx::Error::Relationship(_)),
        ),
        (
            "external-target",
            |package: &mut OpcPackage, slide_name: &PackURI| {
                package
                    .get_part_mut(slide_name)
                    .unwrap()
                    .rels_mut()
                    .remove(ACTION_REL);
                package
                    .get_part_mut(slide_name)
                    .unwrap()
                    .rels_mut()
                    .add_relationship(
                        rt::CUSTOM_XML.to_owned(),
                        "https://example.invalid/action.xml".to_owned(),
                        ACTION_REL.to_owned(),
                        true,
                    );
            },
            |error: litchi_pptx::Error| matches!(error, litchi_pptx::Error::Relationship(_)),
        ),
        (
            "missing-target",
            |package: &mut OpcPackage, _slide_name: &PackURI| {
                package.remove_part(&PackURI::new(ACTION_NAME).unwrap());
            },
            |error: litchi_pptx::Error| matches!(error, litchi_pptx::Error::PartNotFound(_)),
        ),
        (
            "wrong-content-type",
            |package: &mut OpcPackage, _slide_name: &PackURI| {
                package
                    .get_part_mut(&PackURI::new(ACTION_NAME).unwrap())
                    .unwrap()
                    .set_content_type("application/xml".to_owned())
                    .unwrap();
            },
            |error: litchi_pptx::Error| matches!(error, litchi_pptx::Error::ContentType { .. }),
        ),
        (
            "wrong-root",
            |package: &mut OpcPackage, _slide_name: &PackURI| {
                package
                    .get_part_mut(&PackURI::new(ACTION_NAME).unwrap())
                    .unwrap()
                    .set_blob(b"<wrong:root xmlns:wrong=\"urn:wrong\"/>".to_vec());
            },
            |error: litchi_pptx::Error| {
                matches!(
                    error,
                    litchi_pptx::Error::Invalid(_)
                        | litchi_pptx::Error::Xml(_)
                        | litchi_pptx::Error::Drawing(_)
                )
            },
        ),
    ];

    for (label, mutate, expected) in variants {
        let (mut package, slide_name, _action_name) = package_with_action(false);
        mutate(&mut package, &slide_name);
        let staged = package.get_part(&slide_name).unwrap().blob().to_vec();
        let error = load_error(&package, &slide_name);
        assert!(expected(error), "{label}: wrong closure error");
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), staged);
    }
}

fn serialized_opc(package: &OpcPackage) -> Vec<u8> {
    let mut bytes = Vec::new();
    package.to_stream(&mut bytes).unwrap();
    bytes
}

/// Reopen a synthetic package after adding harmless lexical source to selected
/// relationship members.  The decoded graph stays identical while the exact
/// retained `.rels` bytes differ, which exercises source authorization rather
/// than only semantic relationship metadata.
fn with_relationship_source_markers(bytes: &[u8], members: &[&str]) -> Vec<u8> {
    let archive = ArchiveReader::new(bytes).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let mut member = archive.read(name).unwrap();
        if members.contains(&name) {
            let marker = b"<Relationships";
            let start = member
                .windows(marker.len())
                .position(|window| window == marker)
                .expect("relationship root");
            let insertion = member[start..]
                .iter()
                .position(|byte| *byte == b'>')
                .map(|offset| start + offset + 1)
                .expect("relationship root close");
            member.splice(
                insertion..insertion,
                b"<!-- source-marker -->".iter().copied(),
            );
        }
        writer.write_deflated(name, &member).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn reopened_action_package(bytes: Vec<u8>, limits: ReadLimits) -> (OpcPackage, PackURI, PackURI) {
    (
        OpcPackage::from_vec_with_limits(bytes, limits).unwrap(),
        PackURI::new(SLIDE_NAME).unwrap(),
        PackURI::new(ACTION_NAME).unwrap(),
    )
}

fn bounded_read_limits(
    max_parts: usize,
    max_part_bytes: usize,
    max_total_part_bytes: usize,
) -> ReadLimits {
    ReadLimits::builder()
        .max_parts(max_parts)
        .unwrap()
        .max_part_bytes(max_part_bytes as u64)
        .unwrap()
        .max_total_part_bytes(max_total_part_bytes as u64)
        .unwrap()
        .build()
        .unwrap()
}

fn package_with_presentation() -> OpcPackage {
    let (mut package, _slide_name, _action_name) = package_with_action(false);
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRESENTATION_MAIN.to_owned(),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
    );
    presentation.rels_mut().add_relationship(
        rt::SLIDE.to_owned(),
        "slides/slide1.xml".to_owned(),
        "rIdSlide".to_owned(),
        false,
    );
    package.add_part(Box::new(presentation));
    package.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "ppt/presentation.xml".to_owned(),
        "rIdOfficeDocument".to_owned(),
        false,
    );
    package
}

#[test]
fn public_package_facade_reads_edits_publishes_and_reopens_ink_actions() {
    let mut package = Package::from_opc_package(package_with_presentation()).unwrap();
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 1);
    assert_eq!(presentation.slide_references().unwrap().len(), 1);

    let source = package.ink_actions().unwrap().remove(0);
    assert_eq!(source.anchors().len(), 1);
    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_start_time(ActionSelector::ordinal(0), "12.5")?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());

    let published = package
        .apply_ink_actions_patch(commit.patch())
        .expect("facade publishes a checked action patch");
    assert_eq!(published.anchors().len(), 1);
    assert_eq!(
        published.anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .start_time(),
        "12.5"
    );

    let bytes = package.to_bytes().unwrap();
    let reopened = Package::from_vec(bytes).unwrap();
    let snapshots = reopened.ink_actions().unwrap();
    assert_eq!(snapshots.len(), 1);
    assert_eq!(
        snapshots[0].anchors()[0]
            .profile()
            .actions()
            .next()
            .unwrap()
            .start_time(),
        "12.5"
    );
}

#[test]
fn public_presentation_inventory_charges_shared_targets_once_and_distinct_targets_aggregate() {
    let payload_len = action_payload().len();
    let limits = ink_actions::Limits::new(1, payload_len, payload_len, 2).unwrap();

    let shared = Package::from_opc_package(package_with_two_slide_presentation(true)).unwrap();
    let shared_snapshots = shared
        .ink_actions_with_limits(limits)
        .expect("one shared physical target is charged once across slides");
    assert_eq!(shared_snapshots.len(), 2);
    assert_eq!(
        shared_snapshots[0].anchors()[0].target_bytes(),
        shared_snapshots[1].anchors()[0].target_bytes()
    );

    let distinct = Package::from_opc_package(package_with_two_slide_presentation(false)).unwrap();
    let error = distinct
        .ink_actions_with_limits(limits)
        .expect_err("distinct targets must consume the package-wide aggregate cap");
    assert!(
        matches!(
            error,
            litchi_pptx::Error::Limit {
                resource: "ink-action aggregate target bytes",
                limit
            } if limit == payload_len
        ),
        "unexpected multi-slide aggregate error: {error:?}"
    );
}

#[test]
fn retained_opc_read_limits_cover_unrelated_candidate_parts_atomically() {
    let (seed, slide_name, action_name) = package_with_action(false);
    let seed_slide_bytes = seed.get_part(&slide_name).unwrap().blob().len();
    let seed_target_bytes = seed.get_part(&action_name).unwrap().blob().len();
    let seed_total_bytes: usize = seed
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes").blob().len())
        .sum();
    let seed_bytes = serialized_opc(&seed);

    {
        let limits = bounded_read_limits(seed.part_count(), usize::MAX / 4, usize::MAX / 4);
        let (mut package, slide_name, action_name) =
            reopened_action_package(seed_bytes.clone(), limits);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/unrelated.xml").unwrap(),
            "application/xml".to_owned(),
            vec![0xA5],
        )));
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(
            error,
            litchi_pptx::Error::Opc(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                ..
            })
        ));
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }

    let candidate_target_bytes = {
        let (package, slide_name, _action_name) = reopened_action_package(
            seed_bytes.clone(),
            bounded_read_limits(seed.part_count(), usize::MAX / 4, usize::MAX / 4),
        );
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        commit.snapshot().anchors()[0].target_bytes().len()
    };
    let candidate_total_bytes = seed_total_bytes - seed_target_bytes + candidate_target_bytes;
    let max_part_bytes = seed_slide_bytes.max(candidate_target_bytes);

    {
        let limits = bounded_read_limits(seed.part_count() + 1, max_part_bytes, usize::MAX / 4);
        let (mut package, slide_name, action_name) =
            reopened_action_package(seed_bytes.clone(), limits);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/unrelated-large.xml").unwrap(),
            "application/xml".to_owned(),
            vec![0xA5; max_part_bytes + 1],
        )));
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(
            matches!(
                error,
                litchi_pptx::Error::Opc(OpcError::ReadLimit {
                    resource: ReadResource::PartBytes,
                    ..
                })
            ),
            "unexpected per-part candidate error: {error:?}"
        );
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }

    {
        let limits = bounded_read_limits(
            seed.part_count() + 1,
            max_part_bytes.max(seed_target_bytes),
            candidate_total_bytes,
        );
        let (mut package, slide_name, action_name) = reopened_action_package(seed_bytes, limits);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/unrelated-one-byte.xml").unwrap(),
            "application/xml".to_owned(),
            vec![0xA5],
        )));
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(
            error,
            litchi_pptx::Error::Opc(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                ..
            })
        ));
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }
}

#[test]
fn retained_lower_opc_limits_are_checked_during_action_inventory() {
    let (seed, _slide_name, _action_name) = package_with_action(false);
    let seed_part_count = seed.part_count();
    let seed_max_part_bytes = seed
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes").blob().len())
        .max()
        .unwrap();
    let seed_total_bytes: usize = seed
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes").blob().len())
        .sum();
    let seed_bytes = serialized_opc(&seed);

    {
        let limits = bounded_read_limits(seed_part_count, usize::MAX / 4, usize::MAX / 4);
        let (mut package, slide_name, action_name) =
            reopened_action_package(seed_bytes.clone(), limits);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/inventory-extra.xml").unwrap(),
            "application/xml".to_owned(),
            vec![0xA5],
        )));
        let error =
            action_snapshot_with_limits(&package, &slide_name, ink_actions::Limits::default())
                .expect_err("inventory must charge an added part against retained max_parts");
        assert!(
            matches!(
                error,
                litchi_pptx::Error::Opc(OpcError::ReadLimit {
                    resource: ReadResource::Parts,
                    ..
                })
            ),
            "unexpected max-parts inventory error: {error:?}"
        );
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }

    {
        let max_part_bytes = seed_max_part_bytes;
        let limits = bounded_read_limits(seed_part_count + 1, max_part_bytes, usize::MAX / 4);
        let (mut package, slide_name, action_name) =
            reopened_action_package(seed_bytes.clone(), limits);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/inventory-large.xml").unwrap(),
            "application/xml".to_owned(),
            vec![0xA5; max_part_bytes + 1],
        )));
        let error =
            action_snapshot_with_limits(&package, &slide_name, ink_actions::Limits::default())
                .expect_err("inventory must charge an added part against retained max_part_bytes");
        assert!(
            matches!(
                error,
                litchi_pptx::Error::Opc(OpcError::ReadLimit {
                    resource: ReadResource::PartBytes,
                    ..
                })
            ),
            "unexpected per-part inventory error: {error:?}"
        );
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }

    {
        let limits = bounded_read_limits(seed_part_count + 1, usize::MAX / 4, seed_total_bytes);
        let (mut package, slide_name, action_name) = reopened_action_package(seed_bytes, limits);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/inventory-one-byte.xml").unwrap(),
            "application/xml".to_owned(),
            vec![0xA5],
        )));
        let error =
            action_snapshot_with_limits(&package, &slide_name, ink_actions::Limits::default())
                .expect_err("inventory must charge an added part against retained total bytes");
        assert!(
            matches!(
                error,
                litchi_pptx::Error::Opc(OpcError::ReadLimit {
                    resource: ReadResource::TotalPartBytes,
                    ..
                })
            ),
            "unexpected aggregate inventory error: {error:?}"
        );
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }
}

#[test]
fn package_root_incoming_edge_is_retained_through_scalar_edit_and_publish() {
    let (mut package, slide_name, _action_name) = package_with_action(false);
    package.rels_mut().add_relationship(
        rt::CUSTOM_XML.to_owned(),
        "ppt/custom/action.xml".to_owned(),
        "rIdRootAction".to_owned(),
        false,
    );
    let source = action_snapshot(&package, &slide_name);
    let root = PackURI::new("/").unwrap();
    assert_eq!(source.anchors()[0].inbound_references().len(), 2);
    assert!(
        source.anchors()[0]
            .inbound_references()
            .iter()
            .any(|reference| reference.source_part_name() == &root
                && reference.relationship_id() == "rIdRootAction")
    );

    let (_source, commit) = changed_profile_commit(&package, &slide_name);
    ink_actions::apply_patch(&mut package, commit.patch()).unwrap();
    let reopened = action_snapshot(&package, &slide_name);
    assert_eq!(reopened.anchors()[0].inbound_references().len(), 2);
    assert!(
        reopened.anchors()[0]
            .inbound_references()
            .iter()
            .any(|reference| reference.source_part_name() == &root
                && reference.relationship_id() == "rIdRootAction")
    );
    assert_eq!(
        package.rels().get("rIdRootAction").unwrap().target_ref(),
        "ppt/custom/action.xml"
    );
}

#[test]
fn case_equivalent_shared_target_refs_update_one_physical_target_and_all_owners() {
    let (mut package, slide_name, action_name) = package_with_action_refs(
        false,
        &[ACTION_REL, "rIdAction2"],
        &[ACTION_TARGET, "../custom/ACTION.XML"],
    );
    let source = action_snapshot(&package, &slide_name);
    assert_eq!(source.anchors().len(), 2);
    assert_eq!(source.anchors()[0].target_part_name(), &action_name);
    assert_eq!(source.anchors()[1].target_part_name(), &action_name);
    assert_eq!(source.anchors()[0].inbound_references().len(), 2);
    assert_eq!(source.anchors()[1].inbound_references().len(), 2);

    let (_source, commit) = changed_profile_commit(&package, &slide_name);
    ink_actions::apply_patch(&mut package, commit.patch()).unwrap();
    let reopened = action_snapshot(&package, &slide_name);
    assert_eq!(reopened.anchors().len(), 2);
    assert_eq!(reopened.anchors()[0].target_part_name(), &action_name);
    assert_eq!(reopened.anchors()[1].target_part_name(), &action_name);
    assert_eq!(
        reopened.anchors()[0].target_bytes(),
        reopened.anchors()[1].target_bytes()
    );
    for anchor in reopened.anchors() {
        assert_eq!(
            anchor.profile().actions().next().unwrap().properties()[0].value(),
            "stale-check"
        );
    }
    assert_eq!(
        package
            .get_part(&slide_name)
            .unwrap()
            .rels()
            .get("rIdAction2")
            .unwrap()
            .target_ref(),
        "../custom/ACTION.XML"
    );
}

#[test]
fn owner_relationship_metadata_and_read_policy_are_stale_patch_inputs() {
    {
        let (mut package, slide_name, action_name) = package_with_action(false);
        let (_source, commit) = changed_profile_commit(&package, &slide_name);
        let before_owner = package.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = package.get_part(&action_name).unwrap().blob().to_vec();
        package
            .get_part_mut(&slide_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:vendor:owner-readset".to_owned(),
                "https://example.invalid/owner-readset".to_owned(),
                "rIdOwnerReadset".to_owned(),
                true,
            );
        let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
        assert!(matches!(error, litchi_pptx::Error::StaleSource));
        assert_eq!(package.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            package.get_part(&action_name).unwrap().blob(),
            before_target
        );
    }

    let seed = package_with_action(false).0;
    let bytes = serialized_opc(&seed);
    let limits_a = bounded_read_limits(seed.part_count(), usize::MAX / 4, usize::MAX / 4);
    let limits_b = bounded_read_limits(seed.part_count() + 1, usize::MAX / 4, usize::MAX / 4);
    let (package_a, slide_name, _action_name) = reopened_action_package(bytes.clone(), limits_a);
    let (mut package_b, _slide_name, _action_name) = reopened_action_package(bytes, limits_b);
    let source = action_snapshot(&package_a, &slide_name);
    assert_eq!(source.read_limits(), limits_a);
    let selector = source.selector(0).unwrap();
    let mut edit = source.edit();
    edit.edit_profile(selector, |profile| {
        profile.set_start_time(ActionSelector::ordinal(0), "12.75")?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let error = ink_actions::apply_patch(&mut package_b, commit.patch()).unwrap_err();
    assert!(matches!(error, litchi_pptx::Error::StaleSource));
}

#[test]
fn exact_root_and_owner_relationship_source_members_are_stale_patch_inputs() {
    let seed = package_with_presentation();
    let bytes = serialized_opc(&seed);
    let base = OpcPackage::from_vec(bytes.clone()).unwrap();
    let slide_name = PackURI::new(SLIDE_NAME).unwrap();
    let root = PackURI::new("/").unwrap();
    let owner_rels_member = "ppt/slides/_rels/slide1.xml.rels";
    let root_rels_member = "_rels/.rels";

    let base_owner_source = base.source_relationships(&slide_name).unwrap();
    let base_root_source = base.source_relationships(&root).unwrap();
    assert!(base_owner_source.member_present());
    assert!(base_root_source.member_present());
    let base_snapshot = action_snapshot(&base, &slide_name);
    assert_eq!(
        base_snapshot.owner_relationship_source(),
        base_owner_source.bytes()
    );
    assert_eq!(
        base_snapshot.package_relationship_source(),
        base_root_source.bytes()
    );
    assert!(base_snapshot.owner_relationship_member_present());
    assert!(base_snapshot.package_relationship_member_present());

    for member in [owner_rels_member, root_rels_member] {
        let mutated_bytes = with_relationship_source_markers(&bytes, &[member]);
        let mut mutated = OpcPackage::from_vec(mutated_bytes).unwrap();
        let changed_source = if member == owner_rels_member {
            mutated.source_relationships(&slide_name).unwrap()
        } else {
            mutated.source_relationships(&root).unwrap()
        };
        let original_source = if member == owner_rels_member {
            &base_owner_source
        } else {
            &base_root_source
        };
        assert_eq!(
            changed_source.member_present(),
            original_source.member_present()
        );
        assert_ne!(changed_source.bytes(), original_source.bytes());

        let (_source, commit) = changed_profile_commit(&base, &slide_name);
        let before_owner = mutated.get_part(&slide_name).unwrap().blob().to_vec();
        let before_target = mutated
            .get_part(&PackURI::new(ACTION_NAME).unwrap())
            .unwrap()
            .blob()
            .to_vec();
        let error = ink_actions::apply_patch(&mut mutated, commit.patch()).unwrap_err();
        assert!(
            matches!(error, litchi_pptx::Error::StaleSource),
            "{member} lexical source change must be stale: {error:?}"
        );
        assert_eq!(mutated.get_part(&slide_name).unwrap().blob(), before_owner);
        assert_eq!(
            mutated
                .get_part(&PackURI::new(ACTION_NAME).unwrap())
                .unwrap()
                .blob(),
            before_target
        );
    }
}

#[test]
fn exact_root_relationship_readset_rejects_a_boundary_free_hash_collision() {
    let (mut package, slide_name, _action_name) = package_with_action(false);
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/beta").unwrap(),
        "application/octet-stream".to_owned(),
        vec![0x01],
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/eta").unwrap(),
        "application/octet-stream".to_owned(),
        vec![0x02],
    )));
    package.rels_mut().add_relationship(
        "urn:alpha".to_owned(),
        "beta".to_owned(),
        "rIdRootCollision".to_owned(),
        false,
    );
    let (_source, commit) = changed_profile_commit(&package, &slide_name);

    package.rels_mut().remove("rIdRootCollision");
    package.rels_mut().add_relationship(
        "urn:alphab".to_owned(),
        "eta".to_owned(),
        "rIdRootCollision".to_owned(),
        false,
    );
    let error = ink_actions::apply_patch(&mut package, commit.patch()).unwrap_err();
    assert!(matches!(error, litchi_pptx::Error::StaleSource));
}

#[test]
fn changed_profile_over_caller_target_cap_refuses_before_publication() {
    let payload = action_payload();
    let target_cap = payload.len() + 16;
    let limits = ink_actions::Limits::new(1, target_cap, target_cap, 1).unwrap();
    let (package, slide_name, action_name) = package_with_action(false);
    let source = action_snapshot_with_limits(&package, &slide_name, limits).unwrap();
    let selector = source.selector(0).unwrap();
    let before = package.get_part(&action_name).unwrap().blob().to_vec();
    let oversized_value = "x".repeat(128);
    let oversized = String::from_utf8(payload)
        .unwrap()
        .replace(r#"value="old""#, &format!(r#"value="{oversized_value}""#))
        .into_bytes();
    let mut edit = source.edit();
    let error = edit
        .replace_profile_bytes(selector, oversized)
        .expect_err("caller target cap must apply before a replacement is published");
    assert!(
        matches!(
            error,
            litchi_pptx::Error::Limit {
                resource: "ink-action target bytes",
                limit
            } if limit == target_cap
        ),
        "unexpected target-cap error: {error:?}"
    );
    assert_eq!(package.get_part(&action_name).unwrap().blob(), before);
}

#[test]
fn oversized_profile_bytes_hit_the_target_cap_before_profile_parse_or_staging() {
    let payload_len = action_payload().len();
    let target_cap = payload_len + 16;
    let limits = ink_actions::Limits::new(1, target_cap, target_cap, 1).unwrap();
    let (package, slide_name, action_name) = package_with_action(false);
    let source = action_snapshot_with_limits(&package, &slide_name, limits).unwrap();
    let selector = source.selector(0).unwrap();
    let before = package.get_part(&action_name).unwrap().blob().to_vec();
    let mut edit = source.edit();
    let error = edit
        .replace_profile_bytes(selector, vec![b'x'; target_cap + 1])
        .expect_err("the owner cap must precede shared-profile parsing");
    assert!(
        matches!(
            error,
            litchi_pptx::Error::Limit {
                resource: "ink-action target bytes",
                limit
            } if limit == target_cap
        ),
        "unexpected early target-cap error: {error:?}"
    );
    assert!(
        !edit.is_changed(),
        "a refused replacement must not be staged"
    );
    assert_eq!(package.get_part(&action_name).unwrap().blob(), before);
}
