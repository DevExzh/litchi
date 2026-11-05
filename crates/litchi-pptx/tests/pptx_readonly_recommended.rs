#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test fixtures and assertions panic on failure by design"
)]

use std::fmt::Write as _;

mod in_place;
use in_place::InPlace as _;

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::{Error, OpenOptions, Package, SourceLimits};

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const P1710_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2017/10/main";
const URI: &str = "{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}";

fn package_with_properties(xml: &[u8]) -> Package {
    let authored = Package::new().expect("new package");
    let bytes = authored.to_bytes().expect("serialize package");
    let mut opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).expect("reopen package");
    let properties = PackURI::new("/ppt/presProps.xml").expect("properties part URI");
    opc.get_part_mut(&properties)
        .expect("properties part")
        .set_blob(xml.to_vec());
    let presentation = PackURI::new("/ppt/presentation.xml").expect("presentation URI");
    let presentation_part = opc.get_part_mut(&presentation).expect("presentation part");
    if !presentation_part
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == rt::PRES_PROPS)
    {
        presentation_part.rels_mut().add_relationship(
            rt::PRES_PROPS.to_owned(),
            "presProps.xml".to_owned(),
            "rIdReadonlyRecommended".to_owned(),
            false,
        );
    }
    Package::from_opc_package(opc).expect("open package")
}

fn xml_event_count(xml: &[u8]) -> usize {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut events = 0;
    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Eof) => return events,
            Ok(_) => events += 1,
            Err(error) => panic!("fixture XML must parse: {error}"),
        }
    }
}

fn max_non_properties_events(package: &Package, properties: &PackURI) -> usize {
    package
        .opc()
        .unwrap()
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .filter(|part| part.partname() != properties)
        .map(|part| xml_event_count(part.blob()))
        .max()
        .unwrap_or(0)
}

#[test]
fn package_creation_round_trip_and_owner_retention() {
    let mut package = Package::new().expect("new package");
    assert_eq!(package.read_only_recommended().unwrap(), None);
    assert_eq!(
        package.put_read_only_recommended_in_place(true).unwrap(),
        None
    );
    assert_eq!(package.read_only_recommended().unwrap(), Some(true));
    let bytes = package.to_bytes().expect("serialize changed package");
    let mut reopened = Package::from_bytes(&bytes).expect("reopen changed package");
    assert_eq!(reopened.read_only_recommended().unwrap(), Some(true));
    assert_eq!(
        reopened.remove_read_only_recommended_in_place().unwrap(),
        Some(true)
    );
    assert_eq!(reopened.read_only_recommended().unwrap(), None);
    assert_eq!(
        reopened.remove_read_only_recommended_in_place().unwrap(),
        None
    );

    let properties = PackURI::new("/ppt/presProps.xml").unwrap();
    assert!(reopened.opc().unwrap().get_part(&properties).is_ok());
}

#[test]
fn package_snapshot_patch_is_source_checked_and_reversible() {
    let xml = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{P1710_NS}"><p:extLst><!-- keep --><p:ext uri="{URI}"><p1710:readonlyRecommended val="1"/></p:ext><p:ext uri="urn:other"><o:data xmlns:o="urn:other"/></p:ext></p:extLst></p:presentationPr>"#
    );
    let mut package = package_with_properties(xml.as_bytes());
    let snapshot = package
        .read_only_recommended_snapshot()
        .unwrap()
        .expect("owner snapshot");
    let original = snapshot.source_xml().to_vec();
    let mut edit = package.edit_read_only_recommended().unwrap();
    assert_eq!(edit.set(false).unwrap(), Some(true));
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    package = commit.into_package();
    assert_eq!(package.read_only_recommended().unwrap(), Some(false));
    package = package
        .apply_read_only_recommended(&inverse)
        .expect("publish inverse")
        .into_package();
    assert_eq!(package.read_only_recommended().unwrap(), Some(true));
    let restored = package
        .read_only_recommended_snapshot()
        .unwrap()
        .expect("restored owner snapshot");
    assert_eq!(restored.source_xml(), original.as_slice());
}

#[test]
fn replay_rechecks_readonly_output_with_the_recipient_profile() {
    let properties = PackURI::new("/ppt/presProps.xml").unwrap();
    let seed = Package::new().unwrap();
    let other_max = max_non_properties_events(&seed, &properties);

    // Keep the source below the recipient event budget while making the
    // typed edit's new ext/readonlyRecommended branch cross it.  The margin
    // admits every unrelated package part and relationship member.
    let target_events = other_max + 100;
    let mut xml = format!(r#"<p:presentationPr xmlns:p="{P_NS}"><p:extLst>"#);
    let mut events = xml_event_count(xml.as_bytes());
    let mut ext_count = 0;
    while events < target_events - 3 {
        write!(xml, r#"<p:ext uri="urn:opaque:{ext_count}"/>"#).unwrap();
        ext_count += 1;
        events += 1;
    }
    xml.push_str("</p:extLst></p:presentationPr>");

    let before = package_with_properties(xml.as_bytes());
    let before_bytes = before.to_bytes().unwrap();
    let before_part = before.opc().unwrap().get_part(&properties).unwrap();
    let before_events = xml_event_count(before_part.blob());

    let mut edit = before.edit_read_only_recommended().unwrap();
    assert_eq!(edit.set(true).unwrap(), None);
    let commit = edit.commit().unwrap();
    let after = commit.package();
    let after_events = xml_event_count(after.opc().unwrap().get_part(&properties).unwrap().blob());
    assert!(after_events > before_events);

    let tight_limits = SourceLimits::builder()
        .max_events(before_events + 1)
        .unwrap()
        .build()
        .unwrap();
    let tight = Package::from_bytes_with_options(
        &before_bytes,
        OpenOptions::default().source_limits(tight_limits),
    )
    .expect("the original owner fits the recipient profile");
    let unchanged = tight.to_bytes().unwrap();
    let refusal = tight.apply_read_only_recommended(commit.patch());
    assert!(
        matches!(refusal, Err(Error::Limit { .. }) | Err(Error::Opc(_))),
        "expanded readonly owner must be refused by a typed bound"
    );
    assert_eq!(tight.to_bytes().unwrap(), unchanged);

    let permissive = Package::from_bytes(&before_bytes).unwrap();
    let replayed = permissive
        .apply_read_only_recommended(commit.patch())
        .expect("a permissive recipient admits the same patch")
        .into_package();
    assert_eq!(replayed.read_only_recommended().unwrap(), Some(true));

    // Build the exact no-op on the recipient that will replay it. A no-op
    // patch is content-rooted just like a changed patch; constructing it from
    // Package::new() would exercise the foreign-base conflict path instead of
    // the unchanged-payload fast path.
    let noop = permissive
        .edit_read_only_recommended()
        .unwrap()
        .commit()
        .unwrap();
    let noop_source = permissive.to_bytes().unwrap();
    let unchanged_noop = permissive
        .apply_read_only_recommended(noop.patch())
        .expect("no-op replay does not require a semantic scan")
        .into_package();
    assert_eq!(unchanged_noop.to_bytes().unwrap(), noop_source);

    // Equal content under a different recipient SourceLimits profile still
    // matches the no-op base. The replay must return the original bytes before
    // any family semantic validation, even though the tight profile rejects
    // the changed patch above.
    let tight_source = tight.to_bytes().unwrap();
    let unchanged_tight = tight
        .apply_read_only_recommended(noop.patch())
        .expect("equal-content no-op replay is profile-independent")
        .into_package();
    assert_eq!(unchanged_tight.to_bytes().unwrap(), tight_source);

    let foreign = Package::new().unwrap();
    let foreign_source = foreign.to_bytes().unwrap();
    let error = foreign
        .apply_read_only_recommended(noop.patch())
        .unwrap_err();
    assert!(matches!(error, Error::PatchConflict));
    assert_eq!(foreign.to_bytes().unwrap(), foreign_source);
}

#[test]
fn exact_signed_noop_is_allowed_but_changed_edit_requires_explicit_policy() {
    let mut package = Package::new().expect("new package");
    assert_eq!(
        package.put_read_only_recommended_in_place(true).unwrap(),
        None
    );
    package
        .edit_opc_in_place(|opc| {
            opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
            Ok(())
        })
        .expect("install inert signature origin");
    assert!(package.opc().unwrap().is_signed());

    assert_eq!(
        package.put_read_only_recommended_in_place(true).unwrap(),
        Some(true)
    );
    assert!(package.opc().unwrap().is_signed());
    let error = package
        .put_read_only_recommended_in_place(false)
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(package.opc().unwrap().is_signed());
}

#[test]
fn duplicate_presentation_properties_parts_are_rejected() {
    let authored = Package::new().expect("new package");
    let bytes = authored.to_bytes().expect("serialize package");
    let mut opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).expect("reopen package");
    let orphan = PackURI::new("/ppt/presProps-orphan.xml").expect("orphan URI");
    opc.try_add_part(Box::new(BlobPart::new(
        orphan,
        litchi_opc::constants::content_type::PML_PRES_PROPS.to_owned(),
        br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
            .to_vec(),
    )))
    .expect("add orphan owner");
    let package = Package::from_opc_package(opc).expect("open package");
    assert!(package.read_only_recommended().is_err());
}
