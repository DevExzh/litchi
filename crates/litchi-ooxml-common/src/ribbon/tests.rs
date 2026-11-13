//! Focused regression coverage for the Ribbon model, codec, and package graph.

use super::model::{Family, LEGACY_RELATIONSHIP, Limits, Version};
use super::package::{CONTENT_TYPE, Edit, load};
use crate::{Error, Result};
use litchi_opc::BlobPart;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI, XmlPart};
use std::sync::Arc;

const XML_2007: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#;
const XML_2010: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>"#;
const XML_UI2: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2007/10/customui"/>"#;

fn put(package: &mut OpcPackage, version: Version, xml: Vec<u8>, limits: &Limits) -> Result<()> {
    let mut edit =
        Edit::new_with_options(package, super::package::EditOptions { limits: *limits })?;
    edit.put(version, xml)?;
    *package = edit.commit()?.into_package();
    Ok(())
}

fn remove(package: &mut OpcPackage, family: Family) -> Result<bool> {
    let mut edit = Edit::new(package)?;
    let changed = edit.remove(family)?;
    *package = edit.commit()?.into_package();
    Ok(changed)
}

#[test]
fn fixed_slots_borrow_package_bytes_and_prefer_modern() {
    let mut package = OpcPackage::new();
    put(
        &mut package,
        Version::V2007,
        XML_2007.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    put(
        &mut package,
        Version::V2010,
        XML_2010.to_vec(),
        &Limits::standard(),
    )
    .unwrap();

    let ribbons = load(&package, &Limits::standard()).unwrap();
    let legacy = ribbons.legacy().unwrap();
    let modern = ribbons.modern().unwrap();
    assert_eq!(ribbons.iter().collect::<Vec<_>>(), [legacy, modern]);
    assert_eq!(ribbons.effective(), Some(modern));
    assert_eq!(legacy.version(), Version::V2007);
    assert_eq!(modern.version(), Version::V2010);
    assert_eq!(
        modern.xml().as_ptr(),
        package.get_part(modern.part()).unwrap().blob().as_ptr()
    );
}

#[test]
fn modern_vocabulary_updates_in_place() {
    let mut package = OpcPackage::new();
    put(
        &mut package,
        Version::Ui2,
        XML_UI2.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    let part = load(&package, &Limits::standard())
        .unwrap()
        .modern()
        .unwrap()
        .part()
        .clone();

    put(
        &mut package,
        Version::V2010,
        XML_2010.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    let modern = load(&package, &Limits::standard())
        .unwrap()
        .modern()
        .unwrap();
    assert_eq!(modern.part(), &part);
    assert_eq!(modern.version(), Version::V2010);
    assert_eq!(package.part_count(), 1);
}

#[test]
fn shared_ribbon_is_forked_before_update() {
    let mut package = OpcPackage::new();
    let original = raw_ribbon(&mut package, Version::V2007, "/addons/ui.xml", XML_2007);
    let image = add_image(&mut package, "/addons/images/icon.png", "image/png");
    package
        .get_part_mut(&original)
        .unwrap()
        .relate_to("images/icon.png", rt::IMAGE);
    let original_bytes = package.get_part(&original).unwrap().blob_arc();
    let source = add_source(&mut package, "/word/document.xml");
    package
        .get_part_mut(&source)
        .unwrap()
        .relate_to("../addons/ui.xml", "urn:shared-ribbon");
    let relationship_id = load(&package, &Limits::standard())
        .unwrap()
        .legacy()
        .unwrap()
        .id()
        .to_owned();
    let replacement =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#
            .to_vec();
    let allocation = replacement.as_ptr();

    put(
        &mut package,
        Version::V2007,
        replacement,
        &Limits::standard(),
    )
    .unwrap();

    let updated = load(&package, &Limits::standard())
        .unwrap()
        .legacy()
        .unwrap();
    assert_ne!(updated.part(), &original);
    assert_eq!(updated.part().as_str(), "/customUI/customUI.xml");
    assert_eq!(updated.id(), relationship_id);
    assert_eq!(updated.xml().as_ptr(), allocation);
    assert!(Arc::ptr_eq(
        &original_bytes,
        &package.get_part(&original).unwrap().blob_arc()
    ));
    assert_eq!(package.get_part(&original).unwrap().blob(), XML_2007);
    let updated_part = package.get_part(updated.part()).unwrap();
    assert_eq!(updated_part.rels().len(), 1);
    let image_relationship = updated_part.rels().iter().next().unwrap();
    assert_eq!(image_relationship.target_ref(), "../addons/images/icon.png");
    assert_eq!(image_relationship.target_partname().unwrap(), image);
}

#[test]
fn forked_image_relationships_support_inverse_without_relaxing_source_checks() {
    use litchi_opc::PackageWriter;
    let mut package = package_with_legacy();
    let original = legacy_part(&package);
    let image = add_image(&mut package, "/customUI/image.png", "image/png");
    package
        .get_part_mut(&original)
        .unwrap()
        .relate_to("image.png", rt::IMAGE);
    let source = add_source(&mut package, "/word/document.xml");
    package
        .get_part_mut(&source)
        .unwrap()
        .relate_to("../customUI/customUI.xml", "urn:shared-ribbon");
    let package = OpcPackage::from_bytes(
        &PackageWriter::to_bytes(&package).unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    let mut edit = Edit::new(&package).unwrap();
    edit.put(Version::V2007, br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#.to_vec()).unwrap();
    let commit = edit.commit().unwrap();
    let fork = legacy_part(commit.package());
    assert_ne!(fork, original);
    let inverse = commit.patch().inverse();
    let reopened = OpcPackage::from_bytes(
        &PackageWriter::to_bytes(commit.package()).unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    for published in [commit.package(), &reopened] {
        let restored = inverse.apply(published).unwrap().into_package();
        assert_eq!(legacy_part(&restored), original);
        assert_eq!(restored.part_count(), package.part_count());
        let limits = restored.read_limits();
        assert_eq!(
            restored.source_content_types(limits).unwrap(),
            package.source_content_types(limits).unwrap()
        );
        let root = PackURI::new("/").unwrap();
        assert_eq!(
            restored.source_relationships(&root, limits).unwrap(),
            package.source_relationships(&root, limits).unwrap()
        );
        for part in package.iter_parts() {
            let name = part.partname();
            let actual = restored.get_part(name).unwrap();
            assert_eq!(actual.content_type(), part.content_type());
            assert_eq!(actual.blob(), package.get_part(name).unwrap().blob());
            assert_eq!(
                restored.source_relationships(name, limits).unwrap(),
                package.source_relationships(name, limits).unwrap()
            );
        }
    }

    // Source equality still includes payload bytes, lexical relationship XML
    // and the difference between an absent and an explicit empty member.
    for change in 0..3 {
        let mut stale = reopened.clone();
        let limits = stale.read_limits();
        match change {
            0 => stale.get_part_mut(&image).unwrap().set_blob(vec![9, 8, 7]),
            1 => {
                let root = PackURI::new("/").unwrap();
                let before = stale.source_relationships(&root, limits).unwrap();
                let mut bytes = before.bytes().to_vec();
                bytes.extend_from_slice(b"<!-- changed source -->");
                stale
                    .try_replace_relationships_bytes(&before, Arc::new(bytes), true, limits)
                    .unwrap();
            },
            2 => {
                let before = stale.source_relationships(&image, limits).unwrap();
                assert!(!before.member_present());
                stale
                    .try_replace_relationships_bytes(&before, before.bytes_arc(), true, limits)
                    .unwrap();
            },
            _ => unreachable!(),
        }
        let bytes = PackageWriter::to_bytes(&stale).unwrap();
        assert!(matches!(inverse.apply(&stale), Err(Error::PatchConflict)));
        assert_eq!(PackageWriter::to_bytes(&stale).unwrap(), bytes);
    }
}

#[test]
fn identical_put_is_a_signature_and_allocation_preserving_noop() {
    let mut package = OpcPackage::new();
    put(
        &mut package,
        Version::V2007,
        XML_2007.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    let part = load(&package, &Limits::standard())
        .unwrap()
        .legacy()
        .unwrap()
        .part()
        .clone();
    let before = package.get_part(&part).unwrap().blob_arc();
    sign_marker(&mut package);

    put(
        &mut package,
        Version::V2007,
        XML_2007.to_vec(),
        &Limits::standard(),
    )
    .unwrap();

    assert!(package.is_signed());
    let after = package.get_part(&part).unwrap().blob_arc();
    assert!(Arc::ptr_eq(&before, &after));

    // A changed Ribbon refuses the signed package and leaves it untouched
    // until the signature is removed explicitly (ADR 0006).
    let changed = || {
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#
            .to_vec()
    };
    assert!(matches!(
        put(&mut package, Version::V2007, changed(), &Limits::standard()),
        Err(Error::Signed)
    ));
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &before,
        &package.get_part(&part).unwrap().blob_arc()
    ));
    package.unsign();
    put(&mut package, Version::V2007, changed(), &Limits::standard()).unwrap();
    assert!(!package.is_signed());
}

#[test]
fn name_collisions_do_not_copy_the_moved_payload() {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/CUSTOMui/CUSTOMui.XML").unwrap(),
        "application/octet-stream".into(),
        vec![7],
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/customUI/customUI1.xml/child").unwrap(),
        "application/octet-stream".into(),
        vec![8],
    )));
    let xml = XML_2007.to_vec();
    let allocation = xml.as_ptr();

    put(&mut package, Version::V2007, xml, &Limits::standard()).unwrap();

    let ui = load(&package, &Limits::standard())
        .unwrap()
        .legacy()
        .unwrap();
    assert_eq!(ui.part().as_str(), "/customUI/customUI2.xml");
    assert_eq!(ui.xml().as_ptr(), allocation);
}

#[test]
fn xml_bytes_depth_nodes_entities_and_namespaces_are_bounded() {
    let mut package = OpcPackage::new();
    let tiny_bytes = Limits {
        xml_bytes: XML_2007.len() - 1,
        ..Limits::standard()
    };
    assert!(matches!(
        put(&mut package, Version::V2007, XML_2007.to_vec(), &tiny_bytes),
        Err(Error::Limit {
            kind: crate::LimitKind::Ribbon(_),
            ..
        })
    ));
    assert_eq!(package.part_count(), 0);

    let shallow = Limits {
        depth: 1,
        ..Limits::standard()
    };
    let nested = br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#;
    assert!(matches!(
        put(&mut package, Version::V2007, nested.to_vec(), &shallow),
        Err(Error::Limit {
            kind: crate::LimitKind::Ribbon(_),
            ..
        })
    ));

    let few_nodes = Limits {
        nodes: 1,
        ..Limits::standard()
    };
    assert!(matches!(
        put(&mut package, Version::V2007, XML_2007.to_vec(), &few_nodes),
        Err(Error::Limit {
            kind: crate::LimitKind::Ribbon(_),
            ..
        })
    ));
    assert!(put(
        &mut package,
        Version::V2007,
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">&madeUp;</customUI>"#
            .to_vec(), &Limits::standard())
    .is_err());
    assert!(put(
        &mut package,
        Version::V2007,
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><bad:node/></customUI>"#
            .to_vec(), &Limits::standard())
    .is_err());
    assert!(put(
        &mut package,
        Version::V2007,
        br#"<!DOCTYPE customUI><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
            .to_vec(), &Limits::standard())
    .is_err());
    assert_eq!(package.part_count(), 0);
}

#[test]
fn xml_declarations_characters_namespaces_and_inert_instructions_are_strict() {
    let mut valid = OpcPackage::new();
    let normalized_namespace = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?safe inert?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customu&#x69;"><?inside retained?></customUI>"#;
    put(
        &mut valid,
        Version::V2007,
        normalized_namespace.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    assert_eq!(
        load(&valid, &Limits::standard())
            .unwrap()
            .legacy()
            .unwrap()
            .xml(),
        normalized_namespace
    );

    let replacement = br#"<?xml version="1.0"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customu&#x69;"><?changed retained?></customUI>"#;
    put(
        &mut valid,
        Version::V2007,
        replacement.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    assert_eq!(
        load(&valid, &Limits::standard())
            .unwrap()
            .legacy()
            .unwrap()
            .xml(),
        replacement
    );

    for invalid in [
        br#"<?xml?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
            .as_slice(),
        br#"<?xml version="1.1"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
            .as_slice(),
        br#"<?xml version="1.0" encoding="UTF-16"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
            .as_slice(),
        b"<?xml version=\"1.0\" encoding=\"US-ASCII\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\" label=\"caf\xC3\xA9\"/>"
            .as_slice(),
        br#"<!--before--><?xml version="1.0"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
            .as_slice(),
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">&#1;</customUI>"#
            .as_slice(),
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" label="&#xB;"/>"#
            .as_slice(),
        br#"<?xml illegal?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
            .as_slice(),
    ] {
        let mut package = OpcPackage::new();
        assert!(put(&mut package, Version::V2007, invalid.to_vec(), &Limits::standard()).is_err());
        assert_eq!(package.part_count(), 0);
    }

    let mut duplicate_expanded = OpcPackage::new();
    assert!(
        put(
            &mut duplicate_expanded,
            Version::V2007,
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" xmlns:a="urn:x" xmlns:b="urn:&#x78;" a:value="1" b:value="2"/>"#
                .to_vec(), &Limits::standard(),
        )
        .is_err()
    );

    let mut raw_control =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">"#.to_vec();
    raw_control.push(1);
    raw_control.extend_from_slice(b"</customUI>");
    let mut package = OpcPackage::new();
    assert!(
        put(
            &mut package,
            Version::V2007,
            raw_control,
            &Limits::standard()
        )
        .is_err()
    );
}

#[test]
fn relationship_location_cardinality_and_root_family_are_strict() {
    let mut external = OpcPackage::new();
    external.relate_to_external("https://example.invalid/ui.xml", LEGACY_RELATIONSHIP);
    assert!(matches!(
        load(&external, &Limits::standard()),
        Err(Error::Relationship(_))
    ));

    let mut duplicate = OpcPackage::new();
    raw_ribbon(
        &mut duplicate,
        Version::V2010,
        "/customUI/one.xml",
        XML_2010,
    );
    raw_ribbon(&mut duplicate, Version::Ui2, "/customUI/two.xml", XML_UI2);
    assert!(matches!(
        load(&duplicate, &Limits::standard()),
        Err(Error::Relationship(_))
    ));

    let mut mismatched = OpcPackage::new();
    raw_ribbon(
        &mut mismatched,
        Version::V2007,
        "/customUI/customUI.xml",
        XML_2010,
    );
    assert!(matches!(
        load(&mismatched, &Limits::standard()),
        Err(Error::Invalid(_))
    ));

    let mut part_sourced = OpcPackage::new();
    let source = PackURI::new("/word/document.xml").unwrap();
    part_sourced.add_part(Box::new(XmlPart::new(
        source.clone(),
        "application/xml".into(),
        b"<document/>".to_vec(),
    )));
    part_sourced
        .get_part_mut(&source)
        .unwrap()
        .relate_to("../customUI/customUI.xml", LEGACY_RELATIONSHIP);
    assert!(matches!(
        load(&part_sourced, &Limits::standard()),
        Err(Error::Relationship(_))
    ));
}

#[test]
fn ribbon_outbound_relationships_are_internal_images_only() {
    let mut wrong_type = package_with_legacy();
    let ribbon = legacy_part(&wrong_type);
    let data = PackURI::new("/customUI/data.bin").unwrap();
    wrong_type.add_part(Box::new(BlobPart::new(
        data,
        "application/octet-stream".into(),
        vec![1],
    )));
    wrong_type
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("data.bin", "urn:not-an-image");
    assert!(matches!(
        load(&wrong_type, &Limits::standard()),
        Err(Error::Relationship(_))
    ));

    let mut wrong_media = package_with_legacy();
    let ribbon = legacy_part(&wrong_media);
    let data = PackURI::new("/customUI/image.bin").unwrap();
    wrong_media.add_part(Box::new(BlobPart::new(
        data,
        "application/octet-stream".into(),
        vec![1],
    )));
    wrong_media
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("image.bin", rt::IMAGE);
    assert!(matches!(
        load(&wrong_media, &Limits::standard()),
        Err(Error::ContentType { .. })
    ));

    let mut external = package_with_legacy();
    let ribbon = legacy_part(&external);
    external
        .get_part_mut(&ribbon)
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::IMAGE.into(),
            "https://example.invalid/image.png".into(),
            "rIdImage".into(),
            true,
        );
    assert!(matches!(
        load(&external, &Limits::standard()),
        Err(Error::Relationship(_))
    ));

    let mut queried = package_with_legacy();
    let ribbon = legacy_part(&queried);
    add_image(&mut queried, "/customUI/image.png", "image/png");
    queried
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("image.png?variant=2", rt::IMAGE);
    assert!(matches!(
        load(&queried, &Limits::standard()),
        Err(Error::Relationship(_))
    ));

    for content_type in ["image/ png", "image/png;garbage"] {
        let mut malformed = package_with_legacy();
        let ribbon = legacy_part(&malformed);
        add_image(&mut malformed, "/customUI/image.png", content_type);
        malformed
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.png", rt::IMAGE);
        assert!(matches!(
            load(&malformed, &Limits::standard()),
            Err(Error::ContentType { .. })
        ));
    }
}

#[test]
fn aggregate_image_relationships_are_bounded() {
    let mut package = package_with_legacy();
    let ribbon = legacy_part(&package);
    add_image(&mut package, "/customUI/image.png", "IMAGE/PNG");
    package
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("image.png", rt::STRICT_IMAGE);
    let limits = Limits {
        images: 0,
        ..Limits::standard()
    };
    assert!(matches!(
        load(&package, &limits),
        Err(Error::Limit {
            kind: crate::LimitKind::Ribbon(_),
            ..
        })
    ));
}

#[test]
fn remove_collects_unreferenced_ribbon_and_image_parts() {
    let mut package = package_with_legacy();
    let ribbon = legacy_part(&package);
    let image = add_image(&mut package, "/customUI/image.png", "image/png");
    package
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("image.png", rt::IMAGE);
    sign_marker(&mut package);

    // Removing a present family refuses the signed package before any part
    // is removed, until the signature is removed explicitly (ADR 0006).
    assert!(matches!(
        remove(&mut package, Family::Legacy),
        Err(Error::Signed)
    ));
    assert!(package.is_signed());
    assert!(package.get_part(&ribbon).is_ok());
    assert!(package.get_part(&image).is_ok());
    package.unsign();
    assert!(remove(&mut package, Family::Legacy).unwrap());
    assert!(package.get_part(&ribbon).is_err());
    assert!(package.get_part(&image).is_err());
    assert!(!package.is_signed());
    assert!(!remove(&mut package, Family::Legacy).unwrap());
}

#[test]
fn remove_preserves_shared_ribbon_and_image_parts() {
    let mut shared_image = package_with_legacy();
    let ribbon = legacy_part(&shared_image);
    let image = add_image(&mut shared_image, "/customUI/image.png", "image/png");
    shared_image
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("image.png", rt::IMAGE);
    let source = add_source(&mut shared_image, "/word/document.xml");
    shared_image
        .get_part_mut(&source)
        .unwrap()
        .relate_to("../customUI/image.png", rt::IMAGE);
    assert!(remove(&mut shared_image, Family::Legacy).unwrap());
    assert!(shared_image.get_part(&ribbon).is_err());
    assert!(shared_image.get_part(&image).is_ok());

    let mut shared_ribbon = package_with_legacy();
    let ribbon = legacy_part(&shared_ribbon);
    let image = add_image(&mut shared_ribbon, "/customUI/image.png", "image/png");
    shared_ribbon
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("image.png", rt::IMAGE);
    let source = add_source(&mut shared_ribbon, "/word/document.xml");
    shared_ribbon
        .get_part_mut(&source)
        .unwrap()
        .relate_to("../customUI/customUI.xml", "urn:shared-ribbon");
    assert!(remove(&mut shared_ribbon, Family::Legacy).unwrap());
    assert!(shared_ribbon.get_part(&ribbon).is_ok());
    assert!(shared_ribbon.get_part(&image).is_ok());
}

#[test]
fn image_cycles_are_collected_unless_reachable_from_a_kept_part() {
    let mut unanchored = package_with_image_cycle();
    let first = PackURI::new("/customUI/one.png").unwrap();
    let second = PackURI::new("/customUI/two.png").unwrap();
    remove(&mut unanchored, Family::Legacy).unwrap();
    assert!(unanchored.get_part(&first).is_err());
    assert!(unanchored.get_part(&second).is_err());

    let mut anchored = package_with_image_cycle();
    anchored.relate_to("customUI/one.png", "urn:keep-image");
    remove(&mut anchored, Family::Legacy).unwrap();
    assert!(anchored.get_part(&first).is_ok());
    assert!(anchored.get_part(&second).is_ok());
}

#[test]
fn failed_mutations_leave_graph_bytes_and_signatures_untouched() {
    let mut package = package_with_legacy();
    let ribbon = legacy_part(&package);
    let before = package.get_part(&ribbon).unwrap().blob_arc();
    sign_marker(&mut package);
    assert!(
        put(
            &mut package,
            Version::V2010,
            b"<broken".to_vec(),
            &Limits::standard()
        )
        .is_err()
    );
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &before,
        &package.get_part(&ribbon).unwrap().blob_arc()
    ));

    let data = PackURI::new("/customUI/data.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        data,
        "application/octet-stream".into(),
        vec![1],
    )));
    package
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("data.bin", "urn:not-an-image");
    assert!(remove(&mut package, Family::Legacy).is_err());
    assert!(package.is_signed());
    assert!(package.get_part(&ribbon).is_ok());
    assert!(package.rels().iter().any(|relationship| {
        Family::from_relationship(relationship.reltype()) == Some(Family::Legacy)
    }));

    let mut absent = OpcPackage::new();
    sign_marker(&mut absent);
    assert!(!remove(&mut absent, Family::Modern).unwrap());
    assert!(absent.is_signed());
}

fn package_with_legacy() -> OpcPackage {
    let mut package = OpcPackage::new();
    put(
        &mut package,
        Version::V2007,
        XML_2007.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    package
}

fn package_with_image_cycle() -> OpcPackage {
    let mut package = package_with_legacy();
    let ribbon = legacy_part(&package);
    let first = add_image(&mut package, "/customUI/one.png", "image/png");
    let second = add_image(&mut package, "/customUI/two.png", "image/png");
    package
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("one.png", rt::IMAGE);
    package
        .get_part_mut(&ribbon)
        .unwrap()
        .relate_to("two.png", rt::IMAGE);
    package
        .get_part_mut(&first)
        .unwrap()
        .relate_to("two.png", rt::IMAGE);
    package
        .get_part_mut(&second)
        .unwrap()
        .relate_to("one.png", rt::IMAGE);
    package
}

fn raw_ribbon(package: &mut OpcPackage, version: Version, name: &str, xml: &[u8]) -> PackURI {
    let part = PackURI::new(name).unwrap();
    package.add_part(Box::new(XmlPart::new(
        part.clone(),
        CONTENT_TYPE.into(),
        xml.to_vec(),
    )));
    package.relate_to(name.trim_start_matches('/'), version.relationship());
    part
}

fn legacy_part(package: &OpcPackage) -> PackURI {
    load(package, &Limits::standard())
        .unwrap()
        .legacy()
        .unwrap()
        .part()
        .clone()
}

fn add_image(package: &mut OpcPackage, name: &str, content_type: &str) -> PackURI {
    let part = PackURI::new(name).unwrap();
    package.add_part(Box::new(BlobPart::new(
        part.clone(),
        content_type.into(),
        vec![1, 2, 3],
    )));
    part
}

fn add_source(package: &mut OpcPackage, name: &str) -> PackURI {
    let part = PackURI::new(name).unwrap();
    package.add_part(Box::new(XmlPart::new(
        part.clone(),
        "application/xml".into(),
        b"<source/>".to_vec(),
    )));
    part
}

fn sign_marker(package: &mut OpcPackage) {
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
}

const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

/// One reader's root element in the forms the rows are built from.
struct Root<'a> {
    /// The start tag.
    open: &'a str,
    /// Content that makes the element valid for its reader.
    body: &'a str,
    /// The end tag.
    close: &'a str,
    /// The element as one empty-element tag.
    empty: &'a str,
}

fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
    let Root {
        open,
        body,
        close,
        empty,
    } = *root;
    let document = format!("{open}{body}{close}");
    let d = document.as_str();
    vec![
        ("empty input", String::new()),
        ("whitespace only", " \n".to_owned()),
        ("declaration only", DECLARATION.to_owned()),
        ("empty-element root", empty.to_owned()),
        ("start and end root", format!("{open}{close}")),
        ("root then whitespace", format!("{d}\n ")),
        ("root then comment", format!("{d}<!--c-->")),
        ("root then PI", format!("{d}<?pi x?>")),
        ("root then text", format!("{d}text")),
        ("root then second root", format!("{d}{d}")),
        ("root then foreign start", format!("{d}<foreign></foreign>")),
        ("root then foreign empty", format!("{d}<foreign/>")),
        ("declaration after root", format!("{d}{DECLARATION}")),
        ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
        (
            "comment before declaration",
            format!("<!--c-->{DECLARATION}{d}"),
        ),
        ("text before root", format!("text{d}")),
        ("unterminated root", format!("{open}{body}")),
        ("stray end tag", format!("{d}</foreign>")),
        ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
        ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
        ("declaration then root", format!("{DECLARATION}{d}")),
        (
            "whitespace before declaration",
            format!(" {DECLARATION}{d}"),
        ),
        ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
        ("comment then root", format!("<!--c-->{d}")),
        ("PI then root", format!("<?pi x?>{d}")),
        ("root then text then second root", format!("{d}text{d}")),
        ("empty-element root then root", format!("{empty}{d}")),
        (
            "declaration after empty-element root",
            format!("{empty}{DECLARATION}"),
        ),
    ]
}

fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
    let actual = rows(root)
        .into_iter()
        .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
        .collect::<Vec<_>>();
    assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
}

fn outcome<T, E: std::fmt::Debug>(
    result: std::result::Result<T, E>,
    summary: impl Fn(T) -> String,
) -> String {
    match result {
        Ok(value) => format!("Ok({})", summary(value)),
        Err(error) => format!("{error:?}"),
    }
}

/// What the Ribbon validator accepts before, around and after `customUI`,
/// and the exact error it reports otherwise. Each expected line is the row
/// label followed by the result.
#[test]
fn custom_ui_envelope() {
    let root = Root {
        open: r#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">"#,
        body: "<ribbon/>",
        close: "</customUI>",
        empty: r#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>"#,
    };
    assert_matrix(
        &root,
        &[
            "empty input: Invalid(\"Ribbon XML must contain one complete customUI root\")",
            "whitespace only: Invalid(\"Ribbon XML must contain one complete customUI root\")",
            "declaration only: Invalid(\"Ribbon XML must contain one complete customUI root\")",
            "empty-element root: Ok(V2010)",
            "start and end root: Ok(V2010)",
            "root then whitespace: Ok(V2010)",
            "root then comment: Ok(V2010)",
            "root then PI: Ok(V2010)",
            "root then text: Invalid(\"Ribbon XML contains text outside its root\")",
            "root then second root: Invalid(\"Ribbon XML must contain exactly one root\")",
            "root then foreign start: Invalid(\"Ribbon XML must contain exactly one root\")",
            "root then foreign empty: Invalid(\"Ribbon XML must contain exactly one root\")",
            "declaration after root: Invalid(\"Ribbon XML declaration must be the first event and occur once\")",
            "two declarations: Invalid(\"Ribbon XML declaration must be the first event and occur once\")",
            "comment before declaration: Invalid(\"Ribbon XML declaration must be the first event and occur once\")",
            "text before root: Invalid(\"Ribbon XML contains text outside its root\")",
            "unterminated root: Invalid(\"Ribbon XML must contain one complete customUI root\")",
            "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
            "DOCTYPE: Invalid(\"DTD is forbidden in Ribbon XML\")",
            "top-level CDATA: Invalid(\"Ribbon XML contains CDATA outside its root\")",
            "declaration then root: Ok(V2010)",
            "whitespace before declaration: Invalid(\"Ribbon XML declaration must be the first event and occur once\")",
            "PI before declaration: Invalid(\"Ribbon XML declaration must be the first event and occur once\")",
            "comment then root: Ok(V2010)",
            "PI then root: Ok(V2010)",
            "root then text then second root: Invalid(\"Ribbon XML contains text outside its root\")",
            "empty-element root then root: Invalid(\"Ribbon XML must contain exactly one root\")",
            "declaration after empty-element root: Invalid(\"Ribbon XML declaration must be the first event and occur once\")",
        ],
        |xml| {
            outcome(
                super::codec::validate_xml(xml, Family::Modern, &Limits::standard()),
                |version| format!("{version:?}"),
            )
        },
    );
}
