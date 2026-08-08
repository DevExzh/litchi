//! Transactional XLSB slicer/timeline package-owner regressions.

use litchi_opc::{BlobPart, PackURI, PackageWriter, Part, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::slicer::{
    Cache as SlicerCache, CrossFilter, Item, Native, View as SlicerView, Views as SlicerViews,
};
use litchi_xlsb::timeline::{
    Cache as TimelineCache, FilterType, Level, Range, State, View as TimelineView,
    Views as TimelineViews,
};
use std::io::Cursor;

fn slicer_cache() -> SlicerCache {
    let mut native = Native::new(0);
    native.cross_filter = CrossFilter::ShowItemsWithoutData;
    native.items = vec![
        Item {
            cache_index: 0,
            selected: true,
            no_data: false,
        },
        Item {
            cache_index: 1,
            selected: false,
            no_data: true,
        },
    ];
    SlicerCache::native("Slicer_State", "State", native)
}

fn timeline_cache() -> TimelineCache {
    let bounds = Range::new("2024-01-01T00:00:00Z", "2024-12-31T23:59:59Z").unwrap();
    TimelineCache::new(
        "TimelineCache1",
        "Date",
        State::new(bounds, 0, FilterType::DateBetween),
    )
}

fn package_snapshot(package: &litchi_opc::OpcPackage) -> Package {
    Package::from_opc(package.clone()).unwrap()
}

fn apply_slicer(package: &mut litchi_opc::OpcPackage, patch: &litchi_xlsb::slicer::Patch) {
    *package = package_snapshot(package)
        .apply_slicer_patch(patch)
        .unwrap()
        .into_opc();
}

fn apply_timeline(package: &mut litchi_opc::OpcPackage, patch: &litchi_xlsb::timeline::Patch) {
    *package = package_snapshot(package)
        .apply_timeline_patch(patch)
        .unwrap()
        .into_opc();
}

fn package_bytes(package: &litchi_opc::OpcPackage) -> Vec<u8> {
    PackageWriter::to_bytes(package).unwrap()
}

fn add_inbound(package: &mut litchi_opc::OpcPackage, source: &PackURI, target: &PackURI, id: &str) {
    package
        .get_part_mut(source)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test:shared-owner".to_string(),
            target.relative_ref(source.base_uri()),
            id.to_string(),
            TargetMode::Internal,
        )
        .unwrap();
}

fn add_signature_marker(package: &mut litchi_opc::OpcPackage) {
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_string(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature".to_string(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    package.add_part(Box::new(origin_part));
    package.add_part(Box::new(BlobPart::new(
        signature,
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_string(),
        b"<Signature/>".to_vec(),
    )));
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_string(),
            origin.as_str().to_string(),
            "rIdSignatureOrigin".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
}

fn relocate_workbook(package: &mut litchi_opc::OpcPackage, destination: &PackURI) {
    let source = package.main_document_part().unwrap().partname().clone();
    let (content_type, blob, relationships) = {
        let workbook = package.get_part(&source).unwrap();
        let relationships = workbook
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.reltype().to_string(),
                    relationship.target_ref().to_string(),
                    relationship.r_id().to_string(),
                    relationship.target_mode(),
                )
            })
            .collect::<Vec<_>>();
        (
            workbook.content_type().to_string(),
            workbook.blob_arc(),
            relationships,
        )
    };
    let (root_id, root_type) = package
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::OFFICE_DOCUMENT
                    | litchi_opc::constants::relationship_type::STRICT_OFFICE_DOCUMENT
            )
        })
        .map(|relationship| {
            (
                relationship.r_id().to_string(),
                relationship.reltype().to_string(),
            )
        })
        .unwrap();

    let mut workbook = BlobPart::new_shared(destination.clone(), content_type, blob);
    for (relationship_type, target, id, mode) in relationships {
        workbook
            .rels_mut()
            .try_add_relationship(relationship_type, target, id, mode)
            .unwrap();
    }
    assert!(package.remove_part(&source));
    package.add_part(Box::new(workbook));
    package.rels_mut().remove(&root_id);
    package
        .rels_mut()
        .try_add_relationship(
            root_type,
            destination.as_str().trim_start_matches('/').to_string(),
            root_id,
            TargetMode::Internal,
        )
        .unwrap();
}

#[test]
fn semantic_facades_resolve_one_actual_root_workbook_relationship() {
    let mut missing = Package::create().unwrap().into_opc();
    let root_id = missing
        .rels()
        .iter()
        .find(|relationship| {
            relationship.reltype() == litchi_opc::constants::relationship_type::OFFICE_DOCUMENT
        })
        .unwrap()
        .r_id()
        .to_string();
    missing.rels_mut().remove(&root_id);
    let missing = Package::from(missing);
    assert!(missing.slicer_caches().is_err());
    assert!(missing.timeline_caches().is_err());

    let mut ambiguous = Package::create().unwrap().into_opc();
    ambiguous
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::STRICT_OFFICE_DOCUMENT.to_string(),
            "xl/workbook.bin".to_string(),
            "rIdSecondWorkbook".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    let ambiguous = Package::from(ambiguous);
    assert!(ambiguous.slicer_caches().is_err());
    assert!(ambiguous.timeline_caches().is_err());

    let mut alternate = Package::create().unwrap().into_opc();
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    litchi_xlsb::slicer::package::store_caches(&mut alternate, &workbook, &[slicer_cache()])
        .unwrap();
    litchi_xlsb::timeline::package::store_caches(&mut alternate, &workbook, &[timeline_cache()])
        .unwrap();
    litchi_xlsb::slicer::package::store_views(
        &mut alternate,
        &worksheet,
        &SlicerViews {
            items: vec![SlicerView::new("StateView", "Slicer_State")],
        },
    )
    .unwrap();
    litchi_xlsb::timeline::package::store_views(
        &mut alternate,
        &worksheet,
        &TimelineViews {
            items: vec![TimelineView::new(
                "Timeline1",
                "TimelineCache1",
                Level::Month,
            )],
        },
    )
    .unwrap();
    relocate_workbook(
        &mut alternate,
        &PackURI::new("/xl/alternate-workbook.bin").unwrap(),
    );

    let alternate = Package::from_opc(alternate).unwrap();
    assert_eq!(
        alternate.slicer_caches().unwrap().caches().unwrap().len(),
        1
    );
    assert_eq!(
        alternate.timeline_caches().unwrap().caches().unwrap().len(),
        1
    );
    assert!(alternate.slicer_views(0).unwrap().is_some());
    assert!(alternate.timeline_views(0).unwrap().is_some());
}

#[test]
fn workbook_facade_round_trips_and_removes_both_inert_owners() {
    let mut workbook = Package::create().unwrap().into_workbook().unwrap();
    let cache = slicer_cache();
    workbook.set_slicer_caches(vec![cache.clone()]).unwrap();

    let mut slicer_view = SlicerView::new("StateView", "Slicer_State");
    slicer_view.caption = Some("State".to_string());
    workbook
        .set_slicers(
            0,
            SlicerViews {
                items: vec![slicer_view],
            },
        )
        .unwrap();

    let timeline = timeline_cache();
    workbook
        .set_timeline_caches(vec![timeline.clone()])
        .unwrap();
    workbook
        .set_timelines(
            0,
            TimelineViews {
                items: vec![TimelineView::new(
                    "Timeline1",
                    "TimelineCache1",
                    Level::Month,
                )],
            },
        )
        .unwrap();

    assert_eq!(workbook.slicer_caches().unwrap(), vec![cache.clone()]);
    assert_eq!(workbook.timeline_caches().unwrap(), vec![timeline.clone()]);
    assert!(workbook.slicers(0).unwrap().is_some());
    assert!(workbook.timelines(0).unwrap().is_some());

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).unwrap();
    let reopened = Package::from_bytes(bytes.into_inner())
        .unwrap()
        .into_workbook()
        .unwrap();
    assert_eq!(reopened.slicer_caches().unwrap(), vec![cache]);
    assert_eq!(reopened.timeline_caches().unwrap(), vec![timeline]);
    assert!(reopened.slicers(0).unwrap().is_some());
    assert!(reopened.timelines(0).unwrap().is_some());

    assert!(workbook.remove_slicers(0).unwrap());
    assert!(workbook.remove_slicer_caches().unwrap());
    assert!(workbook.remove_timelines(0).unwrap());
    assert!(workbook.remove_timeline_caches().unwrap());
    assert!(workbook.slicer_caches().unwrap().is_empty());
    assert!(workbook.timeline_caches().unwrap().is_empty());
    assert!(workbook.slicers(0).unwrap().is_none());
    assert!(workbook.timelines(0).unwrap().is_none());
}

#[test]
fn slicer_patch_is_source_checked_reversible_and_dependency_safe() {
    let mut package = Package::create().unwrap().into_opc();
    let snapshot = package_snapshot(&package).slicer_caches().unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.replace_caches(vec![slicer_cache()]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_empty());
    apply_slicer(&mut package, commit.patch());
    assert!(
        package_snapshot(&package)
            .apply_slicer_patch(commit.patch())
            .is_err()
    );

    let views = package_snapshot(&package).slicer_views(0).unwrap().unwrap();
    let mut edit = views.edit().unwrap();
    edit.replace_views(SlicerViews {
        items: vec![SlicerView::new("StateView", "Slicer_State")],
    })
    .unwrap();
    apply_slicer(&mut package, edit.commit().unwrap().patch());

    let caches = package_snapshot(&package).slicer_caches().unwrap();
    let mut dangling = caches.edit().unwrap();
    dangling.remove_cache("Slicer_State").unwrap();
    assert!(dangling.commit().is_err());

    let mut views = package_snapshot(&package)
        .slicer_views(0)
        .unwrap()
        .unwrap()
        .edit()
        .unwrap();
    views.replace_views(SlicerViews::new()).unwrap();
    apply_slicer(&mut package, views.commit().unwrap().patch());
    apply_slicer(&mut package, &commit.patch().inverse());
    assert!(
        package_snapshot(&package)
            .slicer_caches()
            .unwrap()
            .caches()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn timeline_patch_preserves_noop_and_emits_compact_xml() {
    let mut package = Package::create().unwrap().into_opc();
    let before = Package::from_opc(package.clone())
        .unwrap()
        .to_bytes()
        .unwrap();
    let snapshot = package_snapshot(&package).timeline_caches().unwrap();
    let noop = snapshot.edit().unwrap().commit().unwrap();
    assert!(noop.patch().is_empty());
    apply_timeline(&mut package, noop.patch());
    assert_eq!(
        Package::from_opc(package.clone())
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );

    let mut caches = package_snapshot(&package)
        .timeline_caches()
        .unwrap()
        .edit()
        .unwrap();
    caches.replace_caches(vec![timeline_cache()]).unwrap();
    apply_timeline(&mut package, caches.commit().unwrap().patch());

    let mut decorated = package.clone();
    let cache_part = decorated
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/timelineCaches/"))
        .unwrap()
        .partname()
        .clone();
    let mut xml = b"<?xml version=\"1.0\"?>".to_vec();
    xml.extend_from_slice(decorated.get_part(&cache_part).unwrap().blob());
    decorated.get_part_mut(&cache_part).unwrap().set_blob(xml);
    let mut lexical = package_snapshot(&decorated)
        .timeline_caches()
        .unwrap()
        .edit()
        .unwrap();
    let mut renamed = timeline_cache();
    renamed.name = "TimelineCache2".to_string();
    lexical.replace_caches(vec![renamed]).unwrap();
    assert!(lexical.commit().is_err());

    let mut views = package_snapshot(&package)
        .timeline_views(0)
        .unwrap()
        .unwrap()
        .edit()
        .unwrap();
    views
        .replace_views(TimelineViews {
            items: vec![TimelineView::new(
                "Timeline1",
                "TimelineCache1",
                Level::Month,
            )],
        })
        .unwrap();
    apply_timeline(&mut package, views.commit().unwrap().patch());

    for part in package.iter_parts().filter(|part| {
        part.partname().as_str().contains("/timelineCaches/")
            || part.partname().as_str().contains("/timelines/")
    }) {
        let xml = part.blob();
        assert!(!xml.contains(&b'\n'));
        assert!(!xml.contains(&b'\r'));
        assert!(!xml.windows(3).any(|window| window == b"> <"));
    }
}

#[test]
fn shared_slicer_targets_are_refused_before_package_mutation() {
    let mut package = Package::create().unwrap().into_opc();
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    litchi_xlsb::slicer::package::store_caches(&mut package, &workbook, &[slicer_cache()]).unwrap();

    let cache_target = package
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/slicerCaches/"))
        .unwrap()
        .partname()
        .clone();
    add_inbound(
        &mut package,
        &worksheet,
        &cache_target,
        "rIdSharedSlicerCache",
    );
    let before = package_bytes(&package);
    let mut extra = slicer_cache();
    extra.name = "Slicer_Province".to_string();
    assert!(
        litchi_xlsb::slicer::package::store_caches(
            &mut package,
            &workbook,
            &[slicer_cache(), extra]
        )
        .is_err()
    );
    assert!(litchi_xlsb::slicer::package::store_caches(&mut package, &workbook, &[]).is_err());
    assert_eq!(package_bytes(&package), before);
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .remove("rIdSharedSlicerCache");

    let views = SlicerViews {
        items: vec![SlicerView::new("StateView", "Slicer_State")],
    };
    litchi_xlsb::slicer::package::store_views(&mut package, &worksheet, &views).unwrap();
    let view_target = package
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/slicers/"))
        .unwrap()
        .partname()
        .clone();
    add_inbound(&mut package, &workbook, &view_target, "rIdSharedSlicerView");
    let before = package_bytes(&package);
    let mut replacement = views.clone();
    replacement.items[0].caption = Some("Shared".to_string());
    assert!(
        litchi_xlsb::slicer::package::store_views(&mut package, &worksheet, &replacement).is_err()
    );
    assert!(
        litchi_xlsb::slicer::package::store_views(&mut package, &worksheet, &SlicerViews::new())
            .is_err()
    );
    assert_eq!(package_bytes(&package), before);
}

#[test]
fn shared_timeline_targets_are_refused_before_package_mutation() {
    let mut package = Package::create().unwrap().into_opc();
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    litchi_xlsb::timeline::package::store_caches(&mut package, &workbook, &[timeline_cache()])
        .unwrap();

    let cache_target = package
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/timelineCaches/"))
        .unwrap()
        .partname()
        .clone();
    add_inbound(
        &mut package,
        &worksheet,
        &cache_target,
        "rIdSharedTimelineCache",
    );
    let before = package_bytes(&package);
    let mut extra = timeline_cache();
    extra.name = "TimelineCache2".to_string();
    assert!(
        litchi_xlsb::timeline::package::store_caches(
            &mut package,
            &workbook,
            &[timeline_cache(), extra]
        )
        .is_err()
    );
    assert!(litchi_xlsb::timeline::package::store_caches(&mut package, &workbook, &[]).is_err());
    assert_eq!(package_bytes(&package), before);
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .remove("rIdSharedTimelineCache");

    let views = TimelineViews {
        items: vec![TimelineView::new(
            "Timeline1",
            "TimelineCache1",
            Level::Month,
        )],
    };
    litchi_xlsb::timeline::package::store_views(&mut package, &worksheet, &views).unwrap();
    let view_target = package
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/timelines/"))
        .unwrap()
        .partname()
        .clone();
    add_inbound(
        &mut package,
        &workbook,
        &view_target,
        "rIdSharedTimelineView",
    );
    let before = package_bytes(&package);
    let mut replacement = views.clone();
    replacement.items[0].caption = Some("Shared".to_string());
    assert!(
        litchi_xlsb::timeline::package::store_views(&mut package, &worksheet, &replacement)
            .is_err()
    );
    assert!(
        litchi_xlsb::timeline::package::store_views(
            &mut package,
            &worksheet,
            &TimelineViews::new()
        )
        .is_err()
    );
    assert_eq!(package_bytes(&package), before);
}

#[test]
fn patch_binds_root_relationships_and_inverse_restores_signature_state() {
    let mut package = Package::create().unwrap().into_opc();
    add_signature_marker(&mut package);
    let signed = package_bytes(&package);
    let snapshot = package_snapshot(&package).slicer_caches().unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.replace_caches(vec![slicer_cache()]).unwrap();
    let commit = edit.commit().unwrap();

    let mut stale = package.clone();
    stale
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test:root-change".to_string(),
            "/xl/workbook.bin".to_string(),
            "rIdRootChange".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(
        package_snapshot(&stale)
            .apply_slicer_patch(commit.patch())
            .is_err()
    );

    let changed = package_snapshot(&package)
        .apply_slicer_patch(commit.patch())
        .unwrap();
    assert!(
        !changed
            .opc_package()
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
    );
    let restored = changed
        .apply_slicer_patch(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.to_bytes().unwrap(), signed);
}

#[test]
fn package_error_is_forward_compatible_and_keeps_runtime_semantics() {
    use litchi_xlsb::package::PackageError;
    use std::error::Error as _;

    fn classify(error: &PackageError) -> &'static str {
        match error {
            PackageError::PasswordProtected => "password",
            _ => "other",
        }
    }

    let error = PackageError::PasswordProtected;
    assert_eq!(classify(&error), "password");
    assert_eq!(error.to_string(), "Workbook is password protected");
    assert!(error.source().is_none());
}
