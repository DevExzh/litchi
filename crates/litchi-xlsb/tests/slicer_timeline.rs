#![allow(
    clippy::pedantic,
    clippy::expect_used,
    clippy::map_err_ignore,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    clippy::wildcard_enum_match_arm,
    reason = "integration tests use panic-on-failure extraction and successive names that mirror package transaction states"
)]

//! Transactional XLSB slicer/timeline package-owner regressions.

use litchi_opc::{BlobPart, PackURI, Part, TargetMode};
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

fn authored_package() -> litchi_opc::OpcPackage {
    let source = Package::new().unwrap().into_opc();
    let mut package = litchi_opc::OpcPackage::new();
    for part in source.try_iter_parts() {
        package.add_part(part.unwrap().clone_part());
    }
    *package.rels_mut() = source.rels().clone();
    package
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

fn slicer_views() -> SlicerViews {
    SlicerViews {
        items: vec![SlicerView::new("StateView", "Slicer_State")],
    }
}

fn timeline_views() -> TimelineViews {
    TimelineViews {
        items: vec![TimelineView::new(
            "Timeline1",
            "TimelineCache1",
            Level::Month,
        )],
    }
}

/// `package` with the slicer and timeline caches and views of sheet 0.
fn with_owners(package: &Package) -> Package {
    let mut slicers = package.edit_slicers().unwrap();
    slicers.set_caches(vec![slicer_cache()]).unwrap();
    slicers.set_views(0, slicer_views()).unwrap();
    let package = slicers.commit().unwrap().into_package();
    let mut timelines = package.edit_timelines().unwrap();
    timelines.set_caches(vec![timeline_cache()]).unwrap();
    timelines.set_views(0, timeline_views()).unwrap();
    timelines.commit().unwrap().into_package()
}

#[test]
fn semantic_facades_resolve_one_actual_root_workbook_relationship() {
    let mut missing = Package::new().unwrap().into_opc();
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
    let slicers = missing.edit_slicers().expect("cold slicer entry");
    assert!(slicers.caches().is_err());
    let timelines = missing.edit_timelines().expect("cold timeline entry");
    assert!(timelines.caches().is_err());

    let mut ambiguous = Package::new().unwrap().into_opc();
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
    let slicers = ambiguous.edit_slicers().expect("cold slicer entry");
    assert!(slicers.caches().is_err());
    let timelines = ambiguous.edit_timelines().expect("cold timeline entry");
    assert!(timelines.caches().is_err());

    let mut alternate = with_owners(&Package::new().unwrap()).into_opc();
    relocate_workbook(
        &mut alternate,
        &PackURI::new("/xl/alternate-workbook.bin").unwrap(),
    );

    let alternate = Package::from_opc(alternate).unwrap();
    let slicers = alternate.edit_slicers().unwrap();
    assert_eq!(slicers.caches().unwrap().len(), 1);
    assert!(!slicers.views(0).unwrap().items.is_empty());
    let timelines = alternate.edit_timelines().unwrap();
    assert_eq!(timelines.caches().unwrap().len(), 1);
    assert!(!timelines.views(0).unwrap().items.is_empty());
}

#[test]
fn workbook_facade_round_trips_and_removes_both_inert_owners() {
    let cache = slicer_cache();
    let mut slicer_view = SlicerView::new("StateView", "Slicer_State");
    slicer_view.caption = Some("State".to_string());
    let mut slicers = Package::new().unwrap().edit_slicers().unwrap();
    slicers.set_caches(vec![cache.clone()]).unwrap();
    slicers
        .set_views(
            0,
            SlicerViews {
                items: vec![slicer_view],
            },
        )
        .unwrap();
    let package = slicers.commit().unwrap().into_package();

    let timeline = timeline_cache();
    let mut timelines = package.edit_timelines().unwrap();
    timelines.set_caches(vec![timeline.clone()]).unwrap();
    timelines.set_views(0, timeline_views()).unwrap();
    let package = timelines.commit().unwrap().into_package();

    let workbook = package.workbook().unwrap();
    assert_eq!(workbook.slicer_caches().unwrap(), vec![cache.clone()]);
    assert_eq!(workbook.timeline_caches().unwrap(), vec![timeline.clone()]);
    assert!(workbook.slicers(0).unwrap().is_some());
    assert!(workbook.timelines(0).unwrap().is_some());

    let mut bytes = Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let reopened = Package::from_bytes(&bytes.into_inner()).unwrap();
    let reopened_workbook = reopened.workbook().unwrap();
    assert_eq!(reopened_workbook.slicer_caches().unwrap(), vec![cache]);
    assert_eq!(reopened_workbook.timeline_caches().unwrap(), vec![timeline]);
    assert!(reopened_workbook.slicers(0).unwrap().is_some());
    assert!(reopened_workbook.timelines(0).unwrap().is_some());

    let mut slicers = package.edit_slicers().unwrap();
    assert!(slicers.remove_views(0).unwrap());
    assert!(slicers.remove_caches().unwrap());
    let package = slicers.commit().unwrap().into_package();
    let mut timelines = package.edit_timelines().unwrap();
    assert!(timelines.remove_views(0).unwrap());
    assert!(timelines.remove_caches().unwrap());
    let package = timelines.commit().unwrap().into_package();
    let workbook = package.workbook().unwrap();
    assert!(workbook.slicer_caches().unwrap().is_empty());
    assert!(workbook.timeline_caches().unwrap().is_empty());
    assert!(workbook.slicers(0).unwrap().is_none());
    assert!(workbook.timelines(0).unwrap().is_none());
}

#[test]
fn slicer_patch_is_source_checked_reversible_and_dependency_safe() {
    let source = Package::new().unwrap();
    let mut edit = source.edit_slicers().unwrap();
    edit.set_caches(vec![slicer_cache()]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_noop());
    let package = source.apply_slicers(commit.patch()).unwrap().into_package();
    assert!(package.apply_slicers(commit.patch()).is_err());

    let mut edit = package.edit_slicers().unwrap();
    edit.set_views(0, slicer_views()).unwrap();
    let package = edit.commit().unwrap().into_package();

    let mut dangling = package.edit_slicers().unwrap();
    assert!(dangling.remove_cache("Slicer_State").is_err());

    let mut views = package.edit_slicers().unwrap();
    views.set_views(0, SlicerViews::new()).unwrap();
    let package = views.commit().unwrap().into_package();
    let package = package
        .apply_slicers(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert!(
        package
            .workbook()
            .unwrap()
            .slicer_caches()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn timeline_patch_preserves_noop_and_emits_compact_xml() {
    let source = Package::new().unwrap();
    let before = source.to_bytes().unwrap();
    let noop = source.edit_timelines().unwrap().commit().unwrap();
    assert!(noop.patch().is_noop());
    let package = source.apply_timelines(noop.patch()).unwrap().into_package();
    assert_eq!(package.to_bytes().unwrap(), before);

    let mut caches = package.edit_timelines().unwrap();
    caches.set_caches(vec![timeline_cache()]).unwrap();
    let package = caches.commit().unwrap().into_package();

    let mut decorated = package.opc_package().clone();
    let cache_part = decorated
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/timelineCaches/"))
        .unwrap()
        .partname()
        .clone();
    let mut xml = b"<?xml version=\"1.0\"?>".to_vec();
    xml.extend_from_slice(decorated.get_part(&cache_part).unwrap().blob());
    decorated.get_part_mut(&cache_part).unwrap().set_blob(xml);
    let decorated = Package::from_opc(decorated).unwrap();
    let mut lexical = decorated.edit_timelines().unwrap();
    let mut renamed = timeline_cache();
    renamed.name = "TimelineCache2".to_string();
    assert!(lexical.set_caches(vec![renamed]).is_err());

    let mut views = package.edit_timelines().unwrap();
    views.set_views(0, timeline_views()).unwrap();
    let package = views.commit().unwrap().into_package();

    let opc = package.opc_package();
    let timeline_names = opc
        .iter_parts()
        .filter(|part| {
            part.partname().as_str().contains("/timelineCaches/")
                || part.partname().as_str().contains("/timelines/")
        })
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    for name in &timeline_names {
        let part = opc.get_part(name).unwrap();
        let xml = part.blob();
        assert!(!xml.contains(&b'\n'));
        assert!(!xml.contains(&b'\r'));
        assert!(!xml.windows(3).any(|window| window == b"> <"));
    }
}

#[test]
fn shared_slicer_targets_are_refused_before_package_mutation() {
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let mut edit = Package::new().unwrap().edit_slicers().unwrap();
    edit.set_caches(vec![slicer_cache()]).unwrap();
    let mut package = edit.commit().unwrap().into_package().into_opc();

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
    let shared = Package::from_opc(package.clone()).unwrap();
    let before = shared.to_bytes().unwrap();
    let mut extra = slicer_cache();
    extra.name = "Slicer_Province".to_string();
    let mut edit = shared.edit_slicers().unwrap();
    assert!(edit.set_caches(vec![slicer_cache(), extra]).is_err());
    assert!(edit.set_caches(Vec::new()).is_err());
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .remove("rIdSharedSlicerCache");

    let mut edit = Package::from_opc(package).unwrap().edit_slicers().unwrap();
    edit.set_views(0, slicer_views()).unwrap();
    let mut package = edit.commit().unwrap().into_package().into_opc();
    let view_target = package
        .iter_parts()
        .find(|part| part.partname().as_str().contains("/slicers/"))
        .unwrap()
        .partname()
        .clone();
    add_inbound(&mut package, &workbook, &view_target, "rIdSharedSlicerView");
    let shared = Package::from_opc(package).unwrap();
    let before = shared.to_bytes().unwrap();
    let mut replacement = slicer_views();
    replacement.items[0].caption = Some("Shared".to_string());
    let mut edit = shared.edit_slicers().unwrap();
    assert!(edit.set_views(0, replacement).is_err());
    assert!(edit.set_views(0, SlicerViews::new()).is_err());
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
}

#[test]
fn shared_timeline_targets_are_refused_before_package_mutation() {
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let mut edit = Package::new().unwrap().edit_timelines().unwrap();
    edit.set_caches(vec![timeline_cache()]).unwrap();
    let mut package = edit.commit().unwrap().into_package().into_opc();

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
    let shared = Package::from_opc(package.clone()).unwrap();
    let before = shared.to_bytes().unwrap();
    let mut extra = timeline_cache();
    extra.name = "TimelineCache2".to_string();
    let mut edit = shared.edit_timelines().unwrap();
    assert!(edit.set_caches(vec![timeline_cache(), extra]).is_err());
    assert!(edit.set_caches(Vec::new()).is_err());
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .remove("rIdSharedTimelineCache");

    let mut edit = Package::from_opc(package)
        .unwrap()
        .edit_timelines()
        .unwrap();
    edit.set_views(0, timeline_views()).unwrap();
    let mut package = edit.commit().unwrap().into_package().into_opc();
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
    let shared = Package::from_opc(package).unwrap();
    let before = shared.to_bytes().unwrap();
    let mut replacement = timeline_views();
    replacement.items[0].caption = Some("Shared".to_string());
    let mut edit = shared.edit_timelines().unwrap();
    assert!(edit.set_views(0, replacement).is_err());
    assert!(edit.set_views(0, TimelineViews::new()).is_err());
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
}

/// A changed slicer commit over a signed package is refused until the
/// signature is removed explicitly (ADR 0006); the patch's inverse restores
/// the explicitly unsigned source byte for byte.
#[test]
fn patch_needs_an_explicit_unsign_and_inverse_is_exact() {
    let mut package = authored_package();
    add_signature_marker(&mut package);
    let signed = Package::from_opc(package).unwrap();
    let mut edit = signed.edit_slicers().unwrap();
    let staged = edit.set_caches(vec![slicer_cache()]);
    assert!(matches!(
        staged.and_then(|_| edit.commit().map(|_| ())),
        Err(litchi_xlsb::package::PackageError::Signed)
    ));

    let unsigned_package = signed.unsign();
    let unsigned = unsigned_package.to_bytes().unwrap();
    let mut edit = unsigned_package.edit_slicers().unwrap();
    edit.set_caches(vec![slicer_cache()]).unwrap();
    let commit = edit.commit().unwrap();

    let changed = unsigned_package
        .apply_slicers(commit.patch())
        .unwrap()
        .into_package();
    assert!(
        !changed
            .opc_package()
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
    );
    let restored = changed
        .apply_slicers(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), unsigned);
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
