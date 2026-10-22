use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, SourceLimits};
use litchi_xlsx::slicer::{self, Definition, Slicer};
use litchi_xlsx::timeline::{CacheDefinition, FilterType, Level, Range, State, View};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

fn package() -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let mut workbook_part = BlobPart::new(
        workbook,
        ct::SML_SHEET_MAIN.into(),
        format!(r#"<workbook xmlns="{SML}" xmlns:r="{R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets><extLst><ext uri="urn:unrelated"><vendor xmlns="urn:test" value="kept"/></ext></extLst></workbook>"#).into_bytes(),
    );
    workbook_part.rels_mut().add_relationship(
        rt::WORKSHEET.into(),
        "worksheets/sheet1.xml".into(),
        "rIdSheet".into(),
        false,
    );
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(BlobPart::new(
        worksheet.clone(),
        ct::SML_WORKSHEET.into(),
        format!(r#"<worksheet xmlns="{SML}"><sheetData/><extLst><ext uri="urn:sheet-unrelated"><vendor xmlns="urn:test" value="kept"/></ext></extLst></worksheet>"#).into_bytes(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    (package, worksheet)
}

fn slicer_cache(name: &str, source: &str) -> Definition {
    slicer::read(slicer_cache_xml(name, source).as_bytes()).unwrap()
}

fn slicer_cache_xml(name: &str, source: &str) -> String {
    format!(
        r#"<slicerCacheDefinition xmlns="{X14}" name="{name}" sourceName="{source}"><data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>"#
    )
}

fn timeline_cache(name: &str, source: &str) -> CacheDefinition {
    CacheDefinition {
        name: name.into(),
        uid: None,
        source_name: source.into(),
        pivot_tables: Vec::new(),
        state: State {
            selection: Some(Range::new("2026-01-01T00:00:00Z", "2026-01-31T23:59:59Z").unwrap()),
            bounds: Range::new("2026-01-01T00:00:00Z", "2026-12-31T23:59:59Z").unwrap(),
            extension_list: None,
            single_range_filter_state: Some(true),
            minimal_refresh_version: 0,
            last_refresh_version: 1,
            pivot_cache_id: 1,
            filter_type: FilterType::DateBetween,
        },
        timeline_pivot_filter: None,
        extension_list: None,
    }
}

fn timeline(name: &str, cache: &str) -> View {
    View {
        name: name.into(),
        uid: None,
        cache: cache.into(),
        caption: Some(name.into()),
        show_header: Some(true),
        show_selection_label: Some(true),
        show_time_level: Some(true),
        show_horizontal_scrollbar: Some(true),
        level: Level::Month,
        selection_level: Level::Day,
        scroll_position: Some("2026-01-01T00:00:00Z".into()),
        style: Some("TimelineStyleLight1".into()),
        extension_list: None,
    }
}

#[test]
fn raw_slicer_insert_preserves_identity_replay_inverse_and_untouched_bytes() {
    let (mut package, worksheet) = package();
    let unrelated = PackURI::new("/xl/unrelated.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        unrelated.clone(),
        "application/octet-stream".into(),
        b"unchanged physical payload".to_vec(),
    )));
    let before_workbook = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let before_worksheet = package.get_part(&worksheet).unwrap().blob().to_vec();
    let before_unrelated = package.get_part(&unrelated).unwrap().blob().to_vec();
    let source = litchi_xlsx::Package::from_opc(package).unwrap();
    let limits = SourceLimits::builder()
        .max_relationship_target_bytes(29)
        .unwrap()
        .build()
        .unwrap();
    let tight_limits = SourceLimits::builder()
        .max_total_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let mut limited = source
        .edit_slicers_with_options(slicer::EditOptions::default().limits(tight_limits))
        .unwrap();
    assert!(
        limited
            .insert_cache(slicer::raw::Cache {
                relationship_id: "rIdRejectedSlicerCache".into(),
                part_name: "/xl/slicerCaches/rejected.xml".into(),
                definition: slicer_cache("RejectedCache", "State"),
            },)
            .is_err()
    );
    assert!(limited.cache("RejectedCache").unwrap().is_none());
    let existing_max = before_workbook
        .len()
        .max(before_worksheet.len())
        .max(before_unrelated.len());
    let part_limited = SourceLimits::builder()
        .max_part_bytes((existing_max + 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    let mut writer_limited = source
        .edit_slicers_with_options(slicer::EditOptions::default().limits(part_limited))
        .unwrap();
    assert!(
        writer_limited
            .insert_cache(slicer::raw::Cache {
                relationship_id: "rIdRejectedWriter".into(),
                part_name: "/xl/slicerCaches/rejected-writer.xml".into(),
                definition: slicer_cache("RejectedWriter", "State"),
            },)
            .is_err()
    );
    assert!(writer_limited.cache("RejectedWriter").unwrap().is_none());
    let mut edit = source
        .edit_slicers_with_options(slicer::EditOptions::default().limits(limits))
        .unwrap();
    edit.insert_cache(slicer::raw::Cache {
        relationship_id: "rIdNativeSlicerCache".into(),
        part_name: "/xl/slicerCaches/native-cache.xml".into(),
        definition: slicer::raw::parse_cache(
            slicer_cache_xml("NativeCache", "State").as_bytes(),
            &limits,
        )
        .unwrap(),
    })
    .unwrap();
    edit.insert_part(
        &worksheet,
        slicer::raw::Part {
            relationship_id: "rIdNativeSlicer".into(),
            part_name: "/xl/slicers/native.xml".into(),
            slicers: slicer::raw::parse_part(
                format!(r#"<x14:slicers xmlns:x14="{X14}"><x14:slicer name="NativeView" cache="NativeCache" rowHeight="228600"/></x14:slicers>"#).as_bytes(),
                &limits,
            )
            .unwrap(),
        },
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let patch = commit.patch().clone();
    let after = commit.into_package();
    let after_opc = after.clone().into_plain_opc();
    let workbook = after_opc
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    let cache_relationship = workbook.rels().get("rIdNativeSlicerCache").unwrap();
    assert_eq!(
        cache_relationship.target_partname().unwrap().as_str(),
        "/xl/slicerCaches/native-cache.xml"
    );
    let worksheet_part = after_opc.get_part(&worksheet).unwrap();
    let slicer_relationship = worksheet_part.rels().get("rIdNativeSlicer").unwrap();
    assert_eq!(
        slicer_relationship.target_partname().unwrap().as_str(),
        "/xl/slicers/native.xml"
    );
    assert_eq!(
        after_opc
            .get_part(&PackURI::new("/xl/slicerCaches/native-cache.xml").unwrap())
            .unwrap()
            .content_type(),
        slicer::CACHE_CONTENT_TYPE
    );
    assert_eq!(
        after_opc
            .get_part(&PackURI::new("/xl/slicers/native.xml").unwrap())
            .unwrap()
            .content_type(),
        "application/vnd.ms-excel.slicer+xml"
    );
    assert!(
        std::str::from_utf8(workbook.blob())
            .unwrap()
            .contains("urn:unrelated")
    );
    assert!(
        std::str::from_utf8(worksheet_part.blob())
            .unwrap()
            .contains("urn:sheet-unrelated")
    );
    assert_eq!(
        after_opc.get_part(&unrelated).unwrap().blob(),
        before_unrelated.as_slice()
    );

    let replayed = source.apply_slicers(&patch).unwrap().into_package();
    assert!(replayed.same_content(&after).unwrap());
    let restored = after
        .apply_slicers(&patch.inverse())
        .unwrap()
        .into_package();
    let restored_opc = restored.into_plain_opc();
    assert_eq!(
        restored_opc
            .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .blob(),
        before_workbook.as_slice()
    );
    assert_eq!(
        restored_opc.get_part(&worksheet).unwrap().blob(),
        before_worksheet.as_slice()
    );
    assert_eq!(
        restored_opc.get_part(&unrelated).unwrap().blob(),
        before_unrelated.as_slice()
    );

    let noop = source.edit_slicers().unwrap().commit().unwrap();
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
}

#[test]
fn raw_timeline_insert_preserves_identity_replay_inverse_and_untouched_bytes() {
    let (mut package, worksheet) = package();
    let unrelated = PackURI::new("/xl/unrelated.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        unrelated.clone(),
        "application/octet-stream".into(),
        b"unchanged physical payload".to_vec(),
    )));
    let before_workbook = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let before_worksheet = package.get_part(&worksheet).unwrap().blob().to_vec();
    let before_unrelated = package.get_part(&unrelated).unwrap().blob().to_vec();
    let source = litchi_xlsx::Package::from_opc(package).unwrap();
    let limits = SourceLimits::builder()
        .max_relationship_target_bytes(31)
        .unwrap()
        .build()
        .unwrap();
    let tight_limits = SourceLimits::builder()
        .max_total_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let rejected = litchi_xlsx::timeline::raw::Cache {
        relationship_id: "rIdRejectedTimelineCache".into(),
        part_name: "/xl/timelineCaches/rejected.xml".into(),
        definition: timeline_cache("RejectedTimelineCache", "Date"),
    };
    let mut limited = source
        .edit_timelines_with_options(
            litchi_xlsx::timeline::EditOptions::default().limits(tight_limits),
        )
        .unwrap();
    assert!(limited.insert_caches(&[rejected]).is_err());
    let mut edit = source
        .edit_timelines_with_options(litchi_xlsx::timeline::EditOptions::default().limits(limits))
        .unwrap();
    let cache_value = timeline_cache("NativeTimelineCache", "Date");
    let cache_xml = litchi_xlsx::timeline::codec::write_cache(&cache_value).unwrap();
    edit.insert_caches(&[litchi_xlsx::timeline::raw::Cache {
        relationship_id: "rIdNativeTimelineCache".into(),
        part_name: "/xl/timelineCaches/native-cache.xml".into(),
        definition: litchi_xlsx::timeline::raw::parse_cache(&cache_xml, &limits).unwrap(),
    }])
    .unwrap();
    let views_value = litchi_xlsx::timeline::Views {
        timelines: vec![timeline("NativeTimeline", "NativeTimelineCache")],
    };
    let views_xml = litchi_xlsx::timeline::codec::write_views(&views_value).unwrap();
    edit.insert_part(litchi_xlsx::timeline::raw::Part {
        worksheet_part_name: worksheet.to_string(),
        relationship_id: "rIdNativeTimeline".into(),
        part_name: "/xl/timelines/native.xml".into(),
        timelines: litchi_xlsx::timeline::raw::parse_part(&views_xml, &limits).unwrap(),
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let patch = commit.patch().clone();
    let after = commit.into_package();
    let after_opc = after.clone().into_plain_opc();
    let worksheet_part = after_opc.get_part(&worksheet).unwrap();
    let relationship = worksheet_part.rels().get("rIdNativeTimeline").unwrap();
    assert_eq!(
        relationship.target_partname().unwrap().as_str(),
        "/xl/timelines/native.xml"
    );
    let workbook = after_opc
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    let cache_relationship = workbook.rels().get("rIdNativeTimelineCache").unwrap();
    assert_eq!(
        cache_relationship.target_partname().unwrap().as_str(),
        "/xl/timelineCaches/native-cache.xml"
    );
    assert!(
        std::str::from_utf8(workbook.blob())
            .unwrap()
            .contains("urn:unrelated")
    );
    assert!(
        std::str::from_utf8(worksheet_part.blob())
            .unwrap()
            .contains("urn:sheet-unrelated")
    );
    assert_eq!(
        after_opc.get_part(&unrelated).unwrap().blob(),
        before_unrelated.as_slice()
    );

    let replayed = source.apply_timelines(&patch).unwrap().into_package();
    assert!(replayed.same_content(&after).unwrap());
    let restored = after
        .apply_timelines(&patch.inverse())
        .unwrap()
        .into_package();
    let restored_opc = restored.into_plain_opc();
    assert_eq!(
        restored_opc
            .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .blob(),
        before_workbook.as_slice()
    );
    assert_eq!(
        restored_opc.get_part(&worksheet).unwrap().blob(),
        before_worksheet.as_slice()
    );
    assert_eq!(
        restored_opc.get_part(&unrelated).unwrap().blob(),
        before_unrelated.as_slice()
    );

    let noop = source.edit_timelines().unwrap().commit().unwrap();
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
}

#[test]
fn raw_identity_admission_uses_archive_member_spelling_and_global_relationship_budget() {
    let (package, _) = package();
    let source = litchi_xlsx::Package::from_opc(package).unwrap();
    let target =
        PackURI::new("/xl/slicerCaches/member-name-that-is-longer-than-existing-rels.xml").unwrap();
    let member_limit = target.membername().len() as u64;
    let value = || slicer::raw::Cache {
        relationship_id: "rIdMemberName".into(),
        part_name: target.to_string(),
        definition: slicer_cache("MemberName", "State"),
    };

    // The absolute URI is one byte longer than the archive member spelling;
    // a limit equal to the member must still admit this valid OPC path.
    let mut valid = source
        .edit_slicers_with_options(
            slicer::EditOptions::default().limits(
                SourceLimits::builder()
                    .max_archive_member_name_bytes(member_limit)
                    .unwrap()
                    .build()
                    .unwrap(),
            ),
        )
        .unwrap();
    valid.insert_cache(value()).unwrap();

    let mut rejected = source
        .edit_slicers_with_options(
            slicer::EditOptions::default().limits(
                SourceLimits::builder()
                    .max_archive_member_name_bytes(member_limit - 1)
                    .unwrap()
                    .build()
                    .unwrap(),
            ),
        )
        .unwrap();
    assert!(rejected.insert_cache(value()).is_err());

    // The root-to-workbook and workbook-to-sheet edges already consume the
    // two-edge policy. Admission must include other owners when charging the
    // new workbook relationship.
    let mut aggregate = source
        .edit_slicers_with_options(
            slicer::EditOptions::default().limits(
                SourceLimits::builder()
                    .max_total_relationships(2)
                    .unwrap()
                    .build()
                    .unwrap(),
            ),
        )
        .unwrap();
    assert!(aggregate.insert_cache(value()).is_err());
}

#[test]
fn slicer_cache_and_view_crud_preserves_state_extensions_and_shared_targets() {
    let (mut package, worksheet) = package();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/slicerCaches/slicerCache1.xml").unwrap(),
        "application/octet-stream".into(),
        b"occupied".to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/slicers/slicer1.xml").unwrap(),
        "application/octet-stream".into(),
        b"occupied".to_vec(),
    )));
    package
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:occupied".into(),
            "occupied.bin".into(),
            "rIdSlicerCache1".into(),
            false,
        );
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:occupied".into(),
            "occupied.bin".into(),
            "rIdSlicer1".into(),
            false,
        );

    let mut edit = litchi_xlsx::Package::from_opc(package.clone())
        .unwrap()
        .edit_slicers()
        .unwrap();
    let cache_a = edit.add_cache(slicer_cache("Cache_A", "State")).unwrap();
    let cache_b = edit.add_cache(slicer_cache("Cache_B", "City")).unwrap();
    assert_ne!(cache_a.part_name, "/xl/slicerCaches/slicerCache1.xml");
    edit.reorder_caches(&["Cache_B".into(), "Cache_A".into()])
        .unwrap();
    let mut view_a = slicer::read_views(
        format!(r#"<x14:slicers xmlns:x14="{X14}"><x14:slicer name="View_A" cache="Cache_A" rowHeight="228600"><x14:extLst><x14:ext uri="urn:test"><v:payload xmlns:v="urn:vendor"/></x14:ext></x14:extLst></x14:slicer></x14:slicers>"#).as_bytes(),
    )
    .unwrap()
    .slicers
    .remove(0);
    view_a.caption = Some("State".into());
    let view_part = edit.add_view(&worksheet, view_a).unwrap();
    assert_ne!(view_part.part_name, "/xl/slicers/slicer1.xml");
    edit.add_view(&worksheet, Slicer::new("View_B", "Cache_B", 228_600))
        .unwrap();
    edit.update_view(&worksheet, "View_A", |view| {
        view.caption = Some("Updated State".into());
        view.column_count = 2;
    })
    .unwrap();
    let mut replacement = Slicer::new("View_B", "Cache_B", 300_000);
    replacement.style = Some("SlicerStyleLight2".into());
    edit.replace_view(&worksheet, "View_B", replacement)
        .unwrap();
    edit.reorder_views(&worksheet, &["View_B".into(), "View_A".into()])
        .unwrap();
    assert!(
        edit.view(&worksheet, "View_A")
            .unwrap()
            .unwrap()
            .extension_list
            .is_some()
    );
    assert!(edit.remove_cache("Cache_A").is_err());
    edit.update_cache("Cache_A", |cache| cache.source_name = "Region".into())
        .unwrap();
    let mut cache_b_replacement = slicer_cache("Cache_B", "Town");
    cache_b_replacement.uid = None;
    edit.replace_cache("Cache_B", cache_b_replacement).unwrap();
    assert_eq!(
        edit.cache("Cache_A")
            .unwrap()
            .unwrap()
            .definition
            .source_name,
        "Region"
    );
    assert!(edit.remove_view(&worksheet, "View_B").unwrap());
    assert!(edit.remove_cache("Cache_B").unwrap());
    assert!(edit.remove_view(&worksheet, "View_A").unwrap());
    assert!(edit.remove_cache("Cache_A").unwrap());
    package = edit.commit().unwrap().into_package().into_plain_opc();

    let target = PackURI::new(&view_part.part_name).unwrap();
    assert!(package.get_part(&target).is_err());
    assert!(cache_b.part_name.contains("slicerCache"));
    assert!(
        std::str::from_utf8(
            package
                .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
                .unwrap()
                .blob()
        )
        .unwrap()
        .contains("urn:unrelated")
    );
}

#[test]
fn timeline_cache_and_view_crud_preserves_selection_and_reference_integrity() {
    let (mut package, worksheet) = package();
    let mut edit = litchi_xlsx::Package::from_opc(package.clone())
        .unwrap()
        .edit_timelines()
        .unwrap();
    let cache_a = edit
        .add_cache(timeline_cache("Timeline_A", "Date"))
        .unwrap();
    edit.add_cache(timeline_cache("Timeline_B", "ShipDate"))
        .unwrap();
    edit.reorder_caches(&["Timeline_B".into(), "Timeline_A".into()])
        .unwrap();
    edit.add_view(&worksheet, timeline("Timeline_View_A", "Timeline_A"))
        .unwrap();
    edit.add_view(&worksheet, timeline("Timeline_View_B", "Timeline_B"))
        .unwrap();
    edit.update_view(&worksheet, "Timeline_View_A", |view| {
        view.caption = Some("Updated Date".into());
        view.selection_level = Level::Month;
    })
    .unwrap();
    edit.replace_view(
        &worksheet,
        "Timeline_View_B",
        timeline("Timeline_View_B", "Timeline_B"),
    )
    .unwrap();
    edit.reorder_views(
        &worksheet,
        &["Timeline_View_B".into(), "Timeline_View_A".into()],
    )
    .unwrap();
    assert_eq!(
        edit.view(&worksheet, "Timeline_View_A")
            .unwrap()
            .unwrap()
            .caption
            .as_deref(),
        Some("Updated Date")
    );
    assert!(edit.remove_cache("Timeline_A").is_err());
    edit.update_cache("Timeline_A", |cache| {
        cache.state.last_refresh_version = 2;
    })
    .unwrap();
    edit.replace_cache("Timeline_B", timeline_cache("Timeline_B", "OrderDate"))
        .unwrap();
    assert_eq!(
        edit.cache("Timeline_A")
            .unwrap()
            .unwrap()
            .definition
            .state
            .selection
            .unwrap()
            .start_date,
        "2026-01-01T00:00:00Z"
    );
    assert!(edit.remove_view(&worksheet, "Timeline_View_B").unwrap());
    assert!(edit.remove_cache("Timeline_B").unwrap());
    assert!(edit.remove_view(&worksheet, "Timeline_View_A").unwrap());
    assert!(edit.remove_cache("Timeline_A").unwrap());
    package = edit.commit().unwrap().into_package().into_plain_opc();
    assert!(cache_a.part_name.contains("timelineCache"));
    let workbook_xml = std::str::from_utf8(
        package
            .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(workbook_xml.contains("urn:unrelated"));
    let worksheet_xml = std::str::from_utf8(package.get_part(&worksheet).unwrap().blob()).unwrap();
    assert!(worksheet_xml.contains("urn:sheet-unrelated"));
}

#[test]
fn removing_last_slicer_keeps_a_target_shared_by_an_unrelated_owner() {
    let (mut package, worksheet) = package();
    let mut edit = litchi_xlsx::Package::from_opc(package.clone())
        .unwrap()
        .edit_slicers()
        .unwrap();
    edit.add_cache(slicer_cache("Cache_A", "State")).unwrap();
    let view = edit
        .add_view(&worksheet, Slicer::new("View_A", "Cache_A", 228_600))
        .unwrap();
    package = edit.commit().unwrap().into_package().into_plain_opc();
    let target = PackURI::new(&view.part_name).unwrap();
    let mut owner = BlobPart::new(
        PackURI::new("/xl/shared-slicer-owner.xml").unwrap(),
        "application/xml".into(),
        Vec::new(),
    );
    owner.relate_to(
        &target.relative_ref("/xl/"),
        "urn:test:shared-slicer-target",
    );
    package.add_part(Box::new(owner));
    let mut edit = litchi_xlsx::Package::from_opc(package.clone())
        .unwrap()
        .edit_slicers()
        .unwrap();
    assert!(edit.remove_view(&worksheet, "View_A").unwrap());
    package = edit.commit().unwrap().into_package().into_plain_opc();
    assert!(package.get_part(&target).is_ok());
}

#[test]
fn feature_transactions_rollback_and_refuse_host_behavior() {
    let (package, worksheet) = package();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let before_workbook = package.get_part(&workbook).unwrap().blob().to_vec();
    {
        let mut edit = litchi_xlsx::Package::from_opc(package.clone())
            .unwrap()
            .edit_slicers()
            .unwrap();
        edit.add_cache(slicer_cache("Cache_A", "State")).unwrap();
        assert!(
            edit.add_view(&worksheet, Slicer::new("View_A", "Missing", 228_600))
                .is_err()
        );
        assert!(matches!(
            edit.apply_filter(),
            Err(litchi_xlsx::Error::Unsupported { .. })
        ));
    }
    assert_eq!(
        package.get_part(&workbook).unwrap().blob(),
        before_workbook.as_slice()
    );

    let mut edit = litchi_xlsx::Package::from_opc(package.clone())
        .unwrap()
        .edit_timelines()
        .unwrap();
    assert!(matches!(
        edit.apply_filter(),
        Err(litchi_xlsx::Error::Unsupported { .. })
    ));
}
