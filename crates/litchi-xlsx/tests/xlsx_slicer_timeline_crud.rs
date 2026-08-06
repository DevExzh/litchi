use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use litchi_xlsx::slicer::{self, Definition, Slicer};
use litchi_xlsx::timeline::{self, CacheDefinition, FilterType, Level, Range, State, View};

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
    slicer::read(
        format!(r#"<slicerCacheDefinition xmlns="{X14}" name="{name}" sourceName="{source}"><data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>"#).as_bytes(),
    )
    .unwrap()
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

    let mut edit = slicer::Transaction::new(&mut package).unwrap();
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
    edit.commit().unwrap();

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
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let mut edit = timeline::Transaction::new(&mut package, &workbook).unwrap();
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
    edit.commit().unwrap();
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
    let mut edit = slicer::Transaction::new(&mut package).unwrap();
    edit.add_cache(slicer_cache("Cache_A", "State")).unwrap();
    let view = edit
        .add_view(&worksheet, Slicer::new("View_A", "Cache_A", 228_600))
        .unwrap();
    edit.commit().unwrap();
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
    let mut edit = slicer::Transaction::new(&mut package).unwrap();
    assert!(edit.remove_view(&worksheet, "View_A").unwrap());
    edit.commit().unwrap();
    assert!(package.get_part(&target).is_ok());
}

#[test]
fn feature_transactions_rollback_and_refuse_host_behavior() {
    let (mut package, worksheet) = package();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let before_workbook = package.get_part(&workbook).unwrap().blob().to_vec();
    {
        let mut edit = slicer::Transaction::new(&mut package).unwrap();
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

    let mut edit = timeline::Transaction::new(&mut package, &workbook).unwrap();
    assert!(matches!(
        edit.apply_filter(),
        Err(litchi_xlsx::Error::Unsupported { .. })
    ));
}
