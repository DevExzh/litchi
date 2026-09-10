use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_xlsx::{
    Error, REFRESH_INTERVALS_NAMESPACE, RefreshInterval, RefreshIntervals, TypeRefreshIntervals,
    apply_rich_value_refresh_patch, edit_rich_value_refresh, load_rich_value_refresh,
    parse_refresh_intervals, write_refresh_intervals,
};

const RICH_DATA_2: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RICH_TYPES_CONTENT_TYPE: &str = "application/vnd.ms-excel.rdRichValuetypes+xml";

fn types_xml() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}" xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}" xmlns:v="urn:vendor">
<rd2:types>
  <rd2:type name="entity"><rd2:keyFlags/><x:extLst>
    <x:ext uri="urn:keep" v:marker="untouched"><v:future a="&quot;">opaque</v:future></x:ext>
    <x:ext uri="urn:refresh"><rr:refreshIntervals><rr:refreshInterval resourceIdInt="7" interval="-1"/><rr:refreshInterval resourceIdStr="a&amp;b" interval="30"/></rr:refreshIntervals></x:ext>
  </x:extLst></rd2:type>
  <rd2:type name="empty"><x:extLst><x:ext uri="urn:other"/></x:extLst></rd2:type>
</rd2:types>
<x:extLst><x:ext uri="urn:root"><v:future/></x:ext></x:extLst>
</rd2:rvTypesInfo>"#
    )
    .into_bytes()
}

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let types_name = PackURI::new("/xl/richValueTypes.xml").unwrap();
    let mut types = BlobPart::new(
        types_name.clone(),
        RICH_TYPES_CONTENT_TYPE.into(),
        types_xml(),
    );
    types
        .rels_mut()
        .try_add_relationship(
            "urn:test:rich-type-resource".into(),
            "../media/image1.png".into(),
            "rIdResource".into(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(types)).unwrap();

    let mut structures = BlobPart::new(
        PackURI::new("/xl/richStructures.xml").unwrap(),
        "application/vnd.ms-excel.rdRichValueStructure+xml".into(),
        b"<rd:rvStructures xmlns:rd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"0\"/>".to_vec(),
    );
    structures
        .rels_mut()
        .try_add_relationship(
            "urn:test:rich-types-owner".into(),
            "richValueTypes.xml".into(),
            "rIdTypes".into(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(structures)).unwrap();
    package
}

fn interval_int(id: i32, interval: i32) -> RefreshInterval {
    RefreshInterval::new(Some(id), None, interval).unwrap()
}

#[test]
fn spec_codec_enforces_identity_and_preserves_interval_semantics() {
    let source = format!(
        r#"<rr:refreshIntervals xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}">
            <rr:refreshInterval resourceIdStr="" interval="0"/>
            <rr:refreshInterval resourceIdInt="-9" resourceIdStr="both" interval="42"/>
        </rr:refreshIntervals>"#
    );
    let parsed = parse_refresh_intervals(source.as_bytes()).unwrap();
    assert_eq!(parsed.len(), 2);
    assert_eq!(parsed.intervals()[0].resource_id_str(), Some(""));
    assert_eq!(parsed.intervals()[1].resource_id_int(), Some(-9));
    assert_eq!(parsed.intervals()[1].interval(), 42);
    let spaced = format!(
        r#"<rr:refreshIntervals xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}"><rr:refreshInterval resourceIdInt=" 9 " interval=" 42 "/></rr:refreshIntervals>"#
    );
    assert_eq!(
        parse_refresh_intervals(spaced.as_bytes())
            .unwrap()
            .intervals()[0]
            .interval(),
        42
    );
    let non_xml_space = format!(
        "<rr:refreshIntervals xmlns:rr=\"{REFRESH_INTERVALS_NAMESPACE}\"><rr:refreshInterval resourceIdInt=\"\u{00a0}9\" interval=\"1\"/></rr:refreshIntervals>"
    );
    assert!(parse_refresh_intervals(non_xml_space.as_bytes()).is_err());
    assert_eq!(
        parse_refresh_intervals(&write_refresh_intervals(&parsed).unwrap()).unwrap(),
        parsed
    );

    assert!(RefreshInterval::new(None, None, 0).is_err());
    assert!(RefreshInterval::new(Some(1), Some("bad\u{0}".into()), 0).is_err());
    assert!(parse_refresh_intervals(
        format!(
            r#"<rr:refreshIntervals xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}"><rr:refreshInterval resourceIdInt="1" interval="&bogus;"/></rr:refreshIntervals>"#
        )
        .as_bytes()
    )
    .is_err());
    let malformed_owner = format!(
        r#"<rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}" xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}"><rd2:types><rd2:type name="entity"><x:extLst><x:ext><rr:refreshIntervals><rr:refreshInterval resourceIdInt="1" interval="1"><rr:unexpected/></rr:refreshInterval></rr:refreshIntervals></x:ext></x:extLst></rd2:type></rd2:types></rd2:rvTypesInfo>"#
    );
    assert!(load_rich_value_refresh(&package_from_xml(malformed_owner)).is_err());
}

#[test]
fn writer_rejects_oversized_aggregate_before_output_materialization() {
    let intervals = (0..700_000)
        .map(|id| RefreshInterval::new(Some(id), None, 30).unwrap())
        .collect::<Vec<_>>();
    let intervals = RefreshIntervals::new(intervals).unwrap();
    assert!(write_refresh_intervals(&intervals).is_err());
}

#[test]
fn owner_scanner_rejects_duplicate_or_stray_refresh_contexts_atomically() {
    let duplicate_types = format!(
        r#"<rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}"><rd2:types/><rd2:types/></rd2:rvTypesInfo>"#
    );
    assert_rejected_owner(duplicate_types);

    let duplicate_ext_lists = format!(
        r#"<rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}"><rd2:types><rd2:type name="entity"><x:extLst/><x:extLst/></rd2:type></rd2:types></rd2:rvTypesInfo>"#
    );
    assert_rejected_owner(duplicate_ext_lists);

    let stray_payload = format!(
        r#"<rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}" xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}"><rd2:types><rd2:type name="entity"><rr:refreshIntervals><rr:refreshInterval resourceIdInt="1" interval="30"/></rr:refreshIntervals><x:extLst><x:ext uri="urn:owner"/></x:extLst></rd2:type></rd2:types></rd2:rvTypesInfo>"#
    );
    assert_rejected_owner(stray_payload);

    let nested_payload = format!(
        r#"<rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}" xmlns:rr="{REFRESH_INTERVALS_NAMESPACE}"><rd2:types><rd2:type name="outer"><x:extLst><x:ext uri="urn:opaque"><rd2:types><rd2:type name="nested"><x:extLst><x:ext><rr:refreshIntervals><rr:refreshInterval resourceIdInt="2" interval="30"/></rr:refreshIntervals></x:ext></x:extLst></rd2:type></rd2:types></x:ext></x:extLst></rd2:type></rd2:types></rd2:rvTypesInfo>"#
    );
    assert_rejected_owner(nested_payload);
}

fn assert_rejected_owner(xml: String) {
    let mut package = package_from_xml(xml);
    let before = package
        .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    assert!(load_rich_value_refresh(&package).is_err());
    assert!(edit_rich_value_refresh(&mut package).is_err());
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
            .unwrap()
            .blob(),
        before.as_slice()
    );
}

fn package_from_xml(xml: String) -> OpcPackage {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/richValueTypes.xml").unwrap(),
            RICH_TYPES_CONTENT_TYPE.into(),
            xml.into_bytes(),
        )))
        .unwrap();
    package
}

#[test]
fn snapshot_and_noop_are_rich_types_scoped_and_source_shared() {
    let mut package = package();
    let before = load_rich_value_refresh(&package).unwrap();
    assert_eq!(before.types().len(), 2);
    assert_eq!(before.types()[0].name(), "entity");
    assert_eq!(before.types()[0].intervals().unwrap().len(), 2);
    assert!(before.types()[1].intervals().is_none());

    // A worksheet-like part is outside this owner and cannot become a
    // refresh metadata target merely because it has a similarly named node.
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
            "application/xml".into(),
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML}"><refreshIntervals xmlns="{REFRESH_INTERVALS_NAMESPACE}"/></worksheet>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    let before = load_rich_value_refresh(&package).unwrap();
    let source = before.source_arc().unwrap();
    let commit = edit_rich_value_refresh(&mut package)
        .unwrap()
        .commit()
        .unwrap();
    assert!(!commit.changed());
    assert!(Arc::ptr_eq(
        &source,
        &commit.snapshot().source_arc().unwrap()
    ));
}

#[test]
fn add_remove_and_inverse_preserve_unknown_extension_bytes_and_relationships() {
    let mut package = package();
    let original = package
        .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let mut transaction = edit_rich_value_refresh(&mut package).unwrap();
    assert!(transaction.add(1, interval_int(99, 120)).unwrap());
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed = package
        .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    assert_ne!(changed, original);
    assert!(
        changed
            .windows(b"urn:keep".len())
            .any(|window| window == b"urn:keep")
    );
    assert!(
        changed
            .windows(b"opaque".len())
            .any(|window| window == b"opaque")
    );
    assert!(
        changed
            .windows(b"resourceIdInt=\"99\"".len())
            .any(|window| window == b"resourceIdInt=\"99\"")
    );
    assert_eq!(commit.snapshot().types()[1].intervals().unwrap().len(), 1);
    assert_eq!(
        commit.patch().before().types()[0]
            .intervals()
            .unwrap()
            .len(),
        2
    );

    apply_rich_value_refresh_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
            .unwrap()
            .blob(),
        original.as_slice()
    );

    let mut transaction = edit_rich_value_refresh(&mut package).unwrap();
    let removed = transaction.remove(0, 0).unwrap().unwrap();
    assert_eq!(removed.resource_id_int(), Some(7));
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(
        String::from_utf8_lossy(
            package
                .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
                .unwrap()
                .blob()
        )
        .contains("resourceIdStr=\"a&amp;b\"")
    );
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
            .unwrap()
            .blob(),
        original.as_slice()
    );
}

#[test]
fn stale_incoming_relationship_and_missing_extension_refuse_atomically() {
    let mut source_package = package();
    let before = source_package
        .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let patch = {
        let mut transaction = edit_rich_value_refresh(&mut source_package).unwrap();
        transaction.add(1, interval_int(4, 10)).unwrap();
        transaction.commit().unwrap().patch().clone()
    };
    let mut stale = package();
    stale
        .get_part_mut(&PackURI::new("/xl/richStructures.xml").unwrap())
        .unwrap()
        .rels_mut()
        .remove("rIdTypes");
    assert!(matches!(
        patch.apply(&mut stale),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(
        stale
            .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
            .unwrap()
            .blob(),
        before.as_slice()
    );

    let mut no_owner = package_without_extension_owner();
    let original = no_owner
        .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let mut transaction = edit_rich_value_refresh(&mut no_owner).unwrap();
    transaction.add(0, interval_int(8, 10)).unwrap();
    let error = transaction.commit();
    assert!(error.is_err());
    assert_eq!(
        no_owner
            .get_part(&PackURI::new("/xl/richValueTypes.xml").unwrap())
            .unwrap()
            .blob(),
        original.as_slice()
    );
}

fn package_without_extension_owner() -> OpcPackage {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/richValueTypes.xml").unwrap(),
            RICH_TYPES_CONTENT_TYPE.into(),
            format!(
                r#"<rd2:rvTypesInfo xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}"><rd2:types><rd2:type name="empty"/></rd2:types></rd2:rvTypesInfo>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    package
}

#[allow(dead_code)]
fn _public_shape(_: &TypeRefreshIntervals) {}
