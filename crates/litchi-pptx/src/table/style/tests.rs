#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::semantic_xml_eq;
use super::model::{Conformance, Def, Id, List, Parts};
use super::validation::{load, present, put, remove};
use super::{A, AS, P, PS, default_xml};
use litchi_opc::constants::{content_type as ct, relationship_type, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};

const DEFAULT: &str = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const FIRST: &str = "{5940675A-B579-460E-94D1-54222C63F5DA}";
const SECOND: &str = "{11111111-2222-3333-4444-555555555555}";

#[test]
fn ids_are_typed_case_insensitive_and_canonical() {
    let upper = Id::parse(FIRST).unwrap();
    let lower = Id::parse("{5940675a-b579-460e-94d1-54222c63f5da}").unwrap();
    assert_eq!(upper, lower);
    assert_eq!(upper.to_string(), FIRST);
    for invalid in [
        "5940675A-B579-460E-94D1-54222C63F5DA",
        "{5940675A-B579-460E-94D1-54222C63F5D}",
        "{5940675A_B579-460E-94D1-54222C63F5DA}",
        "{5940675Z-B579-460E-94D1-54222C63F5DA}",
    ] {
        assert!(Id::parse(invalid).is_err(), "accepted {invalid}");
    }
}

#[test]
fn source_bytes_duplicate_names_background_and_extensions_are_preserved() {
    let xml = format!(
        r#"<?xml version="1.0"?><d:tblStyleLst xmlns:d="{A}" xmlns:x="urn:test" def="{DEFAULT}">
  <d:tblStyle styleId="{FIRST}" styleName=""><d:tblBg><d:fill/></d:tblBg><d:extLst><d:ext uri="x"><x:data/></d:ext></d:extLst></d:tblStyle>
  <d:tblStyle styleId="{SECOND}" styleName=""><d:wholeTbl/></d:tblStyle>
</d:tblStyleLst>"#,
    );
    let mut list = List::parse(xml.as_bytes().to_vec()).unwrap();
    assert_eq!(list.conformance(), Conformance::Transitional);
    assert_eq!(list.named("").count(), 2);
    assert!(
        list.get(Id::parse(FIRST).unwrap())
            .unwrap()
            .has(Parts::BACKGROUND)
    );
    assert_eq!(list.source_xml(), Some(xml.as_bytes()));

    let unchanged = list.into_xml().unwrap();
    assert_eq!(unchanged, xml.as_bytes());

    list = List::parse(unchanged).unwrap();
    list.rename(Id::parse(FIRST).unwrap(), "Renamed").unwrap();
    let changed = list.into_xml().unwrap();
    let parsed = List::parse(changed).unwrap();
    let renamed = parsed.get(Id::parse(FIRST).unwrap()).unwrap();
    assert_eq!(renamed.name(), "Renamed");
    assert!(renamed.has(Parts::BACKGROUND));
    assert!(
        parsed
            .source_xml()
            .unwrap()
            .windows(6)
            .any(|value| value == b"extLst")
    );
}

#[test]
fn semantic_noop_keeps_preserved_and_opaque_whitespace_distinct() {
    let preserved = format!(
        r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x" xml:space="preserve"> </a:tblStyle></a:tblStyleLst>"#,
    );
    let changed_preserved = format!(
        r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x" xml:space="preserve">  </a:tblStyle></a:tblStyleLst>"#,
    );
    List::parse(preserved.as_bytes().to_vec()).unwrap();
    List::parse(changed_preserved.as_bytes().to_vec()).unwrap();
    assert!(!semantic_xml_eq(preserved.as_bytes(), changed_preserved.as_bytes()).unwrap());

    let opaque = format!(
        r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><a:ext uri="x"><x:data xmlns:x="urn:test"> </x:data></a:ext></a:extLst></a:tblStyle></a:tblStyleLst>"#,
    );
    let changed_opaque = format!(
        r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><a:ext uri="x"><x:data xmlns:x="urn:test">  </x:data></a:ext></a:extLst></a:tblStyle></a:tblStyleLst>"#,
    );
    let mut package = synthetic(
        ct::PML_PRESENTATION_MAIN,
        Conformance::Transitional,
        Some(opaque.as_bytes()),
    );
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());

    let candidate = List::parse(changed_opaque.as_bytes().to_vec()).unwrap();
    assert!(put(&mut package, candidate).unwrap());
    assert!(!package.is_signed());
    assert_eq!(
        load(&package).unwrap().unwrap().source_xml(),
        Some(changed_opaque.as_bytes())
    );
}

#[test]
fn semantic_comparison_publishes_changed_opaque_comments() {
    let original = format!(
        r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><!-- producer note --><a:ext uri="x"/></a:extLst></a:tblStyle></a:tblStyleLst>"#,
    );
    let changed = format!(
        r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><!-- updated producer note --><a:ext uri="x"/></a:extLst></a:tblStyle></a:tblStyleLst>"#,
    );
    let mut package = synthetic(
        ct::PML_PRESENTATION_MAIN,
        Conformance::Transitional,
        Some(original.as_bytes()),
    );

    assert!(
        put(
            &mut package,
            List::parse(changed.as_bytes().to_vec()).unwrap()
        )
        .unwrap()
    );
    assert_eq!(
        load(&package).unwrap().unwrap().source_xml(),
        Some(changed.as_bytes())
    );
}

#[test]
fn semantic_crud_is_checked_and_failure_atomic() {
    let default = Id::parse(DEFAULT).unwrap();
    let first = Id::parse(FIRST).unwrap();
    let second = Id::parse(SECOND).unwrap();
    let mut list = List::new(Conformance::Transitional, default);
    let mut a = Def::new(first, "Shared name").unwrap();
    a.reset_parts(Parts::WHOLE | Parts::FIRST_ROW);
    list.add(a).unwrap();
    list.add(Def::new(second, "Shared name").unwrap()).unwrap();
    assert_eq!(list.named("Shared name").count(), 2);

    let before = list.len();
    assert!(list.add(Def::new(first, "Duplicate ID").unwrap()).is_err());
    assert_eq!(list.len(), before);
    assert!(list.remove(default).is_err());
    assert_eq!(list.len(), before);

    let old = list.rename(first, "Renamed").unwrap();
    assert_eq!(old, "Shared name");
    let removed = list.remove(second).unwrap().unwrap();
    assert_eq!(removed.id(), second);
    assert_eq!(list.len(), 1);
    let round_trip = List::parse(list.into_xml().unwrap()).unwrap();
    assert_eq!(round_trip.get(first).unwrap().name(), "Renamed");
    assert!(round_trip.get(first).unwrap().has(Parts::FIRST_ROW));
}

#[test]
fn parser_rejects_requiredness_duplicates_mixed_dialects_and_forbidden_markup() {
    let cases = [
        format!(r#"<a:tblStyleLst xmlns:a="{A}"/>"#),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleName="x"/></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}"/></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="a"/><a:tblStyle styleId="{FIRST}" styleName="b"/></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" xmlns:s="{AS}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><s:firstRow/></a:tblStyle></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:firstRow/><a:firstRow/></a:tblStyle></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:firstCol/><a:lastCol/></a:tblStyle></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst/><a:firstRow/></a:tblStyle></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x">text</a:tblStyle></a:tblStyleLst>"#
        ),
        format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x">&amp;</a:tblStyle></a:tblStyleLst>"#
        ),
        format!(r#"<!DOCTYPE x><a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"/>"#),
    ];
    for xml in cases {
        assert!(List::parse(xml.into_bytes()).is_err());
    }
}

#[test]
fn package_crud_round_trips_all_main_profiles_and_both_conformances() {
    for content_type in [
        ct::PML_PRESENTATION_MAIN,
        ct::PML_SLIDESHOW_MAIN,
        ct::PML_TEMPLATE_MAIN,
        ct::PML_PRES_MACRO_MAIN,
        ct::PML_SLIDESHOW_MACRO_MAIN,
        ct::PML_TEMPLATE_MACRO_MAIN,
    ] {
        for conformance in [Conformance::Transitional, Conformance::Strict] {
            let mut package = synthetic(content_type, conformance, None);
            let mut list = List::new(conformance, Id::parse(DEFAULT).unwrap());
            let mut style = Def::new(Id::parse(FIRST).unwrap(), "Created").unwrap();
            style.reset_parts(Parts::BACKGROUND | Parts::WHOLE | Parts::FIRST_ROW);
            list.add(style).unwrap();
            assert!(put(&mut package, list).unwrap());

            let loaded = load(&package).unwrap().unwrap();
            assert_eq!(loaded.conformance(), conformance);
            assert!(
                loaded
                    .get(Id::parse(FIRST).unwrap())
                    .unwrap()
                    .has(Parts::BACKGROUND)
            );
            package.relate_to(
                "_xmlsignatures/origin.sigs",
                relationship_type::DIGITAL_SIGNATURE_ORIGIN,
            );
            assert!(package.is_signed());
            assert!(!put(&mut package, loaded).unwrap());
            assert!(package.is_signed());
            let semantic = List::parse(
                    format!(
                        r#"<d:tblStyleLst def="{DEFAULT}" xmlns:d="{}">
  <d:tblStyle styleName="Created" styleId="{FIRST}"><d:tblBg></d:tblBg><d:wholeTbl></d:wholeTbl><d:firstRow></d:firstRow></d:tblStyle>
</d:tblStyleLst>"#,
                        conformance.drawing(),
                    )
                    .into_bytes(),
                )
                .unwrap();
            assert!(!put(&mut package, semantic).unwrap());
            assert!(package.is_signed());

            let mut changed = load(&package).unwrap().unwrap();
            changed
                .rename(Id::parse(FIRST).unwrap(), "Changed")
                .unwrap();
            assert!(put(&mut package, changed).unwrap());
            assert!(!package.is_signed());
            assert_eq!(
                load(&package)
                    .unwrap()
                    .unwrap()
                    .get(Id::parse(FIRST).unwrap())
                    .unwrap()
                    .name(),
                "Changed"
            );

            let removed = remove(&mut package).unwrap().unwrap();
            assert_eq!(removed.conformance(), conformance);
            assert!(load(&package).unwrap().is_none());
            assert!(remove(&mut package).unwrap().is_none());
        }
    }
}

#[test]
fn shared_inbound_topology_is_rejected_before_mutation() {
    let mut package = synthetic(
        ct::PML_PRESENTATION_MAIN,
        Conformance::Transitional,
        Some(default_xml().as_bytes()),
    );
    let observer_name = PackURI::new("/ppt/observer.xml").unwrap();
    let mut observer = BlobPart::new(
        observer_name,
        "application/xml".into(),
        b"<observer/>".to_vec(),
    );
    observer.rels_mut().add_relationship(
        "urn:test:shared".into(),
        "tableStyles.xml".into(),
        "rIdShared".into(),
        false,
    );
    package.add_part(Box::new(observer));
    let part_count = package.part_count();
    let presentation = package.main_document_part().unwrap().partname().clone();
    let rel_count = package.get_part(&presentation).unwrap().rels().len();

    assert!(load(&package).is_err());
    assert!(remove(&mut package).is_err());
    assert_eq!(package.part_count(), part_count);
    assert_eq!(
        package.get_part(&presentation).unwrap().rels().len(),
        rel_count
    );
}

#[test]
fn orphan_catalog_is_rejected_even_when_another_catalog_is_attached() {
    let mut package = synthetic(
        ct::PML_PRESENTATION_MAIN,
        Conformance::Transitional,
        Some(default_xml().as_bytes()),
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/orphanTableStyles.xml").unwrap(),
        ct::PML_TABLE_STYLES.into(),
        default_xml().as_bytes().to_vec(),
    )));
    let part_count = package.part_count();
    let presentation = package.main_document_part().unwrap().partname().clone();
    let rel_count = package.get_part(&presentation).unwrap().rels().len();

    assert!(present(&package).is_err());
    assert!(load(&package).is_err());
    assert!(remove(&mut package).is_err());
    assert!(
        put(
            &mut package,
            List::new(Conformance::Transitional, Id::parse(DEFAULT).unwrap()),
        )
        .is_err()
    );
    assert_eq!(package.part_count(), part_count);
    assert_eq!(
        package.get_part(&presentation).unwrap().rels().len(),
        rel_count
    );
}

#[test]
fn root_relationship_must_match_the_presentation_dialect() {
    let mut package = synthetic(
        ct::PML_PRESENTATION_MAIN,
        Conformance::Strict,
        Some(format!(r#"<a:tblStyleLst xmlns:a="{AS}" def="{DEFAULT}"/>"#).as_bytes()),
    );
    let relationship_id = package.rels().iter().next().unwrap().r_id().to_owned();
    assert!(package.rels_mut().remove(&relationship_id).is_some());
    package.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.into(),
        "ppt/presentation.xml".into(),
        relationship_id,
        false,
    );

    assert!(present(&package).is_err());
    assert!(load(&package).is_err());
}

fn synthetic(content_type: &str, conformance: Conformance, catalog: Option<&[u8]>) -> OpcPackage {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        content_type.into(),
        format!(
            "<p:presentation xmlns:p=\"{}\"/>",
            match conformance {
                Conformance::Transitional => P,
                Conformance::Strict => PS,
            }
        )
        .into_bytes(),
    );
    let mut package = OpcPackage::new();
    if let Some(catalog) = catalog {
        presentation.rels_mut().add_relationship(
            conformance.relationship().into(),
            "tableStyles.xml".into(),
            "rIdStyles".into(),
            false,
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/tableStyles.xml").unwrap(),
            ct::PML_TABLE_STYLES.into(),
            catalog.to_vec(),
        )));
    }
    package.add_part(Box::new(presentation));
    package.relate_to("ppt/presentation.xml", conformance.office_document());
    package
}
