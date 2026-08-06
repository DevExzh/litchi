use litchi_sheet::Cell as Address;

use super::{Cell, Collection, Conformance, Property, Tag, parse, replace_worksheet, write};
use crate::Package;

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

fn sample() -> Collection {
    let mut contact = Tag::new(7).unwrap();
    contact
        .set_xml_based(true)
        .add_property(Property::new("kind", "contact & person").unwrap())
        .unwrap();
    let mut retired = Tag::new(32_768).unwrap();
    retired.set_deleted(true);
    Collection::new(vec![
        Cell::new("$C$3", vec![retired]).unwrap(),
        Cell::new("A1", vec![contact]).unwrap(),
    ])
    .unwrap()
}

fn worksheet(fragment: &str) -> Vec<u8> {
    format!(
        r#"<worksheet xmlns="{NS}"><dimension ref="A1"/><sheetData/><ignoredErrors><ignoredError sqref="A1"/></ignoredErrors>{fragment}<drawing xmlns:r="urn:r" r:id="rId1"/><extLst/></worksheet>"#
    )
    .into_bytes()
}

#[test]
fn parses_typed_values_and_office_defaults() {
    let xml = worksheet(
        r#"<smartTags><cellSmartTags r="$C$3"><cellSmartTag type="7"><cellSmartTagPr key="kind" val="contact"/></cellSmartTag></cellSmartTags></smartTags>"#,
    );
    let value = parse(&xml).unwrap().unwrap();
    assert_eq!(value.len(), 1);
    let cell = value.get("C3").unwrap().unwrap();
    assert_eq!(cell.address(), Address::from_a1("C3").unwrap());
    assert_eq!(cell.tags()[0].type_id(), 7);
    assert!(!cell.tags()[0].is_deleted());
    assert!(!cell.tags()[0].is_xml_based());
    assert_eq!(cell.tags()[0].properties()[0].key(), "kind");
}

#[test]
fn canonical_writer_round_trips_both_conformance_families() {
    let value = sample();
    for (conformance, namespace) in [
        (Conformance::Transitional, NS),
        (Conformance::Strict, STRICT),
    ] {
        let fragment = write(&value, conformance).unwrap();
        let document = format!(
            r#"<worksheet xmlns="{namespace}"><sheetData/>{}</worksheet>"#,
            String::from_utf8(fragment).unwrap()
        );
        assert_eq!(parse(document.as_bytes()).unwrap().unwrap(), value);
    }
}

#[test]
fn replacement_preserves_unrelated_source_bytes_and_schema_order() {
    let source = worksheet("");
    let changed = replace_worksheet(&source, Some(&sample())).unwrap();
    let changed_text = String::from_utf8(changed.clone()).unwrap();
    assert!(changed_text.contains(
        r#"<ignoredErrors><ignoredError sqref="A1"/></ignoredErrors><smartTags xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#
    ));
    assert!(changed_text.contains(r#"</smartTags><drawing xmlns:r="urn:r" r:id="rId1"/>"#));
    assert_eq!(parse(&changed).unwrap().unwrap(), sample());

    let removed = replace_worksheet(&changed, None).unwrap();
    assert_eq!(removed, source);
}

#[test]
fn rejects_ambiguous_or_unsafe_metadata() {
    assert!(Tag::new(32_769).is_err());
    assert!(Property::new("", "value").is_err());
    assert!(Cell::new("XFE1", vec![Tag::new(1).unwrap()]).is_err());
    let first = Cell::new("A1", vec![Tag::new(1).unwrap()]).unwrap();
    let second = Cell::new("$A$1", vec![Tag::new(2).unwrap()]).unwrap();
    assert!(Collection::new(vec![first, second]).is_err());

    for fragment in [
        "<smartTags/>",
        "<smartTags><cellSmartTags r=\"A1\"/></smartTags>",
        "<smartTags><cellSmartTags r=\"A1\"><cellSmartTag/></cellSmartTags></smartTags>",
        "<smartTags><cellSmartTags r=\"A1\"><cellSmartTag type=\"32769\"/></cellSmartTags></smartTags>",
        "<smartTags><cellSmartTags r=\"A1\"><cellSmartTag type=\"1\" deleted=\"yes\"/></cellSmartTags></smartTags>",
        "<smartTags><cellSmartTags r=\"A1\"><cellSmartTag type=\"1\"><cellSmartTagPr key=\"x\" val=\"1\"/><cellSmartTagPr key=\"x\" val=\"2\"/></cellSmartTag></cellSmartTags></smartTags>",
    ] {
        assert!(parse(&worksheet(fragment)).is_err(), "accepted {fragment}");
    }
    assert!(parse(format!(r#"<!DOCTYPE worksheet><worksheet xmlns="{NS}"/>"#).as_bytes()).is_err());
}

#[test]
fn package_transaction_is_semantic_atomic_and_ergonomic() {
    let mut package = Package::create().unwrap();
    {
        let mut edit = package.edit_smart_tags("Sheet1").unwrap();
        edit.set(Cell::new("B2", vec![Tag::new(3).unwrap()]).unwrap())
            .unwrap();
        assert_eq!(edit.collection().unwrap().len(), 1);
    }
    assert!(package.smart_tags("Sheet1").unwrap().is_none());

    {
        let mut edit = package.edit_smart_tags(0usize).unwrap();
        edit.replace(sample()).unwrap();
        edit.commit().unwrap();
    }
    let value = package.smart_tags("sheet1").unwrap().unwrap();
    assert_eq!(value, sample());
    assert_eq!(
        package
            .workbook()
            .unwrap()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .smart_tags()
            .unwrap(),
        Some(sample())
    );

    let mut edit = package.edit_smart_tags("Sheet1").unwrap();
    assert!(edit.remove("A1").unwrap().is_some());
    edit.commit().unwrap();
    assert!(
        package
            .smart_tags("Sheet1")
            .unwrap()
            .unwrap()
            .get("A1")
            .unwrap()
            .is_none()
    );
}
