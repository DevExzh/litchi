use super::model::{
    CONTENT_TYPE, MAX_ATTRIBUTE_BYTES, MAX_CELL_CONTENT_BYTES, RELATIONSHIP, STRICT_NS,
    STRICT_RELATIONSHIP, TRANSITIONAL_NS,
};
use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::Cell as Address;

use super::package::{MAX_WORKBOOK_SHEETS, process_workbook_mce_with_limit};
use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};

#[test]
fn parses_writes_typed_sheets_steps_strict_and_extensions() {
    let xml = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="urn:test" x:root="v"><c r="A1" i="2" l="1" t="true" a="1" x:cell="kept"/><c r="B2" s="1"/><c r="C3"/><extLst><ext uri="urn:test"><x:data value="inert"/></ext></extLst></calcChain>"#;
    let chain = read(xml).unwrap();
    assert_eq!(chain.len(), 3);
    assert_eq!(chain.at(1).unwrap().sheet().get(), 2);
    assert_eq!(chain.at(0).unwrap().step(), Step::Level);
    assert_eq!(chain.at(1).unwrap().step(), Step::Child);
    assert_eq!(chain.at(2).unwrap().step(), Step::Same);
    assert!(chain.at(0).unwrap().flags().contains(Flags::THREAD));
    assert!(chain.at(0).unwrap().flags().contains(Flags::ARRAY));
    assert_eq!(
        chain.get(Sheet::new(2).unwrap(), "B2").unwrap(),
        chain.cells().get(1)
    );

    let strict = String::from_utf8(write(&chain, Conformance::Strict).unwrap()).unwrap();
    assert!(strict.contains(STRICT_NS));
    assert!(strict.contains("x:cell=\"kept\""));
    assert!(strict.contains("<extLst>"));
    let reparsed = read(strict.as_bytes()).unwrap();
    assert_eq!(reparsed, chain);
    assert_eq!(
        write(&reparsed, Conformance::Strict).unwrap(),
        strict.as_bytes()
    );
}

#[test]
fn preprocesses_mce_and_rejects_malformed_records() {
    let mce = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:c/></mc:Choice><mc:Fallback><c r="C3" i="1"/></mc:Fallback></mc:AlternateContent></calcChain>"#;
    assert_eq!(read(mce).unwrap().cells()[0].reference(), "C3");
    let invalid = [
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"/>"#),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c/></calcChain>"#),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1"/></calcChain>"#),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="XFE1" i="1"/></calcChain>"#),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1" l="yes"/></calcChain>"#),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="65535"/></calcChain>"#),
        format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1" l="1" s="1"/></calcChain>"#
        ),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><extLst/><c r="A1" i="1"/></calcChain>"#),
        format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"><c r="B1"/></c></calcChain>"#
        ),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1" bogus="1"/></calcChain>"#),
    ];
    for xml in invalid {
        assert!(read(xml.as_bytes()).is_err(), "accepted {xml}");
    }
    assert!(Sheet::new(0).is_err());
    assert!(Sheet::new(65_535).is_err());
    assert_eq!(Sheet::new(65_534).unwrap().get(), 65_534);
}

#[test]
fn semantic_and_positional_crud_is_checked_and_failure_atomic() {
    let sheet = Sheet::new(1).unwrap();
    let first = Cell::new(sheet, "A1").unwrap();
    let mut chain = Chain::new(first.clone());
    chain.push(Cell::new(sheet, "C3").unwrap()).unwrap();
    chain.insert(1, Cell::new(sheet, "B2").unwrap()).unwrap();
    assert_eq!(chain.at(1).unwrap().reference(), "B2");
    assert_eq!(
        chain.get(sheet, "C3").unwrap().unwrap().address(),
        Address::from_a1("C3").unwrap()
    );

    let before = chain.clone();
    assert!(chain.push(Cell::new(sheet, "B2").unwrap()).is_err());
    assert!(chain.insert(9, Cell::new(sheet, "D4").unwrap()).is_err());
    assert_eq!(chain, before);

    let replaced = chain
        .replace_at(1, Cell::new(sheet, "D4").unwrap())
        .unwrap();
    assert_eq!(replaced.reference(), "B2");
    let replaced = chain.put(Cell::new(sheet, "D4").unwrap()).unwrap().unwrap();
    assert_eq!(replaced.reference(), "D4");
    assert_eq!(
        chain.remove(sheet, "C3").unwrap().unwrap().reference(),
        "C3"
    );
    chain.move_at(1, 0).unwrap();
    assert_eq!(chain.at(0).unwrap().reference(), "D4");
    assert_eq!(chain.remove_at(1).unwrap().reference(), "A1");
    let before = chain.clone();
    assert!(chain.remove_at(0).is_err());
    assert_eq!(chain, before);
    assert!(!chain.is_empty());
}

#[test]
fn malformed_duplicate_keys_are_retained_but_never_selected_silently() {
    let xml =
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/><c r="A1"/></calcChain>"#);
    let chain = read(xml.as_bytes()).unwrap();
    assert_eq!(chain.len(), 2);
    assert!(chain.get(Sheet::new(1).unwrap(), "A1").is_err());
}

#[test]
fn semantic_mutations_reject_ambiguous_imports_without_changing_order() {
    let xml = format!(
        r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/><c r="A1"/><c r="B2"/></calcChain>"#
    );
    let mut chain = read(xml.as_bytes()).unwrap();
    let before = chain.clone();
    let sheet = Sheet::new(1).unwrap();

    assert!(chain.get(sheet, "B2").is_err());
    assert!(chain.put(Cell::new(sheet, "C3").unwrap()).is_err());
    assert!(chain.push(Cell::new(sheet, "C3").unwrap()).is_err());
    assert!(chain.insert(1, Cell::new(sheet, "C3").unwrap()).is_err());
    assert!(chain.remove(sheet, "B2").is_err());
    assert_eq!(chain, before);
}

#[test]
fn positional_repairs_refresh_ambiguous_import_state() {
    let xml = format!(
        r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/><c r="A1"/><c r="B2"/></calcChain>"#
    );
    let mut chain = read(xml.as_bytes()).unwrap();
    let sheet = Sheet::new(1).unwrap();

    chain.remove_at(1).unwrap();
    assert_eq!(chain.get(sheet, "B2").unwrap().unwrap().reference(), "B2");

    let mut chain = read(xml.as_bytes()).unwrap();
    chain
        .replace_at(1, Cell::new(sheet, "C3").unwrap())
        .unwrap();
    assert_eq!(chain.get(sheet, "C3").unwrap().unwrap().reference(), "C3");
}

#[test]
fn rejects_oversized_nested_calculation_content_before_decoding() {
    let mut xml = format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1">"#).into_bytes();
    xml.extend(std::iter::repeat_n(
        b' ',
        MAX_CELL_CONTENT_BYTES.saturating_add(1),
    ));
    xml.extend_from_slice(b"</c></calcChain>");

    assert!(read(&xml).is_err());
}

#[test]
fn rejects_raw_oversized_attributes_before_entity_decoding() {
    let encoded = "&amp;".repeat(MAX_ATTRIBUTE_BYTES / 5 + 1);
    let xml = format!(
        r#"<calcChain xmlns="{TRANSITIONAL_NS}" xmlns:x="urn:test" x:data="{encoded}"><c r="A1" i="1"/></calcChain>"#
    );
    let error = read(xml.as_bytes()).unwrap_err();
    assert!(matches!(
        error,
        crate::Error::Invalid(message)
            if message == format!("calculation-chain attribute exceeds {MAX_ATTRIBUTE_BYTES} bytes")
    ));
}

#[test]
fn adversarial_xml_returns_errors_without_unwinding() {
    let inputs: [&[u8]; 4] = [
        b"<",
        b"\xff<calcChain/>",
        br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="0"/></calcChain>"#,
        br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1" \xff="bad"/></calcChain>"#,
    ];
    for input in inputs {
        let result = std::panic::catch_unwind(|| read(input));
        assert!(result.is_ok(), "reader unwound for {input:?}");
        assert!(result.unwrap().is_err());
    }
}

#[test]
fn stores_rewrites_and_removes_inert_calculation_chain_parts() {
    let mut package = workbook_package();
    let mut first = Cell::new(Sheet::new(1).unwrap(), "B2").unwrap();
    first.set_step(Step::Level);
    let chain = Chain::new(first);

    assert!(put(&mut package, &chain, Conformance::Transitional).unwrap());
    assert_eq!(
        load(&package).unwrap(),
        Some((chain.clone(), Conformance::Transitional))
    );

    let workbook = package.main_document_part().unwrap();
    let relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == RELATIONSHIP)
        .unwrap();
    let relationship_id = relationship.r_id().to_string();
    let part_name = relationship.target_partname().unwrap();
    assert_eq!(part_name, PackURI::new("/xl/calcChain.xml").unwrap());
    assert!(
        std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
            .unwrap()
            .contains(TRANSITIONAL_NS)
    );
    let before = package.get_part(&part_name).unwrap().blob_arc();
    assert!(!put(&mut package, &chain, Conformance::Transitional).unwrap());
    let after = package.get_part(&part_name).unwrap().blob_arc();
    assert!(std::sync::Arc::ptr_eq(&before, &after));

    let replacement = Chain::new(Cell::new(Sheet::new(1).unwrap(), "C3").unwrap());
    assert!(put(&mut package, &replacement, Conformance::Strict).unwrap());
    let workbook = package.main_document_part().unwrap();
    let relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.r_id() == relationship_id)
        .unwrap();
    assert_eq!(relationship.reltype(), STRICT_RELATIONSHIP);
    assert_eq!(relationship.target_partname().unwrap(), part_name);
    assert!(
        std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
            .unwrap()
            .contains(STRICT_NS)
    );
    assert_eq!(
        load(&package).unwrap(),
        Some((replacement, Conformance::Strict))
    );

    assert!(remove(&mut package).unwrap());
    assert!(package.get_part(&part_name).is_err());
    assert_eq!(load(&package).unwrap(), None);
    assert!(!remove(&mut package).unwrap());
}

#[test]
fn removal_retains_a_calculation_chain_part_referenced_elsewhere() {
    let mut package = workbook_package();
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "F6").unwrap());
    put(&mut package, &chain, Conformance::Transitional).unwrap();

    let part_name = PackURI::new("/xl/calcChain.xml").unwrap();
    let mut referring_part = BlobPart::new(
        PackURI::new("/xl/retained-reference.xml").unwrap(),
        ct::XML.into(),
        b"<reference/>".to_vec(),
    );
    referring_part.relate_to("calcChain.xml", "urn:litchi:test:calc-chain-reference");
    package.add_part(Box::new(referring_part));

    assert!(remove(&mut package).unwrap());
    assert!(package.get_part(&part_name).is_ok());
    assert!(load(&package).is_err());
    assert!(put(&mut package, &chain, Conformance::Transitional).is_err());
}

#[test]
fn signed_publication_allows_exact_noops_and_rejects_changes() {
    let mut package = workbook_package();
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "A1").unwrap());
    put(&mut package, &chain, Conformance::Transitional).unwrap();
    package.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.into(),
        "_xmlsignatures/origin.sigs".into(),
        "rIdSignature".into(),
        false,
    );
    let part = PackURI::new("/xl/calcChain.xml").unwrap();
    let before = package.get_part(&part).unwrap().blob_arc();

    assert!(!put(&mut package, &chain, Conformance::Transitional).unwrap());
    assert!(package.is_signed());
    assert!(std::sync::Arc::ptr_eq(
        &before,
        &package.get_part(&part).unwrap().blob_arc()
    ));

    let changed = Chain::new(Cell::new(Sheet::new(1).unwrap(), "B2").unwrap());
    assert!(matches!(
        put(&mut package, &changed, Conformance::Transitional),
        Err(crate::Error::Signed)
    ));
    assert!(matches!(remove(&mut package), Err(crate::Error::Signed)));
    assert!(package.is_signed());
    assert!(std::sync::Arc::ptr_eq(
        &before,
        &package.get_part(&part).unwrap().blob_arc()
    ));

    let mut absent = workbook_package();
    absent.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.into(),
        "_xmlsignatures/origin.sigs".into(),
        "rIdSignature".into(),
        false,
    );
    assert!(!remove(&mut absent).unwrap());
    assert!(absent.is_signed());
}

#[test]
fn refuses_to_canonicalize_mce_projected_chain_sources() {
    let sources = [
        format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:c/></mc:Choice><mc:Fallback><c r="C3" i="1"/></mc:Fallback></mc:AlternateContent></calcChain>"#
        ),
        format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" mc:ProcessContent="x:wrap"><x:wrap><c r="C3" i="1"/></x:wrap></calcChain>"#
        ),
    ];
    for source in sources {
        let mut package = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
        let part = PackURI::new("/xl/calcChain.xml").unwrap();
        package
            .get_part_mut(&part)
            .unwrap()
            .set_blob(source.into_bytes());
        let (chain, conformance) = load(&package).unwrap().unwrap();
        assert_eq!(chain.cells()[0].reference(), "C3");
        let before = package.get_part(&part).unwrap().blob_arc();
        let error = put(&mut package, &chain, conformance).unwrap_err();
        assert!(matches!(
            error,
            crate::Error::Invalid(message)
                if message == "cannot replace calculation-chain source projected through MCE"
        ));
        assert!(std::sync::Arc::ptr_eq(
            &before,
            &package.get_part(&part).unwrap().blob_arc()
        ));
    }
}

#[test]
fn rejects_chain_sheet_ids_absent_from_workbook_catalog_exactly() {
    let mut package = workbook_package();
    let chain = Chain::new(Cell::new(Sheet::new(2).unwrap(), "A1").unwrap());
    let error = put(&mut package, &chain, Conformance::Transitional).unwrap_err();
    assert!(matches!(
        error,
        crate::Error::Invalid(message)
            if message == "calculation-chain sheet ID 2 does not resolve to a workbook sheet"
    ));
    assert!(load(&package).unwrap().is_none());
}

#[test]
fn workbook_sheet_catalog_enforces_its_exact_projected_boundary() {
    let mut boundary = workbook_package_with_sheet_count(MAX_WORKBOOK_SHEETS, false);
    let chain = Chain::new(
        Cell::new(
            Sheet::new(u32::try_from(MAX_WORKBOOK_SHEETS).unwrap()).unwrap(),
            "A1",
        )
        .unwrap(),
    );
    assert!(put(&mut boundary, &chain, Conformance::Transitional).unwrap());

    let mut oversized = workbook_package_with_sheet_count(MAX_WORKBOOK_SHEETS + 1, true);
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "A1").unwrap());
    let error = put(&mut oversized, &chain, Conformance::Transitional).unwrap_err();
    assert!(matches!(
        error,
        crate::Error::Invalid(message)
            if message == format!(
                "workbook sheet catalog exceeds {MAX_WORKBOOK_SHEETS} entries"
            )
    ));
    assert!(load(&oversized).unwrap().is_none());
}

#[test]
fn workbook_mce_expansion_is_rejected_during_bounded_processing() {
    let source = format!(
        r#"<workbook xmlns="{TRANSITIONAL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheets><sheet name="Sheet1" sheetId="1"/></sheets></workbook>"#
    );
    let error = process_workbook_mce_with_limit(source.as_bytes(), source.len()).unwrap_err();
    assert!(matches!(
        error,
        crate::Error::Invalid(message)
            if message == "workbook MCE error: markup compatibility resource limit exceeded: output bytes"
    ));
}

#[test]
fn package_calculation_chain_mutators_reject_invalid_existing_graphs() {
    let mut package = synthetic_package(RELATIONSHIP, false, ct::XML, false);
    let chain_part = PackURI::new("/xl/calcChain.xml").unwrap();
    let original = package.get_part(&chain_part).unwrap().blob().to_vec();
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "E5").unwrap());

    assert!(put(&mut package, &chain, Conformance::Transitional).is_err());
    assert_eq!(package.get_part(&chain_part).unwrap().blob(), original);
    assert!(remove(&mut package).is_err());
    assert!(package.get_part(&chain_part).is_ok());

    let mut duplicate = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
    duplicate
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            RELATIONSHIP.into(),
            "calcChain.xml".into(),
            "rIdDuplicateCalcChain".into(),
            false,
        );
    assert!(put(&mut duplicate, &chain, Conformance::Transitional).is_err());
    assert!(remove(&mut duplicate).is_err());

    let mut duplicate_part = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
    duplicate_part.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/calcChainExtra.xml").unwrap(),
        CONTENT_TYPE.into(),
        format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="F6"/></calcChain>"#).into_bytes(),
    )));
    assert!(load(&duplicate_part).is_err());
    assert!(put(&mut duplicate_part, &chain, Conformance::Transitional).is_err());
    assert!(remove(&mut duplicate_part).is_err());

    let mut external = synthetic_package(RELATIONSHIP, true, CONTENT_TYPE, false);
    assert!(put(&mut external, &chain, Conformance::Transitional).is_err());
    assert!(remove(&mut external).is_err());
}

#[test]
fn loads_real_poi_and_synthetic_packages_and_validates_relationships() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..//test-data/poi/test-data/spreadsheet/62834.xlsx");
    let package = OpcPackage::open(path).unwrap();
    let (chain, _) = load(&package).unwrap().unwrap();
    assert_eq!(chain.cells().len(), 3);
    assert_eq!(chain.cells()[0].reference(), "A5");
    assert_eq!(chain.cells()[0].step(), Step::Level);
    assert_eq!(chain.cells()[2].step(), Step::Child);

    let package = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
    assert_eq!(
        load(&package).unwrap().unwrap().0.cells()[0].reference(),
        "A1"
    );

    assert!(load(&synthetic_package(RELATIONSHIP, true, CONTENT_TYPE, false)).is_err());
    assert!(load(&synthetic_package(RELATIONSHIP, false, ct::XML, false)).is_err());
    assert!(load(&synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, true)).is_err());
}

fn workbook_package() -> OpcPackage {
    workbook_package_with_sheet_count(1, false)
}

fn workbook_package_with_sheet_count(count: usize, projected: bool) -> OpcPackage {
    use std::fmt::Write as _;

    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let mut sheets = String::new();
    sheets.try_reserve(count.saturating_mul(48)).unwrap();
    for sheet_id in 1..=count {
        write!(
            sheets,
            r#"<sheet name="Sheet{sheet_id}" sheetId="{sheet_id}"/>"#
        )
        .unwrap();
    }
    let workbook_xml = if projected {
        format!(
            r#"<workbook xmlns="{TRANSITIONAL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" mc:ProcessContent="x:wrap"><x:wrap><sheets>{sheets}</sheets></x:wrap></workbook>"#
        )
    } else {
        format!(r#"<workbook xmlns="{TRANSITIONAL_NS}"><sheets>{sheets}</sheets></workbook>"#)
    };
    let workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        workbook_xml.into_bytes(),
    );
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    package
}

fn synthetic_package(
    relationship_type: &str,
    external: bool,
    content_type: &str,
    outbound: bool,
) -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let mut workbook = BlobPart::new(
        workbook_uri.clone(),
        ct::SML_SHEET_MAIN.into(),
        format!(r#"<workbook xmlns="{TRANSITIONAL_NS}"><sheets><sheet name="Sheet1" sheetId="1"/></sheets></workbook>"#).into_bytes(),
    );
    if external {
        workbook.relate_to_ext("https://example.invalid/calcChain.xml", relationship_type);
    } else {
        workbook.relate_to("calcChain.xml", relationship_type);
    }
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    if !external {
        let mut chain = BlobPart::new(
            PackURI::new("/xl/calcChain.xml").unwrap(),
            content_type.into(),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/></calcChain>"#)
                .into_bytes(),
        );
        if outbound {
            chain.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        }
        package.add_part(Box::new(chain));
    }
    package
}
