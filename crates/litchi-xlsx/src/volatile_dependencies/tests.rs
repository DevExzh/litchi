//! Regression tests for the typed volatile-dependencies model and package service.

use super::model::{CONTENT_TYPE, MAX_PART_BYTES, NS, REL, STRICT_NS, STRICT_REL};
use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

fn sample(ns: &str) -> Vec<u8> {
    format!(r#"<volTypes xmlns="{ns}"><volType type="realTimeData"><main first="server.id"><tp t="s"><v>ready</v><stp>ticker</stp><tr r="$A$1" s="0"/></tp></main></volType><volType type="olapFunctions"><main first="cube"><tp t="n"><v>42.5</v><tr r="B2" s="1"/></tp></main></volType></volTypes>"#).into_bytes()
}

fn value() -> VolatileDependencies {
    VolatileDependencies::parse(&sample(std::str::from_utf8(NS).unwrap())).unwrap()
}

fn workbook_package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        format!(
            r#"<workbook xmlns="{}"><sheets/></workbook>"#,
            std::str::from_utf8(NS).unwrap()
        )
        .into_bytes(),
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
        format!(
            r#"<workbook xmlns="{}"><sheets/></workbook>"#,
            std::str::from_utf8(NS).unwrap()
        )
        .into_bytes(),
    );
    if external {
        workbook.relate_to_ext(
            "https://example.invalid/volatileDependencies.xml",
            relationship_type,
        );
    } else {
        workbook.relate_to("volatileDependencies.xml", relationship_type);
    }
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    if !external {
        let mut dependencies = BlobPart::new(
            PackURI::new("/xl/volatileDependencies.xml").unwrap(),
            content_type.into(),
            sample(std::str::from_utf8(NS).unwrap()),
        );
        if outbound {
            dependencies.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        }
        package.add_part(Box::new(dependencies));
    }
    package
}

#[test]
fn parses_and_writes_transitional_and_strict() {
    for ns in [
        std::str::from_utf8(NS).unwrap(),
        std::str::from_utf8(STRICT_NS).unwrap(),
    ] {
        let parsed = VolatileDependencies::parse(&sample(ns)).unwrap();
        assert_eq!(parsed.types.len(), 2);
        let strict = parsed.to_xml(true).unwrap();
        assert_eq!(VolatileDependencies::parse(&strict).unwrap(), parsed);
    }
}

#[test]
fn serializer_rejects_oversized_output_before_appending_extension() {
    let mut value = value();
    value.extension_list_xml = Some(vec![b'x'; MAX_PART_BYTES]);

    let error = value.to_xml(false).unwrap_err();
    assert_eq!(
        error.to_string(),
        "serialized volatile-dependencies part exceeds 8 MiB"
    );
}

#[test]
fn applies_mce_fallback() {
    let xml = format!(
        r#"<volTypes xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:item/></mc:Choice><mc:Fallback><volType type="realTimeData"><main first="srv"><tp t="b"><v>true</v><tr r="A1" s="0"/></tp></main></volType></mc:Fallback></mc:AlternateContent></volTypes>"#,
        std::str::from_utf8(NS).unwrap()
    );
    assert_eq!(
        VolatileDependencies::parse(xml.as_bytes())
            .unwrap()
            .types
            .len(),
        1
    );
}

#[test]
fn preserves_non_empty_extension_list_and_inherited_namespaces() {
    let xml = format!(
        r#"<volTypes xmlns="{}" xmlns:x14="urn:shadowed" xmlns:x15="urn:inherited"><volType type="realTimeData"><main first="srv"><tp><v>ok</v><tr r="A1" s="0"/></tp></main></volType><extLst xmlns:x14="urn:local"><ext uri="urn:test"><x14:payload value="kept"/><x15:payload/></ext></extLst></volTypes>"#,
        std::str::from_utf8(NS).unwrap()
    );
    let parsed = VolatileDependencies::parse(xml.as_bytes()).unwrap();
    let written = parsed.to_xml(true).unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.contains("xmlns:x14=\"urn:local\""));
    assert!(text.contains("xmlns:x15=\"urn:inherited\""));
    assert!(!text.contains("urn:shadowed"));
    assert!(text.contains("x14:payload"));
    assert_eq!(VolatileDependencies::parse(&written).unwrap(), parsed);
}

#[test]
fn rejects_malformed_and_unsafe_input() {
    for xml in [
        format!(
            r#"<volTypes xmlns="{}"/>"#,
            std::str::from_utf8(NS).unwrap()
        ),
        format!(
            r#"<volTypes xmlns="{}"><volType type="bad"><main first="x"><tp><v/><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
            std::str::from_utf8(NS).unwrap()
        ),
        format!(
            r#"<volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp t="b"><v>maybe</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
            std::str::from_utf8(NS).unwrap()
        ),
        format!(
            r#"<volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp t="x"><v>future</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
            std::str::from_utf8(NS).unwrap()
        ),
        format!(
            r#"<!DOCTYPE x [<!ENTITY e "boom">]><volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp><v>&e;</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
            std::str::from_utf8(NS).unwrap()
        ),
    ] {
        assert!(
            VolatileDependencies::parse(xml.as_bytes()).is_err(),
            "accepted {xml}"
        );
    }
}

#[test]
fn resolves_package_relationship_and_rejects_outbound_relationships() {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
    let mut workbook = BlobPart::new(
        workbook_uri,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
        Vec::new(),
    );
    workbook.relate_to("deps.xml", REL);
    package.relate_to(
        "custom/book.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    );
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/custom/deps.xml").unwrap(),
        CONTENT_TYPE.into(),
        sample(std::str::from_utf8(NS).unwrap()),
    )));
    assert_eq!(load_from_package(&package).unwrap().unwrap().types.len(), 2);
    package
        .get_part_mut(&PackURI::new("/custom/deps.xml").unwrap())
        .unwrap()
        .relate_to("other.xml", "urn:forbidden");
    assert!(load_from_package(&package).is_err());
}

#[test]
fn stores_rewrites_and_removes_inert_volatile_dependencies_parts() {
    let mut package = workbook_package();
    let value = value();

    store_in_package(
        &mut package,
        &value,
        VolatileDependenciesConformance::Transitional,
    )
    .unwrap();
    assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));
    assert_eq!(
        load_from_package_with_conformance(&package).unwrap(),
        Some((value.clone(), VolatileDependenciesConformance::Transitional))
    );

    let workbook = package.main_document_part().unwrap();
    let relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == REL)
        .unwrap();
    let relationship_id = relationship.r_id().to_string();
    let part_name = relationship.target_partname().unwrap();
    assert_eq!(
        part_name,
        PackURI::new("/xl/volatileDependencies.xml").unwrap()
    );
    assert!(
        std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
            .unwrap()
            .contains(std::str::from_utf8(NS).unwrap())
    );

    let mut replacement = value.clone();
    replacement.types[0].mains[0].first = "replacement.server".into();
    store_in_package(
        &mut package,
        &replacement,
        VolatileDependenciesConformance::Strict,
    )
    .unwrap();
    let workbook = package.main_document_part().unwrap();
    let relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.r_id() == relationship_id)
        .unwrap();
    assert_eq!(relationship.reltype(), STRICT_REL);
    assert_eq!(relationship.target_partname().unwrap(), part_name);
    assert!(
        std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
            .unwrap()
            .contains(std::str::from_utf8(STRICT_NS).unwrap())
    );
    assert_eq!(
        load_from_package_with_conformance(&package).unwrap(),
        Some((replacement, VolatileDependenciesConformance::Strict))
    );

    assert!(remove_from_package(&mut package).unwrap());
    assert!(package.get_part(&part_name).is_err());
    assert_eq!(load_from_package(&package).unwrap(), None);
    assert!(!remove_from_package(&mut package).unwrap());
}

#[test]
fn removal_retains_volatile_dependencies_part_referenced_elsewhere() {
    let mut package = workbook_package();
    let value = value();
    store_in_package(
        &mut package,
        &value,
        VolatileDependenciesConformance::Transitional,
    )
    .unwrap();

    let part_name = PackURI::new("/xl/volatileDependencies.xml").unwrap();
    let mut referring_part = BlobPart::new(
        PackURI::new("/xl/retained-reference.xml").unwrap(),
        ct::XML.into(),
        b"<reference/>".to_vec(),
    );
    referring_part.relate_to(
        "volatileDependencies.xml",
        "urn:litchi:test:volatile-dependencies-reference",
    );
    package.add_part(Box::new(referring_part));

    assert!(remove_from_package(&mut package).unwrap());
    assert!(package.get_part(&part_name).is_ok());
    assert!(load_from_package(&package).is_err());
    assert!(
        store_in_package(
            &mut package,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
}

#[test]
fn package_volatile_dependencies_mutators_reject_invalid_existing_graphs() {
    let value = value();
    let mut wrong_content_type = synthetic_package(REL, false, ct::SML_STYLES, false);
    let part_name = PackURI::new("/xl/volatileDependencies.xml").unwrap();
    let original = wrong_content_type
        .get_part(&part_name)
        .unwrap()
        .blob()
        .to_vec();
    assert!(
        store_in_package(
            &mut wrong_content_type,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
    assert_eq!(
        wrong_content_type.get_part(&part_name).unwrap().blob(),
        original
    );
    assert!(remove_from_package(&mut wrong_content_type).is_err());

    let mut duplicate = synthetic_package(REL, false, CONTENT_TYPE, false);
    duplicate
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            REL.into(),
            "volatileDependencies.xml".into(),
            "rIdDuplicateVolatileDependencies".into(),
            false,
        );
    assert!(
        store_in_package(
            &mut duplicate,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
    assert!(remove_from_package(&mut duplicate).is_err());

    let mut duplicate_part = synthetic_package(REL, false, CONTENT_TYPE, false);
    duplicate_part.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/volatileDependenciesExtra.xml").unwrap(),
        CONTENT_TYPE.into(),
        sample(std::str::from_utf8(NS).unwrap()),
    )));
    assert!(load_from_package(&duplicate_part).is_err());
    assert!(
        store_in_package(
            &mut duplicate_part,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
    assert!(remove_from_package(&mut duplicate_part).is_err());

    let mut external = synthetic_package(REL, true, CONTENT_TYPE, false);
    assert!(
        store_in_package(
            &mut external,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
    assert!(remove_from_package(&mut external).is_err());

    let mut outbound = synthetic_package(REL, false, CONTENT_TYPE, true);
    assert!(
        store_in_package(
            &mut outbound,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
    assert!(remove_from_package(&mut outbound).is_err());

    let mut root_relationship = workbook_package();
    root_relationship.relate_to("xl/volatileDependencies.xml", REL);
    assert!(
        store_in_package(
            &mut root_relationship,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .is_err()
    );
}
