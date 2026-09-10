//! Regression tests for the layered external-link owner.

use super::*;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[test]
fn parses_sparse_workbook_cache_and_keeps_target_inert() {
    let xml = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}"><externalBook r:id="rId1"><sheetNames><sheetName val="Data"/></sheetNames><sheetDataSet><sheetData sheetId="1"><row r="1"><cell r="A1" t="str"><v>001.2300</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#
    );
    let mut part = BlobPart::new(
        PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
        litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
        xml.into_bytes(),
    );
    part.relate_to_ext(
        "https://127.0.0.1:9/never-open.xlsx",
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH,
    );
    let link = load_external_link(&part, "bookRel".into(), 1).unwrap();
    let Link::Workbook(book) = link.link else {
        panic!("expected workbook link")
    };
    assert_eq!(book.target.target, "https://127.0.0.1:9/never-open.xlsx");
    assert_eq!(book.sheet_names, ["Data"]);
    assert_eq!(
        book.cached_sheets[0].rows[0].cells[0].raw_value.as_deref(),
        Some("001.2300")
    );
}

#[test]
fn typed_dde_round_trips_without_target_relationships() {
    let value = Link::Dde(Dde {
        service: "Excel".into(),
        topic: "opaque-source.xlsx".into(),
        items: vec![DdeItem {
            name: Some("R1C1".into()),
            use_ole: false,
            advise: true,
            prefer_picture: false,
            values: Some(DdeValues {
                rows: 1,
                columns: 1,
                values: vec![DdeValue {
                    value_type: DdeValueType::String,
                    raw_value: "<&>".into(),
                }],
            }),
        }],
    });
    let xml = value.to_xml().unwrap();
    assert!(
        std::str::from_utf8(&xml)
            .unwrap()
            .contains("opaque-source.xlsx")
    );
    let parsed = parse_external_link(&xml).unwrap();
    assert_eq!(parsed, value);
}

fn workbook_with_alternate_urls() -> Link {
    Link::Workbook(Workbook {
        target: Target {
            relationship_id: "rId1".into(),
            target: "https://primary.invalid/source.xlsx".into(),
            relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        },
        sheet_names: Vec::new(),
        defined_names: Vec::new(),
        cached_sheets: Vec::new(),
        alternate_urls: Some(AlternateUrls {
            drive_id: Some("drive-1".into()),
            item_id: Some("item-1".into()),
            absolute_url: Some(AlternateUrl {
                relationship_id: "rId2".into(),
                target: "https://absolute.invalid/source.xlsx".into(),
                relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                    .into(),
            }),
            relative_url: Some(AlternateUrl {
                relationship_id: "rId3".into(),
                target: "../relative/source.xlsx".into(),
                relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                    .into(),
            }),
        }),
    })
}

#[test]
fn alternate_urls_use_the_normative_namespace_and_relationship_closure() {
    let link = workbook_with_alternate_urls();
    let xml = link.to_xml().unwrap();
    let text = std::str::from_utf8(&xml).unwrap();
    assert!(text.contains(
        "<alternateUrls xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021\""
    ));
    assert!(text.contains("driveId=\"drive-1\""));
    assert!(text.contains("itemId=\"item-1\""));

    let part = build_external_link_part(
        PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
        &link,
    )
    .unwrap();
    assert_eq!(
        part.rels().get("rId1").unwrap().target_ref(),
        "https://primary.invalid/source.xlsx"
    );
    assert_eq!(
        part.rels().get("rId2").unwrap().target_ref(),
        "https://absolute.invalid/source.xlsx"
    );
    assert_eq!(
        part.rels().get("rId3").unwrap().target_ref(),
        "../relative/source.xlsx"
    );
    let loaded = load_external_link(&part, "bookRel".into(), 0).unwrap();
    assert_eq!(loaded.link, link);
}

#[test]
fn alternate_urls_reject_duplicate_children_and_oversized_service_ids() {
    let duplicate = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}"><externalBook r:id="rId1"><alternateUrls xmlns="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><absoluteUrl r:id="rId2"/><absoluteUrl r:id="rId3"/></alternateUrls></externalBook></externalLink>"#
    );
    assert!(parse_external_link(duplicate.as_bytes()).is_err());
    let out_of_order = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}"><externalBook r:id="rId1"><alternateUrls xmlns="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><relativeUrl r:id="rId2"/><absoluteUrl r:id="rId3"/></alternateUrls></externalBook></externalLink>"#
    );
    assert!(parse_external_link(out_of_order.as_bytes()).is_err());
    let invalid_id = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}"><externalBook r:id="rId1"><alternateUrls xmlns="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><absoluteUrl r:id="bad id"/></alternateUrls></externalBook></externalLink>"#
    );
    assert!(parse_external_link(invalid_id.as_bytes()).is_err());

    let mut link = workbook_with_alternate_urls();
    let Link::Workbook(book) = &mut link else {
        unreachable!()
    };
    book.alternate_urls.as_mut().unwrap().drive_id = Some("x".repeat(65 * 1024));
    assert!(link.to_xml().is_err());
}

#[test]
fn rejects_external_cached_matrix_mismatch() {
    let xml = format!(
        r#"<externalLink xmlns="{SML}"><ddeLink ddeService="x" ddeTopic="y"><ddeItems><ddeItem><values rows="2"><value><val>x</val></value></values></ddeItem></ddeItems></ddeLink></externalLink>"#
    );
    assert!(parse_external_link(xml.as_bytes()).is_err());
}

#[test]
fn canonical_writer_validates_target_relationship_metadata() {
    let value = Link::Ole(Ole {
        target: Target {
            relationship_id: "rId1".into(),
            target: "opaque.bin".into(),
            relationship_type: litchi_opc::constants::relationship_type::OLE_OBJECT.into(),
        },
        program_id: "Excel.Sheet.12".into(),
        items: Vec::new(),
    });
    let part = build_external_link_part(
        PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
        &value,
    )
    .unwrap();
    assert!(part.rels().get("rId1").unwrap().is_external());
}

fn workbook_package() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::OFFICE_DOCUMENT.into(),
        "xl/workbook.xml".into(),
        "rId1".into(),
        false,
    );
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        workbook_uri.clone(),
        litchi_opc::constants::content_type::SML_SHEET_MAIN.into(),
        br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec(),
    )));
    let xml = format!(
        r#"<?xml version="1.0"?><externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:x="urn:future"><externalBook r:id="rId7"><sheetNames><sheetName val="Data"/></sheetNames><x:future marker="keep"/><definedNames><definedName name="DataName" refersTo="[Book.xlsx]Data!$A$1"/></definedNames><sheetDataSet><sheetData sheetId="1"><row r="1"><cell r="A1" t="str"><v>001.2300</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#
    );
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    let mut external = BlobPart::new(
        external_uri.clone(),
        litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
        xml.into_bytes(),
    );
    external.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        "https://127.0.0.1:9/never-open.xlsx".into(),
        "rId7".into(),
        true,
    );
    package.add_part(Box::new(external));
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            litchi_opc::constants::relationship_type::EXTERNAL_LINK.into(),
            external_uri.relative_ref(workbook_uri.base_uri()),
            "rId2".into(),
            false,
        );
    package
}

fn alternate_workbook_package() -> OpcPackage {
    let mut package = workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    let part = package.get_part_mut(&external_uri).unwrap();
    part.set_blob(
        format!(
            r#"<?xml version="1.0"?><externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021" xmlns:x="urn:future"><externalBook r:id="rId7"><x:future marker="keep"/><a:alternateUrls driveId="drive-1" itemId="item-1"><a:absoluteUrl r:id="rId8"/><a:relativeUrl r:id="rId9"/></a:alternateUrls></externalBook></externalLink>"#
        )
        .into_bytes(),
    );
    part.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        "https://absolute.invalid/source.xlsx".into(),
        "rId8".into(),
        true,
    );
    part.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        "../relative/source.xlsx".into(),
        "rId9".into(),
        true,
    );
    package
}

#[test]
fn alternate_url_edits_preserve_opaque_children_and_inverse_exactly() {
    let mut package = alternate_workbook_package();
    let original = package
        .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .edit(0, |link| {
            let Link::Workbook(book) = link else {
                panic!("expected workbook link")
            };
            let alternate_urls = book.alternate_urls.as_mut().unwrap();
            alternate_urls.drive_id = Some("drive-2".into());
            alternate_urls
                .absolute_url
                .as_mut()
                .unwrap()
                .relationship_id = "rId10".into();
            Ok(())
        })
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let part = package
        .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap();
    let source = std::str::from_utf8(part.blob()).unwrap();
    assert!(source.contains("x:future marker=\"keep\""));
    assert!(source.contains("a:alternateUrls"));
    assert!(source.contains("driveId=\"drive-2\""));
    assert!(source.contains("a:absoluteUrl r:id=\"rId10\""));
    assert!(part.rels().get("rId8").is_none());
    assert_eq!(
        part.rels().get("rId10").unwrap().target_ref(),
        "https://absolute.invalid/source.xlsx"
    );

    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .blob(),
        original.as_slice()
    );
}

#[test]
fn alternate_url_add_remove_and_child_crud_preserve_unknown_source() {
    let mut package = workbook_package();
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .edit(0, |link| {
            let Link::Workbook(book) = link else {
                panic!("expected workbook link")
            };
            book.alternate_urls = Some(AlternateUrls {
                drive_id: Some("drive-added".into()),
                item_id: None,
                absolute_url: Some(AlternateUrl {
                    relationship_id: "rId8".into(),
                    target: "https://absolute.invalid/added.xlsx".into(),
                    relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                        .into(),
                }),
                relative_url: None,
            });
            Ok(())
        })
        .unwrap();
    transaction.commit().unwrap();
    let source = std::str::from_utf8(
        package
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(source.contains("x:future marker=\"keep\""));
    assert!(source.contains("alternateUrls"));
    assert!(source.contains("absoluteUrl"));
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .rels()
            .get("rId8")
            .unwrap()
            .target_ref(),
        "https://absolute.invalid/added.xlsx"
    );

    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .edit(0, |link| {
            let Link::Workbook(book) = link else {
                panic!("expected workbook link")
            };
            book.alternate_urls = None;
            Ok(())
        })
        .unwrap();
    transaction.commit().unwrap();
    let source = std::str::from_utf8(
        package
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(!source.contains("alternateUrls"));
    assert!(source.contains("x:future marker=\"keep\""));
    assert!(
        package
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .rels()
            .get("rId8")
            .is_none()
    );

    let mut package = alternate_workbook_package();
    let part = package
        .get_part_mut(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap();
    part.set_blob(
        format!(
            r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><externalBook r:id="rId7"><a:alternateUrls driveId="drive"><a:future/></a:alternateUrls></externalBook></externalLink>"#
        )
        .into_bytes(),
    );
    part.rels_mut().remove("rId8");
    part.rels_mut().remove("rId9");
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .edit(0, |link| {
            let Link::Workbook(book) = link else {
                panic!("expected workbook link")
            };
            book.alternate_urls.as_mut().unwrap().absolute_url = Some(AlternateUrl {
                relationship_id: "rId8".into(),
                target: "https://absolute.invalid/added.xlsx".into(),
                relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                    .into(),
            });
            Ok(())
        })
        .unwrap();
    transaction.commit().unwrap();
    let source = std::str::from_utf8(
        package
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(source.contains("a:future"));
    assert!(source.contains("absoluteUrl"));
}

#[test]
fn alternate_url_child_insertions_keep_schema_order_and_expand_empty_sources() {
    let mut package = alternate_workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    let part = package.get_part_mut(&external_uri).unwrap();
    part.set_blob(
        format!(
            r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><externalBook r:id="rId7"><a:alternateUrls driveId="drive"><a:relativeUrl r:id="rId9"/></a:alternateUrls></externalBook></externalLink>"#
        )
        .into_bytes(),
    );
    part.rels_mut().remove("rId8");
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .edit(0, |link| {
            let Link::Workbook(book) = link else {
                panic!("expected workbook link")
            };
            book.alternate_urls.as_mut().unwrap().absolute_url = Some(AlternateUrl {
                relationship_id: "rId8".into(),
                target: "https://absolute.invalid/added.xlsx".into(),
                relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                    .into(),
            });
            Ok(())
        })
        .unwrap();
    transaction.commit().unwrap();
    let source = std::str::from_utf8(package.get_part(&external_uri).unwrap().blob()).unwrap();
    assert!(source.find("absoluteUrl").unwrap() < source.find("relativeUrl").unwrap());

    let part = package.get_part_mut(&external_uri).unwrap();
    part.set_blob(
        format!(
            r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><externalBook r:id="rId7"><a:alternateUrls driveId="drive"/></externalBook></externalLink>"#
        )
        .into_bytes(),
    );
    part.rels_mut().remove("rId8");
    part.rels_mut().remove("rId9");
    let original = part.blob().to_vec();
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .edit(0, |link| {
            let Link::Workbook(book) = link else {
                panic!("expected workbook link")
            };
            let alternate_urls = book.alternate_urls.as_mut().unwrap();
            alternate_urls.absolute_url = Some(AlternateUrl {
                relationship_id: "rId8".into(),
                target: "https://absolute.invalid/added.xlsx".into(),
                relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                    .into(),
            });
            alternate_urls.relative_url = Some(AlternateUrl {
                relationship_id: "rId9".into(),
                target: "../relative/added.xlsx".into(),
                relationship_type: litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH
                    .into(),
            });
            Ok(())
        })
        .unwrap();
    let commit = transaction.commit().unwrap();
    let source = std::str::from_utf8(package.get_part(&external_uri).unwrap().blob()).unwrap();
    assert!(source.find("absoluteUrl").unwrap() < source.find("relativeUrl").unwrap());
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(package.get_part(&external_uri).unwrap().blob(), original);
}

#[test]
fn alternate_url_same_relationship_id_replaces_the_physical_target() {
    let mut package = alternate_workbook_package();
    let mut replacement = load_external_links(&package).unwrap()[0].link.clone();
    let Link::Workbook(book) = &mut replacement else {
        unreachable!()
    };
    book.alternate_urls
        .as_mut()
        .unwrap()
        .absolute_url
        .as_mut()
        .unwrap()
        .target = "https://absolute-new.invalid/source.xlsx".into();

    let entry =
        replace_external_link(&mut package, 0, replacement, Conformance::Transitional).unwrap();
    let Link::Workbook(book) = entry.link else {
        unreachable!()
    };
    assert_eq!(
        book.alternate_urls.unwrap().absolute_url.unwrap().target,
        "https://absolute-new.invalid/source.xlsx"
    );
    let part = package
        .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap();
    assert_eq!(
        part.rels().get("rId8").unwrap().target_ref(),
        "https://absolute-new.invalid/source.xlsx"
    );
}

#[test]
fn alternate_url_patch_matches_expanded_relationship_namespace_not_foreign_id() {
    let mut package = alternate_workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    package
        .get_part_mut(&external_uri)
        .unwrap()
        .set_blob(
            format!(
                r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:rel="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><externalBook r:id="rId7"><a:alternateUrls driveId="drive"><a:absoluteUrl id="future" rel:id="rId8"/><a:relativeUrl rel:id="rId9"/></a:alternateUrls></externalBook></externalLink>"#
            )
            .into_bytes(),
        );
    let mut replacement = load_external_links(&package).unwrap()[0].link.clone();
    let Link::Workbook(book) = &mut replacement else {
        unreachable!()
    };
    book.alternate_urls
        .as_mut()
        .unwrap()
        .absolute_url
        .as_mut()
        .unwrap()
        .relationship_id = "rId10".into();
    let entry =
        replace_external_link(&mut package, 0, replacement, Conformance::Transitional).unwrap();
    let Link::Workbook(book) = entry.link else {
        unreachable!()
    };
    assert_eq!(
        book.alternate_urls
            .unwrap()
            .absolute_url
            .unwrap()
            .relationship_id,
        "rId10"
    );
    let source = std::str::from_utf8(package.get_part(&external_uri).unwrap().blob()).unwrap();
    assert!(source.contains(r#"id="future" rel:id="rId10""#));
    assert!(
        package
            .get_part(&external_uri)
            .unwrap()
            .rels()
            .get("rId8")
            .is_none()
    );
    assert!(
        package
            .get_part(&external_uri)
            .unwrap()
            .rels()
            .get("rId10")
            .is_some()
    );
}

#[test]
fn alternate_url_relationship_targets_and_metadata_are_bounded_before_retention() {
    let oversized_metadata = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}"><externalBook r:id="rId1"><alternateUrls xmlns="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021" driveId="{}"/></externalBook></externalLink>"#,
        "x".repeat(model::MAX_ALTERNATE_URL_METADATA_BYTES + 1)
    );
    assert!(parse_external_link(oversized_metadata.as_bytes()).is_err());

    let mut package = alternate_workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    let part = package.get_part_mut(&external_uri).unwrap();
    part.rels_mut().remove("rId8");
    part.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        "x".repeat(model::MAX_EXTERNAL_TARGET_BYTES + 1),
        "rId8".into(),
        true,
    );
    assert!(load_external_links(&package).is_err());
}

#[test]
fn alternate_url_relationship_id_is_bounded_before_entity_decoding() {
    let encoded = "&amp;".repeat(205);
    let xml = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}"><externalBook r:id="rId1"><alternateUrls xmlns="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><absoluteUrl r:id="{encoded}"/></alternateUrls></externalBook></externalLink>"#
    );
    assert!(parse_external_link(xml.as_bytes()).is_err());
}

#[test]
fn alternate_url_source_matching_uses_expanded_element_namespaces() {
    let mut package = workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    package
        .get_part_mut(&external_uri)
        .unwrap()
        .set_blob(
            format!(
                r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021" xmlns:f="urn:foreign"><f:externalBook><f:alternateUrls><f:absoluteUrl id="foreign"/></f:alternateUrls></f:externalBook><externalBook r:id="rId7"><f:alternateUrls><f:absoluteUrl id="foreign-child"/></f:alternateUrls><a:alternateUrls driveId="drive"><a:absoluteUrl r:id="rId8"/><a:relativeUrl r:id="rId9"/></a:alternateUrls></externalBook></externalLink>"#
            )
            .into_bytes(),
        );
    let part = package.get_part_mut(&external_uri).unwrap();
    part.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        "https://absolute.invalid/source.xlsx".into(),
        "rId8".into(),
        true,
    );
    part.rels_mut().add_relationship(
        litchi_opc::constants::relationship_type::EXTERNAL_LINK_PATH.into(),
        "../relative/source.xlsx".into(),
        "rId9".into(),
        true,
    );
    let mut replacement = load_external_links(&package).unwrap()[0].link.clone();
    let Link::Workbook(book) = &mut replacement else {
        unreachable!()
    };
    book.alternate_urls
        .as_mut()
        .unwrap()
        .absolute_url
        .as_mut()
        .unwrap()
        .relationship_id = "rId10".into();
    replace_external_link(&mut package, 0, replacement, Conformance::Transitional).unwrap();
    let source = std::str::from_utf8(package.get_part(&external_uri).unwrap().blob()).unwrap();
    assert!(source.contains(r#"<f:externalBook><f:alternateUrls><f:absoluteUrl id="foreign"/></f:alternateUrls></f:externalBook>"#));
    assert!(
        source
            .contains(r#"<f:alternateUrls><f:absoluteUrl id="foreign-child"/></f:alternateUrls>"#)
    );
    assert!(source.contains(r#"<a:absoluteUrl r:id="rId10"/>"#));
    assert!(!source.contains(r#"<a:absoluteUrl r:id="rId8"/>"#));
}

#[test]
fn alternate_url_removal_refuses_to_discard_opaque_descendants() {
    let mut package = alternate_workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    package
        .get_part_mut(&external_uri)
        .unwrap()
        .set_blob(
            format!(
                r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><externalBook r:id="rId7"><a:alternateUrls driveId="drive"><a:future marker="keep"/></a:alternateUrls></externalBook></externalLink>"#
            )
            .into_bytes(),
        );
    let part = package.get_part_mut(&external_uri).unwrap();
    part.rels_mut().remove("rId8");
    part.rels_mut().remove("rId9");
    let original = part.blob().to_vec();
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction.set_alternate_urls(0, None).unwrap();
    assert!(transaction.commit().is_err());
    assert_eq!(package.get_part(&external_uri).unwrap().blob(), original);
}

#[test]
fn oversized_opaque_external_link_source_is_rejected_before_retention() {
    let mut package = workbook_package();
    let external_uri = PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap();
    let mut xml = format!(
        r#"<externalLink xmlns="{SML}" xmlns:r="{REL}" xmlns:x="urn:future"><externalBook r:id="rId7"><x:future>"#
    )
    .into_bytes();
    xml.extend(std::iter::repeat_n(b'x', model::MAX_CACHE_TEXT_BYTES + 1));
    xml.extend_from_slice(b"</x:future></externalBook></externalLink>");
    package.get_part_mut(&external_uri).unwrap().set_blob(xml);
    assert!(load_external_links(&package).is_err());
}

#[test]
fn transaction_edits_known_metadata_without_rebuilding_opaque_xml() {
    let mut package = workbook_package();
    let original = package
        .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(
        transaction
            .edit(0, |link| {
                let Link::Workbook(link) = link else {
                    panic!("expected workbook link")
                };
                link.sheet_names[0] = "Renamed".into();
                link.cached_sheets[0].rows[0].cells[0].raw_value = Some("2.50".into());
                Ok(())
            })
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let part = package
        .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap();
    let source = std::str::from_utf8(part.blob()).unwrap();
    assert!(source.contains("x:future marker=\"keep\""));
    assert!(source.contains("val=\"Renamed\""));
    assert!(source.contains(">2.50</v>"));
    assert!(!part.blob().eq(original.as_slice()));
    assert_eq!(
        part.rels().get("rId7").unwrap().target_ref(),
        "https://127.0.0.1:9/never-open.xlsx"
    );
}

#[test]
fn transaction_noop_and_inverse_are_exact_and_source_checked() {
    let mut package = workbook_package();
    let before = Snapshot::load(&package).unwrap();
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(
        !transaction
            .edit(0, |link| {
                let Link::Workbook(link) = link else {
                    panic!("expected workbook link")
                };
                assert_eq!(link.sheet_names[0], "Data");
                Ok(())
            })
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(Snapshot::load(&package).unwrap(), before);

    let mut changed = Transaction::new(&mut package).unwrap();
    changed
        .edit(0, |link| {
            let Link::Workbook(link) = link else {
                panic!("expected workbook link")
            };
            link.sheet_names[0] = "Changed".into();
            Ok(())
        })
        .unwrap();
    let commit = changed.commit().unwrap();
    let patch = commit.patch().clone();
    patch.inverse().apply(&mut package).unwrap();
    assert_eq!(Snapshot::load(&package).unwrap(), before);

    let mut stale = workbook_package();
    stale
        .get_part_mut(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
        .unwrap()
        .set_blob(b"<externalLink/>".to_vec());
    assert!(patch.apply(&mut stale).is_err());
    assert_eq!(
        stale
            .get_part(&PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap())
            .unwrap()
            .blob(),
        b"<externalLink/>"
    );
}

#[test]
fn package_crud_keeps_relationship_graph_bounded_and_inert() {
    let mut package = workbook_package();
    let added = add_external_link(
        &mut package,
        Link::Dde(Dde {
            service: "Excel".into(),
            topic: "https://127.0.0.1:9/never-open.xlsx".into(),
            items: Vec::new(),
        }),
        Conformance::Transitional,
    )
    .unwrap();
    assert!(
        load_external_links(&package)
            .unwrap()
            .iter()
            .any(|entry| entry.relationship_id == added.relationship_id)
    );
    let index = load_external_links(&package)
        .unwrap()
        .iter()
        .position(|entry| entry.relationship_id == added.relationship_id)
        .unwrap();
    let removed = remove_external_link(&mut package, index).unwrap().unwrap();
    assert_eq!(removed.relationship_id, added.relationship_id);
    assert_eq!(load_external_links(&package).unwrap().len(), 1);
    assert!(
        package.get_part(&removed.part_uri).is_err(),
        "removed external-link part must not become an orphan"
    );
}
