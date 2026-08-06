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
