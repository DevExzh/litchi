//! Regression tests for the layered external-link owner.

use super::*;
use litchi_opc::part::BlobPart;
use litchi_opc::{PackURI, Part};

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
