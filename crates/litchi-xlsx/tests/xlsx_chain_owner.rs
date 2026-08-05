use litchi_ooxml::xlsx::Workbook;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsx::chain::{self, Cell, Chain, Conformance, Sheet};

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

#[test]
fn legacy_host_caches_the_canonical_chain_owner() {
    let mut workbook = Workbook::new(workbook_package()).expect("open workbook");
    let chain =
        Chain::new(Cell::new(Sheet::new(1).expect("sheet"), "D4").expect("calculation cell"));

    workbook
        .put_chain(chain.clone(), Conformance::Strict)
        .expect("put chain");
    assert_eq!(workbook.chain(), Some(&chain));
    assert_eq!(workbook.chain_conformance(), Some(Conformance::Strict));
    assert_eq!(
        chain::load(workbook.opc_package()).expect("load canonical chain"),
        Some((chain, Conformance::Strict))
    );

    assert!(workbook.remove_chain().expect("remove chain"));
    assert_eq!(workbook.chain(), None);
    assert!(!workbook.remove_chain().expect("idempotent remove"));
}

#[test]
fn legacy_host_restores_the_leaf_chain_after_writer_materialization() {
    let mut workbook = Workbook::create().expect("create workbook");
    let chain =
        Chain::new(Cell::new(Sheet::new(1).expect("sheet"), "A1").expect("calculation cell"));
    workbook
        .put_chain(chain.clone(), Conformance::Transitional)
        .expect("put chain");

    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("chain-owner.xlsx");
    workbook.save(&path).expect("save workbook");
    let reopened = Workbook::open(&path).expect("reopen workbook");

    assert_eq!(reopened.chain(), Some(&chain));
    assert_eq!(
        reopened.chain_conformance(),
        Some(Conformance::Transitional)
    );
    assert_eq!(
        reopened
            .opc_package()
            .iter_parts()
            .filter(|part| part.content_type() == CHAIN_CONTENT_TYPE)
            .count(),
        1
    );
}

fn workbook_package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook = BlobPart::new(
        PackURI::new("/xl/workbook.xml").expect("workbook URI"),
        ct::SML_SHEET_MAIN.into(),
        format!(r#"<workbook xmlns="{S}"><sheets/></workbook>"#).into_bytes(),
    );
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    package
}
