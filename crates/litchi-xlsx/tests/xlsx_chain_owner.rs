use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsx::Package;
use litchi_xlsx::chain::{self, Cell, Chain, Conformance, Sheet};

const CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

fn package() -> OpcPackage {
    Package::create()
        .expect("create workbook package")
        .into_plain_opc()
}

#[test]
fn calculation_chain_package_operations_preserve_the_canonical_owner() {
    let mut package = package();
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "D4").unwrap());

    assert!(chain::put(&mut package, &chain, Conformance::Strict).unwrap());
    assert_eq!(
        chain::load(&package).unwrap(),
        Some((chain.clone(), Conformance::Strict))
    );
    assert!(chain::remove(&mut package).unwrap());
    assert_eq!(chain::load(&package).unwrap(), None);
    assert!(!chain::remove(&mut package).unwrap());
}

#[test]
fn calculation_chain_survives_package_publication_and_reopen() {
    let mut package = package();
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "A1").unwrap());
    chain::put(&mut package, &chain, Conformance::Transitional).unwrap();

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("chain-owner.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path).unwrap();

    assert_eq!(
        chain::load(&reopened).unwrap(),
        Some((chain, Conformance::Transitional))
    );
    assert_eq!(
        reopened
            .iter_parts()
            .filter(|part| part.content_type() == CHAIN_CONTENT_TYPE)
            .count(),
        1
    );
    assert!(
        reopened
            .get_part(&PackURI::new("/xl/calcChain.xml").unwrap())
            .is_ok()
    );
}
