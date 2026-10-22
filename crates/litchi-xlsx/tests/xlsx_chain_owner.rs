use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsx::Package;
use litchi_xlsx::chain::{self, Cell, Chain, Conformance, Sheet};

const CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

fn package() -> Package {
    Package::new().expect("create workbook package")
}

#[test]
fn calculation_chain_package_operations_preserve_the_canonical_owner() {
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "D4").unwrap());

    let mut edit = package().edit_calculation_chain().unwrap();
    assert!(edit.set(&chain, Conformance::Strict).unwrap());
    assert_eq!(
        edit.chain().unwrap(),
        Some((chain.clone(), Conformance::Strict))
    );
    let package = edit.commit().unwrap().into_package();
    assert_eq!(
        chain::load(&package.clone().into_plain_opc()).unwrap(),
        Some((chain.clone(), Conformance::Strict))
    );
    let mut edit = package.edit_calculation_chain().unwrap();
    assert!(edit.remove().unwrap());
    assert_eq!(edit.chain().unwrap(), None);
    assert!(!edit.remove().unwrap());
    let package = edit.commit().unwrap().into_package();
    assert_eq!(chain::load(&package.into_plain_opc()).unwrap(), None);
}

#[test]
fn calculation_chain_survives_package_publication_and_reopen() {
    let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "A1").unwrap());
    let mut edit = package().edit_calculation_chain().unwrap();
    edit.set(&chain, Conformance::Transitional).unwrap();
    let package = edit.commit().unwrap().into_package().into_plain_opc();

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("chain-owner.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path, litchi_opc::SourceLimits::default()).unwrap();

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
