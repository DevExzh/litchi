//! OLE equation objects added to and removed from an Excel-written workbook
//! that holds other embedded objects.
#![cfg(feature = "equation")]

use litchi_xlsx::Package;
use litchi_xlsx::equation_objects::Source;

const EXCEL: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/excel/excel-links-objects.xlsx");

#[test]
fn an_equation_object_joins_and_leaves_excel_embedded_objects() {
    let package = Package::from_bytes(EXCEL).unwrap();
    assert!(package.equation_objects(0).unwrap().is_empty());
    let limits = litchi_equation::Limits::DEFAULT;

    let mut edit = package.edit_equation_objects().unwrap();
    let written = edit
        .add(0, 24, 3, Source::Latex(r"\frac{a}{b}"), &limits)
        .unwrap();
    let added = edit.commit().unwrap().into_package();
    let objects = added.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].prog_id(), "Equation.DSMT4");
    assert_eq!(objects[0].latex(&limits).unwrap(), r"\frac{a}{b}");
    // Excel's embedded document stays an OLE object beside the equation.
    let mut ole = added.edit_ole_objects(0usize).unwrap();
    assert_eq!(ole.objects().unwrap().unwrap().objects.len(), 2);

    let mut edit = added.edit_equation_objects().unwrap();
    edit.remove(0, written.shape_id()).unwrap();
    let removed = edit.commit().unwrap().into_package();
    assert!(removed.equation_objects(0).unwrap().is_empty());
    let reopened = Package::from_bytes(&removed.to_bytes().unwrap()).unwrap();
    let mut ole = reopened.edit_ole_objects(0usize).unwrap();
    assert_eq!(ole.objects().unwrap().unwrap().objects.len(), 1);
    assert_eq!(
        removed.to_bytes().unwrap(),
        Package::from_bytes(EXCEL).unwrap().to_bytes().unwrap()
    );
}
