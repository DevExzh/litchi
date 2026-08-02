use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::{DocumentAssociatedStrings, Package};
use std::fs::File;
use std::path::Path;

#[test]
fn apache_poi_associated_strings_integrate_and_round_trip_exactly() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document/47950_normal.doc");

    let mut package = Package::from_reader(File::open(&path).unwrap()).unwrap();
    let document = package.document().unwrap();
    let integrated = document.associated_strings().unwrap();
    assert_eq!(integrated.author(), "Ross Johnson");
    assert_eq!(integrated.last_revised_by(), "Ross Johnson");

    let mut ole = OleFile::open(File::open(path).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    let parsed = DocumentAssociatedStrings::parse(&fib, &table_stream)
        .unwrap()
        .unwrap();
    let (offset, length) = fib.get_table_pointer(32).unwrap();
    let start = offset as usize;
    let end = start + length as usize;
    assert_eq!(parsed.to_bytes().unwrap(), table_stream[start..end]);
}
