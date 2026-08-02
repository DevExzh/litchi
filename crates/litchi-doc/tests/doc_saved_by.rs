use litchi_cfb::OleFile;
use litchi_doc::SavedByTable;
use litchi_doc::parts::fib::FileInformationBlock;
use std::fs::File;
use std::path::Path;

#[test]
fn apache_poi_saved_by_table_is_exact_and_round_trips() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document/saved-by-table.doc");
    let mut ole = OleFile::open(File::open(path).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    let saved_by = SavedByTable::parse(&fib, &table_stream).unwrap();

    let expected = [
        (
            "cic22",
            "C:\\DOCUME~1\\phamill\\LOCALS~1\\Temp\\AutoRecovery save of Iraq - security.asd",
        ),
        (
            "cic22",
            "C:\\DOCUME~1\\phamill\\LOCALS~1\\Temp\\AutoRecovery save of Iraq - security.asd",
        ),
        (
            "cic22",
            "C:\\DOCUME~1\\phamill\\LOCALS~1\\Temp\\AutoRecovery save of Iraq - security.asd",
        ),
        ("JPratt", "C:\\TEMP\\Iraq - security.doc"),
        ("JPratt", "A:\\Iraq - security.doc"),
        ("ablackshaw", "C:\\ABlackshaw\\Iraq - security.doc"),
        ("ablackshaw", "C:\\ABlackshaw\\A;Iraq - security.doc"),
        ("ablackshaw", "A:\\Iraq - security.doc"),
        ("MKhan", "C:\\TEMP\\Iraq - security.doc"),
        ("MKhan", "C:\\WINNT\\Profiles\\mkhan\\Desktop\\Iraq.doc"),
    ];
    assert_eq!(saved_by.entries().len(), expected.len());
    for (entry, (author, location)) in saved_by.entries().iter().zip(expected) {
        assert_eq!(entry.author(), author);
        assert_eq!(entry.location(), location);
    }

    let (offset, length) = fib.get_table_pointer(71).unwrap();
    let start = offset as usize;
    let end = start + length as usize;
    assert_eq!(saved_by.to_bytes().unwrap(), table_stream[start..end]);
}
