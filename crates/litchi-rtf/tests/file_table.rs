use litchi_rtf::{FileLocation, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r#"{\rtf1\ansi\ansicpg1250\uc1
{\*\filetbl
{\file\fid1\frelative2\fosnum1\fvaliddos\fvalidntfs C:\\dir\\\'8a\u20320?.txt;}
{\file\fid7\fvalidmac\fnetwork \\server\\share\\remote.doc;}
{\file\fid9\fnonfilesys Printer Queue;}}
Body}"#;

#[test]
fn parses_decodes_and_round_trips_file_table() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let table = doc.file_table().unwrap();
    assert_eq!(table.entries().len(), 3);

    let local = table.get(1).unwrap();
    assert_eq!(local.name, "C:\\dir\\Š你.txt");
    assert_eq!(local.relative_path_level, Some(2));
    assert_eq!(local.operating_system, Some(1));
    assert!(local.valid_on.dos);
    assert!(local.valid_on.ntfs);
    assert_eq!(local.location, FileLocation::Local);

    let network = table.get(7).unwrap();
    assert_eq!(network.name, "\\server\\share\\remote.doc");
    assert!(network.valid_on.mac);
    assert_eq!(network.location, FileLocation::Network);
    assert_eq!(table.get(9).unwrap().location, FileLocation::NonFileSystem);
    assert_eq!(doc.text().trim(), "Body");

    let mut first_bytes = Vec::new();
    RtfWriter::new(&mut first_bytes).write_document(&doc).unwrap();
    let reparsed = RtfDocument::parse_bytes(&first_bytes).unwrap();
    assert_eq!(table, reparsed.file_table().unwrap());
    let mut second_bytes = Vec::new();
    RtfWriter::new(&mut second_bytes)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_bytes, second_bytes);
}

#[test]
fn rejects_malformed_file_tables_and_active_content() {
    let malformed = [
        r#"{\rtf1{\filetbl{\file\fid1 x;}}}"#,
        r#"{\rtf1{\*\filetbl}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1 x;}}{\*\filetbl{\file\fid2 y;}}}"#,
        r#"{\rtf1 Body{\*\filetbl{\file\fid1 x;}}}"#,
        r#"{\rtf1{\file\fid1 x;}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1 x;}{\file\fid1 y;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid2 x;}{\file\fid1 y;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid-1 x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1 ;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1 x}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1\frelative-1 x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1\frelative256 x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1\fosnum256 x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1\fnetwork\fnonfilesys x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1\fvaliddos\fvaliddos x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1{\field X}x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1{\object X}x;}}}"#,
        r#"{\rtf1{\*\filetbl{\file\fid1\bin2 ABx;}}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}

#[test]
fn rejects_overlong_file_name() {
    let name = "x".repeat(4097);
    let source = format!(r#"{{\rtf1{{\*\filetbl{{\file\fid1 {name};}}}}}}"#);
    assert!(RtfDocument::parse(&source).is_err());
}
