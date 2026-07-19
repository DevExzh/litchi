use litchi_rtf::{RtfDocument, RtfWriter, TableDistanceUnit};

#[test]
fn parses_row_cell_units_resets_groups_and_destinations() {
    let source = r#"{\rtf1\trowd\trpaddl100\trpaddfl3\trspdt20\trspdft3\clpadr30\clpadfr3\clspdb40\clspdfb3\cellx1000\intbl A\cell\row\trowd{\trpaddl200\trpaddfl3}\clpadt10\clpadft3\cellx1000\intbl B\cell{\*\unknown\trpaddl999999\trpaddfl9 ignored}\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(rows[0].padding().left.value, Some(100));
    assert_eq!(rows[0].padding().left.unit, Some(TableDistanceUnit::Twips));
    assert_eq!(rows[0].spacing().top.value, Some(20));
    assert_eq!(rows[0].cells()[0].padding().right.value, Some(30));
    assert_eq!(rows[0].cells()[0].spacing().bottom.value, Some(40));
    assert_eq!(rows[1].padding().left.value, None);
    assert_eq!(rows[1].cells()[0].padding().top.value, Some(10));
}

#[test]
fn writer_round_trips_deterministically() {
    let document=RtfDocument::parse(r#"{\rtf1\trowd\trpaddl70\trpaddfl3\trspdr21\trspdfr3\clpadt57\clpadft3\cellx1200\intbl Cell\cell\row}"#).unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first = String::from_utf8(first).unwrap();
    assert!(first.contains("\\trpaddl70\\trpaddfl3"));
    assert!(first.contains("\\clpadt57\\clpadft3"));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(
        reparsed.tables()[0].rows()[0].padding(),
        document.tables()[0].rows()[0].padding()
    );
    assert_eq!(
        reparsed.tables()[0].rows()[0].cells()[0].padding(),
        document.tables()[0].rows()[0].cells()[0].padding()
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, String::from_utf8(second).unwrap());
}

#[test]
fn parses_real_libreoffice_row_and_cell_distances() {
    let row_source = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf165923.rtf"),
    )
    .unwrap();
    let row_doc = RtfDocument::parse(&row_source).unwrap();
    assert!(
        row_doc
            .tables()
            .iter()
            .flat_map(|t| t.rows())
            .any(|row| row.padding().left.value == Some(57)
                && row.padding().left.unit == Some(TableDistanceUnit::Twips))
    );
    let cell_source=std::fs::read_to_string(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../3rdparty/libreoffice-core/sw/qa/writerfilter/dmapper/data/floattable-vertical-frame-offset.rtf")).unwrap();
    let cell_doc = RtfDocument::parse(&cell_source).unwrap();
    assert!(
        cell_doc
            .tables()
            .iter()
            .flat_map(|t| t.rows())
            .flat_map(|r| r.cells())
            .any(|cell| cell.padding().top.value == Some(57))
    );
}

#[test]
fn rejects_malformed_values_and_units() {
    for source in [
        r#"{\rtf1\trowd\trpaddl X}"#,
        r#"{\rtf1\trowd\trpaddl-1 X}"#,
        r#"{\rtf1\trowd\trpaddl31681 X}"#,
        r#"{\rtf1\trowd\trpaddfl X}"#,
        r#"{\rtf1\trowd\trpaddfl1 X}"#,
        r#"{\rtf1\trowd\clspdfb4 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    assert!(RtfDocument::parse(r#"{\rtf1{\*\unknown\trpaddl-1\trpaddfl9 bad}Visible}"#).is_ok());
}
