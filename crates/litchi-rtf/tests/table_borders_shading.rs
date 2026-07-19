use litchi_rtf::{BorderStyle, RtfDocument, RtfWriter, ShadingPattern};

fn write(document: &RtfDocument) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

#[test]
fn parses_complete_row_and_cell_family_and_writes_canonically() {
    let source = r#"{\rtf1\trowd\trbrdrt\brdrs\brdrw11\brdrcf1\brsp2\brdrsh\trbrdrl\brdrth\brdrw12\trbrdrb\brdrdashsm\brdrw13\trbrdrr\brdrdashd\brdrw14\trbrdrh\brdrdashdd\brdrw15\trbrdrv\brdrdb\brdrw16\trbgdkcross\trcfpat1\trcbpat2\trshdng3750\clbrdrt\brdrtriple\brdrw17\clbrdrl\brdrtnthsg\brdrw18\clbrdrb\brdrthtnsg\brdrw19\clbrdrr\brdrtnthtnsg\brdrw20\cldglu\brdrwavy\brdrw21\cldgll\brdrwavydb\brdrw22\clbgfdiag\clcfpat2\clcbpat1\clshdng6250\cellx1000\intbl X\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let row = &document.tables()[0].rows()[0];
    assert_eq!(row.borders().top.unwrap().style, BorderStyle::Single);
    assert_eq!(row.borders().left.unwrap().style, BorderStyle::Thick);
    assert_eq!(
        row.borders().horizontal.unwrap().style,
        BorderStyle::DotDotDash
    );
    assert_eq!(row.shading().pattern, Some(ShadingPattern::DarkCross));
    assert_eq!(row.shading().amount, Some(3750));
    let cell = &row.cells()[0];
    assert_eq!(
        cell.borders().upper_left_to_lower_right.unwrap().style,
        BorderStyle::Wavy
    );
    assert_eq!(
        cell.borders().upper_right_to_lower_left.unwrap().style,
        BorderStyle::WavyDouble
    );
    assert_eq!(
        cell.shading().pattern,
        Some(ShadingPattern::ForwardDiagonal)
    );
    assert_eq!(cell.shading().amount, Some(6250));
    let first = write(&document);
    assert!(first.contains("\\trbrdrt\\brdrs\\brdrw11\\brdrcf1\\brsp2\\brdrsh\\trbrdrl\\brdrth"));
    assert!(first.contains("\\trbgdkcross\\trcfpat1\\trcbpat2\\trshdng3750"));
    assert!(first.contains("\\cldglu\\brdrwavy\\brdrw21\\brdrcf0\\brsp0\\cldgll"));
    assert!(first.contains("\\clbgfdiag\\clcfpat2\\clcbpat1\\clshdng6250"));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(first, write(&reparsed));
    assert_eq!(reparsed.tables()[0].rows()[0].borders(), row.borders());
}

#[test]
fn restores_groups_resets_trowd_and_snapshots_at_cellx() {
    let source = r#"{\rtf1\trowd{\trbrdrt\brdrs\brdrw10\trshdng5000\clbrdrt\brdrs\brdrw11\clshdng2500}\trbrdrb\brdrdb\brdrw12\clbrdrl\brdrdot\brdrw13\clcfpat2\cellx1000\clbrdrr\brdrs\brdrw14\clcbpat3\cellx2000\intbl A\cell B\cell\row\trowd\cellx1000\intbl C\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert!(rows[0].borders().top.is_none());
    assert_eq!(rows[0].borders().bottom.unwrap().style, BorderStyle::Double);
    assert!(rows[0].shading().amount.is_none());
    assert_eq!(
        rows[0].cells()[0].borders().left.unwrap().style,
        BorderStyle::Dotted
    );
    assert!(rows[0].cells()[0].borders().right.is_none());
    assert_eq!(rows[0].cells()[0].shading().foreground_color, Some(2));
    assert_eq!(rows[0].cells()[1].borders().right.unwrap().width, 14);
    assert_eq!(rows[0].cells()[1].shading().background_color, Some(3));
    assert_eq!(*rows[1].borders(), Default::default());
    assert_eq!(rows[1].shading(), Default::default());
    assert_eq!(*rows[1].cells()[0].borders(), Default::default());
}

#[test]
fn applies_end_defined_nested_decorations_without_outer_leakage() {
    let source = r#"{\rtf1\trowd\trbrdrt\brdrs\brdrw10\clbrdrt\brdrs\brdrw11\cellx5000\intbl\itap2 Inner\nestcell{\*\nesttableprops\itap2\trowd\trbrdrh\brdrdb\brdrw12\trbghoriz\trshdng1500\cldglu\brdrdot\brdrw13\clbgvert\clshdng2500\cellx1000\nestrow}\intbl\itap1\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let outer = &document.tables()[0].rows()[0];
    assert!(outer.borders().horizontal.is_none());
    let nested = &outer.cells()[0].nested_tables()[0].table.rows()[0];
    assert_eq!(
        nested.borders().horizontal.unwrap().style,
        BorderStyle::Double
    );
    assert_eq!(nested.shading().pattern, Some(ShadingPattern::Horizontal));
    assert_eq!(
        nested.cells()[0]
            .borders()
            .upper_left_to_lower_right
            .unwrap()
            .style,
        BorderStyle::Dotted
    );
    assert_eq!(
        nested.cells()[0].shading().pattern,
        Some(ShadingPattern::Vertical)
    );
    assert_eq!(
        RtfDocument::parse(&write(&document)).unwrap().tables()[0].rows()[0].cells()[0]
            .nested_tables()[0]
            .table
            .rows()[0]
            .borders(),
        nested.borders()
    );
}

#[test]
fn parses_real_libreoffice_row_borders() {
    let source =
        include_str!("../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/fdo44715.rtf");
    let document = RtfDocument::parse(source).unwrap();
    let row = document
        .tables()
        .iter()
        .flat_map(|table| table.rows())
        .find(|row| row.borders().top.is_some())
        .unwrap();
    assert_eq!(row.borders().top.unwrap().width, 45);
    assert_eq!(row.borders().left.unwrap().style, BorderStyle::Single);
}

#[test]
fn rejects_malformed_parameters_order_duplicates_bounds_and_caps() {
    for source in [
        r#"{\rtf1\trowd\trbrdrt1\brdrs\cellx1}"#,
        r#"{\rtf1\trowd\trbrdrt\brdrw10\brdrs\cellx1}"#,
        r#"{\rtf1\trowd\trbrdrt\brdrs\brdrs\cellx1}"#,
        r#"{\rtf1\trowd\trbrdrt\brdrs\brdrw76\cellx1}"#,
        r#"{\rtf1\trowd\trbrdrt\brdrs\brsp31681\cellx1}"#,
        r#"{\rtf1\trowd\trbrdrt\brdrs\brdrcf-1\cellx1}"#,
        r#"{\rtf1\trowd\trbrdrt\brdrs\brdrcf65536\cellx1}"#,
        r#"{\rtf1\trowd\trshdng\cellx1}"#,
        r#"{\rtf1\trowd\trshdng10001\cellx1}"#,
        r#"{\rtf1\trowd\trbghoriz1\cellx1}"#,
        r#"{\rtf1\trowd\trbghoriz\trbgvert\cellx1}"#,
        r#"{\rtf1\trowd\clcfpat-1\cellx1}"#,
        r#"{\rtf1\trowd\clcbpat65536\cellx1}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    let mut source = String::from("{\\rtf1\\trowd");
    for index in 0..=4096 {
        source.push_str("\\clbrdrt\\brdrs\\clshdng1\\cellx");
        source.push_str(&(index + 1).to_string());
    }
    source.push('}');
    assert!(RtfDocument::parse(&source).is_err());
}
