use litchi_rtf::{
    RtfDocument, RtfWriter, TableCellTextFlow, TableCellVerticalAlignment, TableRowAlignment,
    TextDirection,
};

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_complete_layout_family_and_round_trips_deterministically() {
    let source = r#"{\rtf1\trowd\trhdr\trkeep\trkeepfollow\trqc\rtlrow\clvertalt\cltxlrtb\clFitText\clNoWrap\clhidemark\cellx1000\clvertalc\cltxtbrl\cellx2000\clvertalb\cltxbtlr\cellx3000\cltxlrtbv\cellx4000\cltxtbrlv\cellx5000\intbl A\cell B\cell C\cell D\cell E\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let row = &document.tables()[0].rows()[0];
    assert_eq!(row.direction(), Some(TextDirection::RightToLeft));
    assert_eq!(row.layout().alignment, Some(TableRowAlignment::Center));
    assert!(row.layout().header && row.layout().keep_together && row.layout().keep_with_following);
    let cells = row.cells();
    assert_eq!(
        cells[0].layout().vertical_alignment,
        Some(TableCellVerticalAlignment::Top)
    );
    assert_eq!(
        cells[1].layout().vertical_alignment,
        Some(TableCellVerticalAlignment::Center)
    );
    assert_eq!(
        cells[2].layout().vertical_alignment,
        Some(TableCellVerticalAlignment::Bottom)
    );
    assert!(cells[0].layout().fit_text && cells[0].layout().no_wrap && cells[0].layout().hide_mark);
    assert_eq!(
        cells
            .iter()
            .map(|cell| cell.layout().text_flow)
            .collect::<Vec<_>>(),
        vec![
            Some(TableCellTextFlow::LeftToRightTopToBottom),
            Some(TableCellTextFlow::RightToLeftTopToBottom),
            Some(TableCellTextFlow::LeftToRightBottomToTop),
            Some(TableCellTextFlow::LeftToRightTopToBottomVertical),
            Some(TableCellTextFlow::TopToBottomRightToLeftVertical)
        ]
    );
    let first = write(&document);
    let reparsed = RtfDocument::parse(&first).unwrap();
    let second = write(&reparsed);
    assert_eq!(first, second);
    assert_eq!(reparsed.tables()[0].rows()[0].layout(), row.layout());
}

#[test]
fn resets_trowd_and_restores_groups_and_inert_destinations() {
    let source = r#"{\rtf1\trowd{\trhdr\trqr\rtlrow\clvertalb\clNoWrap}{\*\unknown\trhdr\clNoWrap}\trql\cellx1000\intbl A\cell\row\trowd\cellx1000\intbl B\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(rows[0].layout().alignment, Some(TableRowAlignment::Left));
    assert!(!rows[0].layout().header);
    assert_eq!(rows[0].direction(), None);
    assert_eq!(*rows[0].cells()[0].layout(), Default::default());
    assert_eq!(*rows[1].layout(), Default::default());
    assert_eq!(*rows[1].cells()[0].layout(), Default::default());
}

#[test]
fn applies_end_defined_nested_layout_without_leaking_outer_state() {
    let source = r#"{\rtf1\trowd\trhdr\trqc\clvertalc\cellx5000\intbl\itap2 Inner\nestcell{\*\nesttableprops\itap2\trowd\trkeep\trqr\ltrrow\clvertalb\cltxbtlr\clNoWrap\cellx1000\nestrow}\intbl\itap1\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let outer = &document.tables()[0].rows()[0];
    assert!(outer.layout().header);
    assert_eq!(outer.layout().alignment, Some(TableRowAlignment::Center));
    assert_eq!(
        outer.cells()[0].layout().vertical_alignment,
        Some(TableCellVerticalAlignment::Center)
    );
    let nested = &outer.cells()[0].nested_tables()[0].table.rows()[0];
    assert!(nested.layout().keep_together);
    assert_eq!(nested.layout().alignment, Some(TableRowAlignment::Right));
    assert_eq!(nested.direction(), Some(TextDirection::LeftToRight));
    assert_eq!(
        nested.cells()[0].layout().vertical_alignment,
        Some(TableCellVerticalAlignment::Bottom)
    );
    assert_eq!(
        nested.cells()[0].layout().text_flow,
        Some(TableCellTextFlow::LeftToRightBottomToTop)
    );
    assert!(nested.cells()[0].layout().no_wrap);
    assert_eq!(
        RtfDocument::parse(&write(&document)).unwrap().tables()[0].rows()[0].cells()[0]
            .nested_tables()[0]
            .table
            .rows()[0]
            .layout(),
        nested.layout()
    );
}

#[test]
fn parses_real_libreoffice_row_and_text_flow_fixtures() {
    let base = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sw/qa");
    let header = RtfDocument::parse(
        &std::fs::read_to_string(base.join("extras/rtfimport/data/tblrepeat.rtf")).unwrap(),
    )
    .unwrap();
    assert!(header.tables()[0].rows()[0].layout().header);
    let flow = RtfDocument::parse(
        &std::fs::read_to_string(base.join("extras/rtfexport/data/btlr-cell.rtf")).unwrap(),
    )
    .unwrap();
    let cells = flow.tables()[0].rows()[0].cells();
    assert_eq!(
        cells[0].layout().text_flow,
        Some(TableCellTextFlow::LeftToRightBottomToTop)
    );
    assert_eq!(cells[1].layout().text_flow, None);
    assert_eq!(
        cells[2].layout().text_flow,
        Some(TableCellTextFlow::RightToLeftTopToBottom)
    );
}

#[test]
fn rejects_parameters_and_enforces_existing_cell_cap() {
    for word in [
        "ltrrow1",
        "rtlrow0",
        "trhdr1",
        "trkeep0",
        "trkeepfollow1",
        "trql1",
        "trqc0",
        "trqr1",
        "clvertalt1",
        "clvertalc0",
        "clvertalb1",
        "cltxlrtb1",
        "cltxtbrl0",
        "cltxbtlr1",
        "cltxlrtbv0",
        "cltxtbrlv1",
        "clFitText0",
        "clNoWrap1",
        "clhidemark0",
    ] {
        let source = format!("{{\\rtf1\\trowd\\{word}\\cellx1000\\intbl X\\cell\\row}}");
        assert!(RtfDocument::parse(&source).is_err(), "accepted {word}");
    }
    let mut source = String::from("{\\rtf1\\trowd");
    for index in 0..=4096 {
        source.push_str("\\clvertalt\\cellx");
        source.push_str(&(index + 1).to_string());
    }
    source.push('}');
    assert!(RtfDocument::parse(&source).is_err());
}
