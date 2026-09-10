#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use std::borrow::Cow;

use litchi_rtf::{
    Cell, ParagraphFrameHorizontalPosition, ParagraphFrameHorizontalReference,
    ParagraphFrameTextFlow, ParagraphFrameVerticalPosition, ParagraphFrameVerticalReference,
    ParagraphFrameWrap, RtfDocument, RtfWriter, StyleBlock,
};

fn block<'a>(document: &'a RtfDocument<'a>, needle: &str) -> &'a StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(needle))
        .unwrap()
}

#[test]
fn parses_positioned_paragraph_frame_controls() {
    let source = concat!(
        r#"{\rtf1\ansi\pard\absw1440\absh-720\phmrg\posnegx-120"#,
        r#"\pvpg\posy240\abslock1\nowrap\dxfrtext90\dfrmtxtx10\dfrmtxty20"#,
        r#"\wrapthrough\overlay\absnoovrlp1\frmtxtbrlv Framed\par}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let frame = block(&document, "Framed").paragraph.frame.unwrap();
    assert_eq!(frame.width_twips, Some(1440));
    assert_eq!(frame.height_twips, Some(-720));
    assert_eq!(
        frame.horizontal_reference,
        ParagraphFrameHorizontalReference::Margin
    );
    assert_eq!(
        frame.horizontal_position,
        ParagraphFrameHorizontalPosition::NegativeOffset(-120)
    );
    assert_eq!(
        frame.vertical_reference,
        ParagraphFrameVerticalReference::Page
    );
    assert_eq!(
        frame.vertical_position,
        ParagraphFrameVerticalPosition::Offset(240)
    );
    assert_eq!(frame.anchor_locked, Some(true));
    assert!(frame.no_wrap);
    assert_eq!(frame.horizontal_text_distance_twips, Some(90));
    assert_eq!(frame.horizontal_text_offset_twips, Some(10));
    assert_eq!(frame.vertical_text_offset_twips, Some(20));
    assert_eq!(frame.wrap, ParagraphFrameWrap::Through);
    assert!(frame.overlay);
    assert_eq!(frame.no_overlap, Some(true));
    assert_eq!(
        frame.text_flow,
        ParagraphFrameTextFlow::TopToBottomRightToLeftVertical
    );

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    assert_eq!(block(&reparsed, "Framed").paragraph.frame, Some(frame));
    assert!(String::from_utf8(bytes).unwrap().contains(r"\absw1440"));
}

#[test]
fn framed_drop_cap_is_written_before_text_flow_controls() {
    let document = RtfDocument::parse(
        r#"{\rtf1\dropcapli2\dropcapt1\absw720\phmrg\posxl\pvmrg\posyt\frmtxtbrl Framed}"#,
    )
    .unwrap();
    let paragraph = block(&document, "Framed").paragraph;
    assert!(paragraph.drop_cap.is_some());
    assert!(paragraph.frame.is_some());

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let text = String::from_utf8(bytes.clone()).unwrap();
    assert!(text.find(r"\dropcapli2").unwrap() < text.find(r"\frmtxtbrl").unwrap());
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let reparsed_paragraph = block(&reparsed, "Framed").paragraph;
    assert_eq!(reparsed_paragraph.drop_cap, paragraph.drop_cap);
    assert_eq!(reparsed_paragraph.frame, paragraph.frame);
}

#[test]
fn frame_state_inherits_scopes_and_pard_resets_it() {
    let document = RtfDocument::parse(
        r#"{\rtf1\ansi\absw720 Outer\par {\absw1440 Inner\par }Tail\par {\pard Reset\par }}"#,
    )
    .unwrap();
    assert_eq!(
        block(&document, "Outer")
            .paragraph
            .frame
            .unwrap()
            .width_twips,
        Some(720)
    );
    assert_eq!(
        block(&document, "Inner")
            .paragraph
            .frame
            .unwrap()
            .width_twips,
        Some(1440)
    );
    assert_eq!(
        block(&document, "Tail")
            .paragraph
            .frame
            .unwrap()
            .width_twips,
        Some(720)
    );
    assert_eq!(block(&document, "Reset").paragraph.frame, None);
}

#[test]
fn frame_controls_are_retained_in_styles_and_default_paragraphs() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\stylesheet{\s7\absw960\phcol\posxc\pvpara\posyc"#,
        r#"\wraparound\frmtxlrtb Style;}}"#,
        r#"{\*\defpap\absw480\phmrg\posxr\pvmrg\posyb\wraptight\frmtxtbrl}"#,
        r#"\s7 Styled}"#,
    ))
    .unwrap();
    let style_frame = document
        .stylesheet()
        .get(7)
        .unwrap()
        .paragraph
        .unwrap()
        .frame
        .unwrap();
    assert_eq!(style_frame.width_twips, Some(960));
    assert_eq!(
        style_frame.horizontal_position,
        ParagraphFrameHorizontalPosition::Center
    );
    assert_eq!(style_frame.wrap, ParagraphFrameWrap::Around);
    assert_eq!(
        document
            .default_formatting()
            .paragraph()
            .unwrap()
            .paragraph
            .frame
            .unwrap()
            .width_twips,
        Some(480)
    );
}

#[test]
fn header_and_footer_paragraph_frames_round_trip() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\header\absw720\pvpara\posyt Header\par}"#,
        r#"{\footer\absw960\phmrg\posyb Footer\par}Body}"#,
    ))
    .unwrap();
    let header = &document.sections()[0].headers_footers[0].paragraphs[0];
    let footer = &document.sections()[0].headers_footers[1].paragraphs[0];
    assert_eq!(header.paragraph.frame.unwrap().width_twips, Some(720));
    assert_eq!(
        header.paragraph.frame.unwrap().vertical_reference,
        ParagraphFrameVerticalReference::Paragraph
    );
    assert_eq!(footer.paragraph.frame.unwrap().width_twips, Some(960));
    assert_eq!(
        footer.paragraph.frame.unwrap().vertical_position,
        ParagraphFrameVerticalPosition::Bottom
    );

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    assert_eq!(
        reparsed.sections()[0].headers_footers[0].paragraphs[0]
            .paragraph
            .frame,
        header.paragraph.frame
    );
    assert_eq!(
        reparsed.sections()[0].headers_footers[1].paragraphs[0]
            .paragraph
            .frame,
        footer.paragraph.frame
    );
}

#[test]
fn zero_frame_toggle_values_are_retained_in_paragraph_and_style() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\stylesheet{\s9\absw960\abslock0\absnoovrlp0 Style;}}"#,
        r#"\absw720\abslock0\absnoovrlp0 Paragraph}"#,
    ))
    .unwrap();
    let paragraph = block(&document, "Paragraph").paragraph.frame.unwrap();
    assert_eq!(paragraph.anchor_locked, Some(false));
    assert_eq!(paragraph.no_overlap, Some(false));
    let style = document
        .stylesheet()
        .get(9)
        .unwrap()
        .paragraph
        .unwrap()
        .frame
        .unwrap();
    assert_eq!(style.anchor_locked, Some(false));
    assert_eq!(style.no_overlap, Some(false));
}

#[test]
fn rejects_missing_or_oversized_frame_parameters() {
    for source in [
        r#"{\rtf1\absw Missing}"#,
        r#"{\rtf1\posx Missing}"#,
        r#"{\rtf1\phmrg1 Missing}"#,
        r#"{\rtf1\absw10000001 Missing}"#,
        r#"{\rtf1\posy-10000001 Missing}"#,
        r#"{\rtf1\posx-1 NegativeHorizontal}"#,
        r#"{\rtf1\posy-1 NegativeVertical}"#,
        r#"{\rtf1\abslock Missing}"#,
        r#"{\rtf1\absnoovrlp Missing}"#,
        r#"{\rtf1\abslock2 Missing}"#,
        r#"{\rtf1\absnoovrlp2 Missing}"#,
        r#"{\rtf1{\stylesheet{\s7\abslock Bad;}}Text}"#,
        r#"{\rtf1{\stylesheet{\s7\absnoovrlp Bad;}}Text}"#,
        r#"{\rtf1{\stylesheet{\s7\abslock2 Bad;}}Text}"#,
        r#"{\rtf1{\stylesheet{\s7\absnoovrlp2 Bad;}}Text}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn table_cell_frames_round_trip_and_preserve_style_frames() {
    let source = r#"{\rtf1\trowd\cellx1000\intbl\absw720\phmrg\posxl\pvmrg\posyt Cell\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let paragraph = &document.tables()[0].rows()[0].cells()[0].paragraphs()[0];
    assert_eq!(paragraph.text_range(), 0..4);
    assert_eq!(paragraph.frame().unwrap().width_twips, Some(720));
    assert!(!paragraph.has_paragraph_break());

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    assert_eq!(
        reparsed.tables()[0].rows()[0].cells()[0].paragraphs()[0]
            .frame()
            .unwrap()
            .width_twips,
        Some(720)
    );

    let styled = RtfDocument::parse(
        r#"{\rtf1{\stylesheet{\s7\absw720\phmrg\posxl\pvmrg\posyt Framed;}}\trowd\cellx1000\intbl\s7 Cell\cell\row}"#,
    )
    .unwrap();
    assert_eq!(
        styled.tables()[0].rows()[0].cells()[0].paragraphs()[0]
            .frame()
            .unwrap()
            .width_twips,
        Some(720)
    );
}

#[test]
fn table_cell_frames_inherit_controls_before_intbl_and_itap() {
    let outer =
        RtfDocument::parse(r#"{\rtf1\absw720\trowd\cellx1000\intbl Cell\cell\row}"#).unwrap();
    assert_eq!(
        outer.tables()[0].rows()[0].cells()[0].paragraphs()[0]
            .frame()
            .unwrap()
            .width_twips,
        Some(720)
    );

    let nested = RtfDocument::parse(
        r#"{\rtf1\absw960\trowd\cellx5000\intbl\itap1 Outer \intbl\itap2 Inner\nestcell{\*\nesttableprops\trowd\cellx1000\nestrow}\intbl\itap1\cell\row}"#,
    )
    .unwrap();
    assert_eq!(
        nested.tables()[0].rows()[0].cells()[0].nested_tables()[0]
            .table
            .rows()[0]
            .cells()[0]
            .paragraphs()[0]
            .frame()
            .unwrap()
            .width_twips,
        Some(960)
    );
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(&nested).unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let outer_cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_eq!(
        outer_cell.paragraphs()[0].frame().unwrap().width_twips,
        Some(960)
    );
    assert_eq!(
        outer_cell.nested_tables()[0].table.rows()[0].cells()[0].paragraphs()[0]
            .frame()
            .unwrap()
            .width_twips,
        Some(960)
    );
}

#[test]
fn table_wrapdefault_keeps_default_line_breaking_without_creating_a_frame() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/165717.rtf");
    let source = std::fs::read_to_string(path).unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    let paragraphs = document
        .tables()
        .iter()
        .flat_map(|table| table.rows())
        .flat_map(litchi_rtf::raw::Row::cells)
        .flat_map(Cell::paragraphs)
        .collect::<Vec<_>>();
    assert!(!paragraphs.is_empty());
    assert!(
        paragraphs
            .iter()
            .all(|paragraph| paragraph.frame().is_none())
    );
}

#[test]
fn table_row_frame_controls_must_be_consistent() {
    for source in [
        r#"{\rtf1\trowd\cellx1000\cellx2000\intbl\absw720 A\cell\pard\intbl B\cell\row}"#,
        r#"{\rtf1\trowd\cellx1000\cellx2000\intbl\absw720 A\cell\intbl\absw960 B\cell\row}"#,
    ] {
        let error = match RtfDocument::parse(source) {
            Ok(_) => String::new(),
            Err(error) => error.to_string(),
        };
        assert!(
            error.contains("identical for every paragraph in a table row"),
            "misclassified {source}: {error}"
        );
    }
}

#[test]
fn constructed_implicit_unframed_cells_cannot_join_a_framed_row() {
    for text in ["plain", ""] {
        let mut document =
            RtfDocument::parse(r#"{\rtf1\trowd\cellx1000\intbl\absw720 Framed\cell\row}"#).unwrap();
        document.tables_mut()[0].rows_mut()[0].add_cell(Cell::new(Cow::Borrowed(text)));

        let mut bytes = Vec::new();
        let error = RtfWriter::new(&mut bytes)
            .write_document(&document)
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("identical for every paragraph in a table row")
        );
    }
}

#[test]
fn constructed_implicit_unframed_nested_cells_obey_the_row_rule() {
    let mut document = RtfDocument::parse(
        r#"{\rtf1\absw720\trowd\cellx5000\intbl Outer\intbl\itap2\absw720 Inner\nestcell{\*\nesttableprops\trowd\cellx1000\nestrow}\intbl\itap1\cell\row}"#,
    )
    .unwrap();
    let nested_row = &mut document.tables_mut()[0].rows_mut()[0].cells_mut()[0].nested_tables_mut()
        [0]
    .table
    .rows_mut()[0];
    nested_row.add_cell(Cell::new(Cow::Borrowed("legacy")));

    let mut bytes = Vec::new();
    let error = RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap_err();
    assert!(
        error
            .to_string()
            .contains("identical for every paragraph in a table row")
    );
}

#[test]
fn table_cell_multiple_paragraph_frames_round_trip() {
    let source = r#"{\rtf1\trowd\cellx1000\intbl\absw720 A\par B\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let paragraphs = document.tables()[0].rows()[0].cells()[0].paragraphs();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0].text_range(), 0..1);
    assert!(paragraphs[0].has_paragraph_break());
    assert_eq!(paragraphs[1].text_range(), 2..3);
    assert_eq!(paragraphs[1].frame().unwrap().width_twips, Some(720));

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let reparsed_paragraphs = reparsed.tables()[0].rows()[0].cells()[0].paragraphs();
    assert_eq!(reparsed_paragraphs.len(), 2);
    assert_eq!(
        reparsed_paragraphs[0].frame().unwrap().width_twips,
        Some(720)
    );
    assert_eq!(
        reparsed_paragraphs[1].frame().unwrap().width_twips,
        Some(720)
    );
}

#[test]
fn table_cell_text_edits_remap_frame_spans() {
    let mut document =
        RtfDocument::parse(r#"{\rtf1\trowd\cellx1000\intbl\absw720 Cell\cell\row}"#).unwrap();
    let cell = &mut document.tables_mut()[0].rows_mut()[0].cells_mut()[0];
    cell.set_text(Cow::Borrowed("Changed")).unwrap();
    assert_eq!(cell.paragraphs()[0].text_range(), 0..7);
    assert_eq!(cell.paragraphs()[0].frame().unwrap().width_twips, Some(720));

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let reparsed_cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_eq!(reparsed_cell.text(), "Changed");
    assert_eq!(
        reparsed_cell.paragraphs()[0].frame().unwrap().width_twips,
        Some(720)
    );
}

#[test]
fn table_cell_soft_line_noop_preserves_one_framed_paragraph() {
    let mut document =
        RtfDocument::parse(r#"{\rtf1\trowd\cellx1000\intbl\absw720 A\line B\cell\row}"#).unwrap();
    let cell = &mut document.tables_mut()[0].rows_mut()[0].cells_mut()[0];
    let before = cell.paragraphs().to_vec();
    let text = cell.text().to_owned();
    cell.set_text(Cow::Owned(text)).unwrap();
    assert_eq!(cell.paragraphs(), before.as_slice());

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let written = String::from_utf8(bytes.clone()).unwrap();
    assert!(written.contains(r"\line "));
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let reparsed_cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_eq!(reparsed_cell.text(), "A\nB");
    assert_eq!(reparsed_cell.paragraphs(), before.as_slice());
}

#[test]
fn table_cell_soft_line_and_paragraph_boundaries_round_trip() {
    let mut document =
        RtfDocument::parse(r#"{\rtf1\trowd\cellx1000\intbl\absw720 A\line B\par C\cell\row}"#)
            .unwrap();
    let cell = &mut document.tables_mut()[0].rows_mut()[0].cells_mut()[0];
    assert_eq!(cell.paragraphs().len(), 2);
    assert!(cell.paragraphs()[0].has_paragraph_break());
    assert_eq!(cell.paragraphs()[1].text_range(), 4..5);

    let text = cell.text().to_owned();
    cell.set_text(Cow::Owned(text)).unwrap();
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let reparsed_cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_eq!(reparsed_cell.text(), "A\nB\nC");
    assert_eq!(reparsed_cell.paragraphs().len(), 2);
    assert!(reparsed_cell.paragraphs()[0].has_paragraph_break());
    assert!(!reparsed_cell.paragraphs()[1].has_paragraph_break());
}

#[test]
fn changed_table_cell_text_remaps_soft_and_paragraph_break_kinds() {
    let mut document =
        RtfDocument::parse(r#"{\rtf1\trowd\cellx1000\intbl\absw720 A\line B\par C\cell\row}"#)
            .unwrap();
    let cell = &mut document.tables_mut()[0].rows_mut()[0].cells_mut()[0];
    cell.set_text(Cow::Owned("Changed\nB\nTail".to_string()))
        .unwrap();
    assert_eq!(cell.paragraphs().len(), 2);
    assert!(cell.paragraphs()[0].has_paragraph_break());
    assert_eq!(cell.paragraphs()[0].text_range(), 0..9);
    assert!(!cell.paragraphs()[1].has_paragraph_break());

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&bytes).unwrap();
    let reparsed_cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_eq!(reparsed_cell.text(), "Changed\nB\nTail");
    assert_eq!(reparsed_cell.paragraphs().len(), 2);
    assert!(reparsed_cell.paragraphs()[0].has_paragraph_break());
}
