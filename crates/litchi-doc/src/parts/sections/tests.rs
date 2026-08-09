use super::SectionsTable;
use crate::package::Result;
use crate::parts::numbering::NumberFormat;
use crate::parts::revisions::RevisionAuthorTable;
use crate::section::borders::{self, Borders};
use crate::section::columns::{Column, Layout};
use crate::section::{
    BreakKind, ChapterNumberSeparator, FootnotePosition, LineNumberRestart, NoteNumberRestart,
    PageGrid, PageGridMode, PageOrientation, Protection, TextFlow, VerticalJustification,
    VerticalMargin,
};
use crate::{OpenOptions, Package};
use std::path::{Path, PathBuf};

fn fixed_sprm(opcode: u16, operand: &[u8]) -> Vec<u8> {
    let mut bytes = opcode.to_le_bytes().to_vec();
    bytes.extend_from_slice(operand);
    bytes
}

fn variable_sprm(opcode: u16, operand: &[u8]) -> Vec<u8> {
    let mut bytes = opcode.to_le_bytes().to_vec();
    bytes.push(operand.len() as u8);
    bytes.extend_from_slice(operand);
    bytes
}

fn build_section_data(grpprls: &[Option<Vec<u8>>]) -> (Vec<u8>, Vec<u8>) {
    let mut data = Vec::new();
    for cp in 0..=grpprls.len() {
        data.extend_from_slice(&((cp * 10) as u32).to_le_bytes());
    }
    let mut word_document = vec![0; 8];
    for grpprl in grpprls {
        data.extend_from_slice(&0u16.to_le_bytes());
        let fc_sepx = if let Some(grpprl) = grpprl {
            let offset = word_document.len() as i32;
            word_document.extend_from_slice(&(grpprl.len() as i16).to_le_bytes());
            word_document.extend_from_slice(grpprl);
            offset
        } else {
            -1
        };
        data.extend_from_slice(&fc_sepx.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
    }
    (data, word_document)
}

fn parse_synthetic(grpprls: &[Option<Vec<u8>>]) -> Result<SectionsTable> {
    let (data, word_document) = build_section_data(grpprls);
    SectionsTable::parse_data(
        &data,
        &word_document,
        &RevisionAuthorTable::from_authors(&["Unknown", "Editor"]),
    )
}

#[test]
fn parses_defaults_all_breaks_and_binary_lookup() {
    let mut grpprls = Vec::new();
    for value in 0u8..=4 {
        grpprls.push(Some(fixed_sprm(0x3009, &[value])));
    }
    grpprls.push(None);
    let parsed = parse_synthetic(&grpprls).unwrap();
    assert_eq!(parsed.sections().len(), 6);
    assert_eq!(parsed.sections()[0].break_kind, BreakKind::Continuous);
    assert_eq!(parsed.sections()[1].break_kind, BreakKind::NewColumn);
    assert_eq!(parsed.sections()[2].break_kind, BreakKind::NewPage);
    assert_eq!(parsed.sections()[3].break_kind, BreakKind::EvenPage);
    assert_eq!(parsed.sections()[4].break_kind, BreakKind::OddPage);
    let defaults = &parsed.sections()[5];
    assert_eq!(defaults.break_kind, BreakKind::NewPage);
    assert_eq!(
        (defaults.page.width_twips, defaults.page.height_twips),
        (12_240, 15_840)
    );
    assert_eq!(defaults.page.orientation, PageOrientation::Portrait);
    assert_eq!(defaults.page.margins.left_twips, 1_800);
    assert_eq!(defaults.columns.count(), 1);
    assert_eq!(parsed.section_at_cp(0).unwrap().start_cp, 0);
    assert_eq!(parsed.section_at_cp(10).unwrap().start_cp, 10);
    assert_eq!(parsed.section_at_cp(59).unwrap().end_cp, 60);
    assert!(parsed.section_at_cp(60).is_none());
}

#[test]
fn parses_orientation_signed_margins_and_later_wins() {
    let mut grpprl = Vec::new();
    grpprl.extend(fixed_sprm(0x301D, &[1]));
    grpprl.extend(fixed_sprm(0x301D, &[2]));
    grpprl.extend(fixed_sprm(0xB01F, &15_840u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0xB020, &12_240u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0xB021, &1_000u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0xB022, &1_100u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x9023, &(-1_200i16).to_le_bytes()));
    grpprl.extend(fixed_sprm(0x9024, &1_300i16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0xB025, &200u16.to_le_bytes()));
    let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
    let page = parsed.sections()[0].page;
    assert_eq!(page.orientation, PageOrientation::Landscape);
    assert_eq!((page.width_twips, page.height_twips), (15_840, 12_240));
    assert_eq!(page.margins.top, VerticalMargin::Fixed(1_200));
    assert_eq!(page.margins.top.signed_twips(), -1_200);
    assert_eq!(page.margins.bottom, VerticalMargin::Minimum(1_300));
    assert_eq!(page.margins.gutter_twips, 200);
}

#[test]
fn parses_even_and_unequal_columns() {
    let mut even = Vec::new();
    even.extend(fixed_sprm(0x500B, &2u16.to_le_bytes()));
    even.extend(fixed_sprm(0x900C, &900i16.to_le_bytes()));
    even.extend(fixed_sprm(0x3019, &[1]));

    let mut unequal = Vec::new();
    unequal.extend(fixed_sprm(0x500B, &2u16.to_le_bytes()));
    unequal.extend(fixed_sprm(0x3005, &[0]));
    for (index, width) in [1_000u16, 1_100, 1_200].into_iter().enumerate() {
        let mut operand = vec![index as u8];
        operand.extend_from_slice(&width.to_le_bytes());
        unequal.extend(fixed_sprm(0xF203, &operand));
    }
    for (index, spacing) in [100u16, 200].into_iter().enumerate() {
        let mut operand = vec![index as u8];
        operand.extend_from_slice(&spacing.to_le_bytes());
        unequal.extend(fixed_sprm(0xF204, &operand));
    }
    let parsed = parse_synthetic(&[Some(even), Some(unequal)]).unwrap();
    assert_eq!(
        parsed.sections()[0].columns,
        Layout::even(3, 900, true).unwrap()
    );
    let columns = parsed.sections()[1]
        .columns
        .unequal_columns()
        .expect("parser produced unequal columns");
    assert_eq!(columns.len(), 3);
    assert_eq!(columns[0], Column::new(1_000, Some(100)).unwrap());
    assert_eq!(columns[2], Column::new(1_200, None).unwrap());
}

#[test]
fn unequal_columns_are_order_independent_and_later_indexed_values_win() {
    let mut grpprl = Vec::new();
    for (index, width) in [1_000u16, 1_100, 1_200].into_iter().enumerate() {
        let mut operand = vec![index as u8];
        operand.extend_from_slice(&width.to_le_bytes());
        grpprl.extend(fixed_sprm(0xF203, &operand));
    }
    for (index, spacing) in [100u16, 200].into_iter().enumerate() {
        let mut operand = vec![index as u8];
        operand.extend_from_slice(&spacing.to_le_bytes());
        grpprl.extend(fixed_sprm(0xF204, &operand));
    }
    let mut replacement = vec![0];
    replacement.extend_from_slice(&1_500u16.to_le_bytes());
    grpprl.extend(fixed_sprm(0xF203, &replacement));
    grpprl.extend(fixed_sprm(0x3005, &[0]));
    grpprl.extend(fixed_sprm(0x500B, &2u16.to_le_bytes()));

    let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
    let columns = parsed.sections()[0]
        .columns
        .unequal_columns()
        .expect("parser produced unequal columns");
    assert_eq!(columns[0].width_twips(), 1_500);
    assert_eq!(columns[0].spacing_after_twips(), Some(100));
    assert_eq!(columns[2].width_twips(), 1_200);
}

#[test]
fn rejects_out_of_count_and_final_column_spacing_operands() {
    let invalid = |extra_index: u8| {
        let mut grpprl = Vec::new();
        grpprl.extend(fixed_sprm(0x3005, &[0]));
        grpprl.extend(fixed_sprm(0x500B, &1u16.to_le_bytes()));
        for index in 0..2u8 {
            let mut operand = vec![index];
            operand.extend_from_slice(&1_000u16.to_le_bytes());
            grpprl.extend(fixed_sprm(0xF203, &operand));
        }
        let mut spacing = vec![0];
        spacing.extend_from_slice(&100u16.to_le_bytes());
        grpprl.extend(fixed_sprm(0xF204, &spacing));
        let mut extra = vec![extra_index];
        extra.extend_from_slice(&200u16.to_le_bytes());
        grpprl.extend(fixed_sprm(0xF204, &extra));
        parse_synthetic(&[Some(grpprl)])
    };
    assert!(invalid(1).is_err());
    assert!(invalid(2).is_err());
}

#[test]
fn parses_section_page_line_and_note_numbering() {
    let mut grpprl = Vec::new();
    grpprl.extend(fixed_sprm(0x3000, &[4]));
    grpprl.extend(fixed_sprm(0x3001, &[3]));
    grpprl.extend(fixed_sprm(0x300E, &[2]));
    grpprl.extend(fixed_sprm(0x3011, &[1]));
    grpprl.extend(fixed_sprm(0x501C, &123u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x7044, &70_000u32.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x3012, &[0]));
    grpprl.extend(fixed_sprm(0x3013, &[2]));
    grpprl.extend(fixed_sprm(0x5015, &3u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x9016, &360u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x501B, &6u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x301A, &[3]));
    grpprl.extend(fixed_sprm(0x303B, &[2]));
    grpprl.extend(fixed_sprm(0x303C, &[2]));
    grpprl.extend(fixed_sprm(0x303E, &[1]));
    grpprl.extend(fixed_sprm(0x503F, &6u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5040, &3u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5041, &8u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5042, &4u16.to_le_bytes()));

    let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
    let section = &parsed.sections()[0];
    assert_eq!(
        section.page_numbering.chapter_separator,
        ChapterNumberSeparator::EnDash
    );
    assert_eq!(section.page_numbering.chapter_heading_level, Some(3));
    assert_eq!(
        section.page_numbering.number_format,
        NumberFormat::LowerRoman
    );
    assert!(section.page_numbering.restart);
    assert_eq!(section.page_numbering.start_at, 70_000);
    assert_eq!(section.line_numbering.interval, 3);
    assert_eq!(
        section.line_numbering.restart,
        LineNumberRestart::Continuous
    );
    assert_eq!(section.line_numbering.distance_twips, 360);
    assert_eq!(section.line_numbering.start_at, 7);
    assert_eq!(
        section.page.vertical_justification,
        VerticalJustification::Bottom
    );
    assert!(!section.notes.show_endnotes_at_section_end);
    assert_eq!(
        section.notes.footnote_position,
        FootnotePosition::BeneathText
    );
    assert_eq!(section.notes.footnote_restart, NoteNumberRestart::EachPage);
    assert_eq!(
        section.notes.endnote_restart,
        NoteNumberRestart::EachSection
    );
    assert_eq!(section.notes.footnote_offset_operand, 6);
    assert_eq!(
        section.notes.footnote_number_format,
        NumberFormat::UpperLetter
    );
    assert_eq!(section.notes.endnote_offset_operand, 8);
    assert_eq!(
        section.notes.endnote_number_format,
        NumberFormat::LowerLetter
    );
}

#[test]
fn rejects_invalid_section_numbering_operands() {
    for grpprl in [
        fixed_sprm(0x3000, &[5]),
        fixed_sprm(0x3001, &[10]),
        fixed_sprm(0x300E, &[60]),
        fixed_sprm(0x3013, &[3]),
        fixed_sprm(0x5015, &101u16.to_le_bytes()),
        fixed_sprm(0x501B, &32_767u16.to_le_bytes()),
        fixed_sprm(0x303B, &[0]),
        fixed_sprm(0x303E, &[2]),
        fixed_sprm(0x503F, &16_384u16.to_le_bytes()),
        fixed_sprm(0x5040, &0x0100u16.to_le_bytes()),
        fixed_sprm(0x7044, &2_147_483_647u32.to_le_bytes()),
    ] {
        assert!(parse_synthetic(&[Some(grpprl)]).is_err());
    }
}

#[test]
fn parses_section_behavior_and_paper_settings() {
    let mut grpprl = Vec::new();
    grpprl.extend(fixed_sprm(0x3006, &[0]));
    grpprl.extend(fixed_sprm(0x3006, &[1]));
    grpprl.extend(fixed_sprm(0x5007, &7u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5008, &9u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x300A, &[1]));
    grpprl.extend(fixed_sprm(0x5026, &42u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x3228, &[1]));
    grpprl.extend(fixed_sprm(0x322A, &[1]));
    grpprl.extend(fixed_sprm(0x3239, &[1]));
    grpprl.extend(fixed_sprm(0x703A, &0xAABB_CCDDu32.to_le_bytes()));
    let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
    let section = &parsed.sections()[0];
    assert_eq!(section.behavior.protection, Protection::Unprotected);
    assert!(section.behavior.different_first_page);
    assert!(section.behavior.right_to_left);
    assert!(section.behavior.right_to_left_gutter);
    assert!(section.behavior.preserve_properties_for_revision);
    assert_eq!(section.behavior.revision_save_id, Some(0xAABB_CCDD));
    assert_eq!(section.paper.first_page_source, Some(7));
    assert_eq!(section.paper.other_page_source, Some(9));
    assert_eq!(section.paper.requested_paper_kind, Some(42));
}

#[test]
fn rejects_invalid_section_behavior_booleans() {
    for opcode in [0x3006, 0x300A, 0x3228, 0x322A, 0x3239] {
        assert!(parse_synthetic(&[Some(fixed_sprm(opcode, &[2]))]).is_err());
    }
}

#[test]
fn retains_section_property_revisions() {
    let timestamp: u32 = 30 | (14 << 6) | (12 << 11) | (7 << 16) | (126 << 20) | (1 << 29);
    let mut operand = vec![1];
    operand.extend_from_slice(&1i16.to_le_bytes());
    operand.extend_from_slice(&timestamp.to_le_bytes());
    let parsed = parse_synthetic(&[Some(variable_sprm(0xD243, &operand))]).unwrap();
    let revision = &parsed.revisions()[0];
    assert_eq!((revision.start, revision.end), (0, 10));
    assert_eq!(revision.author_index, 1);
    assert_eq!(revision.author, "Editor");
    assert_eq!(revision.timestamp.unwrap().year, 2026);
}

#[test]
fn rejects_malformed_tables_properties_and_columns() {
    let authors = RevisionAuthorTable::from_authors(&["Unknown"]);
    assert!(SectionsTable::parse_data(&[0; 19], &[], &authors).is_err());
    let (mut duplicate_cps, word_document) = build_section_data(&[None]);
    duplicate_cps[4..8].copy_from_slice(&0u32.to_le_bytes());
    assert!(SectionsTable::parse_data(&duplicate_cps, &word_document, &authors).is_err());

    for grpprl in [
        fixed_sprm(0x3009, &[5]),
        fixed_sprm(0x301D, &[0]),
        fixed_sprm(0xB01F, &143u16.to_le_bytes()),
        fixed_sprm(0x500B, &44u16.to_le_bytes()),
        fixed_sprm(0x9023, &i16::MIN.to_le_bytes()),
    ] {
        assert!(parse_synthetic(&[Some(grpprl)]).is_err());
    }

    let mut incomplete = Vec::new();
    incomplete.extend(fixed_sprm(0x500B, &1u16.to_le_bytes()));
    incomplete.extend(fixed_sprm(0x3005, &[0]));
    assert!(parse_synthetic(&[Some(incomplete)]).is_err());

    let (data, _) = build_section_data(&[Some(Vec::new())]);
    assert!(SectionsTable::parse_data(&data, &[0; 9], &authors).is_err());
}

fn poi_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document")
        .join(name)
}

#[test]
fn opens_poi_bug53453_section_layout() {
    let mut package = Package::open(poi_fixture("Bug53453Section.doc")).unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.sections().len(), 2);
    for section in document.sections() {
        assert_eq!(section.page.margins.left_twips, 1_440);
        assert_eq!(section.page.margins.right_twips, 1_440);
        assert_eq!(section.page.margins.top, VerticalMargin::Minimum(1_440));
        assert_eq!(section.page.margins.bottom, VerticalMargin::Minimum(1_440));
    }
    assert_eq!(document.sections()[0].columns.count(), 1);
    assert_eq!(document.sections()[1].columns.count(), 3);
}

#[test]
fn exposes_layout_after_doc_decryption() {
    for (name, password) in [
        ("password_tika_binaryrc4.doc", "tika"),
        ("password_password_cryptoapi.doc", "password"),
    ] {
        let mut package = Package::open(poi_fixture(name)).unwrap();
        let document = package
            .document_with_options(OpenOptions::default().with_password(password.to_owned().into()))
            .unwrap();
        assert!(!document.sections().is_empty());
        assert!(document.sections().iter().all(|section| {
            (144..=31_680).contains(&section.page.width_twips)
                && (144..=31_680).contains(&section.page.height_twips)
        }));
    }
}

#[test]
fn page_border_defaults_and_later_operands_win() {
    let defaults = parse_synthetic(&[None]).unwrap();
    assert_eq!(defaults.sections()[0].page_borders, Borders::default());

    let mut grpprl = Vec::new();
    grpprl.extend(fixed_sprm(0x702B, &[4, 0x06, 1, 0]));
    grpprl.extend(fixed_sprm(0x702B, &[8, 0x01, 6, 3]));
    grpprl.extend(fixed_sprm(0x702C, &[16, 0x03, 0, 0x64]));
    grpprl.extend(fixed_sprm(0x702D, &[24, 0x40, 16, 0x45]));
    grpprl.extend(fixed_sprm(0x702E, &[0, 0, 0, 0]));
    grpprl.extend(fixed_sprm(0x522F, &[0, 0]));
    grpprl.extend(fixed_sprm(0x522F, &[0x2A, 0]));
    let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
    let borders = parsed.sections()[0].page_borders;

    assert_eq!(
        borders.top,
        Some(borders::Border {
            style: borders::Style::Single,
            width_eighth_points: 8,
            color: borders::Color::Red,
            spacing_points: 3,
            shadow: false,
            frame: false,
        })
    );
    assert_eq!(borders.left.unwrap().spacing_points, 4);
    assert!(borders.left.unwrap().shadow);
    let borders::Style::Art(art) = borders.bottom.unwrap().style else {
        panic!("expected an art page border");
    };
    assert_eq!(art.code(), 0x40);
    assert!(borders.bottom.unwrap().frame);
    assert_eq!(borders.right, None);
    assert_eq!(borders.apply_to, borders::ApplyTo::AllButFirstPage);
    assert_eq!(borders.depth, borders::Depth::Behind);
    assert_eq!(borders.offset_from, borders::Offset::PageEdge);
}

#[test]
fn rejects_malformed_page_border_operands() {
    for border in [
        fixed_sprm(0x702B, &[8, 0x02, 0, 0]),
        fixed_sprm(0x702B, &[8, 0x1A, 0, 0]),
        fixed_sprm(0x702B, &[8, 0x1B, 0, 0]),
        fixed_sprm(0x702B, &[8, 0xE4, 0, 0]),
        fixed_sprm(0x702B, &[8, 0x01, 0x11, 0]),
        fixed_sprm(0x702B, &[8, 0x00, 0x11, 0]),
        fixed_sprm(0x702B, &[8, 0x01, 0, 0x80]),
    ] {
        assert!(parse_synthetic(&[Some(border)]).is_err());
    }

    for properties in [
        fixed_sprm(0x522F, &[0x03, 0]),
        fixed_sprm(0x522F, &[0x10, 0]),
        fixed_sprm(0x522F, &[0x40, 0]),
        fixed_sprm(0x522F, &[0, 1]),
        fixed_sprm(0x522F, &[0]),
    ] {
        assert!(parse_synthetic(&[Some(properties)]).is_err());
    }
}

#[test]
fn parses_page_grid_text_flow_defaults_and_later_overrides() {
    let defaults = parse_synthetic(&[None]).unwrap();
    assert_eq!(defaults.sections()[0].page_grid, PageGrid::default());
    assert_eq!(
        defaults.sections()[0].text_flow,
        TextFlow::HorizontalNonAsian
    );

    let mut grpprl = Vec::new();
    grpprl.extend(fixed_sprm(0x7030, &(-670_925i32).to_le_bytes()));
    grpprl.extend(fixed_sprm(0x7030, &6_144i32.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x9031, &360u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x9031, &480u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5032, &1u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5032, &3u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5033, &0u16.to_le_bytes()));
    grpprl.extend(fixed_sprm(0x5033, &5u16.to_le_bytes()));
    let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
    assert_eq!(
        parsed.sections()[0].page_grid,
        PageGrid {
            mode: PageGridMode::EnforceCharacterGrid,
            character_pitch_adjustment: 6_144,
            line_pitch_twips: Some(480),
        }
    );
    assert_eq!(parsed.sections()[0].text_flow, TextFlow::VerticalNonAsian);

    let mut disabled = Vec::new();
    disabled.extend(fixed_sprm(0x7030, &1_024i32.to_le_bytes()));
    disabled.extend(fixed_sprm(0x9031, &240u16.to_le_bytes()));
    disabled.extend(fixed_sprm(0x5032, &2u16.to_le_bytes()));
    disabled.extend(fixed_sprm(0x5032, &0u16.to_le_bytes()));
    let parsed = parse_synthetic(&[Some(disabled)]).unwrap();
    assert_eq!(parsed.sections()[0].page_grid.mode, PageGridMode::Disabled);
    assert_eq!(
        parsed.sections()[0].page_grid.character_pitch_adjustment,
        1_024
    );
    assert_eq!(parsed.sections()[0].page_grid.line_pitch_twips, Some(240));
}

#[test]
fn rejects_malformed_page_grid_and_text_flow_operands() {
    for grpprl in [
        fixed_sprm(0x7030, &(-670_926i32).to_le_bytes()),
        fixed_sprm(0x7030, &6_488_065i32.to_le_bytes()),
        fixed_sprm(0x9031, &0u16.to_le_bytes()),
        fixed_sprm(0x9031, &31_681u16.to_le_bytes()),
        fixed_sprm(0x5032, &4u16.to_le_bytes()),
        fixed_sprm(0x5033, &6u16.to_le_bytes()),
        fixed_sprm(0x7030, &[0, 0, 0]),
        fixed_sprm(0x9031, &[1]),
    ] {
        assert!(parse_synthetic(&[Some(grpprl)]).is_err());
    }

    assert!(parse_synthetic(&[Some(fixed_sprm(0x5032, &1u16.to_le_bytes()))]).is_err());

    let mut later_enabled = Vec::new();
    later_enabled.extend(fixed_sprm(0x5032, &0u16.to_le_bytes()));
    later_enabled.extend(fixed_sprm(0x5032, &2u16.to_le_bytes()));
    assert!(parse_synthetic(&[Some(later_enabled)]).is_err());
}
