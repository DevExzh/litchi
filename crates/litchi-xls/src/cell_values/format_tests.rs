#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "format fixtures intentionally fail fast"
)]

use super::*;
use crate::{BorderStyle, Color, FillPattern, HorizontalAlignment, VerticalAlignment};

fn fixture() -> Vec<u8> {
    let mut builder = crate::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_number(data, 0, 0, 1.0).unwrap();
    builder.write_number(data, 0, 1, 2.0).unwrap();
    builder.write_string(data, 1, 0, "text").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

fn reader(snapshot: &Snapshot) -> Reader<Cursor<Vec<u8>>> {
    Reader::new(Cursor::new(snapshot.bytes().to_vec())).unwrap()
}

#[test]
fn alignment_border_and_fill_changes_author_one_shared_xf() {
    let bytes = fixture();
    let source = Snapshot::from_bytes(bytes.clone(), Limits::default()).unwrap();
    let xf_count = source.inner.xf_records.len();
    let change = FormatChange {
        horizontal: Some(HorizontalAlignment::Center),
        vertical: Some(VerticalAlignment::Top),
        wrap_text: Some(true),
        indent: Some(2),
        border_bottom: Some(BorderStyle::Thin),
        border_left: Some(BorderStyle::Double),
        fill_pattern: Some(FillPattern::Solid),
        fill_foreground: Some(Color::new(0xff, 0x00, 0x00)),
        ..FormatChange::default()
    };
    let mut edit = source.edit();
    edit.set_cell_format(
        Selector::Position(Position::new(0)),
        Reference::new(0, 0).unwrap(),
        &change,
    )
    .unwrap();
    edit.set_cell_format(
        Selector::Position(Position::new(0)),
        Reference::new(0, 1).unwrap(),
        &change,
    )
    .unwrap();
    // An absent coordinate gains a formatting-only blank cell.
    edit.set_cell_format(
        Selector::Position(Position::new(0)),
        Reference::new(5, 5).unwrap(),
        &change,
    )
    .unwrap();
    let (target, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(target.inner.xf_records.len(), xf_count + 1);
    let reader = reader(&target);
    let worksheet = reader.xls_worksheet(0).unwrap();
    for (row, column) in [(0, 0), (0, 1), (5, 5)] {
        let format = worksheet.format_for_cell(row, column).unwrap();
        assert_eq!(format.alignment().horizontal(), HorizontalAlignment::Center);
        assert_eq!(format.alignment().vertical(), VerticalAlignment::Top);
        assert!(format.alignment().wraps_text());
        assert_eq!(format.alignment().indent(), 2);
        assert_eq!(format.borders().bottom().style(), BorderStyle::Thin);
        assert_eq!(format.borders().left().style(), BorderStyle::Double);
        assert_eq!(format.fill().pattern(), FillPattern::Solid);
        assert_eq!(
            format.fill().foreground_color(reader.palette()),
            Some(Color::new(0xff, 0x00, 0x00))
        );
    }
    assert!(matches!(
        worksheet.get_cell(5, 5).map(|cell| cell.value().clone()),
        Some(CellValue::Empty)
    ));
    // The untouched cell keeps its original XF.
    assert_eq!(
        target
            .worksheet(Selector::Position(Position::new(0)))
            .unwrap()
            .unwrap()
            .cell(Reference::new(1, 0).unwrap())
            .unwrap()
            .unwrap()
            .style(),
        source
            .worksheet(Selector::Position(Position::new(0)))
            .unwrap()
            .unwrap()
            .cell(Reference::new(1, 0).unwrap())
            .unwrap()
            .unwrap()
            .style()
    );
    assert_eq!(patch.inverse().apply(&target).unwrap().bytes(), &bytes[..]);
    let replayed = patch.semantic().apply(&source).unwrap().into_parts().0;
    assert_eq!(replayed.bytes(), target.bytes());
}

#[test]
fn ineffective_and_invalid_format_changes() {
    let source = Snapshot::from_bytes(fixture(), Limits::default()).unwrap();
    let mut edit = source.edit();
    edit.set_cell_format(
        Selector::Position(Position::new(0)),
        Reference::new(0, 0).unwrap(),
        &FormatChange::default(),
    )
    .unwrap();
    let current = {
        let reader = reader(&source);
        reader
            .xls_worksheet(0)
            .unwrap()
            .format_for_cell(0, 0)
            .unwrap()
            .alignment()
            .horizontal()
    };
    let same = FormatChange {
        horizontal: Some(current),
        ..FormatChange::default()
    };
    edit.set_cell_format(
        Selector::Position(Position::new(0)),
        Reference::new(0, 0).unwrap(),
        &same,
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());

    let mut edit = source.edit();
    let change = FormatChange {
        indent: Some(16),
        ..FormatChange::default()
    };
    assert!(
        edit.set_cell_format(
            Selector::Position(Position::new(0)),
            Reference::new(0, 0).unwrap(),
            &change
        )
        .is_err()
    );
    let change = FormatChange {
        fill_foreground: Some(Color::new(0x12, 0x34, 0x56)),
        ..FormatChange::default()
    };
    assert!(matches!(
        edit.set_cell_format(
            Selector::Position(Position::new(0)),
            Reference::new(0, 0).unwrap(),
            &change
        ),
        Err(Error::UnsupportedFeature(_))
    ));
    assert!(edit.structural_changes.is_empty() && edit.resource_changes.is_empty());
}
