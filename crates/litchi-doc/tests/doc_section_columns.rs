use litchi_doc::{
    DocWriter, Package, TextFlow,
    section::columns::{Column, Error, Layout},
};
use std::io::Cursor;

fn round_trip(writer: &mut DocWriter) -> litchi_doc::Section {
    writer.add_paragraph("Columns").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    package.document().unwrap().sections()[0].clone()
}

#[test]
fn omitted_layout_reads_as_the_normative_single_column_default() {
    let section = round_trip(&mut DocWriter::new());
    assert_eq!(section.columns, Layout::even(1, 720, false).unwrap());
    assert!(!section.behavior.right_to_left);
    assert_eq!(section.text_flow, TextFlow::HorizontalNonAsian);
}

#[test]
fn equal_columns_rtl_and_vertical_flow_round_trip() {
    let mut writer = DocWriter::new();
    let mut layout = Layout::even(3, 900, false).unwrap();
    layout.set_line_between(true);
    writer.set_section_columns(layout.clone()).unwrap();
    writer.set_section_right_to_left(true);
    writer.set_section_text_flow(TextFlow::VerticalNonAsian);
    assert_eq!(writer.section_columns(), Some(&layout));
    assert!(writer.section_right_to_left());
    assert_eq!(writer.section_text_flow(), TextFlow::VerticalNonAsian);

    let section = round_trip(&mut writer);
    assert_eq!(section.columns, layout);
    assert!(section.behavior.right_to_left);
    assert_eq!(section.text_flow, TextFlow::VerticalNonAsian);
}

#[test]
fn unequal_width_and_spacing_arrays_round_trip_without_rtl_reversal() {
    let layout = Layout::unequal(
        vec![
            Column::new(1_000, Some(120)).unwrap(),
            Column::new(2_000, Some(240)).unwrap(),
            Column::new(3_000, None).unwrap(),
        ],
        true,
    )
    .unwrap();
    let mut writer = DocWriter::new();
    writer.set_section_columns(layout.clone()).unwrap();
    writer.set_section_right_to_left(true);
    let section = round_trip(&mut writer);
    assert_eq!(section.columns, layout);
    assert!(section.behavior.right_to_left);
}

#[test]
fn mutation_rejects_invalid_counts_widths_and_spacing_dependencies() {
    assert_eq!(
        Layout::even(0, 720, false).unwrap_err(),
        Error::InvalidCount(0)
    );
    assert!(Layout::even(45, 720, false).is_err());
    assert!(Layout::even(2, 31_681, false).is_err());
    assert!(Column::new(717, None).is_err());
    assert!(
        Layout::unequal(
            vec![
                Column::new(1_000, None).unwrap(),
                Column::new(1_000, None).unwrap(),
            ],
            false,
        )
        .is_err()
    );
    assert!(Layout::unequal(vec![Column::new(1_000, Some(100)).unwrap()], false).is_err());

    let mut writer = DocWriter::new();
    writer
        .set_section_columns(Layout::even(2, 720, false).unwrap())
        .unwrap();
    writer.clear_section_columns();
    assert!(writer.section_columns().is_none());
}
