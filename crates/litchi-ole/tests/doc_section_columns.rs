use litchi_ole::doc::{
    DocWriter, Package, SectionColumn, SectionColumnError, SectionColumnLayout, SectionTextFlow,
};
use std::io::Cursor;

fn round_trip(writer: &mut DocWriter) -> litchi_ole::doc::DocSection {
    writer.add_paragraph("Columns").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    package.document().unwrap().sections()[0].clone()
}

#[test]
fn omitted_layout_reads_as_the_normative_single_column_default() {
    let section = round_trip(&mut DocWriter::new());
    assert_eq!(
        section.columns,
        SectionColumnLayout::Even {
            count: 1,
            spacing_twips: 720,
            line_between: false,
        }
    );
    assert!(!section.behavior.right_to_left);
    assert_eq!(section.text_flow, SectionTextFlow::HorizontalNonAsian);
}

#[test]
fn equal_columns_rtl_and_vertical_flow_round_trip() {
    let mut writer = DocWriter::new();
    let mut layout = SectionColumnLayout::even(3, 900, false).unwrap();
    layout.set_line_between(true);
    writer.set_section_columns(layout.clone()).unwrap();
    writer.set_section_right_to_left(true);
    writer.set_section_text_flow(SectionTextFlow::VerticalNonAsian);
    assert_eq!(writer.section_columns(), Some(&layout));
    assert!(writer.section_right_to_left());
    assert_eq!(writer.section_text_flow(), SectionTextFlow::VerticalNonAsian);

    let section = round_trip(&mut writer);
    assert_eq!(section.columns, layout);
    assert!(section.behavior.right_to_left);
    assert_eq!(section.text_flow, SectionTextFlow::VerticalNonAsian);
}

#[test]
fn unequal_width_and_spacing_arrays_round_trip_without_rtl_reversal() {
    let layout = SectionColumnLayout::unequal(
        vec![
            SectionColumn {
                width_twips: 1_000,
                spacing_after_twips: Some(120),
            },
            SectionColumn {
                width_twips: 2_000,
                spacing_after_twips: Some(240),
            },
            SectionColumn {
                width_twips: 3_000,
                spacing_after_twips: None,
            },
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
        SectionColumnLayout::even(0, 720, false).unwrap_err(),
        SectionColumnError::InvalidCount(0)
    );
    assert!(SectionColumnLayout::even(45, 720, false).is_err());
    assert!(SectionColumnLayout::even(2, 31_681, false).is_err());
    assert!(SectionColumnLayout::unequal(
        vec![SectionColumn {
            width_twips: 717,
            spacing_after_twips: None,
        }],
        false,
    )
    .is_err());
    assert!(SectionColumnLayout::unequal(
        vec![
            SectionColumn {
                width_twips: 1_000,
                spacing_after_twips: None,
            },
            SectionColumn {
                width_twips: 1_000,
                spacing_after_twips: None,
            },
        ],
        false,
    )
    .is_err());
    assert!(SectionColumnLayout::unequal(
        vec![SectionColumn {
            width_twips: 1_000,
            spacing_after_twips: Some(100),
        }],
        false,
    )
    .is_err());

    let mut writer = DocWriter::new();
    writer
        .set_section_columns(SectionColumnLayout::even(2, 720, false).unwrap())
        .unwrap();
    writer.clear_section_columns();
    assert!(writer.section_columns().is_none());
}
