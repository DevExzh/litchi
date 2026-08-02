use litchi_doc::{
    DocWriter, Package, SectionPageBorder, SectionPageBorderApplyTo, SectionPageBorderArt,
    SectionPageBorderColor, SectionPageBorderDepth, SectionPageBorderError,
    SectionPageBorderOffsetFrom, SectionPageBorderStyle, SectionPageBorders,
};
use std::io::Cursor;

fn round_trip(writer: &mut DocWriter) -> litchi_doc::DocSection {
    writer.add_paragraph("Borders").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    package.document().unwrap().sections()[0].clone()
}

fn border(style: SectionPageBorderStyle, color: SectionPageBorderColor) -> SectionPageBorder {
    SectionPageBorder {
        style,
        width_eighth_points: 8,
        color,
        spacing_points: 4,
        shadow: false,
        frame: false,
    }
}

#[test]
fn omitted_and_explicit_default_borders_emit_the_normative_default() {
    assert_eq!(
        round_trip(&mut DocWriter::new()).page_borders,
        SectionPageBorders::default()
    );
    let mut writer = DocWriter::new();
    writer
        .set_section_page_borders(SectionPageBorders::default())
        .unwrap();
    assert_eq!(
        round_trip(&mut writer).page_borders,
        SectionPageBorders::default()
    );
}

#[test]
fn four_edges_art_and_shared_placement_round_trip() {
    let expected = SectionPageBorders {
        top: Some(border(
            SectionPageBorderStyle::Single,
            SectionPageBorderColor::Red,
        )),
        left: Some(SectionPageBorder {
            shadow: true,
            ..border(
                SectionPageBorderStyle::Double,
                SectionPageBorderColor::DarkBlue,
            )
        }),
        bottom: Some(SectionPageBorder {
            style: SectionPageBorderStyle::Art(SectionPageBorderArt::try_from(0x40).unwrap()),
            frame: true,
            ..border(
                SectionPageBorderStyle::Single,
                SectionPageBorderColor::LightGray,
            )
        }),
        right: Some(border(
            SectionPageBorderStyle::ThreeDEngrave,
            SectionPageBorderColor::Automatic,
        )),
        apply_to: SectionPageBorderApplyTo::AllButFirstPage,
        depth: SectionPageBorderDepth::Behind,
        offset_from: SectionPageBorderOffsetFrom::PageEdge,
    };
    let mut writer = DocWriter::new();
    writer.set_section_page_borders(expected).unwrap();
    assert_eq!(writer.section_page_borders(), Some(&expected));
    assert_eq!(round_trip(&mut writer).page_borders, expected);
}

#[test]
fn setter_is_atomic_and_rejects_out_of_range_spacing() {
    let valid = SectionPageBorders {
        top: Some(border(
            SectionPageBorderStyle::Dotted,
            SectionPageBorderColor::Black,
        )),
        ..SectionPageBorders::default()
    };
    let mut writer = DocWriter::new();
    writer.set_section_page_borders(valid).unwrap();
    let invalid = SectionPageBorders {
        top: Some(SectionPageBorder {
            spacing_points: 32,
            ..valid.top.unwrap()
        }),
        ..valid
    };
    assert_eq!(
        invalid.validate().unwrap_err(),
        SectionPageBorderError::InvalidSpacing(32)
    );
    assert!(writer.set_section_page_borders(invalid).is_err());
    assert_eq!(writer.section_page_borders(), Some(&valid));
}

#[test]
fn later_set_replaces_and_clear_restores_omission() {
    let mut writer = DocWriter::new();
    writer
        .set_section_page_borders(SectionPageBorders {
            top: Some(border(
                SectionPageBorderStyle::Single,
                SectionPageBorderColor::Blue,
            )),
            ..SectionPageBorders::default()
        })
        .unwrap();
    let replacement = SectionPageBorders {
        right: Some(border(
            SectionPageBorderStyle::Wave,
            SectionPageBorderColor::Green,
        )),
        apply_to: SectionPageBorderApplyTo::FirstPage,
        ..SectionPageBorders::default()
    };
    writer.set_section_page_borders(replacement).unwrap();
    assert_eq!(round_trip(&mut writer).page_borders, replacement);
    writer.clear_section_page_borders();
    assert!(writer.section_page_borders().is_none());
}
