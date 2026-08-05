use litchi_doc::{
    ApplyTo, Art, Border, BorderError, Borders, Color, Depth, Offset, Package, Style, Writer,
};
use std::io::Cursor;

fn round_trip(writer: &mut Writer) -> litchi_doc::Section {
    writer.add_paragraph("Borders").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    package.document().unwrap().sections()[0].clone()
}

fn border(style: Style, color: Color) -> Border {
    Border {
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
        round_trip(&mut Writer::new()).page_borders,
        Borders::default()
    );
    let mut writer = Writer::new();
    writer.set_section_page_borders(Borders::default()).unwrap();
    assert_eq!(round_trip(&mut writer).page_borders, Borders::default());
}

#[test]
fn four_edges_art_and_shared_placement_round_trip() {
    let expected = Borders {
        top: Some(border(Style::Single, Color::Red)),
        left: Some(Border {
            shadow: true,
            ..border(Style::Double, Color::DarkBlue)
        }),
        bottom: Some(Border {
            style: Style::Art(Art::try_from(0x40).unwrap()),
            frame: true,
            ..border(Style::Single, Color::LightGray)
        }),
        right: Some(border(Style::ThreeDEngrave, Color::Automatic)),
        apply_to: ApplyTo::AllButFirstPage,
        depth: Depth::Behind,
        offset_from: Offset::PageEdge,
    };
    let mut writer = Writer::new();
    writer.set_section_page_borders(expected).unwrap();
    assert_eq!(writer.section_page_borders(), Some(&expected));
    assert_eq!(round_trip(&mut writer).page_borders, expected);
}

#[test]
fn setter_is_atomic_and_rejects_out_of_range_spacing() {
    let valid = Borders {
        top: Some(border(Style::Dotted, Color::Black)),
        ..Borders::default()
    };
    let mut writer = Writer::new();
    writer.set_section_page_borders(valid).unwrap();
    let invalid = Borders {
        top: Some(Border {
            spacing_points: 32,
            ..valid.top.unwrap()
        }),
        ..valid
    };
    assert_eq!(
        invalid.validate().unwrap_err(),
        BorderError::InvalidSpacing(32)
    );
    assert!(writer.set_section_page_borders(invalid).is_err());
    assert_eq!(writer.section_page_borders(), Some(&valid));
}

#[test]
fn art_codes_report_typed_validation_errors() {
    assert_eq!(Art::try_from(0x3F), Err(BorderError::InvalidArt(0x3F)));
    assert_eq!(Art::try_from(0xE4), Err(BorderError::InvalidArt(0xE4)));
}

#[test]
fn later_set_replaces_and_clear_restores_omission() {
    let mut writer = Writer::new();
    writer
        .set_section_page_borders(Borders {
            top: Some(border(Style::Single, Color::Blue)),
            ..Borders::default()
        })
        .unwrap();
    let replacement = Borders {
        right: Some(border(Style::Wave, Color::Green)),
        apply_to: ApplyTo::FirstPage,
        ..Borders::default()
    };
    writer.set_section_page_borders(replacement).unwrap();
    assert_eq!(round_trip(&mut writer).page_borders, replacement);
    writer.clear_section_page_borders();
    assert!(writer.section_page_borders().is_none());
}
