use super::super::support::*;

#[test]
fn preserves_ordered_character_property_revision_state() {
    let formatting = CharacterFormatting {
        italic: Some(true),
        preserved_properties_for_revision: Some(Box::new(CharacterFormatting {
            bold: Some(true),
            ..CharacterFormatting::default()
        })),
        ..CharacterFormatting::default()
    };
    let mut fonts = FontTableBuilder::new();
    let grpprl = build_revision_chpx_grpprl(&formatting, &mut fonts, None).unwrap();
    let properties = crate::parts::chp::CharacterProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.is_bold, Some(true));
    assert_eq!(properties.is_italic, Some(true));
    let previous = properties.preserved_properties_for_revision.unwrap();
    assert_eq!(previous.is_bold, Some(true));
    assert_eq!(previous.is_italic, None);

    let mut writer = Writer::new();
    writer
        .add_paragraph_runs(
            vec![("Tracked".to_string(), formatting)],
            ParagraphFormatting::default(),
        )
        .unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    let runs = paragraphs[0].runs().unwrap();
    let properties = runs[0].properties();
    assert_eq!(properties.is_bold, Some(true));
    assert_eq!(properties.is_italic, Some(true));
    let previous = properties
        .preserved_properties_for_revision
        .as_ref()
        .unwrap();
    assert_eq!(previous.is_bold, Some(true));
    assert_eq!(previous.is_italic, None);
}
