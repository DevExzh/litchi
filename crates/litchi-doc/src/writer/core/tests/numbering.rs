use super::support::*;

#[test]
fn list_tables_round_trip_through_fib_indices() {
    let mut writer = Writer::new();
    let mut list = ListStructure::new(42);
    let mut level = crate::writer::numbering::ListLevel::new(3, NumberFormat::Decimal);
    level.number_text = "%1.😀".to_string();
    list.add_level(level);
    writer.add_list(list);
    writer.add_list_override(ListFormatOverride::new(42, 1));
    writer.set_list_names(ListNamesTable::try_new(vec!["Outline".to_string()]).unwrap());
    let template = crate::ListTemplateCode::BuiltIn {
        format: crate::BuiltInListTemplate::ArabicPeriod,
        language: crate::ListTemplateLanguageId::new(0x0409),
    };
    writer.set_list_templates(ListTemplateTable::try_new(vec![Some([template; 9])]).unwrap());
    writer
        .add_paragraph_runs(
            vec![("List item".to_string(), CharacterFormatting::default())],
            ParagraphFormatting {
                ilvl: Some(0),
                ilfo: Some(1),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let tables = document.list_tables().unwrap();
    assert_eq!(tables.structures().len(), 1);
    assert_eq!(tables.overrides().len(), 1);
    assert_eq!(tables.structures()[0].levels[0].number_text, "%1.😀");
    assert_eq!(document.list_names().unwrap().name(0), Some("Outline"));
    assert_eq!(
        document.list_templates().unwrap().get(0).unwrap(),
        &[template; 9]
    );

    let paragraphs = document.paragraphs().unwrap();
    let info = document.paragraph_list_info(&paragraphs[0]).unwrap();
    assert_eq!(info.start_at, 3);
    assert_eq!(info.number_text, "%1.😀");
}
