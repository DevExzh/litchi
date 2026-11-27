use super::super::model::*;
use super::super::package::*;

#[test]
fn field_text_rejects_unrepresentable_or_reversed_ranges() {
    let overflow = Field {
        story: FieldStory::Main,
        start_cp: u32::MAX,
        separator_cp: None,
        end_cp: u32::MAX,
        field_type: FieldType::If,
        end_flags: FieldEndFlags::default(),
        nesting_depth: 0,
        has_separator: false,
    };
    assert!(FieldText::from_field(&overflow, |_, _| Ok(String::new())).is_err());

    let reversed = Field {
        start_cp: 8,
        end_cp: 8,
        ..overflow
    };
    assert!(FieldText::from_field(&reversed, |_, _| Ok(String::new())).is_err());
}

#[test]
fn fields_table_extracts_text_from_each_field_story() {
    let main = Field {
        story: FieldStory::Main,
        start_cp: 0,
        separator_cp: None,
        end_cp: 4,
        field_type: FieldType::Date,
        end_flags: FieldEndFlags::default(),
        nesting_depth: 0,
        has_separator: false,
    };
    let header = Field {
        story: FieldStory::Header,
        start_cp: 2,
        separator_cp: Some(7),
        end_cp: 9,
        field_type: FieldType::IncludeText,
        end_flags: FieldEndFlags {
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 0,
        has_separator: true,
    };
    let table = FieldsTable {
        stories: vec![
            FieldStoryTable {
                story: FieldStory::Main,
                markers: Vec::new(),
                terminal_cp: 4,
                fields: vec![main],
            },
            FieldStoryTable {
                story: FieldStory::Header,
                markers: Vec::new(),
                terminal_cp: 9,
                fields: vec![header],
            },
        ],
    };

    let text = table
        .field_texts(|story, start, end| {
            Ok(match (story, start, end) {
                (FieldStory::Main, 1, 4) => " DATE ".to_string(),
                (FieldStory::Header, 3, 7) => r#" INCLUDETEXT "draft.doc" "#.to_string(),
                (FieldStory::Header, 8, 9) => "cached".to_string(),
                _ => return Err(corrupted("unexpected field story range")),
            })
        })
        .unwrap();

    assert_eq!(text.len(), 2);
    assert_eq!(text[0].instruction, " DATE ");
    assert_eq!(text[1].instruction, r#" INCLUDETEXT "draft.doc" "#);
    assert_eq!(text[1].result.as_deref(), Some("cached"));
}
