use super::super::super::model::*;

#[test]
fn if_fields_expose_cached_metadata_without_evaluation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::If,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let text = FieldText {
        field: field.clone(),
        instruction: r#" IF "A" = "A" "yes" "no" "#.to_string(),
        result: Some("yes".to_string()),
    };

    let if_field = text.if_field().unwrap();
    assert_eq!(if_field.field(), &field);
    assert_eq!(if_field.instruction(), text.instruction);
    assert_eq!(if_field.expression(), r#""A" = "A" "yes" "no""#);
    assert_eq!(if_field.cached_result(), Some("yes"));
    assert!(if_field.is_dirty());
    assert!(if_field.is_locked());

    let missing_expression = FieldText {
        instruction: "IF".to_string(),
        ..text.clone()
    };
    assert!(missing_expression.if_field().is_none());

    let wrong_keyword = FieldText {
        instruction: r#"IFF "A" = "A" "yes" "no""#.to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.if_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::NextIf,
            ..field
        },
        ..text
    };
    assert!(wrong_type.if_field().is_none());
}

#[test]
fn compare_fields_expose_cached_metadata_without_evaluation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::Compare,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let text = FieldText {
        field: field.clone(),
        instruction: r#" COMPARE "CustomerNumber" >= 4 "#.to_string(),
        result: Some("1".to_string()),
    };

    let compare_field = text.compare_field().unwrap();
    assert_eq!(compare_field.field(), &field);
    assert_eq!(compare_field.instruction(), text.instruction);
    assert_eq!(compare_field.comparison(), r#""CustomerNumber" >= 4"#);
    assert_eq!(compare_field.cached_result(), Some("1"));
    assert!(compare_field.is_dirty());
    assert!(compare_field.is_locked());

    let nested = FieldText {
        instruction: "compare MERGEFIELD CustomerRating <= 9".to_string(),
        ..text.clone()
    };
    let compare_field = nested.compare_field().unwrap();
    assert_eq!(compare_field.comparison(), "MERGEFIELD CustomerRating <= 9");

    let missing_comparison = FieldText {
        instruction: "COMPARE".to_string(),
        ..text.clone()
    };
    assert!(missing_comparison.compare_field().is_none());

    let wrong_keyword = FieldText {
        instruction: "COMPARES Customer = 1".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.compare_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::If,
            ..field
        },
        ..text
    };
    assert!(wrong_type.compare_field().is_none());
}

#[test]
fn prompt_fields_expose_cached_metadata_without_displaying_prompts() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::Ask,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let ask = FieldText {
        field: field.clone(),
        instruction: r#" ASK AskResponse "What is your first name?" \d "" \o "#.to_string(),
        result: Some("cached ask response".to_string()),
    };

    let prompt = ask.prompt_field().unwrap();
    assert_eq!(prompt.field(), &field);
    assert_eq!(prompt.instruction(), ask.instruction);
    assert_eq!(prompt.kind(), PromptFieldKind::Ask);
    assert_eq!(prompt.bookmark(), Some("AskResponse"));
    assert_eq!(prompt.prompt(), Some("What is your first name?"));
    assert_eq!(prompt.default_response(), Some(""));
    assert!(prompt.prompts_once_per_mail_merge());
    assert_eq!(prompt.cached_result(), Some("cached ask response"));
    assert!(prompt.is_dirty());
    assert!(prompt.is_locked());

    let fill_in = FieldText {
        field: Field {
            field_type: FieldType::FillIn,
            ..field.clone()
        },
        instruction: r#"fillin "Enter appointment time" \d "09:00""#.to_string(),
        result: Some("10:30".to_string()),
    };
    let fill_in_prompt = fill_in.prompt_field().unwrap();
    assert_eq!(fill_in_prompt.kind(), PromptFieldKind::FillIn);
    assert_eq!(fill_in_prompt.bookmark(), None);
    assert_eq!(fill_in_prompt.prompt(), Some("Enter appointment time"));
    assert_eq!(fill_in_prompt.default_response(), Some("09:00"));
    assert!(!fill_in_prompt.prompts_once_per_mail_merge());
    assert_eq!(fill_in_prompt.cached_result(), Some("10:30"));

    let default_only = FieldText {
        instruction: r#"FILLIN \d "recent response" \o"#.to_string(),
        result: None,
        ..fill_in.clone()
    };
    let default_only_prompt = default_only.prompt_field().unwrap();
    assert_eq!(default_only_prompt.prompt(), None);
    assert_eq!(
        default_only_prompt.default_response(),
        Some("recent response")
    );
    assert!(default_only_prompt.prompts_once_per_mail_merge());

    for instruction in [
        "ASK",
        r#"ASK "" "Question""#,
        "ASK Answer",
        r#"ASK Answer "Question" \d"#,
        r#"ASK Answer "Question" \o extra"#,
        r#"FILLIN "Question" \x"#,
        r#"FILLIN "Question" \d "first" \d "second""#,
        r#"FILLIN "Question" \o \o"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..ask.clone()
        };
        assert!(malformed.prompt_field().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        instruction: r#"ASKER Answer "Question""#.to_string(),
        ..ask.clone()
    };
    assert!(wrong_keyword.prompt_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::FillIn,
            ..field
        },
        ..ask
    };
    assert!(wrong_type.prompt_field().is_none());
}
