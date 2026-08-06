use super::super::super::codec::*;
use super::super::super::model::*;

#[test]
fn go_to_button_field_exposes_stored_metadata_without_navigation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::GoToButton,
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
        instruction: r#" GOTOBUTTON "f 2" "Footnote" "#.to_string(),
        result: Some("cached footnote button".to_string()),
    };

    let button = text.go_to_button().unwrap();
    assert_eq!(button.field(), &field);
    assert_eq!(button.target(), "f 2");
    assert_eq!(button.button_text(), "Footnote");
    assert_eq!(button.cached_result(), Some("cached footnote button"));
    assert!(button.is_dirty());
    assert!(button.is_locked());

    for instruction in [
        "GOTOBUTTON",
        r#"GOTOBUTTON "" Button"#,
        "GOTOBUTTON Destination",
        r#"GOTOBUTTON Destination """#,
        "GOTOBUTTON Destination Button unexpected",
        r#"GOTOBUTTON Destination Button \* MERGEFORMAT"#,
        r#"GOTOBUTTON Destination "Button \now""#,
    ] {
        let malformed = FieldText {
            field: field.clone(),
            instruction: instruction.to_string(),
            result: None,
        };
        assert!(malformed.go_to_button().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "GOTOBUTTONS Destination Button".to_string(),
        result: None,
    };
    assert!(wrong_keyword.go_to_button().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MacroButton,
            ..field
        },
        instruction: "GOTOBUTTON Destination Button".to_string(),
        result: None,
    };
    assert!(wrong_type.go_to_button().is_none());
}

#[test]
fn document_variable_fields_expose_cached_metadata_without_resolution() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::DocumentVariable,
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
        instruction: r#" DOCVARIABLE "Customer Region" \* MERGEFORMAT "#.to_string(),
        result: Some("cached region".to_string()),
    };

    let variable = text.document_variable().unwrap();
    assert_eq!(variable.field(), &field);
    assert_eq!(variable.instruction(), text.instruction);
    assert_eq!(variable.variable_name(), "Customer Region");
    assert_eq!(variable.cached_result(), Some("cached region"));
    assert!(variable.is_dirty());
    assert!(variable.is_locked());
    assert_eq!(variable.unknown_switches().len(), 1);
    assert_eq!(variable.unknown_switches()[0].name(), '*');
    assert_eq!(
        variable.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );

    let compact = FieldText {
        instruction: r#"DOCVARIABLE"Customer Name"\*MERGEFORMAT"#.to_string(),
        ..text.clone()
    };
    let compact_variable = compact.document_variable().unwrap();
    assert_eq!(compact_variable.variable_name(), "Customer Name");
    assert_eq!(
        compact_variable.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );

    let missing_name = FieldText {
        instruction: r#"DOCVARIABLE \* MERGEFORMAT"#.to_string(),
        ..text.clone()
    };
    assert!(missing_name.document_variable().is_none());

    let unexpected_operand = FieldText {
        instruction: "DOCVARIABLE Customer unexpected".to_string(),
        ..text.clone()
    };
    assert!(unexpected_operand.document_variable().is_none());

    let wrong_keyword = FieldText {
        instruction: "DOCVARIABLES Customer".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.document_variable().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.document_variable().is_none());
}

#[test]
fn document_property_fields_expose_cached_metadata_without_resolution() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::DocumentProperty,
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
        instruction: r#" DOCPROPERTY "Project Name" \* MERGEFORMAT \@ "MMMM d, yyyy" "#.to_string(),
        result: Some("cached project".to_string()),
    };

    let property = text.document_property().unwrap();
    assert_eq!(property.field(), &field);
    assert_eq!(property.instruction(), text.instruction);
    assert_eq!(property.property_name(), "Project Name");
    assert_eq!(property.cached_result(), Some("cached project"));
    assert!(property.is_dirty());
    assert!(property.is_locked());
    assert_eq!(property.switches().len(), 2);
    assert_eq!(property.switches()[0].name(), '*');
    assert_eq!(property.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(property.switches()[1].name(), '@');
    assert_eq!(property.switches()[1].argument(), Some("MMMM d, yyyy"));

    let compact = FieldText {
        instruction: r#"DOCPROPERTY"Project Name"\*MERGEFORMAT"#.to_string(),
        ..text.clone()
    };
    let compact_property = compact.document_property().unwrap();
    assert_eq!(compact_property.property_name(), "Project Name");
    assert_eq!(
        compact_property.switches()[0].argument(),
        Some("MERGEFORMAT")
    );

    for instruction in [
        r#"DOCPROPERTY \* MERGEFORMAT"#,
        r#"DOCPROPERTY """#,
        "DOCPROPERTY Project unexpected",
        r#"DOCPROPERTY Project \"#,
        "DOCPROPERTYS Project",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.document_property().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "DOCPROPERTY {}",
            "x".repeat(MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.document_property().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::DocumentVariable,
            ..field
        },
        ..text
    };
    assert!(wrong_type.document_property().is_none());
}

#[test]
fn info_fields_expose_stored_metadata_without_property_resolution_or_updates() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Info,
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
        instruction: r#" INFO TITLE "Stored title override" \* MERGEFORMAT \@ "opaque format" "#
            .to_string(),
        result: Some("cached title".to_string()),
    };

    let information = text.info_field().unwrap();
    assert_eq!(information.field(), &field);
    assert_eq!(information.instruction(), text.instruction);
    assert_eq!(information.information_type(), "TITLE");
    assert_eq!(information.new_value(), Some("Stored title override"));
    assert_eq!(information.cached_result(), Some("cached title"));
    assert!(information.is_dirty());
    assert!(information.is_locked());
    assert_eq!(information.switches().len(), 2);
    assert_eq!(information.switches()[0].name(), '*');
    assert_eq!(information.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(information.switches()[1].name(), '@');
    assert_eq!(information.switches()[1].argument(), Some("opaque format"));

    let implicit = FieldText {
        instruction: r#" COMMENTS "Stored comment" \* MERGEFORMAT "#.to_string(),
        result: Some("cached comment".to_string()),
        ..text.clone()
    };
    let implicit_information = implicit.info_field().unwrap();
    assert_eq!(implicit_information.information_type(), "COMMENTS");
    assert_eq!(implicit_information.new_value(), Some("Stored comment"));
    assert_eq!(implicit_information.cached_result(), Some("cached comment"));
    assert_eq!(
        implicit_information.switches()[0].argument(),
        Some("MERGEFORMAT")
    );

    let no_replacement = FieldText {
        instruction: "TEMPLATE".to_string(),
        ..text.clone()
    };
    assert_eq!(no_replacement.info_field().unwrap().new_value(), None);

    for instruction in [
        "INFO",
        r#"INFO "" "#,
        r#"INFO TITLE "Stored title" unexpected"#,
        r#"INFO TITLE "unterminated"#,
        r#"INFO TITLE \"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.info_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("INFO {}", "x".repeat(MAX_INFO_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.info_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Title,
            ..field
        },
        ..text
    };
    assert!(wrong_type.info_field().is_none());
}

#[test]
fn document_information_fields_expose_cached_metadata_without_resolution_or_calculation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Title,
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
        instruction: r#" TITLE \* MERGEFORMAT \@ "opaque format" "#.to_string(),
        result: Some("cached title".to_string()),
    };

    let information = text.document_information().unwrap();
    assert_eq!(information.field(), &field);
    assert_eq!(information.instruction(), text.instruction);
    assert_eq!(information.kind(), DocumentInformationFieldKind::Title);
    assert_eq!(information.cached_result(), Some("cached title"));
    assert!(information.is_dirty());
    assert!(information.is_locked());
    assert_eq!(information.switches().len(), 2);
    assert_eq!(information.switches()[0].name(), '*');
    assert_eq!(information.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(information.switches()[1].name(), '@');
    assert_eq!(information.switches()[1].argument(), Some("opaque format"));

    for (field_type, instruction, kind) in [
        (
            FieldType::Title,
            "TITLE",
            DocumentInformationFieldKind::Title,
        ),
        (
            FieldType::Subject,
            "SUBJECT",
            DocumentInformationFieldKind::Subject,
        ),
        (
            FieldType::Author,
            "AUTHOR",
            DocumentInformationFieldKind::Author,
        ),
        (
            FieldType::Keywords,
            "KEYWORDS",
            DocumentInformationFieldKind::Keywords,
        ),
        (
            FieldType::Comments,
            "COMMENTS",
            DocumentInformationFieldKind::Comments,
        ),
        (
            FieldType::LastSavedBy,
            "LASTSAVEDBY",
            DocumentInformationFieldKind::LastSavedBy,
        ),
        (
            FieldType::CreateDate,
            "CREATEDATE",
            DocumentInformationFieldKind::CreateDate,
        ),
        (
            FieldType::SaveDate,
            "SAVEDATE",
            DocumentInformationFieldKind::SaveDate,
        ),
        (
            FieldType::PrintDate,
            "PRINTDATE",
            DocumentInformationFieldKind::PrintDate,
        ),
        (
            FieldType::RevisionNumber,
            "REVNUM",
            DocumentInformationFieldKind::RevisionNumber,
        ),
        (
            FieldType::EditTime,
            "EDITTIME",
            DocumentInformationFieldKind::EditTime,
        ),
        (
            FieldType::NumberOfPages,
            "NUMPAGES",
            DocumentInformationFieldKind::NumberOfPages,
        ),
        (
            FieldType::NumberOfWords,
            "NUMWORDS",
            DocumentInformationFieldKind::NumberOfWords,
        ),
        (
            FieldType::NumberOfCharacters,
            "NUMCHARS",
            DocumentInformationFieldKind::NumberOfCharacters,
        ),
    ] {
        let text = FieldText {
            field: Field {
                field_type,
                ..field.clone()
            },
            instruction: format!("{instruction} \\* MERGEFORMAT"),
            ..text.clone()
        };
        let information = text.document_information().unwrap();
        assert_eq!(information.kind(), kind);
        assert_eq!(information.kind().field_keyword(), instruction);
        assert_eq!(information.switches()[0].name(), '*');
        assert_eq!(information.switches()[0].argument(), Some("MERGEFORMAT"));
    }

    for (field_type, instruction) in [
        (FieldType::Title, "TITLE unexpected"),
        (FieldType::Author, r#"AUTHOR "unterminated"#),
        (FieldType::Comments, "COMMENTS \\"),
        (
            FieldType::LastSavedBy,
            r"LASTSAVEDBY \* MERGEFORMAT unexpected",
        ),
        (FieldType::NumberOfWords, "NUMWORDS unexpected"),
        (FieldType::Author, "AUTHORS"),
    ] {
        let malformed = FieldText {
            field: Field {
                field_type,
                ..field.clone()
            },
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.document_information().is_none(), "{instruction}");
    }

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Author,
            ..field.clone()
        },
        ..text.clone()
    };
    assert!(wrong_type.document_information().is_none());

    let too_long = FieldText {
        instruction: format!(
            "TITLE \\* {}",
            "x".repeat(MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES)
        ),
        ..text
    };
    assert!(too_long.document_information().is_none());
}

#[test]
fn document_context_fields_expose_cached_metadata_without_reading_or_calculating_values() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::FileName,
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
        instruction: r" FILENAME \p ".to_string(),
        result: Some("cached file name".to_string()),
    };

    let context = text.document_context().unwrap();
    assert_eq!(context.field(), &field);
    assert_eq!(context.instruction(), text.instruction);
    assert_eq!(context.kind(), DocumentContextFieldKind::FileName);
    assert_eq!(context.cached_result(), Some("cached file name"));
    assert!(context.is_dirty());
    assert!(context.is_locked());
    assert_eq!(context.switches().len(), 1);
    assert_eq!(context.switches()[0].name(), 'p');
    assert_eq!(context.switches()[0].argument(), None);

    for (field_type, instruction, kind) in [
        (
            FieldType::FileName,
            "FILENAME",
            DocumentContextFieldKind::FileName,
        ),
        (
            FieldType::Template,
            "TEMPLATE",
            DocumentContextFieldKind::Template,
        ),
        (FieldType::Date, "DATE", DocumentContextFieldKind::Date),
        (FieldType::Time, "TIME", DocumentContextFieldKind::Time),
        (FieldType::Page, "PAGE", DocumentContextFieldKind::Page),
        (
            FieldType::FileSize,
            "FILESIZE",
            DocumentContextFieldKind::FileSize,
        ),
        (
            FieldType::Section,
            "SECTION",
            DocumentContextFieldKind::Section,
        ),
        (
            FieldType::SectionPages,
            "SECTIONPAGES",
            DocumentContextFieldKind::SectionPages,
        ),
    ] {
        let text = FieldText {
            field: Field {
                field_type,
                ..field.clone()
            },
            instruction: format!("{instruction} \\* MERGEFORMAT"),
            ..text.clone()
        };
        let context = text.document_context().unwrap();
        assert_eq!(context.kind(), kind);
        assert_eq!(context.kind().field_keyword(), instruction);
        assert_eq!(context.switches()[0].name(), '*');
        assert_eq!(context.switches()[0].argument(), Some("MERGEFORMAT"));
    }

    for (field_type, instruction) in [
        (FieldType::FileName, "FILENAME unexpected"),
        (FieldType::Template, r#"TEMPLATE "unterminated"#),
        (FieldType::FileName, "FILENAME \\"),
        (FieldType::FileName, "FILENAMES"),
        (FieldType::Page, "PAGE unexpected"),
        (FieldType::Page, "PAGES"),
        (FieldType::SectionPages, "SECTIONPAGES unexpected"),
        (FieldType::Section, "SECTIONPAGE"),
    ] {
        let malformed = FieldText {
            field: Field {
                field_type,
                ..field.clone()
            },
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.document_context().is_none(), "{instruction}");
    }

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Template,
            ..field.clone()
        },
        ..text.clone()
    };
    assert!(wrong_type.document_context().is_none());

    let too_long = FieldText {
        instruction: format!(
            "FILENAME \\* {}",
            "x".repeat(MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES)
        ),
        ..text
    };
    assert!(too_long.document_context().is_none());
}

#[test]
fn user_identity_fields_expose_metadata_without_reading_host_identity() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::UserAddress,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let address = FieldText {
        field: field.clone(),
        instruction: r#" USERADDRESS "10 Top Secret Lane" \* Upper "#.to_string(),
        result: Some("10 TOP SECRET LANE".to_string()),
    };

    let address_field = address.user_identity_field().unwrap();
    assert_eq!(address_field.field(), &field);
    assert_eq!(address_field.instruction(), address.instruction);
    assert_eq!(address_field.kind(), UserIdentityFieldKind::Address);
    assert_eq!(address_field.override_value(), Some("10 Top Secret Lane"));
    assert_eq!(
        address_field.formatting(),
        Some(UserIdentityFormatting::Upper)
    );
    assert_eq!(address_field.cached_result(), Some("10 TOP SECRET LANE"));
    assert!(address_field.is_dirty());
    assert!(address_field.is_locked());

    let initials = FieldText {
        field: Field {
            field_type: FieldType::UserInitials,
            ..field.clone()
        },
        instruction: r#"userinitials \* Lower"#.to_string(),
        result: Some("dw".to_string()),
    };
    let initials_field = initials.user_identity_field().unwrap();
    assert_eq!(initials_field.kind(), UserIdentityFieldKind::Initials);
    assert_eq!(initials_field.override_value(), None);
    assert_eq!(
        initials_field.formatting(),
        Some(UserIdentityFormatting::Lower)
    );
    assert_eq!(initials_field.cached_result(), Some("dw"));

    let name = FieldText {
        field: Field {
            field_type: FieldType::UserName,
            ..field.clone()
        },
        instruction: r#"USERNAME "Ada Lovelace" \* FirstCap"#.to_string(),
        result: Some("Ada Lovelace".to_string()),
    };
    let name_field = name.user_identity_field().unwrap();
    assert_eq!(name_field.kind(), UserIdentityFieldKind::Name);
    assert_eq!(name_field.override_value(), Some("Ada Lovelace"));
    assert_eq!(
        name_field.formatting(),
        Some(UserIdentityFormatting::FirstCap)
    );
    assert_eq!(name_field.cached_result(), Some("Ada Lovelace"));

    let blank_override = FieldText {
        field: Field {
            field_type: FieldType::UserName,
            ..field.clone()
        },
        instruction: r#"USERNAME "" \* Caps"#.to_string(),
        result: None,
    };
    let blank_override = blank_override.user_identity_field().unwrap();
    assert_eq!(blank_override.override_value(), Some(""));
    assert_eq!(
        blank_override.formatting(),
        Some(UserIdentityFormatting::Caps)
    );

    for instruction in [
        r#"USERADDRESS \*"#,
        r#"USERADDRESS \* Title"#,
        r#"USERADDRESS Ada \* Upper \* Lower"#,
        r#"USERADDRESS Ada \l 1033"#,
        "USERADDRESS Ada Lovelace",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            result: None,
            ..address.clone()
        };
        assert!(malformed.user_identity_field().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        instruction: "USERADDRESSES Ada".to_string(),
        result: None,
        ..address.clone()
    };
    assert!(wrong_keyword.user_identity_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::UserInitials,
            ..field
        },
        result: None,
        ..address
    };
    assert!(wrong_type.user_identity_field().is_none());
}

#[test]
fn advance_fields_expose_metadata_without_changing_layout() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::Advance,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let advance = FieldText {
        field: field.clone(),
        instruction: r#" ADVANCE \u 6 \d 12 \l 20 \r -4 \x 150 \y "72" \d -3 "#.to_string(),
        result: Some("cached placement".to_string()),
    };

    let advance_field = advance.advance_field().unwrap();
    assert_eq!(advance_field.field(), &field);
    assert_eq!(advance_field.instruction(), advance.instruction);
    let adjustments = advance_field
        .adjustments()
        .iter()
        .map(|adjustment| (adjustment.operation(), adjustment.points()))
        .collect::<Vec<_>>();
    assert_eq!(
        adjustments,
        vec![
            (AdvanceFieldOperation::Up, 6),
            (AdvanceFieldOperation::Down, 12),
            (AdvanceFieldOperation::Left, 20),
            (AdvanceFieldOperation::Right, -4),
            (AdvanceFieldOperation::HorizontalPosition, 150),
            (AdvanceFieldOperation::VerticalPosition, 72),
            (AdvanceFieldOperation::Down, -3),
        ]
    );
    assert_eq!(advance_field.cached_result(), Some("cached placement"));
    assert!(advance_field.is_dirty());
    assert!(advance_field.is_locked());

    let no_adjustments = FieldText {
        instruction: "aDvAnCe".to_string(),
        result: None,
        ..advance.clone()
    };
    assert!(
        no_adjustments
            .advance_field()
            .unwrap()
            .adjustments()
            .is_empty()
    );

    for instruction in [
        r#"ADVANCE \d"#,
        r#"ADVANCE \z 10"#,
        r#"ADVANCE \x 1.5"#,
        r#"ADVANCE \u 9223372036854775808"#,
        "ADVANCE 12",
        r#"ADVANCE \d 6 trailing"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            result: None,
            ..advance.clone()
        };
        assert!(malformed.advance_field().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        instruction: r#"ADVANCER \u 6"#.to_string(),
        result: None,
        ..advance.clone()
    };
    assert!(wrong_keyword.advance_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::UserAddress,
            ..field
        },
        result: None,
        ..advance
    };
    assert!(wrong_type.advance_field().is_none());
}
