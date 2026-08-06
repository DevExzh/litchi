use super::super::super::codec::*;
use super::super::super::model::*;

#[test]
fn reference_fields_preserve_metadata_without_resolution_or_navigation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Ref,
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
        instruction:
            r#" REF "Target Bookmark" \d "-" \f \h \n \p \r \t \w \* MERGEFORMAT \q opaque "#
                .to_string(),
        result: Some("cached reference".to_string()),
    };

    let reference = text.reference_field().unwrap();
    assert_eq!(reference.field(), &field);
    assert_eq!(reference.instruction(), text.instruction);
    assert_eq!(reference.kind(), ReferenceFieldKind::Reference);
    assert_eq!(reference.bookmark(), "Target Bookmark");
    assert_eq!(
        reference.options(),
        &[
            ReferenceFieldOption::SequencePageSeparator("-".to_string()),
            ReferenceFieldOption::ReferencedNoteContent,
            ReferenceFieldOption::Hyperlink,
            ReferenceFieldOption::ParagraphNumberWithoutContext,
            ReferenceFieldOption::RelativePosition,
            ReferenceFieldOption::ParagraphNumberRelativeContext,
            ReferenceFieldOption::SuppressNonNumberText,
            ReferenceFieldOption::ParagraphNumberFullContext,
        ]
    );
    assert_eq!(
        reference.unknown_switches(),
        &[
            MergeFieldSwitch {
                name: '*',
                argument: Some("MERGEFORMAT".to_string()),
            },
            MergeFieldSwitch {
                name: 'q',
                argument: Some("opaque".to_string()),
            },
        ]
    );
    assert_eq!(reference.cached_result(), Some("cached reference"));
    assert!(reference.is_dirty());
    assert!(reference.is_locked());

    let page_reference = FieldText {
        field: Field {
            field_type: FieldType::PageRef,
            ..field.clone()
        },
        instruction: r"PAGEREF PageTarget \h \p".to_string(),
        result: None,
    };
    let page_reference = page_reference.reference_field().unwrap();
    assert_eq!(page_reference.kind(), ReferenceFieldKind::PageReference);
    assert_eq!(page_reference.bookmark(), "PageTarget");
    assert_eq!(
        page_reference.options(),
        &[
            ReferenceFieldOption::Hyperlink,
            ReferenceFieldOption::RelativePosition,
        ]
    );

    let footnote_reference = FieldText {
        field: Field {
            field_type: FieldType::FootnoteRef,
            ..field.clone()
        },
        instruction: r"NOTEREF FootnoteTarget \p \f".to_string(),
        result: None,
    };
    let footnote_reference = footnote_reference.reference_field().unwrap();
    assert_eq!(
        footnote_reference.kind(),
        ReferenceFieldKind::FootnoteReference
    );
    assert_eq!(footnote_reference.bookmark(), "FootnoteTarget");
    assert_eq!(
        footnote_reference.options(),
        &[
            ReferenceFieldOption::RelativePosition,
            ReferenceFieldOption::NoteMarkFormatting,
        ]
    );

    let note_reference = FieldText {
        field: Field {
            field_type: FieldType::NoteRef,
            ..field.clone()
        },
        instruction: r"FTNREF EndnoteTarget \p".to_string(),
        result: None,
    };
    let note_reference = note_reference.reference_field().unwrap();
    assert_eq!(note_reference.kind(), ReferenceFieldKind::NoteReference);
    assert_eq!(note_reference.bookmark(), "EndnoteTarget");

    let reference_without_keyword = FieldText {
        field: Field {
            field_type: FieldType::RefWithoutKeyword,
            ..field.clone()
        },
        instruction: r#""Bare Bookmark" \h"#.to_string(),
        result: None,
    };
    let reference_without_keyword = reference_without_keyword.reference_field().unwrap();
    assert_eq!(
        reference_without_keyword.kind(),
        ReferenceFieldKind::ReferenceWithoutKeyword
    );
    assert_eq!(reference_without_keyword.bookmark(), "Bare Bookmark");
    assert_eq!(
        reference_without_keyword.options(),
        &[ReferenceFieldOption::Hyperlink]
    );

    for instruction in [
        "REF",
        r"REF \h",
        r"REF Bookmark \d",
        "REF Bookmark unexpected",
        r"REF Bookmark \h unexpected",
        r"REF Bookmark \n unexpected",
        "REF Bookmark \\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.reference_field().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "PAGEREF Bookmark".to_string(),
        result: None,
    };
    assert!(wrong_keyword.reference_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.reference_field().is_none());
}

#[test]
fn set_fields_preserve_target_and_expression_without_evaluation_or_state_changes() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Set,
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
        instruction: r#" SET RecipientName "North America" \* MERGEFORMAT"#.to_string(),
        result: Some("cached set result".to_string()),
    };

    let set = text.set_field().unwrap();
    assert_eq!(set.field(), &field);
    assert_eq!(set.instruction(), text.instruction);
    assert_eq!(set.target_name(), "RecipientName");
    assert_eq!(set.expression(), r#""North America" \* MERGEFORMAT"#);
    assert_eq!(set.cached_result(), Some("cached set result"));
    assert!(set.is_dirty());
    assert!(set.is_locked());

    let formula = FieldText {
        instruction: "SET Total =SUM(ABOVE) + 1".to_string(),
        result: None,
        ..text.clone()
    };
    let formula = formula.set_field().unwrap();
    assert_eq!(formula.target_name(), "Total");
    assert_eq!(formula.expression(), "=SUM(ABOVE) + 1");

    for instruction in ["SET", "SET \"\" value", "SET Target", "SET Target   "] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.set_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("SET Target {}", "x".repeat(MAX_SET_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.set_field().is_none());

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "SETX Target value".to_string(),
        result: None,
    };
    assert!(wrong_keyword.set_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.set_field().is_none());
}

#[test]
fn formula_fields_preserve_stored_formulas_without_evaluation_or_cell_reads() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Formula,
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
        instruction: r##" =SUM(ABOVE) \# "#,##0.00""##.to_string(),
        result: Some("125.00".to_string()),
    };

    let formula = text.formula_field().unwrap();
    assert_eq!(formula.field(), &field);
    assert_eq!(formula.instruction(), text.instruction);
    assert_eq!(formula.formula(), Some(r##"SUM(ABOVE) \# "#,##0.00""##));
    assert_eq!(formula.cached_result(), Some("125.00"));
    assert!(formula.is_dirty());
    assert!(formula.is_locked());

    let implicit = FieldText {
        instruction: "=".to_string(),
        result: None,
        ..text.clone()
    };
    assert_eq!(implicit.formula_field().unwrap().formula(), None);

    for instruction in ["", "SUM(ABOVE)", "FORMULA SUM(ABOVE)"] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.formula_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("={}", "x".repeat(MAX_FORMULA_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.formula_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.formula_field().is_none());
}

#[test]
fn equation_fields_preserve_opaque_expressions_without_calculation_or_rendering() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Equation,
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
        instruction: r#" eQ \\o\\ac(\\fs24 Q,\\fs16 R)"#.to_string(),
        result: Some("cached equation".to_string()),
    };

    let equation = text.equation_field().unwrap();
    assert_eq!(equation.field(), &field);
    assert_eq!(equation.instruction(), text.instruction);
    assert_eq!(equation.expression(), r#"\\o\\ac(\\fs24 Q,\\fs16 R)"#);
    assert_eq!(equation.cached_result(), Some("cached equation"));
    assert!(equation.is_dirty());
    assert!(equation.is_locked());

    let empty = FieldText {
        instruction: "EQ".to_string(),
        result: None,
        ..text.clone()
    };
    assert_eq!(empty.equation_field().unwrap().expression(), "");

    for instruction in ["", "EQUAL 1 + 1", "FORMULA 1 + 1"] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.equation_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("EQ {}", "x".repeat(MAX_EQUATION_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.equation_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Formula,
            ..field
        },
        ..text
    };
    assert!(wrong_type.equation_field().is_none());
}

#[test]
fn hyperlink_fields_preserve_stored_metadata_without_resolving_or_opening_targets() {
    let field = Field {
        story: FieldStory::Header,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Hyperlink,
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
        instruction: r#" hYpErLiNk "https://example.test/manual" \L "Target" \o "Tip" \T "_blank" \m \n \z "future" \* MERGEFORMAT"#.to_string(),
        result: Some("cached link text".to_string()),
    };

    let hyperlink = text.hyperlink_field().unwrap();
    assert_eq!(hyperlink.field(), &field);
    assert_eq!(hyperlink.instruction(), text.instruction);
    assert_eq!(
        hyperlink.external_target(),
        Some("https://example.test/manual")
    );
    assert_eq!(hyperlink.bookmark(), Some("Target"));
    assert_eq!(hyperlink.screen_tip(), Some("Tip"));
    assert_eq!(hyperlink.target_frame(), Some("_blank"));
    assert!(hyperlink.appends_image_map_coordinates());
    assert!(hyperlink.opens_new_window());
    assert_eq!(hyperlink.unknown_switches().len(), 2);
    assert_eq!(hyperlink.unknown_switches()[0].name(), 'z');
    assert_eq!(hyperlink.unknown_switches()[0].argument(), Some("future"));
    assert_eq!(hyperlink.unknown_switches()[1].name(), '*');
    assert_eq!(
        hyperlink.unknown_switches()[1].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(hyperlink.cached_result(), Some("cached link text"));
    assert!(hyperlink.is_dirty());
    assert!(hyperlink.is_locked());

    let bookmark_only = FieldText {
        instruction: r#"HYPERLINK \l "Target""#.to_string(),
        result: None,
        ..text.clone()
    };
    let bookmark_only = bookmark_only.hyperlink_field().unwrap();
    assert_eq!(bookmark_only.external_target(), None);
    assert_eq!(bookmark_only.bookmark(), Some("Target"));

    for instruction in [
        "HYPERLINK",
        r#"HYPERLINK "" "#,
        r#"HYPERLINK \l "" "#,
        r#"HYPERLINK target \m "unexpected""#,
        r#"HYPERLINK target \n "unexpected""#,
        r#"HYPERLINK target \l "first" \l "second""#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.hyperlink_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "HYPERLINK {}",
            "x".repeat(MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.hyperlink_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.hyperlink_field().is_none());
}

#[test]
fn quote_fields_preserve_cached_text_without_inserting_or_transforming_it() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Quote,
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
        instruction: r#" QUOTE "Stored literal" \* MERGEFORMAT \# "000" "#.to_string(),
        result: Some("cached literal".to_string()),
    };

    let quote = text.quote_field().unwrap();
    assert_eq!(quote.field(), &field);
    assert_eq!(quote.instruction(), text.instruction);
    assert_eq!(quote.text(), "Stored literal");
    assert_eq!(quote.cached_result(), Some("cached literal"));
    assert!(quote.is_dirty());
    assert!(quote.is_locked());
    assert_eq!(quote.switches().len(), 2);
    assert_eq!(quote.switches()[0].name(), '*');
    assert_eq!(quote.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(quote.switches()[1].name(), '#');
    assert_eq!(quote.switches()[1].argument(), Some("000"));

    let unquoted = FieldText {
        instruction: r#"quote CompatibilityText \@ "MMMM""#.to_string(),
        result: None,
        ..text.clone()
    };
    let unquoted = unquoted.quote_field().unwrap();
    assert_eq!(unquoted.text(), "CompatibilityText");
    assert_eq!(unquoted.switches()[0].name(), '@');
    assert_eq!(unquoted.switches()[0].argument(), Some("MMMM"));

    for instruction in [
        "QUOTE",
        r#"QUOTE \* MERGEFORMAT"#,
        r#"QUOTE "literal" unexpected"#,
        r#"QUOTE "unterminated"#,
        r#"QUOTE \"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.quote_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "QUOTE \"{}\"",
            "x".repeat(MAX_QUOTE_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.quote_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Formula,
            ..field
        },
        ..text
    };
    assert!(wrong_type.quote_field().is_none());
}

#[test]
fn symbol_fields_preserve_cached_metadata_without_mapping_codes_or_inserting_glyphs() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Symbol,
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
        instruction: r#" SYMBOL 0xA9 \f "Symbol" \s 12 \u "#.to_string(),
        result: Some("cached copyright".to_string()),
    };

    let symbol = text.symbol_field().unwrap();
    assert_eq!(symbol.field(), &field);
    assert_eq!(symbol.instruction(), text.instruction);
    assert_eq!(symbol.character_argument(), "0xA9");
    assert_eq!(symbol.cached_result(), Some("cached copyright"));
    assert!(symbol.is_dirty());
    assert!(symbol.is_locked());
    assert_eq!(symbol.switches().len(), 3);
    assert_eq!(symbol.switches()[0].name(), 'f');
    assert_eq!(symbol.switches()[0].argument(), Some("Symbol"));
    assert_eq!(symbol.switches()[1].name(), 's');
    assert_eq!(symbol.switches()[1].argument(), Some("12"));
    assert_eq!(symbol.switches()[2].name(), 'u');
    assert_eq!(symbol.switches()[2].argument(), None);

    let unquoted = FieldText {
        instruction: r"symbol 163 \a \h \j".to_string(),
        result: None,
        ..text.clone()
    };
    let unquoted = unquoted.symbol_field().unwrap();
    assert_eq!(unquoted.character_argument(), "163");
    assert_eq!(unquoted.switches()[0].name(), 'a');
    assert_eq!(unquoted.switches()[1].name(), 'h');
    assert_eq!(unquoted.switches()[2].name(), 'j');

    for instruction in [
        "SYMBOL",
        r#"SYMBOL \f "Symbol""#,
        "SYMBOL 0xA9 unexpected",
        r#"SYMBOL 0xA9 \f "unterminated"#,
        r"SYMBOL \",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.symbol_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("SYMBOL {}", "x".repeat(MAX_SYMBOL_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.symbol_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Quote,
            ..field
        },
        ..text
    };
    assert!(wrong_type.symbol_field().is_none());
}

#[test]
fn automatic_number_fields_preserve_cached_metadata_without_calculating_numbers_or_layout() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::AutoNum,
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
        instruction: r#" AUTONUM \s "." \* MERGEFORMAT "#.to_string(),
        result: Some("7.".to_string()),
    };

    let automatic = text.auto_number_field().unwrap();
    assert_eq!(automatic.field(), &field);
    assert_eq!(automatic.instruction(), text.instruction);
    assert_eq!(automatic.kind(), AutoNumberFieldKind::AutoNum);
    assert_eq!(automatic.kind().field_keyword(), "AUTONUM");
    assert_eq!(automatic.cached_result(), Some("7."));
    assert!(automatic.is_dirty());
    assert!(automatic.is_locked());
    assert_eq!(automatic.switches().len(), 2);
    assert_eq!(automatic.switches()[0].name(), 's');
    assert_eq!(automatic.switches()[0].argument(), Some("."));
    assert_eq!(automatic.switches()[1].name(), '*');
    assert_eq!(automatic.switches()[1].argument(), Some("MERGEFORMAT"));

    let legal = FieldText {
        field: Field {
            field_type: FieldType::AutoNumLegal,
            ..field.clone()
        },
        instruction: r#"autonumlgl \e \s ")" "#.to_string(),
        result: Some("2.4".to_string()),
    };
    let legal = legal.auto_number_field().unwrap();
    assert_eq!(legal.kind(), AutoNumberFieldKind::AutoNumLegal);
    assert_eq!(legal.kind().field_keyword(), "AUTONUMLGL");
    assert_eq!(legal.cached_result(), Some("2.4"));
    assert_eq!(legal.switches()[0].name(), 'e');
    assert_eq!(legal.switches()[0].argument(), None);
    assert_eq!(legal.switches()[1].name(), 's');
    assert_eq!(legal.switches()[1].argument(), Some(")"));

    let outline = FieldText {
        field: Field {
            field_type: FieldType::AutoNumOutline,
            ..field.clone()
        },
        instruction: "AUTONUMOUT".to_string(),
        result: Some("III".to_string()),
    };
    let outline = outline.auto_number_field().unwrap();
    assert_eq!(outline.kind(), AutoNumberFieldKind::AutoNumOutline);
    assert_eq!(outline.kind().field_keyword(), "AUTONUMOUT");
    assert_eq!(outline.cached_result(), Some("III"));
    assert!(outline.switches().is_empty());

    for (field_type, instruction) in [
        (FieldType::AutoNum, "AUTONUM unexpected"),
        (FieldType::AutoNumLegal, r#"AUTONUMLGL \s "unterminated"#),
        (FieldType::AutoNumOutline, "AUTONUMOUT \\"),
    ] {
        let malformed = FieldText {
            field: Field {
                field_type,
                ..field.clone()
            },
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.auto_number_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "AUTONUM {}",
            "x".repeat(MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.auto_number_field().is_none());

    let mismatched_type = FieldText {
        field: Field {
            field_type: FieldType::AutoNumOutline,
            ..field.clone()
        },
        instruction: "AUTONUM".to_string(),
        ..text.clone()
    };
    assert!(mismatched_type.auto_number_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Symbol,
            ..field
        },
        ..text
    };
    assert!(wrong_type.auto_number_field().is_none());
}

#[test]
fn list_number_fields_preserve_cached_metadata_without_reading_lists_or_calculating_numbers() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::ListNumber,
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
        instruction: r#" LISTNUM NumberDefault \l 6 \s 3 \* MERGEFORMAT "#.to_string(),
        result: Some("(iii)".to_string()),
    };

    let numbered = text.list_number_field().unwrap();
    assert_eq!(numbered.field(), &field);
    assert_eq!(numbered.instruction(), text.instruction);
    assert_eq!(numbered.list_name(), Some("NumberDefault"));
    assert_eq!(numbered.cached_result(), Some("(iii)"));
    assert!(numbered.is_dirty());
    assert!(numbered.is_locked());
    assert_eq!(numbered.switches().len(), 3);
    assert_eq!(numbered.switches()[0].name(), 'l');
    assert_eq!(numbered.switches()[0].argument(), Some("6"));
    assert_eq!(numbered.switches()[1].name(), 's');
    assert_eq!(numbered.switches()[1].argument(), Some("3"));
    assert_eq!(numbered.switches()[2].name(), '*');
    assert_eq!(numbered.switches()[2].argument(), Some("MERGEFORMAT"));

    let outline = FieldText {
        instruction: r#"listnum "Outline Default" \l 4"#.to_string(),
        result: Some("c".to_string()),
        ..text.clone()
    };
    let outline = outline.list_number_field().unwrap();
    assert_eq!(outline.list_name(), Some("Outline Default"));
    assert_eq!(outline.cached_result(), Some("c"));
    assert_eq!(outline.switches()[0].name(), 'l');
    assert_eq!(outline.switches()[0].argument(), Some("4"));

    let unnamed = FieldText {
        instruction: r"LISTNUM \l 2".to_string(),
        result: Some("i".to_string()),
        ..text.clone()
    };
    let unnamed = unnamed.list_number_field().unwrap();
    assert_eq!(unnamed.list_name(), None);
    assert_eq!(unnamed.cached_result(), Some("i"));
    assert_eq!(unnamed.switches()[0].name(), 'l');
    assert_eq!(unnamed.switches()[0].argument(), Some("2"));

    for instruction in [
        "LISTNUM NumberDefault unexpected",
        r#"LISTNUM "unterminated"#,
        "LISTNUM \\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.list_number_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "LISTNUM {}",
            "x".repeat(MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.list_number_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::AutoNum,
            ..field
        },
        ..text
    };
    assert!(wrong_type.list_number_field().is_none());
}

#[test]
fn sequence_fields_preserve_metadata_without_bookmark_lookup_or_numbering() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Sequence,
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
        instruction: r#" SEQ Figure CaptionBookmark \r 1 \s 2 \* ARABIC"#.to_string(),
        result: Some("7".to_string()),
    };

    let sequence = text.sequence_field().unwrap();
    assert_eq!(sequence.field(), &field);
    assert_eq!(sequence.instruction(), text.instruction);
    assert_eq!(sequence.identifier(), "Figure");
    assert_eq!(sequence.bookmark(), Some("CaptionBookmark"));
    assert_eq!(sequence.tail(), r#"\r 1 \s 2 \* ARABIC"#);
    assert_eq!(sequence.cached_result(), Some("7"));
    assert!(sequence.is_dirty());
    assert!(sequence.is_locked());

    let no_bookmark = FieldText {
        instruction: r#"SEQ Equation \* ROMAN"#.to_string(),
        result: None,
        ..text.clone()
    };
    let no_bookmark = no_bookmark.sequence_field().unwrap();
    assert_eq!(no_bookmark.identifier(), "Equation");
    assert_eq!(no_bookmark.bookmark(), None);
    assert_eq!(no_bookmark.tail(), r#"\* ROMAN"#);

    for instruction in ["SEQ", "SEQ \"\"", "SEQ Figure \"\""] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.sequence_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "SEQ Figure {}",
            "x".repeat(MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.sequence_field().is_none());

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "SEQUENCE Figure".to_string(),
        result: None,
    };
    assert!(wrong_keyword.sequence_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.sequence_field().is_none());
}

#[test]
fn style_reference_fields_preserve_metadata_without_style_or_layout_resolution() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::StyleRef,
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
        instruction: r#" STYLEREF "Heading 1" \l \n \p \r \t \w \* MERGEFORMAT \q opaque"#
            .to_string(),
        result: Some("Cached heading".to_string()),
    };

    let style_reference = text.style_reference_field().unwrap();
    assert_eq!(style_reference.field(), &field);
    assert_eq!(style_reference.instruction(), text.instruction);
    assert_eq!(style_reference.style_name(), "Heading 1");
    assert_eq!(
        style_reference.options(),
        &[
            StyleReferenceFieldOption::FollowingText,
            StyleReferenceFieldOption::ParagraphNumber,
            StyleReferenceFieldOption::RelativePosition,
            StyleReferenceFieldOption::ParagraphNumberRelativeContext,
            StyleReferenceFieldOption::SuppressNonNumberText,
            StyleReferenceFieldOption::ParagraphNumberFullContext,
        ]
    );
    assert_eq!(
        style_reference.unknown_switches(),
        &[
            MergeFieldSwitch {
                name: '*',
                argument: Some("MERGEFORMAT".to_string()),
            },
            MergeFieldSwitch {
                name: 'q',
                argument: Some("opaque".to_string()),
            },
        ]
    );
    assert_eq!(style_reference.cached_result(), Some("Cached heading"));
    assert!(style_reference.is_dirty());
    assert!(style_reference.is_locked());

    for instruction in [
        "STYLEREF",
        "STYLEREF \"\"",
        "STYLEREF Heading \\l unexpected",
        "STYLEREF Heading \\n unexpected",
        "STYLEREF Heading unexpected",
        "STYLEREF Heading \\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.style_reference_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "STYLEREF Heading {}",
            "x".repeat(MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.style_reference_field().is_none());

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "STYLEREFS Heading".to_string(),
        result: None,
    };
    assert!(wrong_keyword.style_reference_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.style_reference_field().is_none());
}
