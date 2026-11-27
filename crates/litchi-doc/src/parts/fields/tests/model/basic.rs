use super::super::super::codec::*;
use super::super::super::model::*;

#[test]
fn field_text_keeps_instruction_and_cached_result_separate() {
    let field = Field {
        story: FieldStory::Header,
        start_cp: 2,
        separator_cp: Some(17),
        end_cp: 23,
        field_type: FieldType::IncludeText,
        end_flags: FieldEndFlags {
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 0,
        has_separator: true,
    };

    let text = FieldText::from_field(&field, |start, end| match (start, end) {
        (3, 17) => Ok(r#" INCLUDETEXT "file:///draft.doc" "#.to_string()),
        (18, 23) => Ok("cached".to_string()),
        _ => Err(corrupted("unexpected field range")),
    })
    .unwrap();

    assert_eq!(text.field, field);
    assert_eq!(text.instruction, r#" INCLUDETEXT "file:///draft.doc" "#);
    assert_eq!(text.result.as_deref(), Some("cached"));
}

#[test]
fn field_text_reports_absent_separator_as_no_cached_result() {
    let field = Field {
        story: FieldStory::Main,
        start_cp: 4,
        separator_cp: None,
        end_cp: 12,
        field_type: FieldType::MacroButton,
        end_flags: FieldEndFlags::default(),
        nesting_depth: 0,
        has_separator: false,
    };

    let text = FieldText::from_field(&field, |start, end| {
        assert_eq!((start, end), (5, 12));
        Ok(" MACROBUTTON NeverRun Label ".to_string())
    })
    .unwrap();

    assert_eq!(text.instruction, " MACROBUTTON NeverRun Label ");
    assert_eq!(text.result, None);
    let macro_button = text.macro_button().unwrap();
    assert_eq!(macro_button.macro_name(), "NeverRun");
    assert_eq!(macro_button.display_text(), "Label");
    assert_eq!(macro_button.cached_result(), None);
    assert!(!macro_button.is_dirty());
    assert!(!macro_button.is_locked());
}

#[test]
fn macro_button_field_exposes_stored_metadata_without_execution() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 44,
        field_type: FieldType::MacroButton,
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
        instruction: r#" MACROBUTTON "Never Run" "Click \"here\"\\now" "#.to_string(),
        result: Some("cached button".to_string()),
    };

    let macro_button = text.macro_button().unwrap();
    assert_eq!(macro_button.field(), &field);
    assert_eq!(macro_button.macro_name(), "Never Run");
    assert_eq!(macro_button.display_text(), r#"Click "here"\now"#);
    assert_eq!(macro_button.cached_result(), Some("cached button"));
    assert!(macro_button.is_dirty());
    assert!(macro_button.is_locked());

    let compact = FieldText {
        instruction: r#"MACROBUTTON"Never Run""Click""#.to_string(),
        ..text.clone()
    };
    let compact_button = compact.macro_button().unwrap();
    assert_eq!(compact_button.macro_name(), "Never Run");
    assert_eq!(compact_button.display_text(), "Click");

    let missing_button = FieldText {
        instruction: "MACROBUTTON NeverRun".to_string(),
        ..text.clone()
    };
    assert!(missing_button.macro_button().is_none());

    let empty_button = FieldText {
        instruction: r#"MACROBUTTON NeverRun """#.to_string(),
        ..text.clone()
    };
    assert!(empty_button.macro_button().is_none());

    let extra_argument = FieldText {
        instruction: "MACROBUTTON NeverRun Button unexpected".to_string(),
        ..text.clone()
    };
    assert!(extra_argument.macro_button().is_none());

    let invalid_escape = FieldText {
        instruction: r#"MACROBUTTON NeverRun "Click \now""#.to_string(),
        ..text.clone()
    };
    assert!(invalid_escape.macro_button().is_none());

    let wrong_keyword = FieldText {
        instruction: "DOCVARIABLE Customer".to_string(),
        ..text
    };
    assert!(wrong_keyword.macro_button().is_none());
}

#[test]
fn active_content_fields_expose_opaque_metadata_without_activation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::AddIn,
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
        instruction: " ADDIN opaque-add-in-data ".to_string(),
        result: Some("cached add-in result".to_string()),
    };

    let add_in = text.active_content_field().unwrap();
    assert_eq!(add_in.field(), &field);
    assert_eq!(add_in.instruction(), text.instruction);
    assert_eq!(add_in.kind(), ActiveContentFieldKind::AddIn);
    assert_eq!(add_in.cached_result(), Some("cached add-in result"));
    assert!(add_in.is_dirty());
    assert!(add_in.is_locked());

    let ocx = FieldText {
        field: Field {
            field_type: FieldType::Control,
            ..field.clone()
        },
        instruction: " CONTROL opaque-ocx-metadata ".to_string(),
        result: None,
    };
    let ocx = ocx.active_content_field().unwrap();
    assert_eq!(ocx.kind(), ActiveContentFieldKind::OcxControl);
    assert_eq!(ocx.cached_result(), None);

    let html = FieldText {
        field: Field {
            field_type: FieldType::HtmlControl,
            ..field.clone()
        },
        instruction: " HTMLCONTROL opaque-html-control-metadata ".to_string(),
        result: None,
    };
    let html = html.active_content_field().unwrap();
    assert_eq!(html.kind(), ActiveContentFieldKind::HtmlControl);

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.active_content_field().is_none());
}

#[test]
fn print_fields_preserve_opaque_metadata_without_sending_printer_commands() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Print,
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
        instruction: r#" PRINT "ESC&l1O" "#.to_string(),
        result: Some("cached printer result".to_string()),
    };

    let printer = text.print_field().unwrap();
    assert_eq!(printer.field(), &field);
    assert_eq!(printer.instruction(), text.instruction);
    assert_eq!(printer.printer_instructions(), r#""ESC&l1O""#);
    assert_eq!(printer.cached_result(), Some("cached printer result"));
    assert!(printer.is_dirty());
    assert!(printer.is_locked());

    let postscript = FieldText {
        instruction: r#"print \p 2 "0 0 moveto""#.to_string(),
        result: Some("cached PostScript result".to_string()),
        ..text.clone()
    };
    let postscript = postscript.print_field().unwrap();
    assert_eq!(postscript.printer_instructions(), r#"\p 2 "0 0 moveto""#);
    assert_eq!(postscript.cached_result(), Some("cached PostScript result"));

    for instruction in [
        r#"PRINTS "not a print field""#,
        r#"PRINTER "not a print field""#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.print_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("PRINT {}", "x".repeat(MAX_PRINT_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.print_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.print_field().is_none());
}

#[test]
fn embed_fields_preserve_opaque_metadata_without_loading_or_activating_objects() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::EmbeddedObject,
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
        instruction: r#" EMBED Excel.Sheet.12 \* MERGEFORMAT "#.to_string(),
        result: Some("cached worksheet object".to_string()),
    };

    let embedded = text.embed_field().unwrap();
    assert_eq!(embedded.field(), &field);
    assert_eq!(embedded.instruction(), text.instruction);
    assert_eq!(
        embedded.object_instructions(),
        r#"Excel.Sheet.12 \* MERGEFORMAT"#
    );
    assert_eq!(embedded.cached_result(), Some("cached worksheet object"));
    assert!(embedded.is_dirty());
    assert!(embedded.is_locked());

    let equation = FieldText {
        instruction: r#"embed "Equation.DSMT4" \d"#.to_string(),
        result: Some("cached equation object".to_string()),
        ..text.clone()
    };
    let equation = equation.embed_field().unwrap();
    assert_eq!(equation.object_instructions(), r#""Equation.DSMT4" \d"#);
    assert_eq!(equation.cached_result(), Some("cached equation object"));

    let bare = FieldText {
        instruction: "EMBED".to_string(),
        result: None,
        ..text.clone()
    };
    assert_eq!(bare.embed_field().unwrap().object_instructions(), "");

    for instruction in [r#"EMBEDS Excel.Sheet.12"#, r#"EMBEDDED Excel.Sheet.12"#] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.embed_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("EMBED {}", "x".repeat(MAX_EMBED_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.embed_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.embed_field().is_none());
}

#[test]
fn barcode_fields_preserve_opaque_metadata_without_decoding_or_rendering() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::BarCode,
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
        instruction: r#" BARCODE "4901234567894" EAN13 \h 1440 \* MERGEFORMAT "#.to_string(),
        result: Some("cached EAN13 barcode".to_string()),
    };

    let barcode = text.barcode_field().unwrap();
    assert_eq!(barcode.field(), &field);
    assert_eq!(barcode.instruction(), text.instruction);
    assert_eq!(
        barcode.barcode_instructions(),
        r#""4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#
    );
    assert_eq!(barcode.cached_result(), Some("cached EAN13 barcode"));
    assert!(barcode.is_dirty());
    assert!(barcode.is_locked());

    let code_39 = FieldText {
        instruction: r#"barcode "ABC-123" CODE39 \d"#.to_string(),
        result: Some("cached Code39 barcode".to_string()),
        ..text.clone()
    };
    let code_39 = code_39.barcode_field().unwrap();
    assert_eq!(code_39.barcode_instructions(), r#""ABC-123" CODE39 \d"#);
    assert_eq!(code_39.cached_result(), Some("cached Code39 barcode"));

    let bare = FieldText {
        instruction: "BARCODE".to_string(),
        result: None,
        ..text.clone()
    };
    assert_eq!(bare.barcode_field().unwrap().barcode_instructions(), "");

    for instruction in [r#"BARCODES 4901234567894"#, r#"BARCODED 4901234567894"#] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.barcode_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "BARCODE {}",
            "x".repeat(MAX_BARCODE_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.barcode_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.barcode_field().is_none());
}

#[test]
fn bidi_outline_fields_preserve_metadata_without_resolving_numbering_or_layout() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::BidiOutline,
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
        instruction: r#" BIDIOUTLINE \* MERGEFORMAT "#.to_string(),
        result: Some("cached bidi outline number".to_string()),
    };

    let outline = text.bidi_outline_field().unwrap();
    assert_eq!(outline.field(), &field);
    assert_eq!(outline.instruction(), text.instruction);
    assert_eq!(outline.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(outline.cached_result(), Some("cached bidi outline number"));
    assert!(outline.is_dirty());
    assert!(outline.is_locked());

    let bare = FieldText {
        instruction: "bidioutline".to_string(),
        result: Some("cached bare bidi outline".to_string()),
        ..text.clone()
    };
    let bare = bare.bidi_outline_field().unwrap();
    assert_eq!(bare.opaque_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare bidi outline"));

    for instruction in [r#"BIDIOUTLINES"#, r#"BIDIOUTLINED"#] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.bidi_outline_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "BIDIOUTLINE {}",
            "x".repeat(MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.bidi_outline_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.bidi_outline_field().is_none());
}

#[test]
fn shape_fields_preserve_metadata_without_linking_or_rendering_drawings() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Shape,
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
        instruction: r#" SHAPE \* MERGEFORMAT "#.to_string(),
        result: Some("cached drawing anchor".to_string()),
    };

    let shape = text.shape_field().unwrap();
    assert_eq!(shape.field(), &field);
    assert_eq!(shape.instruction(), text.instruction);
    assert_eq!(shape.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(shape.cached_result(), Some("cached drawing anchor"));
    assert!(shape.is_dirty());
    assert!(shape.is_locked());

    let bare = FieldText {
        instruction: "shape".to_string(),
        result: Some("cached bare drawing anchor".to_string()),
        ..text.clone()
    };
    let bare = bare.shape_field().unwrap();
    assert_eq!(bare.opaque_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare drawing anchor"));

    for instruction in [r#"SHAPES"#, r#"SHAPEANCHOR"#] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.shape_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!("SHAPE {}", "x".repeat(MAX_SHAPE_FIELD_INSTRUCTION_BYTES)),
        ..text.clone()
    };
    assert!(too_long.shape_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.shape_field().is_none());
}

#[test]
fn legacy_form_fields_preserve_metadata_without_filling_or_executing() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::FormText,
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
        instruction: r#" FORMTEXT \* MERGEFORMAT "#.to_string(),
        result: Some("cached text field".to_string()),
    };

    let text_field = text.legacy_form_field().unwrap();
    assert_eq!(text_field.field(), &field);
    assert_eq!(text_field.kind(), LegacyFormFieldKind::Text);
    assert_eq!(text_field.instruction(), text.instruction);
    assert_eq!(text_field.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(text_field.cached_result(), Some("cached text field"));
    assert!(text_field.is_dirty());
    assert!(text_field.is_locked());

    let checkbox = FieldText {
        field: Field {
            field_type: FieldType::FormCheckbox,
            ..field.clone()
        },
        instruction: "formcheckbox".to_string(),
        result: Some("cached checkbox".to_string()),
    };
    let checkbox = checkbox.legacy_form_field().unwrap();
    assert_eq!(checkbox.kind(), LegacyFormFieldKind::CheckBox);
    assert_eq!(checkbox.opaque_instructions(), "");
    assert_eq!(checkbox.cached_result(), Some("cached checkbox"));

    let drop_down = FieldText {
        field: Field {
            field_type: FieldType::FormDropdown,
            ..field.clone()
        },
        instruction: r#" FORMDROPDOWN \* MERGEFORMAT "#.to_string(),
        result: Some("cached drop-down selection".to_string()),
    };
    let drop_down = drop_down.legacy_form_field().unwrap();
    assert_eq!(drop_down.kind(), LegacyFormFieldKind::DropDown);
    assert_eq!(drop_down.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(
        drop_down.cached_result(),
        Some("cached drop-down selection")
    );

    for instruction in [r#"FORMTEXTUAL"#, r#"FORMCHECKBOX"#] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.legacy_form_field().is_none(), "{instruction}");
    }

    let too_long = FieldText {
        instruction: format!(
            "FORMTEXT {}",
            "x".repeat(MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES)
        ),
        ..text.clone()
    };
    assert!(too_long.legacy_form_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Equation,
            ..field
        },
        ..text
    };
    assert!(wrong_type.legacy_form_field().is_none());
}
