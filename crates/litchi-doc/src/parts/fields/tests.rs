//! Focused coverage for field models, codecs, and package integration.

use super::codec::*;
use super::model::*;
use super::package::*;

fn plcf(cps: &[u32], descriptors: &[[u8; 2]]) -> Vec<u8> {
    assert_eq!(cps.len(), descriptors.len() + 1);
    let mut data = Vec::new();
    for cp in cps {
        data.extend_from_slice(&cp.to_le_bytes());
    }
    for descriptor in descriptors {
        data.extend_from_slice(descriptor);
    }
    data
}

#[test]
fn nested_fields_flags_and_roundtrip_are_exact() {
    let data = plcf(
        &[1, 3, 5, 7, 9, 11, 13],
        &[
            [0x13, 0x07],
            [0x13, 0x21],
            [0x14, 0xFF],
            [0x15, 0xC1],
            [0x14, 0xA5],
            [0x15, 0xB1],
        ],
    );
    let table = FieldStoryTable::parse_plcf(FieldStory::Main, 20, &data).unwrap();
    assert_eq!(table.to_plcf_bytes().unwrap(), data);
    assert_eq!(table.fields().len(), 2);
    assert_eq!(table.fields()[0].field_type, FieldType::If);
    assert_eq!(table.fields()[0].separator_cp, Some(9));
    assert!(table.fields()[0].end_flags.locked);
    assert_eq!(table.fields()[1].field_type, FieldType::Page);
    assert_eq!(table.fields()[1].nesting_depth, 1);
    assert!(table.fields()[1].end_flags.nested);
}

#[test]
fn malformed_plcf_and_fieldlist_matrix_is_rejected() {
    let valid = plcf(&[1, 3, 5, 7], &[[0x13, 0x21], [0x14, 0], [0x15, 0x80]]);
    assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 7, &valid).is_ok());
    for invalid in [
        Vec::new(),
        vec![0; 5],
        plcf(&[1, 1], &[[0x13, 0x21]]),
        plcf(&[2, 1], &[[0x13, 0x21]]),
        plcf(&[1, 9], &[[0x13, 0x21]]),
        plcf(&[1, 3], &[[0x12, 0x21]]),
        plcf(&[1, 3], &[[0x14, 0]]),
        plcf(&[1, 3], &[[0x15, 0]]),
        plcf(&[1, 3], &[[0x13, 0x21]]),
        plcf(&[1, 3, 5], &[[0x13, 0x21], [0x15, 0x80]]),
        plcf(
            &[1, 2, 3, 4, 5],
            &[[0x13, 0x21], [0x14, 0], [0x14, 0], [0x15, 0x80]],
        ),
        plcf(
            &[1, 2, 3, 4, 5],
            &[[0x13, 0x07], [0x13, 0x21], [0x15, 0], [0x15, 0]],
        ),
    ] {
        assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 7, &invalid).is_err());
    }
}

#[test]
fn all_end_flags_and_reserved_descriptor_bits_are_preserved() {
    let descriptor = FieldDescriptor::from_bytes(&[0xF5, 0xFF]).unwrap();
    assert_eq!(descriptor.reserved_bits, 7);
    let FieldMarkerValue::End(flags) = descriptor.value else {
        panic!("end descriptor");
    };
    assert!(flags.differ && flags.zombie_embed && flags.results_dirty);
    assert!(flags.results_edited && flags.locked && flags.private_result);
    assert!(flags.nested && flags.has_separator);
    assert_eq!(descriptor.to_bytes(), [0xF5, 0xFF]);
}

#[test]
fn field_type_mapping_covers_specified_and_unknown_values() {
    assert_eq!(FieldType::from(0x0E), FieldType::Info);
    assert_eq!(FieldType::Info.as_u8(), 0x0E);
    assert_eq!(FieldType::from(0x3A), FieldType::EmbeddedObject);
    assert_eq!(FieldType::EmbeddedObject.as_u8(), 0x3A);
    assert_eq!(FieldType::from(0x3F), FieldType::BarCode);
    assert_eq!(FieldType::BarCode.as_u8(), 0x3F);
    assert_eq!(FieldType::from(0x5C), FieldType::BidiOutline);
    assert_eq!(FieldType::BidiOutline.as_u8(), 0x5C);
    assert_eq!(FieldType::from(0x5F), FieldType::Shape);
    assert_eq!(FieldType::Shape.as_u8(), 0x5F);
    assert_eq!(FieldType::from(0x46), FieldType::FormText);
    assert_eq!(FieldType::FormText.as_u8(), 0x46);
    assert_eq!(FieldType::from(0x47), FieldType::FormCheckbox);
    assert_eq!(FieldType::FormCheckbox.as_u8(), 0x47);
    assert_eq!(FieldType::from(0x53), FieldType::FormDropdown);
    assert_eq!(FieldType::FormDropdown.as_u8(), 0x53);
    assert_eq!(FieldType::from(0x58), FieldType::Hyperlink);
    assert_eq!(FieldType::from(0x34), FieldType::AutoNumOutline);
    assert_eq!(FieldType::from(0x35), FieldType::AutoNumLegal);
    assert_eq!(FieldType::from(0x36), FieldType::AutoNum);
    assert_eq!(FieldType::from(0x39), FieldType::Symbol);
    assert_eq!(FieldType::from(0x30), FieldType::Print);
    assert_eq!(FieldType::Print.as_u8(), 0x30);
    assert_eq!(FieldType::from(0x5A), FieldType::ListNumber);
    assert_eq!(FieldType::from(0x41), FieldType::Section);
    assert_eq!(FieldType::from(0x42), FieldType::SectionPages);
    assert_eq!(FieldType::from(0x45), FieldType::FileSize);
    assert_eq!(FieldType::from(0x04), FieldType::Unknown(0x04));
    assert_eq!(
        FieldType::from_keyword("hyperlink"),
        Some(FieldType::Hyperlink)
    );
    assert_eq!(FieldType::from_keyword("="), Some(FieldType::Formula));
    assert_eq!(
        FieldType::from_keyword("FTNREF"),
        Some(FieldType::FootnoteRef)
    );
    assert_eq!(FieldType::from_keyword("TC"), None);
    assert_eq!(FieldType::from_keyword("future-field"), None);
}

#[test]
fn unrecognized_field_type_is_preserved_with_valid_boundaries() {
    let data = plcf(&[1, 3, 5, 7], &[[0x13, 0x4E], [0x14, 0], [0x15, 0x80]]);
    let table = FieldStoryTable::parse_plcf(FieldStory::Main, 7, &data).unwrap();

    assert_eq!(table.fields().len(), 1);
    assert_eq!(table.fields()[0].field_type, FieldType::Unknown(0x4E));
    assert_eq!(table.fields()[0].separator_cp, Some(3));
    assert_eq!(table.to_plcf_bytes().unwrap(), data);
}

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

#[test]
fn table_of_contents_fields_preserve_stored_configuration_without_generation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::TableOfContents,
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
        instruction: r#" TOC \a Figure \b "Scope Bookmark" \c Table \d "/" \f A \h \l 1-3 \n "2-3" \o "1-4" \p " — " \s Figure \t "Custom,1,Appendix,2" \u \w \x \z \* MERGEFORMAT \q opaque "#.to_string(),
        result: Some("cached contents".to_string()),
    };

    let toc = text.table_of_contents().unwrap();
    assert_eq!(toc.field(), &field);
    assert_eq!(toc.instruction(), text.instruction);
    assert_eq!(
        toc.options(),
        &[
            TableOfContentsOption::CaptionWithoutLabel("Figure".to_string()),
            TableOfContentsOption::Bookmark("Scope Bookmark".to_string()),
            TableOfContentsOption::CaptionSequence("Table".to_string()),
            TableOfContentsOption::SequencePageSeparator("/".to_string()),
            TableOfContentsOption::TableEntryIdentifier("A".to_string()),
            TableOfContentsOption::Hyperlinks,
            TableOfContentsOption::TableEntryLevels("1-3".to_string()),
            TableOfContentsOption::OmitPageNumbers(Some("2-3".to_string())),
            TableOfContentsOption::HeadingStyleRange(Some("1-4".to_string())),
            TableOfContentsOption::EntryPageNumberSeparator(" — ".to_string()),
            TableOfContentsOption::SequenceIdentifier("Figure".to_string()),
            TableOfContentsOption::StyleMappings("Custom,1,Appendix,2".to_string()),
            TableOfContentsOption::OutlineLevels,
            TableOfContentsOption::PreserveTabs,
            TableOfContentsOption::PreserveNewlines,
            TableOfContentsOption::HidePageNumbersInWebLayout,
        ]
    );
    assert_eq!(
        toc.unknown_switches(),
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
    assert_eq!(toc.cached_result(), Some("cached contents"));
    assert!(toc.is_dirty());
    assert!(toc.is_locked());

    let optional_ranges = FieldText {
        instruction: r"TOC \n \o".to_string(),
        ..text.clone()
    };
    assert_eq!(
        optional_ranges.table_of_contents().unwrap().options(),
        &[
            TableOfContentsOption::OmitPageNumbers(None),
            TableOfContentsOption::HeadingStyleRange(None),
        ]
    );

    for instruction in ["TOC \\a", r"TOC \h unexpected", "TOC unexpected", "TOC \\"] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.table_of_contents().is_none());
    }

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.table_of_contents().is_none());
}

#[test]
fn table_of_contents_entries_reconstruct_omitted_field_markers() {
    let text = concat!(
        "\u{0013} TC \"Illustration 1\" \\f i \\l 4 ",
        "\\n \\* MERGEFORMAT ",
        "\u{0014}cached entry\u{0015}",
        "\u{0013} TCC \"not an entry\"\u{0015}",
        "\u{0013} TC \"missing end\""
    );
    let stored = non_plcf_field_texts(FieldStory::Textbox, text);
    assert_eq!(stored.len(), 2);

    let entries: Vec<_> = stored
        .iter()
        .filter_map(TableOfContentsEntryField::from_non_plcf_field)
        .collect();
    assert_eq!(entries.len(), 1);
    let entry = &entries[0];
    assert_eq!(entry.story(), FieldStory::Textbox);
    assert_eq!(entry.start_position(), 0);
    assert_eq!(
        entry.instruction(),
        " TC \"Illustration 1\" \\f i \\l 4 \\n \\* MERGEFORMAT "
    );
    assert_eq!(entry.entry(), "Illustration 1");
    assert_eq!(
        entry.options(),
        &[
            TableOfContentsEntryOption::ListIdentifier("i".to_string()),
            TableOfContentsEntryOption::Level("4".to_string()),
            TableOfContentsEntryOption::OmitPageNumber,
        ]
    );
    assert_eq!(
        entry.unknown_switches(),
        &[MergeFieldSwitch {
            name: '*',
            argument: Some("MERGEFORMAT".to_string()),
        }]
    );
    assert_eq!(entry.cached_result(), Some("cached entry"));
    assert!(entry.separator_position().is_some());
    assert!(entry.end_position() > entry.start_position());

    let utf16_prefix = "\u{1F980}\u{0013} TC \"Crab\"\u{0015}";
    let prefixed = non_plcf_field_texts(FieldStory::Main, utf16_prefix);
    let prefixed = TableOfContentsEntryField::from_non_plcf_field(&prefixed[0]).unwrap();
    assert_eq!(prefixed.start_position(), 2);

    for instruction in [
        "TC",
        "TC \\f i",
        "TC entry unexpected",
        "TC entry \\n unexpected",
        "TC entry \\f",
        "TC entry \\l",
    ] {
        let text = format!("\u{0013}{instruction}\u{0015}");
        assert!(
            non_plcf_field_texts(FieldStory::Main, &text)
                .iter()
                .all(|field| TableOfContentsEntryField::from_non_plcf_field(field).is_none()),
            "{instruction}"
        );
    }

    let too_long = format!(
        "\u{0013}TC {} \u{0015}",
        "x".repeat(MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES)
    );
    assert!(
        non_plcf_field_texts(FieldStory::Main, &too_long)
            .iter()
            .all(|field| TableOfContentsEntryField::from_non_plcf_field(field).is_none())
    );
}

#[test]
fn table_of_authorities_entries_reconstruct_omitted_field_markers() {
    let text = concat!(
        "\u{0013} TA \\l \"Baldwin v. Alberti\" \\c 1 \\s Baldwin ",
        "\\b \\i \\r PageRange \\* MERGEFORMAT ",
        "\u{0014}cached authority\u{0015}",
        "\u{0013} TAA \\l \"not an entry\"\u{0015}",
        "\u{0013} TA"
    );
    let stored = non_plcf_field_texts(FieldStory::Comment, text);
    assert_eq!(stored.len(), 2);

    let entries: Vec<_> = stored
        .iter()
        .filter_map(TableOfAuthoritiesEntryField::from_non_plcf_field)
        .collect();
    assert_eq!(entries.len(), 1);
    let entry = &entries[0];
    assert_eq!(entry.story(), FieldStory::Comment);
    assert_eq!(entry.start_position(), 0);
    assert_eq!(
        entry.instruction(),
        " TA \\l \"Baldwin v. Alberti\" \\c 1 \\s Baldwin \\b \\i \\r PageRange \\* MERGEFORMAT "
    );
    assert_eq!(
        entry.options(),
        &[
            TableOfAuthoritiesEntryOption::LongCitation("Baldwin v. Alberti".to_string()),
            TableOfAuthoritiesEntryOption::Category("1".to_string()),
            TableOfAuthoritiesEntryOption::ShortCitation("Baldwin".to_string()),
            TableOfAuthoritiesEntryOption::BoldPageNumber,
            TableOfAuthoritiesEntryOption::ItalicPageNumber,
            TableOfAuthoritiesEntryOption::PageRangeBookmark("PageRange".to_string()),
        ]
    );
    assert_eq!(
        entry.unknown_switches(),
        &[MergeFieldSwitch {
            name: '*',
            argument: Some("MERGEFORMAT".to_string()),
        }]
    );
    assert_eq!(entry.cached_result(), Some("cached authority"));
    assert!(entry.separator_position().is_some());
    assert!(entry.end_position() > entry.start_position());

    let no_options = non_plcf_field_texts(FieldStory::Main, "\u{0013} TA\u{0015}");
    let no_options = TableOfAuthoritiesEntryField::from_non_plcf_field(&no_options[0]).unwrap();
    assert!(no_options.options().is_empty());

    for instruction in [
        "TA unexpected",
        "TA \\b unexpected",
        "TA \\c",
        "TA \\i unexpected",
        "TA \\l",
        "TA \\r",
        "TA \\s",
    ] {
        let text = format!("\u{0013}{instruction}\u{0015}");
        assert!(
            non_plcf_field_texts(FieldStory::Main, &text)
                .iter()
                .all(|field| TableOfAuthoritiesEntryField::from_non_plcf_field(field).is_none()),
            "{instruction}"
        );
    }

    let too_long = format!(
        "\u{0013}TA \\l {} \u{0015}",
        "x".repeat(MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_INSTRUCTION_BYTES)
    );
    assert!(
        non_plcf_field_texts(FieldStory::Main, &too_long)
            .iter()
            .all(|field| TableOfAuthoritiesEntryField::from_non_plcf_field(field).is_none())
    );
}

#[test]
fn index_entries_reconstruct_omitted_field_markers() {
    let text = concat!(
        "\u{0013} XE \"Office Open XML:Syntax\" \\b \\f Intro \\i ",
        "\\r PageRange \\t \"See syntax\" \\y Office \\* MERGEFORMAT ",
        "\u{0014}cached entry\u{0015}",
        "\u{0013} XER \"not an entry\"\u{0015}",
        "\u{0013} XE \"missing end\""
    );
    let stored = non_plcf_field_texts(FieldStory::Endnote, text);
    assert_eq!(stored.len(), 2);

    let entries: Vec<_> = stored
        .iter()
        .filter_map(IndexEntryField::from_non_plcf_field)
        .collect();
    assert_eq!(entries.len(), 1);
    let entry = &entries[0];
    assert_eq!(entry.story(), FieldStory::Endnote);
    assert_eq!(entry.start_position(), 0);
    assert_eq!(
        entry.instruction(),
        " XE \"Office Open XML:Syntax\" \\b \\f Intro \\i \\r PageRange \\t \"See syntax\" \\y Office \\* MERGEFORMAT "
    );
    assert_eq!(entry.entry(), "Office Open XML:Syntax");
    assert_eq!(
        entry.options(),
        &[
            IndexEntryOption::BoldPageNumber,
            IndexEntryOption::EntryType("Intro".to_string()),
            IndexEntryOption::ItalicPageNumber,
            IndexEntryOption::PageRangeBookmark("PageRange".to_string()),
            IndexEntryOption::CrossReference("See syntax".to_string()),
            IndexEntryOption::Yomi("Office".to_string()),
        ]
    );
    assert_eq!(
        entry.unknown_switches(),
        &[MergeFieldSwitch {
            name: '*',
            argument: Some("MERGEFORMAT".to_string()),
        }]
    );
    assert_eq!(entry.cached_result(), Some("cached entry"));
    assert!(entry.separator_position().is_some());
    assert!(entry.end_position() > entry.start_position());

    let no_options = non_plcf_field_texts(FieldStory::Main, "\u{0013} XE entry\u{0015}");
    let no_options = IndexEntryField::from_non_plcf_field(&no_options[0]).unwrap();
    assert!(no_options.options().is_empty());

    for instruction in [
        "XE",
        "XE \\f Intro",
        "XE entry unexpected",
        "XE entry \\b unexpected",
        "XE entry \\f",
        "XE entry \\i unexpected",
        "XE entry \\r",
        "XE entry \\t",
        "XE entry \\y",
    ] {
        let text = format!("\u{0013}{instruction}\u{0015}");
        assert!(
            non_plcf_field_texts(FieldStory::Main, &text)
                .iter()
                .all(|field| IndexEntryField::from_non_plcf_field(field).is_none()),
            "{instruction}"
        );
    }

    let too_long = format!(
        "\u{0013}XE {} \u{0015}",
        "x".repeat(MAX_INDEX_ENTRY_FIELD_INSTRUCTION_BYTES)
    );
    assert!(
        non_plcf_field_texts(FieldStory::Main, &too_long)
            .iter()
            .all(|field| IndexEntryField::from_non_plcf_field(field).is_none())
    );
}

#[test]
fn referenced_documents_reconstruct_omitted_field_markers() {
    let text = concat!(
        "\u{0013} RD \"chapters/Chapter 1.doc\" \\f \\* MERGEFORMAT ",
        "\u{0014}cached reference\u{0015}",
        "\u{0013} RDX \"not a reference\"\u{0015}",
        "\u{0013} RD \"missing end\""
    );
    let stored = non_plcf_field_texts(FieldStory::Header, text);
    assert_eq!(stored.len(), 2);

    let references: Vec<_> = stored
        .iter()
        .filter_map(ReferencedDocumentField::from_non_plcf_field)
        .collect();
    assert_eq!(references.len(), 1);
    let reference = &references[0];
    assert_eq!(reference.story(), FieldStory::Header);
    assert_eq!(reference.start_position(), 0);
    assert_eq!(
        reference.instruction(),
        " RD \"chapters/Chapter 1.doc\" \\f \\* MERGEFORMAT "
    );
    assert_eq!(reference.source(), "chapters/Chapter 1.doc");
    assert!(reference.uses_relative_path());
    assert_eq!(
        reference.switches(),
        &[
            MergeFieldSwitch {
                name: 'f',
                argument: None,
            },
            MergeFieldSwitch {
                name: '*',
                argument: Some("MERGEFORMAT".to_string()),
            },
        ]
    );
    assert_eq!(reference.cached_result(), Some("cached reference"));
    assert!(reference.separator_position().is_some());
    assert!(reference.end_position() > reference.start_position());

    let absolute = non_plcf_field_texts(FieldStory::Main, "\u{0013} RD \"appendix.doc\"\u{0015}");
    let absolute = ReferencedDocumentField::from_non_plcf_field(&absolute[0]).unwrap();
    assert_eq!(absolute.source(), "appendix.doc");
    assert!(!absolute.uses_relative_path());

    for instruction in [
        "RD",
        "RD \\f",
        "RD \"\"",
        "RD document unexpected",
        "RD document \\f relative",
        "RD document \\f \\f",
        r"RD document \",
    ] {
        let text = format!("\u{0013}{instruction}\u{0015}");
        assert!(
            non_plcf_field_texts(FieldStory::Main, &text)
                .iter()
                .all(|field| ReferencedDocumentField::from_non_plcf_field(field).is_none()),
            "{instruction}"
        );
    }

    let too_long = format!(
        "\u{0013}RD {} \u{0015}",
        "x".repeat(MAX_REFERENCED_DOCUMENT_FIELD_INSTRUCTION_BYTES)
    );
    assert!(
        non_plcf_field_texts(FieldStory::Main, &too_long)
            .iter()
            .all(|field| ReferencedDocumentField::from_non_plcf_field(field).is_none())
    );
}

#[test]
fn private_fields_reconstruct_omitted_field_markers() {
    let text = concat!(
        "\u{0013} PRIVATE \"converter payload\" \\* MERGEFORMAT ",
        "\u{0014}cached private payload\u{0015}",
        "\u{0013} PRIVATELY not-private\u{0015}",
        "\u{0013} PRIVATE missing end"
    );
    let stored = non_plcf_field_texts(FieldStory::Textbox, text);
    assert_eq!(stored.len(), 2);

    let private_fields: Vec<_> = stored
        .iter()
        .filter_map(PrivateField::from_non_plcf_field)
        .collect();
    assert_eq!(private_fields.len(), 1);
    let private = &private_fields[0];
    assert_eq!(private.story(), FieldStory::Textbox);
    assert_eq!(private.start_position(), 0);
    assert_eq!(
        private.instruction(),
        " PRIVATE \"converter payload\" \\* MERGEFORMAT "
    );
    assert_eq!(
        private.opaque_instructions(),
        "\"converter payload\" \\* MERGEFORMAT"
    );
    assert_eq!(private.cached_result(), Some("cached private payload"));
    assert!(private.separator_position().is_some());
    assert!(private.end_position() > private.start_position());

    let bare = non_plcf_field_texts(FieldStory::Main, "\u{0013} PRIVATE\u{0015}");
    let bare = PrivateField::from_non_plcf_field(&bare[0]).unwrap();
    assert!(bare.opaque_instructions().is_empty());

    for instruction in ["PRIVATEpayload", "PRIVATELY opaque"] {
        let text = format!("\u{0013}{instruction}\u{0015}");
        assert!(
            non_plcf_field_texts(FieldStory::Main, &text)
                .iter()
                .all(|field| PrivateField::from_non_plcf_field(field).is_none()),
            "{instruction}"
        );
    }

    let too_long = format!(
        "\u{0013}PRIVATE {} \u{0015}",
        "x".repeat(MAX_PRIVATE_FIELD_INSTRUCTION_BYTES)
    );
    assert!(
        non_plcf_field_texts(FieldStory::Main, &too_long)
            .iter()
            .all(|field| PrivateField::from_non_plcf_field(field).is_none())
    );
}

#[test]
fn non_plcf_collection_classifies_all_five_excluded_types_once() {
    let main = concat!(
        "\u{0013}TC Contents\u{0015}",
        "\u{0013}TA \\l Citation\u{0015}",
        "\u{0013}XE Entry\u{0015}",
        "\u{0013}RD appendix.doc \\f\u{0015}",
        "\u{0013}PRIVATE opaque\u{0015}",
        "\u{0013}TC missing-end",
    );
    let header = "\u{0013}UNKNOWN ignored\u{0015}";
    let fields =
        NonPlcfFields::from_story_texts([(FieldStory::Main, main), (FieldStory::Header, header)]);

    assert_eq!(fields.len(), 5);
    assert!(!fields.is_empty());
    assert_eq!(fields.table_of_contents_entries().len(), 1);
    assert_eq!(fields.table_of_authorities_entries().len(), 1);
    assert_eq!(fields.index_entries().len(), 1);
    assert_eq!(fields.referenced_documents().len(), 1);
    assert_eq!(fields.private_fields().len(), 1);
    assert_eq!(fields.referenced_documents()[0].story(), FieldStory::Main);

    assert!(NonPlcfFields::from_story_texts([(FieldStory::Main, "\u{0013}TC")]).is_empty());
}

#[test]
fn table_of_authorities_fields_preserve_stored_configuration_without_generation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::TableOfAuthorities,
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
        instruction: r#" TOA \b Authorities \c 2 \d "-" \e " — " \f \g "–" \h \l ", " \p \s Section \* MERGEFORMAT \q opaque "#.to_string(),
        result: Some("cached authorities".to_string()),
    };

    let toa = text.table_of_authorities().unwrap();
    assert_eq!(toa.field(), &field);
    assert_eq!(toa.instruction(), text.instruction);
    assert_eq!(
        toa.options(),
        &[
            TableOfAuthoritiesOption::Bookmark("Authorities".to_string()),
            TableOfAuthoritiesOption::Category("2".to_string()),
            TableOfAuthoritiesOption::SequencePageSeparator("-".to_string()),
            TableOfAuthoritiesOption::EntryPageNumberSeparator(" — ".to_string()),
            TableOfAuthoritiesOption::EntryFormatting,
            TableOfAuthoritiesOption::PageRangeSeparator("–".to_string()),
            TableOfAuthoritiesOption::CategoryHeadings,
            TableOfAuthoritiesOption::PageReferenceSeparator(", ".to_string()),
            TableOfAuthoritiesOption::UsePassim,
            TableOfAuthoritiesOption::SequenceIdentifier("Section".to_string()),
        ]
    );
    assert_eq!(
        toa.unknown_switches(),
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
    assert_eq!(toa.cached_result(), Some("cached authorities"));
    assert!(toa.is_dirty());
    assert!(toa.is_locked());

    for instruction in ["TOA \\b", r"TOA \f unexpected", "TOA unexpected", "TOA \\"] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.table_of_authorities().is_none());
    }

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::TableOfContents,
            ..field
        },
        ..text
    };
    assert!(wrong_type.table_of_authorities().is_none());
}

#[test]
fn index_fields_preserve_stored_configuration_without_generation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Index,
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
        instruction: r#" INDEX \b "Scope Bookmark" \c 2 \d "-" \e ", " \f A \g "–" \h A \k "; " \l ", " \o S \p "A-D" \r \s Chapter \y \z 1033 \* MERGEFORMAT \q opaque "#.to_string(),
        result: Some("cached index".to_string()),
    };

    let index = text.index().unwrap();
    assert_eq!(index.field(), &field);
    assert_eq!(index.instruction(), text.instruction);
    assert_eq!(
        index.options(),
        &[
            IndexOption::Bookmark("Scope Bookmark".to_string()),
            IndexOption::Columns("2".to_string()),
            IndexOption::SequencePageSeparator("-".to_string()),
            IndexOption::EntryPageNumberSeparator(", ".to_string()),
            IndexOption::EntryType("A".to_string()),
            IndexOption::PageRangeSeparator("–".to_string()),
            IndexOption::Heading("A".to_string()),
            IndexOption::CrossReferenceSeparator("; ".to_string()),
            IndexOption::PageNumberSeparator(", ".to_string()),
            IndexOption::EastAsianSortOrder("S".to_string()),
            IndexOption::LetterRange("A-D".to_string()),
            IndexOption::RunIn,
            IndexOption::SequenceIdentifier("Chapter".to_string()),
            IndexOption::UseYomi,
            IndexOption::LanguageId("1033".to_string()),
        ]
    );
    assert_eq!(
        index.unknown_switches(),
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
    assert_eq!(index.cached_result(), Some("cached index"));
    assert!(index.is_dirty());
    assert!(index.is_locked());

    for instruction in [
        "INDEX \\b",
        "INDEX \\o",
        r"INDEX \r unexpected",
        r"INDEX \y unexpected",
        "INDEX unexpected",
        "INDEX \\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.index().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "INDEXES \\b Bookmark".to_string(),
        result: None,
    };
    assert!(wrong_keyword.index().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::TableOfContents,
            ..field
        },
        ..text
    };
    assert!(wrong_type.index().is_none());
}

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

#[test]
fn auto_text_fields_preserve_entries_without_lookup_or_insertion() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Glossary,
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
        instruction: r#" GLOSSARY "Legacy Clause" \* MERGEFORMAT \q opaque "#.to_string(),
        result: Some("cached glossary entry".to_string()),
    };

    let glossary = text.auto_text_field().unwrap();
    assert_eq!(glossary.field(), &field);
    assert_eq!(glossary.instruction(), text.instruction);
    assert_eq!(glossary.kind(), AutoTextFieldKind::Glossary);
    assert_eq!(glossary.entry_name(), "Legacy Clause");
    assert_eq!(
        glossary.unknown_switches(),
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
    assert_eq!(glossary.cached_result(), Some("cached glossary entry"));
    assert!(glossary.is_dirty());
    assert!(glossary.is_locked());

    let auto_text = FieldText {
        field: Field {
            field_type: FieldType::AutoText,
            ..field.clone()
        },
        instruction: r#" AUTOTEXT "Reusable Clause" \* MERGEFORMAT "#.to_string(),
        result: None,
    };
    let auto_text = auto_text.auto_text_field().unwrap();
    assert_eq!(auto_text.kind(), AutoTextFieldKind::AutoText);
    assert_eq!(auto_text.entry_name(), "Reusable Clause");
    assert_eq!(
        auto_text.unknown_switches(),
        &[MergeFieldSwitch {
            name: '*',
            argument: Some("MERGEFORMAT".to_string()),
        }]
    );
    assert_eq!(auto_text.cached_result(), None);

    let historical_alias = FieldText {
        field: field.clone(),
        instruction: r#" AUTOTEXT "Legacy Alias" "#.to_string(),
        result: None,
    };
    let historical_alias = historical_alias.auto_text_field().unwrap();
    assert_eq!(historical_alias.kind(), AutoTextFieldKind::Glossary);
    assert_eq!(historical_alias.entry_name(), "Legacy Alias");

    for instruction in [
        "GLOSSARY",
        r#"GLOSSARY ""#,
        "GLOSSARY Entry unexpected",
        "GLOSSARY Entry \\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.auto_text_field().is_none(), "{instruction}");
    }

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.auto_text_field().is_none());
}

#[test]
fn auto_text_list_fields_preserve_metadata_without_selection() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::AutoTextList,
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
        instruction: r#" AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Right-click to select" \* MERGEFORMAT \q opaque "#.to_string(),
        result: Some("cached selection".to_string()),
    };

    let list = text.auto_text_list_field().unwrap();
    assert_eq!(list.field(), &field);
    assert_eq!(list.instruction(), text.instruction);
    assert_eq!(list.display_text(), Some("Choose a name"));
    assert_eq!(
        list.options(),
        &[
            AutoTextListOption::Style("Name Style".to_string()),
            AutoTextListOption::Tip("Right-click to select".to_string()),
        ]
    );
    assert_eq!(
        list.unknown_switches(),
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
    assert_eq!(list.cached_result(), Some("cached selection"));
    assert!(list.is_dirty());
    assert!(list.is_locked());

    let no_display_text = FieldText {
        instruction: r"AUTOTEXTLIST \s NameStyle".to_string(),
        ..text.clone()
    };
    let no_display_text = no_display_text.auto_text_list_field().unwrap();
    assert_eq!(no_display_text.display_text(), None);
    assert_eq!(
        no_display_text.options(),
        &[AutoTextListOption::Style("NameStyle".to_string())]
    );

    for instruction in [
        "AUTOTEXTLIST \\\\s",
        "AUTOTEXTLIST \\\\t",
        "AUTOTEXTLIST display unexpected",
        "AUTOTEXTLIST \\\\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.auto_text_list_field().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        field: field.clone(),
        instruction: "AUTOTEXTLISTS display".to_string(),
        result: None,
    };
    assert!(wrong_keyword.auto_text_list_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::AutoText,
            ..field
        },
        ..text
    };
    assert!(wrong_type.auto_text_list_field().is_none());
}

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
fn merge_field_exposes_stored_metadata_without_merging() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::MergeField,
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
        instruction: r#" MERGEFIELD "Customer Region" \b "Dear " \f "!" \m \v \* MERGEFORMAT "#
            .to_string(),
        result: Some("cached customer".to_string()),
    };

    let merge = text.merge_field().unwrap();
    assert_eq!(merge.field(), &field);
    assert_eq!(merge.instruction(), text.instruction);
    assert_eq!(merge.field_name(), "Customer Region");
    assert_eq!(merge.cached_result(), Some("cached customer"));
    assert!(merge.is_dirty());
    assert!(merge.is_locked());
    assert_eq!(merge.switches().len(), 5);
    assert_eq!(merge.switches()[0].name(), 'b');
    assert_eq!(merge.switches()[0].argument(), Some("Dear "));
    assert_eq!(merge.switches()[1].name(), 'f');
    assert_eq!(merge.switches()[1].argument(), Some("!"));
    assert!(merge.has_switch('m'));
    assert!(merge.has_switch('v'));
    assert!(merge.has_switch('*'));
    assert_eq!(merge.switches()[4].argument(), Some("MERGEFORMAT"));

    let compact = FieldText {
        instruction: r#"MERGEFIELD"Customer Name"\f" ""#.to_string(),
        ..text.clone()
    };
    let compact_merge = compact.merge_field().unwrap();
    assert_eq!(compact_merge.field_name(), "Customer Name");
    assert_eq!(compact_merge.switches()[0].argument(), Some(" "));

    let missing_name = FieldText {
        instruction: r#"MERGEFIELD \* MERGEFORMAT"#.to_string(),
        ..text.clone()
    };
    assert!(missing_name.merge_field().is_none());

    let unexpected_operand = FieldText {
        instruction: "MERGEFIELD Customer unexpected".to_string(),
        ..text.clone()
    };
    assert!(unexpected_operand.merge_field().is_none());

    let wrong_keyword = FieldText {
        instruction: "MERGEFIELDS Customer".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.merge_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::DocumentVariable,
            ..field
        },
        ..text
    };
    assert!(wrong_type.merge_field().is_none());
}

#[test]
fn mail_merge_data_fields_expose_sources_without_opening_them() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Data,
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
        instruction: r#" DATA "unavailable.csv" "unavailable.hdr" \* MERGEFORMAT \x retained "#
            .to_string(),
        result: Some("cached data source".to_string()),
    };

    let data = text.mail_merge_data().unwrap();
    assert_eq!(data.field(), &field);
    assert_eq!(data.instruction(), text.instruction);
    assert_eq!(data.data_source(), "unavailable.csv");
    assert_eq!(data.header_source(), Some("unavailable.hdr"));
    assert_eq!(data.cached_result(), Some("cached data source"));
    assert!(data.is_dirty());
    assert!(data.is_locked());
    assert_eq!(data.switches().len(), 2);
    assert_eq!(data.switches()[0].name(), '*');
    assert_eq!(data.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(data.switches()[1].name(), 'x');
    assert_eq!(data.switches()[1].argument(), Some("retained"));

    let no_header = FieldText {
        instruction: r#"DATA source.csv \* MERGEFORMAT"#.to_string(),
        ..text.clone()
    };
    let no_header = no_header.mail_merge_data().unwrap();
    assert_eq!(no_header.data_source(), "source.csv");
    assert_eq!(no_header.header_source(), None);

    for instruction in [
        "DATA",
        r#"DATA ""#,
        r#"DATA source.csv """#,
        "DATA source.csv header.hdr unexpected",
        "DATA source.csv \\",
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.mail_merge_data().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        instruction: "DATABASE source.csv".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.mail_merge_data().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeField,
            ..field
        },
        ..text
    };
    assert!(wrong_type.mail_merge_data().is_none());
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
fn dde_links_expose_cached_metadata_without_activation() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Dde,
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
        instruction: r#" DDE Excel "missing.xlsx" "Sheet1!A1" \a \p \x "ignored" "#.to_string(),
        result: Some("cached DDE".to_string()),
    };

    let dde = text.dde_link().unwrap();
    assert_eq!(dde.field(), &field);
    assert_eq!(dde.instruction(), text.instruction);
    assert_eq!(dde.kind(), DdeFieldKind::Dde);
    assert_eq!(dde.application(), "Excel");
    assert_eq!(dde.source(), "missing.xlsx");
    assert_eq!(dde.item(), Some("Sheet1!A1"));
    assert!(dde.requests_automatic_updates());
    assert_eq!(dde.representation(), Some(DdeRepresentation::Picture));
    assert!(!dde.omits_graphic_data());
    assert_eq!(dde.cached_result(), Some("cached DDE"));
    assert!(dde.is_dirty());
    assert!(dde.is_locked());
    assert_eq!(dde.unknown_switches().len(), 1);
    assert_eq!(dde.unknown_switches()[0].name(), 'x');
    assert_eq!(dde.unknown_switches()[0].argument(), Some("ignored"));

    let automatic = FieldText {
        field: Field {
            field_type: FieldType::DdeAuto,
            ..field.clone()
        },
        instruction: r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \t"#.to_string(),
        result: Some("cached auto".to_string()),
    };
    let automatic = automatic.dde_link().unwrap();
    assert_eq!(automatic.kind(), DdeFieldKind::DdeAuto);
    assert_eq!(automatic.item(), Some("Sheet1!A2"));
    assert!(automatic.requests_automatic_updates());
    assert_eq!(automatic.representation(), Some(DdeRepresentation::Text));

    for instruction in [
        r#"DDE Excel source \p \t"#,
        r#"DDEAUTO Excel source \a"#,
        r#"DDE Excel source \d \p"#,
        r#"DDE Excel source \a value"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.dde_link().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        instruction: "DDEAUTOMATED Excel source".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.dde_link().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Link,
            ..field
        },
        ..text
    };
    assert!(wrong_type.dde_link().is_none());
}

#[test]
fn link_fields_expose_cached_metadata_without_activating_sources() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::Link,
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
            r#" LINK Excel.Sheet.12 "missing.xlsx" "Sheet1!A1" \a \p \f 2 \f 9 \x "ignored" "#
                .to_string(),
        result: Some("cached link".to_string()),
    };

    let link = text.link_field().unwrap();
    assert_eq!(link.field(), &field);
    assert_eq!(link.instruction(), text.instruction);
    assert_eq!(link.application_type(), "Excel.Sheet.12");
    assert_eq!(link.source(), "missing.xlsx");
    assert_eq!(link.item(), Some("Sheet1!A1"));
    assert!(link.requests_automatic_updates());
    assert_eq!(link.result_options(), &[LinkResultOption::Picture]);
    assert_eq!(
        link.effective_result_option(),
        Some(LinkResultOption::Picture)
    );
    assert_eq!(
        link.formatting_modes(),
        &[LinkFormatting::Destination, LinkFormatting::Unsupported(9)]
    );
    assert_eq!(link.cached_result(), Some("cached link"));
    assert!(link.is_dirty());
    assert!(link.is_locked());
    assert_eq!(link.switches().len(), 5);
    assert_eq!(link.switches()[4].name(), 'x');
    assert_eq!(link.switches()[4].argument(), Some("ignored"));

    let no_item = FieldText {
        instruction: r#"LINK Excel.Sheet.12 "missing.xlsx" \d \b"#.to_string(),
        ..text.clone()
    };
    let no_item = no_item.link_field().unwrap();
    assert_eq!(no_item.item(), None);
    assert_eq!(
        no_item.result_options(),
        &[LinkResultOption::OmitGraphicData, LinkResultOption::Bitmap]
    );
    assert_eq!(
        no_item.effective_result_option(),
        Some(LinkResultOption::Bitmap)
    );

    for instruction in [
        r#"LINK Excel source \a value"#,
        r#"LINK Excel source \p value"#,
        r#"LINK Excel source \f"#,
        r#"LINK Excel source \f not-an-integer"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.link_field().is_none(), "{instruction}");
    }

    let wrong_keyword = FieldText {
        instruction: "LINKS Excel source".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.link_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::Dde,
            ..field
        },
        ..text
    };
    assert!(wrong_type.link_field().is_none());
}

#[test]
fn external_include_fields_expose_cached_metadata_without_opening_sources() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(37),
        end_cp: 52,
        field_type: FieldType::IncludeText,
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
        instruction: r#" INCLUDETEXT "unavailable.xml" Summary \! \c Word8 \e utf-8 \m application/xml \n "xmlns:a=\"resume-schema\"" \t "file:///unavailable.xsl" \x a:Resume/a:Name \* MERGEFORMAT "#
            .to_string(),
        result: Some("cached include".to_string()),
    };

    let include = text.external_include().unwrap();
    assert_eq!(include.field(), &field);
    assert_eq!(include.instruction(), text.instruction);
    assert_eq!(include.kind(), IncludeFieldKind::Text);
    assert_eq!(include.source(), "unavailable.xml");
    assert_eq!(include.bookmark(), Some("Summary"));
    assert_eq!(include.converter(), Some("Word8"));
    assert!(include.suppresses_nested_field_updates());
    assert!(!include.omits_picture_data());
    assert_eq!(
        include.options(),
        &[
            ExternalIncludeOption::Converter("Word8".to_string()),
            ExternalIncludeOption::Encoding("utf-8".to_string()),
            ExternalIncludeOption::MimeType("application/xml".to_string()),
            ExternalIncludeOption::NamespaceMapping("xmlns:a=\"resume-schema\"".to_string()),
            ExternalIncludeOption::Xslt("file:///unavailable.xsl".to_string()),
            ExternalIncludeOption::XPath("a:Resume/a:Name".to_string()),
        ]
    );
    assert_eq!(include.unknown_switches().len(), 1);
    assert_eq!(include.unknown_switches()[0].name(), '*');
    assert_eq!(
        include.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(include.cached_result(), Some("cached include"));
    assert!(include.is_dirty());
    assert!(include.is_locked());

    let picture = FieldText {
        field: Field {
            field_type: FieldType::IncludePicture,
            ..field.clone()
        },
        instruction: r#"INCLUDEPICTURE "unavailable.gif" \c Pictim32 \d \* MERGEFORMAT"#
            .to_string(),
        result: Some("cached picture".to_string()),
    };
    let picture_include = picture.external_include().unwrap();
    assert_eq!(picture_include.kind(), IncludeFieldKind::Picture);
    assert_eq!(picture_include.source(), "unavailable.gif");
    assert_eq!(picture_include.bookmark(), None);
    assert_eq!(picture_include.converter(), Some("Pictim32"));
    assert_eq!(
        picture_include.options(),
        &[ExternalIncludeOption::Converter("Pictim32".to_string())]
    );
    assert!(!picture_include.suppresses_nested_field_updates());
    assert!(picture_include.omits_picture_data());
    assert_eq!(picture_include.cached_result(), Some("cached picture"));

    let legacy_text = FieldText {
        field: Field {
            field_type: FieldType::Include,
            ..field.clone()
        },
        instruction: r#"INCLUDE "unavailable.docx" LegacySection \!"#.to_string(),
        result: None,
    };
    let legacy_text = legacy_text.external_include().unwrap();
    assert_eq!(legacy_text.kind(), IncludeFieldKind::Text);
    assert_eq!(legacy_text.source(), "unavailable.docx");
    assert_eq!(legacy_text.bookmark(), Some("LegacySection"));
    assert!(legacy_text.suppresses_nested_field_updates());

    let legacy_picture = FieldText {
        field: Field {
            field_type: FieldType::Import,
            ..field.clone()
        },
        instruction: r#"IMPORT "unavailable.wmf" \c GraphicsFilter \d"#.to_string(),
        result: None,
    };
    let legacy_picture = legacy_picture.external_include().unwrap();
    assert_eq!(legacy_picture.kind(), IncludeFieldKind::Picture);
    assert_eq!(legacy_picture.source(), "unavailable.wmf");
    assert_eq!(legacy_picture.converter(), Some("GraphicsFilter"));
    assert!(legacy_picture.omits_picture_data());

    for instruction in [
        "INCLUDETEXT",
        r#"INCLUDETEXT \c Word8"#,
        r#"INCLUDETEXT source \! unexpected"#,
        r#"INCLUDETEXT source \e"#,
        r#"INCLUDETEXT source \! \!"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..text.clone()
        };
        assert!(malformed.external_include().is_none(), "{instruction}");
    }
    for instruction in [
        r#"INCLUDEPICTURE "picture.gif" Selector"#,
        r#"INCLUDEPICTURE "picture.gif" \d extra"#,
        r#"INCLUDEPICTURE "picture.gif" \d \d"#,
        r#"INCLUDEPICTURE "picture.gif" \c"#,
    ] {
        let malformed = FieldText {
            instruction: instruction.to_string(),
            ..picture.clone()
        };
        assert!(malformed.external_include().is_none(), "{instruction}");
    }

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::IncludePicture,
            ..field
        },
        ..text
    };
    assert!(wrong_type.external_include().is_none());
}

#[test]
fn mail_merge_counters_expose_cached_metadata_without_merging() {
    let record_field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(13),
        end_cp: 16,
        field_type: FieldType::MergeRecord,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let record = FieldText {
        field: record_field.clone(),
        instruction: " MERGEREC ".to_string(),
        result: Some("12".to_string()),
    };

    let counter = record.mail_merge_counter().unwrap();
    assert_eq!(counter.field(), &record_field);
    assert_eq!(counter.instruction(), record.instruction);
    assert_eq!(counter.kind(), MailMergeCounterKind::Record);
    assert_eq!(counter.cached_result(), Some("12"));
    assert!(counter.is_dirty());
    assert!(counter.is_locked());

    let sequence = FieldText {
        field: Field {
            field_type: FieldType::MergeSequence,
            ..record_field.clone()
        },
        instruction: "mergeSEQ".to_string(),
        result: Some("3".to_string()),
    };
    let sequence_counter = sequence.mail_merge_counter().unwrap();
    assert_eq!(sequence_counter.kind(), MailMergeCounterKind::Sequence);
    assert_eq!(sequence_counter.cached_result(), Some("3"));

    let unexpected_operand = FieldText {
        instruction: "MERGEREC 12".to_string(),
        ..record.clone()
    };
    assert!(unexpected_operand.mail_merge_counter().is_none());

    let unexpected_switch = FieldText {
        instruction: r"MERGESEQ \* MERGEFORMAT".to_string(),
        ..sequence.clone()
    };
    assert!(unexpected_switch.mail_merge_counter().is_none());

    let wrong_keyword = FieldText {
        instruction: "MERGERECORD".to_string(),
        ..record.clone()
    };
    assert!(wrong_keyword.mail_merge_counter().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::MergeSequence,
            ..record_field
        },
        ..record
    };
    assert!(wrong_type.mail_merge_counter().is_none());
}

#[test]
fn mail_merge_next_fields_expose_cached_metadata_without_advancing_records() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::Next,
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
        instruction: " NEXT ".to_string(),
        result: Some("cached next".to_string()),
    };

    let next = text.mail_merge_next().unwrap();
    assert_eq!(next.field(), &field);
    assert_eq!(next.instruction(), text.instruction);
    assert_eq!(next.cached_result(), Some("cached next"));
    assert!(next.is_dirty());
    assert!(next.is_locked());

    let unexpected_operand = FieldText {
        instruction: "NEXT 12".to_string(),
        ..text.clone()
    };
    assert!(unexpected_operand.mail_merge_next().is_none());

    let unexpected_switch = FieldText {
        instruction: r"NEXT \* MERGEFORMAT".to_string(),
        ..text.clone()
    };
    assert!(unexpected_switch.mail_merge_next().is_none());

    let wrong_keyword = FieldText {
        instruction: "NEXTIF Customer = Ada".to_string(),
        ..text.clone()
    };
    assert!(wrong_keyword.mail_merge_next().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::NextIf,
            ..field
        },
        ..text
    };
    assert!(wrong_type.mail_merge_next().is_none());
}

#[test]
fn conditional_mail_merge_controls_expose_metadata_without_merging() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::NextIf,
        end_flags: FieldEndFlags {
            results_dirty: true,
            locked: true,
            has_separator: true,
            ..FieldEndFlags::default()
        },
        nesting_depth: 1,
        has_separator: true,
    };
    let next_if = FieldText {
        field: field.clone(),
        instruction: r#" NEXTIF Customer = "Ada" "#.to_string(),
        result: Some("cached nextif".to_string()),
    };

    let control = next_if.mail_merge_conditional_control().unwrap();
    assert_eq!(control.field(), &field);
    assert_eq!(control.instruction(), next_if.instruction);
    assert_eq!(control.kind(), MailMergeConditionalControlKind::NextIf);
    assert_eq!(control.comparison(), r#"Customer = "Ada""#);
    assert_eq!(control.cached_result(), Some("cached nextif"));
    assert!(control.is_dirty());
    assert!(control.is_locked());

    let skip_if = FieldText {
        field: Field {
            field_type: FieldType::SkipIf,
            ..field.clone()
        },
        instruction: "skipif MERGEFIELD Order < 100".to_string(),
        result: Some("cached skipif".to_string()),
    };
    let skip_control = skip_if.mail_merge_conditional_control().unwrap();
    assert_eq!(skip_control.kind(), MailMergeConditionalControlKind::SkipIf);
    assert_eq!(skip_control.comparison(), "MERGEFIELD Order < 100");
    assert_eq!(skip_control.cached_result(), Some("cached skipif"));

    let missing_comparison = FieldText {
        instruction: "NEXTIF".to_string(),
        ..next_if.clone()
    };
    assert!(
        missing_comparison
            .mail_merge_conditional_control()
            .is_none()
    );

    let wrong_keyword = FieldText {
        instruction: "NEXTIFF Customer = Ada".to_string(),
        ..next_if.clone()
    };
    assert!(wrong_keyword.mail_merge_conditional_control().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::SkipIf,
            ..field
        },
        ..next_if
    };
    assert!(wrong_type.mail_merge_conditional_control().is_none());
}

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

#[test]
fn recipient_fields_expose_layout_metadata_without_merging() {
    let field = Field {
        story: FieldStory::Textbox,
        start_cp: 4,
        separator_cp: Some(9),
        end_cp: 22,
        field_type: FieldType::AddressBlock,
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
        instruction: r#" ADDRESSBLOCK \c 2 \d \e "United States" \e Canada \f "<<_FIRST0_>> <<_LAST0_>>" \l 1033 \* MERGEFORMAT "#
            .to_string(),
        result: Some("cached address".to_string()),
    };

    let address = address.mail_merge_recipient_field().unwrap();
    assert_eq!(address.field(), &field);
    assert_eq!(address.kind(), MailMergeRecipientFieldKind::AddressBlock);
    assert_eq!(
        address.country_inclusion(),
        Some(AddressBlockCountryInclusion::UnlessExcluded)
    );
    assert!(address.formats_using_recipient_country());
    let excluded = address
        .excluded_countries()
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(excluded, vec!["United States", "Canada"]);
    assert_eq!(address.format_template(), Some("<<_FIRST0_>> <<_LAST0_>>"));
    assert_eq!(address.language(), Some("1033"));
    assert_eq!(address.greeting_fallback_text(), None);
    assert_eq!(address.unknown_switches().len(), 1);
    assert_eq!(address.unknown_switches()[0].name(), '*');
    assert_eq!(
        address.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(address.cached_result(), Some("cached address"));
    assert!(address.is_dirty());
    assert!(address.is_locked());

    let greeting = FieldText {
        field: Field {
            field_type: FieldType::GreetingLine,
            ..field.clone()
        },
        instruction: r#"greetingline \f "Dear <<_FIRST0_>>," \e "To Whom It May Concern" \l en-US"#
            .to_string(),
        result: Some("Dear Ada,".to_string()),
    };
    let greeting = greeting.mail_merge_recipient_field().unwrap();
    assert_eq!(greeting.kind(), MailMergeRecipientFieldKind::GreetingLine);
    assert_eq!(greeting.country_inclusion(), None);
    assert!(!greeting.formats_using_recipient_country());
    assert!(greeting.excluded_countries().is_empty());
    assert_eq!(greeting.format_template(), Some("Dear <<_FIRST0_>>,"));
    assert_eq!(greeting.language(), Some("en-US"));
    assert_eq!(
        greeting.greeting_fallback_text(),
        Some("To Whom It May Concern")
    );
    assert_eq!(greeting.cached_result(), Some("Dear Ada,"));

    for instruction in [
        "ADDRESSBLOCK text",
        r"ADDRESSBLOCK \c",
        r"ADDRESSBLOCK \c 3",
        r"ADDRESSBLOCK \d 1",
        r"ADDRESSBLOCK \d \d",
        r"ADDRESSBLOCK \f",
        r#"GREETINGLINE \f "Dear" \f "Hello""#,
        r"GREETINGLINE \l",
        r#"GREETINGLINE \c "First" \e "Second""#,
    ] {
        let malformed = FieldText {
            field: field.clone(),
            instruction: instruction.to_string(),
            result: None,
        };
        assert!(
            malformed.mail_merge_recipient_field().is_none(),
            "{instruction}"
        );
    }

    let wrong_keyword = FieldText {
        instruction: r"ADDRESSBLOCKING \c 1".to_string(),
        field: field.clone(),
        result: None,
    };
    assert!(wrong_keyword.mail_merge_recipient_field().is_none());

    let wrong_type = FieldText {
        field: Field {
            field_type: FieldType::GreetingLine,
            ..field
        },
        instruction: r"ADDRESSBLOCK \c 1".to_string(),
        result: None,
    };
    assert!(wrong_type.mail_merge_recipient_field().is_none());
}

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

#[cfg(test)]
mod terminal_cp_regression_tests {
    use super::{FieldStory, FieldStoryTable};

    fn field_plcf(cps: &[u32]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&[0x13, 0x21]);
        data.extend_from_slice(&[0x15, 0x00]);
        data
    }

    #[test]
    fn terminal_cp_is_not_story_bounded_but_marker_cps_are() {
        let data = field_plcf(&[1, 3, u32::MAX]);
        let parsed = FieldStoryTable::parse_plcf(FieldStory::Main, 3, &data)
            .expect("undefined terminal CP may exceed the story length");
        assert_eq!(parsed.terminal_cp, u32::MAX);
        assert_eq!(parsed.markers[0].position, 1);
        assert_eq!(parsed.markers[1].position, 3);

        let marker_outside_story = field_plcf(&[1, 4, 5]);
        assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 3, &marker_outside_story).is_err());
    }
}
