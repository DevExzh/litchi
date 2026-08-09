#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]
#![allow(
    clippy::let_underscore_must_use,
    reason = "totality smoke tests deliberately discard view results"
)]

use super::*;
use std::borrow::Cow;

fn exercise_field_views(instruction: &str) {
    let field = Field::parse_instruction(instruction);
    let _ = field.parsed_code();
    let _ = field.extract_url();
    let _ = field.extract_bookmark();
    let _ = field.equation();
    let _ = field.macro_button();
    let _ = field.go_to_button();
    let _ = field.print_field();
    let _ = field.embed_field();
    let _ = field.barcode_field();
    let _ = field.barcode_display_field();
    let _ = field.bidi_outline_field();
    let _ = field.shape_field();
    let _ = field.legacy_form_field();
    let _ = field.private_field();
    let _ = field.active_content_field();
    let _ = field.auto_text_field();
    let _ = field.auto_text_list_field();
    let _ = field.dde_link();
    let _ = field.link_field();
    let _ = field.external_include();
    let _ = field.referenced_document();
    let _ = field.table_of_contents();
    let _ = field.table_of_contents_entry();
    let _ = field.table_of_authorities_entry();
    let _ = field.table_of_authorities();
    let _ = field.index();
    let _ = field.index_entry();
    let _ = field.citation();
    let _ = field.bibliography();
    let _ = field.document_variable();
    let _ = field.document_property();
    let _ = field.info_field();
    let _ = field.document_information();
    let _ = field.document_context();
    let _ = field.merge_field();
    let _ = field.database_field();
    let _ = field.mail_merge_data();
    let _ = field.mail_merge_counter();
    let _ = field.mail_merge_next();
    let _ = field.mail_merge_conditional_control();
    let _ = field.if_field();
    let _ = field.compare_field();
    let _ = field.set_field();
    let _ = field.sequence_field();
    let _ = field.formula_field();
    let _ = field.quote_field();
    let _ = field.symbol_field();
    let _ = field.auto_number_field();
    let _ = field.list_number_field();
    let _ = field.style_reference_field();
    let _ = field.prompt_field();
    let _ = field.user_identity_field();
    let _ = field.advance_field();
    let _ = field.mail_merge_recipient_field();
    let _ = field.hyperlink();
    let _ = field.reference_field();
}

#[test]
fn field_views_are_total_for_truncated_and_mutated_instructions() {
    let seeds = [
        r#"HYPERLINK "https://例.example/a b" \l mark \n"#,
        r"REF mark \h \p",
        r#"DISPLAYBARCODE "123 456" QR \q 3"#,
        r#"MACROBUTTON NoMacro "Click here now""#,
        r#"GOTOBUTTON mark "Jump here""#,
        r"AUTOTEXT Clause \x value",
        r#"AUTOTEXTLIST "Choose" \s "Name Style" \x value"#,
        r"DDE Excel book item \p \x value",
        r"LINK Excel.Sheet.8 book item \f 4 \x value",
        r"INCLUDETEXT source bookmark \c converter \! \q value",
        r"RD source.doc \f \x value",
        r#"TOC \o "1-3" \h \x value"#,
        r#"TC "entry" \l 2 \n \x value"#,
        r#"TA \l "long citation" \c 1 \b \x value"#,
        r"TOA \c 1 \h \x value",
        r"INDEX \c 2 \r \x value",
        r#"XE "entry" \b \t "see also" \x value"#,
        r"CITATION tag \l 1033 \n \x value",
        r"BIBLIOGRAPHY \l 1033 \x value",
        r"DOCVARIABLE name \x value",
        r"DOCPROPERTY Title \* MERGEFORMAT",
        r"INFO TITLE replacement \* MERGEFORMAT",
        r"TITLE \* MERGEFORMAT",
        r"FILENAME \p \* MERGEFORMAT",
        r#"MERGEFIELD Name \b "Dear ""#,
        r"DATA source.csv headers.csv \* MERGEFORMAT",
        r#"ASK mark "Prompt?" \d default \o"#,
        r#"USERNAME "Ada Lovelace" \* Upper"#,
        r"ADVANCE \d 1 \x 2",
        r"ADDRESSBLOCK \c 2 \d \e US \f fmt \l 1033 \x value",
        "MERGEREC",
        "NEXT",
        r#"NEXTIF "A" = "B""#,
        r#"SET name "value with spaces""#,
        r"SEQ Figure mark \r 1",
        r#"QUOTE "text" \* MERGEFORMAT"#,
        r"SYMBOL 65 \f Symbol",
        r"AUTONUM \s .",
        r"LISTNUM Name \l 2",
        r"STYLEREF Heading \l \x value",
    ];
    let replacements = ['\0', ' ', '"', '\\', '例'];

    for seed in seeds {
        for end in seed
            .char_indices()
            .map(|(index, _)| index)
            .chain(std::iter::once(seed.len()))
        {
            if let Some(prefix) = seed.get(..end) {
                exercise_field_views(prefix);
            }
        }

        for (start, character) in seed.char_indices() {
            let end = start + character.len_utf8();
            let Some(prefix) = seed.get(..start) else {
                continue;
            };
            let Some(suffix) = seed.get(end..) else {
                continue;
            };
            for replacement in replacements {
                let mut mutated =
                    String::with_capacity(prefix.len() + replacement.len_utf8() + suffix.len());
                mutated.push_str(prefix);
                mutated.push(replacement);
                mutated.push_str(suffix);
                exercise_field_views(&mutated);
            }
        }
    }
}

#[test]
fn field_code_limits_return_typed_malformed_results() {
    let too_long = "x".repeat(MAX_INSTRUCTION_LEN + 1);
    assert_eq!(
        parse_field_code(&too_long),
        ParsedFieldCode::Malformed(FieldCodeError::InstructionTooLong)
    );

    let too_many_tokens = (0..=MAX_TOKENS).map(|_| "x").collect::<Vec<_>>().join(" ");
    assert_eq!(
        parse_field_code(&too_many_tokens),
        ParsedFieldCode::Malformed(FieldCodeError::TooManyTokens)
    );
}

#[test]
fn story_validation_rejects_missing_drawing_references() {
    let drawing = crate::StoryDrawing::Shape(0);
    let error = validate_story_events(
        "",
        &[],
        &[],
        &[],
        &[StoryEvent::Drawing(drawing)],
        "test field",
    )
    .expect_err("an unresolved drawing reference must be rejected");
    assert!(matches!(
        error,
        crate::RtfError::MalformedDocument(message)
            if message.contains("invalid drawing reference")
    ));
}

#[test]
fn exact_case_insensitive_keywords_and_distinct_references() {
    assert!(matches!(
        parse_field_code("hyperlink \"https://e\""),
        ParsedFieldCode::Hyperlink(_)
    ));
    for invalid in ["HYPERLINKER x", "REFRESH x", "PAGEREFERENCE x"] {
        assert!(matches!(
            parse_field_code(invalid),
            ParsedFieldCode::Other { .. }
        ));
        assert_eq!(
            Field::parse_instruction(invalid).field_type,
            FieldType::Unknown
        );
    }
    assert!(matches!(
        parse_field_code("REF mark \\h"),
        ParsedFieldCode::Reference(_)
    ));
    assert!(matches!(
        parse_field_code("PAGEREF mark \\p"),
        ParsedFieldCode::PageReference(_)
    ));
    assert!(matches!(
        parse_field_code("NOTEREF mark \\f"),
        ParsedFieldCode::NoteReference(_)
    ));
}

#[test]
fn parses_internal_external_and_switch_semantics() {
    let ParsedFieldCode::Hyperlink(code) =
        parse_field_code(r#"HYPERLINK "https://example/a b" \l "_Toc1" \o "Tip" \t "_blank" \n"#)
    else {
        panic!("expected hyperlink");
    };
    assert_eq!(code.external_target.as_deref(), Some("https://example/a b"));
    assert_eq!(code.bookmark.as_deref(), Some("_Toc1"));
    assert_eq!(code.screen_tip.as_deref(), Some("Tip"));
    assert_eq!(code.target_frame.as_deref(), Some("_blank"));
    assert!(code.new_window);
    let field = Field::parse_instruction(r#"HYPERLINK \l "_Toc1""#);
    assert_eq!(field.extract_url().as_deref(), Some("#_Toc1"));
    assert_eq!(field.extract_bookmark().as_deref(), Some("_Toc1"));
}

#[test]
fn quoted_field_tokens_preserve_multibyte_characters() {
    let ParsedFieldCode::Hyperlink(code) = parse_field_code(r#"HYPERLINK "https://例.example/文""#)
    else {
        panic!("expected hyperlink");
    };

    assert_eq!(
        code.external_target.as_deref(),
        Some("https://例.example/文")
    );
}

#[test]
fn writer_operand_cannot_inject_switches_and_round_trips_specials() {
    let target = "c:\\docs\\a \" \\l \"attacker{one}";
    let instruction = format!("HYPERLINK {}", quoted_field_operand(target));
    let ParsedFieldCode::Hyperlink(code) = parse_field_code(&instruction) else {
        panic!("expected hyperlink");
    };
    assert_eq!(code.external_target.as_deref(), Some(target));
    assert!(code.bookmark.is_none());

    let mut rtf = br"{\rtf1\ansi ".to_vec();
    crate::RtfWriter::new(&mut rtf)
        .write_hyperlink(target, "safe link")
        .unwrap();
    rtf.push(b'}');
    let document = crate::RtfDocument::from_bytes(&rtf).unwrap();
    let ParsedFieldCode::Hyperlink(code) = document.fields()[0].parsed_code() else {
        panic!("expected serialized hyperlink");
    };
    assert_eq!(code.external_target.as_deref(), Some(target));
    assert!(code.bookmark.is_none());
}

#[test]
fn malformed_recognized_fields_are_non_actionable() {
    for instruction in [
        "HYPERLINK",
        r#"HYPERLINK "unterminated"#,
        r"HYPERLINK \l",
        r"HYPERLINK x \l a \l b",
        "REF",
        r"REF a \h \h",
    ] {
        assert!(matches!(
            parse_field_code(instruction),
            ParsedFieldCode::Malformed(_)
        ));
    }
}

#[test]
fn equation_fields_preserve_opaque_expression_metadata() {
    let mut field = Field::parse_instruction(r"EQ \o\ac(\fs24 Q,\fs16 R)");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    let equation = field.equation().unwrap();
    assert_eq!(equation.instruction(), r"EQ \o\ac(\fs24 Q,\fs16 R)");
    assert_eq!(equation.expression(), r"\o\ac(\fs24 Q,\fs16 R)");
    assert_eq!(equation.cached_result(), None);
    assert!(equation.is_dirty());
    assert!(equation.is_locked());
    assert_eq!(equation.owner(), FieldOwner::Body);
    assert_eq!(equation.position(), 4);

    let authored = Field::new_equation(r"\f(1,2)").unwrap();
    assert_eq!(authored.field_type, FieldType::Equation);
    assert_eq!(authored.equation().unwrap().expression(), r"\f(1,2)");
    assert!(Field::new_equation("x".repeat(MAX_INSTRUCTION_LEN)).is_err());
}

#[test]
fn macro_button_fields_expose_stored_metadata_without_execution() {
    let mut field = Field::parse_instruction(r#"MACROBUTTON NoMacro "Click here""#);
    field.result = Cow::Borrowed("Click here");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::MacroButton);
    let macro_button = field.macro_button().unwrap();
    assert_eq!(
        macro_button.instruction(),
        r#"MACROBUTTON NoMacro "Click here""#
    );
    assert_eq!(macro_button.macro_name(), "NoMacro");
    assert_eq!(macro_button.display_text(), Some("Click here"));
    assert_eq!(macro_button.cached_result(), Some("Click here"));
    assert!(macro_button.is_dirty());
    assert!(macro_button.is_locked());
    assert_eq!(macro_button.owner(), FieldOwner::Body);
    assert_eq!(macro_button.position(), 4);

    let multiword = Field::parse_instruction("MACROBUTTON NoMacro Click here now");
    assert_eq!(
        multiword.macro_button().unwrap().display_text(),
        Some("Click here now")
    );
    assert!(
        Field::parse_instruction("MACROBUTTON")
            .macro_button()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"MACROBUTTON "" "button""#)
            .macro_button()
            .is_none()
    );
}

#[test]
fn go_to_button_fields_expose_stored_metadata_without_navigation() {
    let mut field = Field::parse_instruction(r#"GOTOBUTTON "f 2" "Footnote""#);
    field.result = Cow::Borrowed("cached footnote button");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::GoToButton);
    let button = field.go_to_button().unwrap();
    assert_eq!(button.instruction(), r#"GOTOBUTTON "f 2" "Footnote""#);
    assert_eq!(button.target(), "f 2");
    assert_eq!(button.button_text(), "Footnote");
    assert_eq!(button.cached_result(), Some("cached footnote button"));
    assert!(button.is_dirty());
    assert!(button.is_locked());
    assert_eq!(button.owner(), FieldOwner::Body);
    assert_eq!(button.position(), 4);

    for instruction in [
        "GOTOBUTTON",
        r#"GOTOBUTTON "" Button"#,
        "GOTOBUTTON Destination",
        r#"GOTOBUTTON Destination """#,
        "GOTOBUTTON Destination Button unexpected",
        r"GOTOBUTTON Destination Button \* MERGEFORMAT",
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .go_to_button()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction("GOTOBUTTONS Destination Button").field_type,
        FieldType::Unknown
    );
}

#[test]
fn active_content_fields_expose_opaque_metadata_without_activation() {
    let mut add_in = Field::parse_instruction("ADDIN opaque-add-in-data");
    add_in.result = Cow::Borrowed("cached add-in result");
    add_in.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    add_in.owner = FieldOwner::Body;
    add_in.position = 4;

    assert_eq!(add_in.field_type, FieldType::AddIn);
    let add_in = add_in.active_content_field().unwrap();
    assert_eq!(add_in.instruction(), "ADDIN opaque-add-in-data");
    assert_eq!(add_in.kind(), ActiveContentFieldKind::AddIn);
    assert_eq!(add_in.cached_result(), Some("cached add-in result"));
    assert!(add_in.is_dirty());
    assert!(add_in.is_locked());
    assert_eq!(add_in.owner(), FieldOwner::Body);
    assert_eq!(add_in.position(), 4);

    let control = Field::parse_instruction("control opaque-ocx-metadata");
    assert_eq!(control.field_type, FieldType::Control);
    let control = control.active_content_field().unwrap();
    assert_eq!(control.kind(), ActiveContentFieldKind::OcxControl);
    assert_eq!(control.cached_result(), None);

    let html = Field::parse_instruction("HTMLCONTROL opaque-html-control-metadata");
    assert_eq!(html.field_type, FieldType::HtmlControl);
    let html = html.active_content_field().unwrap();
    assert_eq!(html.kind(), ActiveContentFieldKind::HtmlControl);

    assert_eq!(
        Field::parse_instruction("ADDINS not-an-add-in").field_type,
        FieldType::Unknown
    );
    assert!(
        Field::parse_instruction("MACROBUTTON NoMacro Button")
            .active_content_field()
            .is_none()
    );
}

#[test]
fn print_fields_preserve_opaque_metadata_without_sending_printer_commands() {
    let mut printer = Field::parse_instruction(r#"PRINT "ESC&l1O""#);
    printer.result = Cow::Borrowed("cached printer result");
    printer.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    printer.owner = FieldOwner::Body;
    printer.position = 4;

    assert_eq!(printer.field_type, FieldType::Print);
    let printer = printer.print_field().unwrap();
    assert_eq!(printer.instruction(), r#"PRINT "ESC&l1O""#);
    assert_eq!(printer.printer_instructions(), r#""ESC&l1O""#);
    assert_eq!(printer.cached_result(), Some("cached printer result"));
    assert!(printer.is_dirty());
    assert!(printer.is_locked());
    assert_eq!(printer.owner(), FieldOwner::Body);
    assert_eq!(printer.position(), 4);

    let postscript = Field::parse_instruction(r#"print \p 2 "0 0 moveto""#);
    assert_eq!(postscript.field_type, FieldType::Print);
    let postscript = postscript.print_field().unwrap();
    assert_eq!(postscript.printer_instructions(), r#"\p 2 "0 0 moveto""#);

    assert_eq!(
        Field::parse_instruction(r#"PRINTS "not a print field""#).field_type,
        FieldType::Unknown
    );
    assert!(
        Field::parse_instruction("ADDIN opaque-metadata")
            .print_field()
            .is_none()
    );
    let too_long = Field::new(
        FieldType::Print,
        Cow::Owned(format!("PRINT {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.print_field().is_none());
}

#[test]
fn embed_fields_preserve_opaque_metadata_without_loading_or_activating_objects() {
    let mut embedded = Field::parse_instruction(r"EMBED Excel.Sheet.12 \* MERGEFORMAT");
    embedded.result = Cow::Borrowed("cached worksheet object");
    embedded.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    embedded.owner = FieldOwner::Body;
    embedded.position = 4;

    assert_eq!(embedded.field_type, FieldType::Embed);
    let embedded = embedded.embed_field().unwrap();
    assert_eq!(
        embedded.instruction(),
        r"EMBED Excel.Sheet.12 \* MERGEFORMAT"
    );
    assert_eq!(
        embedded.object_instructions(),
        r"Excel.Sheet.12 \* MERGEFORMAT"
    );
    assert_eq!(embedded.cached_result(), Some("cached worksheet object"));
    assert!(embedded.is_dirty());
    assert!(embedded.is_locked());
    assert_eq!(embedded.owner(), FieldOwner::Body);
    assert_eq!(embedded.position(), 4);

    let equation = Field::parse_instruction(r#"embed "Equation.DSMT4" \d"#);
    assert_eq!(equation.field_type, FieldType::Embed);
    assert_eq!(
        equation.embed_field().unwrap().object_instructions(),
        r#""Equation.DSMT4" \d"#
    );

    let bare = Field::parse_instruction("EMBED");
    assert_eq!(bare.embed_field().unwrap().object_instructions(), "");
    assert_eq!(
        Field::parse_instruction("EMBEDS Excel.Sheet.12").field_type,
        FieldType::Unknown
    );
    assert!(
        Field::parse_instruction("EMBEDS Excel.Sheet.12")
            .embed_field()
            .is_none()
    );
    let too_long = Field::new(
        FieldType::Embed,
        Cow::Owned(format!("EMBED {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.embed_field().is_none());
}

#[test]
fn barcode_fields_preserve_opaque_metadata_without_decoding_or_rendering() {
    let mut barcode =
        Field::parse_instruction(r#"BARCODE "4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#);
    barcode.result = Cow::Borrowed("cached barcode");
    barcode.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    barcode.owner = FieldOwner::Body;
    barcode.position = 4;

    assert_eq!(barcode.field_type, FieldType::Barcode);
    let barcode = barcode.barcode_field().unwrap();
    assert_eq!(
        barcode.instruction(),
        r#"BARCODE "4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#
    );
    assert_eq!(
        barcode.barcode_instructions(),
        r#""4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#
    );
    assert_eq!(barcode.cached_result(), Some("cached barcode"));
    assert!(barcode.is_dirty());
    assert!(barcode.is_locked());
    assert_eq!(barcode.owner(), FieldOwner::Body);
    assert_eq!(barcode.position(), 4);

    let code_39 = Field::parse_instruction(r#"barcode "ABC-123" CODE39 \d"#);
    assert_eq!(code_39.field_type, FieldType::Barcode);
    assert_eq!(
        code_39.barcode_field().unwrap().barcode_instructions(),
        r#""ABC-123" CODE39 \d"#
    );

    let bare = Field::parse_instruction("BARCODE");
    assert_eq!(bare.barcode_field().unwrap().barcode_instructions(), "");
    assert_eq!(
        Field::parse_instruction("BARCODES 4901234567894").field_type,
        FieldType::Unknown
    );
    assert!(
        Field::parse_instruction("BARCODES 4901234567894")
            .barcode_field()
            .is_none()
    );
    let too_long = Field::new(
        FieldType::Barcode,
        Cow::Owned(format!("BARCODE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.barcode_field().is_none());
}

#[test]
fn bidi_outline_fields_preserve_metadata_without_resolving_numbering_or_layout() {
    let mut outline = Field::parse_instruction(r"BIDIOUTLINE \* MERGEFORMAT");
    outline.result = Cow::Borrowed("cached bidi outline number");
    outline.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    outline.owner = FieldOwner::Body;
    outline.position = 4;

    assert_eq!(outline.field_type, FieldType::BidiOutline);
    let outline = outline.bidi_outline_field().unwrap();
    assert_eq!(outline.instruction(), r"BIDIOUTLINE \* MERGEFORMAT");
    assert_eq!(outline.opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(outline.cached_result(), Some("cached bidi outline number"));
    assert!(outline.is_dirty());
    assert!(outline.is_locked());
    assert_eq!(outline.owner(), FieldOwner::Body);
    assert_eq!(outline.position(), 4);

    let bare = Field::parse_instruction("bidioutline");
    assert_eq!(bare.field_type, FieldType::BidiOutline);
    assert_eq!(bare.bidi_outline_field().unwrap().opaque_instructions(), "");
    assert_eq!(
        Field::parse_instruction("BIDIOUTLINES").field_type,
        FieldType::Unknown
    );
    assert!(
        Field::parse_instruction("BIDIOUTLINES")
            .bidi_outline_field()
            .is_none()
    );
    let too_long = Field::new(
        FieldType::BidiOutline,
        Cow::Owned(format!("BIDIOUTLINE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.bidi_outline_field().is_none());
}

#[test]
fn shape_fields_preserve_metadata_without_linking_or_rendering_drawings() {
    let mut shape = Field::parse_instruction(r"SHAPE \* MERGEFORMAT");
    shape.result = Cow::Borrowed("cached drawing anchor");
    shape.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    shape.owner = FieldOwner::Body;
    shape.position = 4;

    assert_eq!(shape.field_type, FieldType::Shape);
    let shape = shape.shape_field().unwrap();
    assert_eq!(shape.instruction(), r"SHAPE \* MERGEFORMAT");
    assert_eq!(shape.opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(shape.cached_result(), Some("cached drawing anchor"));
    assert!(shape.is_dirty());
    assert!(shape.is_locked());
    assert_eq!(shape.owner(), FieldOwner::Body);
    assert_eq!(shape.position(), 4);

    let bare = Field::parse_instruction("shape");
    assert_eq!(bare.field_type, FieldType::Shape);
    assert_eq!(bare.shape_field().unwrap().opaque_instructions(), "");
    assert_eq!(
        Field::parse_instruction("SHAPES").field_type,
        FieldType::Unknown
    );
    assert!(Field::parse_instruction("SHAPES").shape_field().is_none());
    let too_long = Field::new(
        FieldType::Shape,
        Cow::Owned(format!("SHAPE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.shape_field().is_none());
}

#[test]
fn legacy_form_fields_preserve_metadata_without_filling_or_executing() {
    let mut text = Field::parse_instruction(r"FORMTEXT \* MERGEFORMAT");
    text.result = Cow::Borrowed("cached text field");
    text.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    text.owner = FieldOwner::Body;
    text.position = 4;

    assert_eq!(text.field_type, FieldType::FormText);
    let text_field = text.legacy_form_field().unwrap();
    assert_eq!(text_field.kind(), LegacyFormFieldKind::Text);
    assert_eq!(text_field.instruction(), r"FORMTEXT \* MERGEFORMAT");
    assert_eq!(text_field.opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(text_field.cached_result(), Some("cached text field"));
    assert!(text_field.is_dirty());
    assert!(text_field.is_locked());
    assert_eq!(text_field.owner(), FieldOwner::Body);
    assert_eq!(text_field.position(), 4);

    let checkbox = Field::parse_instruction("formcheckbox");
    assert_eq!(checkbox.field_type, FieldType::FormCheckbox);
    let checkbox = checkbox.legacy_form_field().unwrap();
    assert_eq!(checkbox.kind(), LegacyFormFieldKind::CheckBox);
    assert_eq!(checkbox.opaque_instructions(), "");

    let drop_down = Field::parse_instruction(r"FORMDROPDOWN \* MERGEFORMAT");
    assert_eq!(drop_down.field_type, FieldType::FormDropdown);
    let drop_down = drop_down.legacy_form_field().unwrap();
    assert_eq!(drop_down.kind(), LegacyFormFieldKind::DropDown);
    assert_eq!(drop_down.opaque_instructions(), r"\* MERGEFORMAT");

    for instruction in [r"FORMTEXTUAL", r"FORMCHECKBOXLIST"] {
        assert_eq!(
            Field::parse_instruction(instruction).field_type,
            FieldType::Unknown
        );
        assert!(
            Field::parse_instruction(instruction)
                .legacy_form_field()
                .is_none()
        );
    }

    let mismatched_kind = Field::new(
        FieldType::FormText,
        Cow::Borrowed("FORMCHECKBOX"),
        Cow::Borrowed(""),
    );
    assert!(mismatched_kind.legacy_form_field().is_none());

    let too_long = Field::new(
        FieldType::FormText,
        Cow::Owned(format!("FORMTEXT {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.legacy_form_field().is_none());
}

#[test]
fn private_fields_preserve_conversion_data_without_conversion_or_layout() {
    let mut private = Field::parse_instruction(r"PRIVATE \* MERGEFORMAT");
    private.result = Cow::Borrowed("opaque converter payload");
    private.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    private.owner = FieldOwner::Body;
    private.position = 4;

    assert_eq!(private.field_type, FieldType::Private);
    let private_field = private.private_field().unwrap();
    assert_eq!(private_field.instruction(), r"PRIVATE \* MERGEFORMAT");
    assert_eq!(private_field.opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(
        private_field.cached_result(),
        Some("opaque converter payload")
    );
    assert!(private_field.is_dirty());
    assert!(private_field.is_locked());
    assert_eq!(private_field.owner(), FieldOwner::Body);
    assert_eq!(private_field.position(), 4);

    let bare = Field::parse_instruction("private");
    assert_eq!(bare.field_type, FieldType::Private);
    assert_eq!(bare.private_field().unwrap().opaque_instructions(), "");

    let nonmatching = Field::parse_instruction("PRIVATELY");
    assert_eq!(nonmatching.field_type, FieldType::Unknown);
    assert!(nonmatching.private_field().is_none());

    let too_long = Field::new(
        FieldType::Private,
        Cow::Owned(format!("PRIVATE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.private_field().is_none());
}

#[test]
fn auto_text_fields_preserve_metadata_without_lookup_or_insertion() {
    let mut glossary =
        Field::parse_instruction(r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT \q opaque"#);
    glossary.result = Cow::Borrowed("cached glossary entry");
    glossary.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    glossary.owner = FieldOwner::Body;
    glossary.position = 4;

    assert_eq!(glossary.field_type, FieldType::Glossary);
    let glossary = glossary.auto_text_field().unwrap();
    assert_eq!(
        glossary.instruction(),
        r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT \q opaque"#
    );
    assert_eq!(glossary.kind(), AutoTextFieldKind::Glossary);
    assert_eq!(glossary.entry_name(), "Legacy Clause");
    assert_eq!(glossary.unknown_switches().len(), 2);
    assert_eq!(glossary.unknown_switches()[0].name, "*");
    assert_eq!(
        glossary.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(glossary.unknown_switches()[1].name, "q");
    assert_eq!(
        glossary.unknown_switches()[1].value.as_deref(),
        Some("opaque")
    );
    assert_eq!(glossary.cached_result(), Some("cached glossary entry"));
    assert!(glossary.is_dirty());
    assert!(glossary.is_locked());
    assert_eq!(glossary.owner(), FieldOwner::Body);
    assert_eq!(glossary.position(), 4);

    let auto_text = Field::parse_instruction(r#"autotext "Reusable Clause" \* MERGEFORMAT"#);
    assert_eq!(auto_text.field_type, FieldType::AutoText);
    let auto_text = auto_text.auto_text_field().unwrap();
    assert_eq!(auto_text.kind(), AutoTextFieldKind::AutoText);
    assert_eq!(auto_text.entry_name(), "Reusable Clause");
    assert_eq!(auto_text.unknown_switches().len(), 1);
    assert_eq!(auto_text.unknown_switches()[0].name, "*");
    assert_eq!(
        auto_text.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(auto_text.cached_result(), None);

    let auto_text_list = Field::parse_instruction("AUTOTEXTLIST display");
    assert_eq!(auto_text_list.field_type, FieldType::AutoTextList);
    assert!(auto_text_list.auto_text_field().is_none());
    for instruction in [
        "GLOSSARY",
        r#"GLOSSARY ""#,
        "GLOSSARY Entry unexpected",
        r"GLOSSARY Entry \",
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .auto_text_field()
                .is_none(),
            "{instruction}"
        );
    }
    let too_long = Field::new(
        FieldType::Glossary,
        Cow::Owned(format!(
            "GLOSSARY Entry {}",
            "x".repeat(MAX_INSTRUCTION_LEN)
        )),
        Cow::Borrowed(""),
    );
    assert!(too_long.auto_text_field().is_none());
}

#[test]
fn auto_text_list_fields_preserve_metadata_without_selection_or_insertion() {
    let mut list = Field::parse_instruction(
        r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Right-click to select" \* MERGEFORMAT \q opaque"#,
    );
    list.result = Cow::Borrowed("cached selection");
    list.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    list.owner = FieldOwner::Body;
    list.position = 4;

    assert_eq!(list.field_type, FieldType::AutoTextList);
    let list = list.auto_text_list_field().unwrap();
    assert_eq!(
        list.instruction(),
        r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Right-click to select" \* MERGEFORMAT \q opaque"#
    );
    assert_eq!(list.display_text(), Some("Choose a name"));
    assert_eq!(
        list.options(),
        &[
            AutoTextListOption::Style(Cow::Borrowed("Name Style")),
            AutoTextListOption::Tip(Cow::Borrowed("Right-click to select")),
        ]
    );
    assert_eq!(list.unknown_switches().len(), 2);
    assert_eq!(list.unknown_switches()[0].name, "*");
    assert_eq!(
        list.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(list.unknown_switches()[1].name, "q");
    assert_eq!(list.unknown_switches()[1].value.as_deref(), Some("opaque"));
    assert_eq!(list.cached_result(), Some("cached selection"));
    assert!(list.is_dirty());
    assert!(list.is_locked());
    assert_eq!(list.owner(), FieldOwner::Body);
    assert_eq!(list.position(), 4);

    let style_only = Field::parse_instruction(r"autotextlist \s NameStyle");
    assert_eq!(style_only.field_type, FieldType::AutoTextList);
    let style_only = style_only.auto_text_list_field().unwrap();
    assert_eq!(style_only.display_text(), None);
    assert_eq!(
        style_only.options(),
        &[AutoTextListOption::Style(Cow::Borrowed("NameStyle"))]
    );
    assert_eq!(style_only.cached_result(), None);

    let empty_display = Field::parse_instruction(r#"AUTOTEXTLIST "" \s NameStyle"#);
    let empty_display = empty_display.auto_text_list_field().unwrap();
    assert_eq!(empty_display.display_text(), Some(""));

    assert_eq!(
        Field::parse_instruction("AUTOTEXTLISTS display").field_type,
        FieldType::Unknown
    );
    for instruction in [
        r"AUTOTEXTLIST \s",
        r"AUTOTEXTLIST \t",
        r"AUTOTEXTLIST \s \",
        "AUTOTEXTLIST display unexpected",
        r"AUTOTEXTLIST \",
        r#"AUTOTEXTLIST "unterminated"#,
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .auto_text_list_field()
                .is_none(),
            "{instruction}"
        );
    }
    let too_long = Field::new(
        FieldType::AutoTextList,
        Cow::Owned(format!("AUTOTEXTLIST {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.auto_text_list_field().is_none());
}

#[test]
fn dde_fields_expose_stored_metadata_without_contacting_sources() {
    let mut field = Field::parse_instruction(
        r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached DDE result");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Dde);
    let dde = field.dde_link().unwrap();
    assert_eq!(
        dde.instruction(),
        r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#
    );
    assert_eq!(dde.kind(), DdeFieldKind::Dde);
    assert_eq!(dde.application(), "Excel");
    assert_eq!(dde.source(), r"C:\no-contact\source.xlsx");
    assert_eq!(dde.item(), Some("Sheet1!R1C1:R4C4"));
    assert!(dde.requests_automatic_updates());
    assert_eq!(dde.representation(), Some(DdeRepresentation::Picture));
    assert!(!dde.omits_graphic_data());
    assert_eq!(dde.cached_result(), Some("cached DDE result"));
    assert!(dde.is_dirty());
    assert!(dde.is_locked());
    assert_eq!(dde.owner(), FieldOwner::Body);
    assert_eq!(dde.position(), 4);
    assert_eq!(dde.unknown_switches().len(), 1);
    assert_eq!(dde.unknown_switches()[0].name, "*");
    assert_eq!(
        dde.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    let automatic = Field::parse_instruction(r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A1" \t"#);
    assert_eq!(automatic.field_type, FieldType::DdeAuto);
    let automatic = automatic.dde_link().unwrap();
    assert_eq!(automatic.kind(), DdeFieldKind::DdeAuto);
    assert!(automatic.requests_automatic_updates());
    assert_eq!(automatic.representation(), Some(DdeRepresentation::Text));
    assert!(!automatic.omits_graphic_data());

    let omit_graphics = Field::parse_instruction(r"DDE Excel source \a \d");
    let omit_graphics = omit_graphics.dde_link().unwrap();
    assert!(omit_graphics.requests_automatic_updates());
    assert_eq!(omit_graphics.representation(), None);
    assert!(omit_graphics.omits_graphic_data());

    assert!(Field::parse_instruction("DDE").dde_link().is_none());
    assert!(
        Field::parse_instruction(r"DDE Excel \p")
            .dde_link()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DDE Excel source \p unexpected")
            .dde_link()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DDE Excel source \p \t")
            .dde_link()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DDEAUTO Excel source \p \t")
            .dde_link()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DDEAUTO Excel source \a")
            .dde_link()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("DDEAUTOMATED Excel source").field_type,
        FieldType::Unknown
    );
}

#[test]
fn link_fields_expose_stored_metadata_without_activating_sources() {
    let mut field = Field::parse_instruction(
        r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached LINK result");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Link);
    let link = field.link_field().unwrap();
    assert_eq!(
        link.instruction(),
        r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#
    );
    assert_eq!(link.application_type(), "Excel.Sheet.8");
    assert_eq!(link.source(), r"C:\no-contact\source.xlsx");
    assert_eq!(link.item(), Some("Sheet1!R1C1:R4C4"));
    assert!(link.requests_automatic_updates());
    assert_eq!(
        link.result_options(),
        &[LinkResultOption::Picture, LinkResultOption::OmitGraphicData]
    );
    assert_eq!(
        link.effective_result_option(),
        Some(LinkResultOption::OmitGraphicData)
    );
    assert_eq!(
        link.formatting_modes(),
        &[LinkFormatting::SpreadsheetSource]
    );
    assert_eq!(link.cached_result(), Some("cached LINK result"));
    assert!(link.is_dirty());
    assert!(link.is_locked());
    assert_eq!(link.owner(), FieldOwner::Body);
    assert_eq!(link.position(), 4);
    assert_eq!(link.unknown_switches().len(), 1);
    assert_eq!(link.unknown_switches()[0].name, "*");
    assert_eq!(
        link.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    let destination =
        Field::parse_instruction(r#"LINK Word.Document.8 "missing.docx" Bookmark \f 2 \t"#);
    let destination = destination.link_field().unwrap();
    assert!(!destination.requests_automatic_updates());
    assert_eq!(
        destination.formatting_modes(),
        &[LinkFormatting::Destination]
    );
    assert_eq!(
        destination.effective_result_option(),
        Some(LinkResultOption::Text)
    );

    let unsupported = Field::parse_instruction(r"LINK Package source \f 1");
    assert_eq!(
        unsupported.link_field().unwrap().formatting_modes(),
        &[LinkFormatting::Unsupported(1)]
    );

    let multiple_formatting = Field::parse_instruction(r"LINK Excel.Sheet.8 source \f 0 \f 2");
    assert_eq!(
        multiple_formatting.link_field().unwrap().formatting_modes(),
        &[LinkFormatting::Source, LinkFormatting::Destination]
    );

    let repeated_updates = Field::parse_instruction(r"LINK Excel.Sheet.8 source \a \a");
    assert!(
        repeated_updates
            .link_field()
            .unwrap()
            .requests_automatic_updates()
    );

    assert!(Field::parse_instruction("LINK").link_field().is_none());
    assert!(
        Field::parse_instruction(r"LINK Excel.Sheet.8 \p")
            .link_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"LINK Excel.Sheet.8 source \f")
            .link_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"LINK Excel.Sheet.8 source \f invalid")
            .link_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"LINK Excel.Sheet.8 source \p unexpected")
            .link_field()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("LINKAGE Excel.Sheet.8 source").field_type,
        FieldType::Unknown
    );
}

#[test]
fn external_include_fields_expose_stored_metadata_without_resolution() {
    let mut include_text = Field::parse_instruction(
        r#"INCLUDETEXT "missing source.xml" Summary \! \c Word8 \e utf-8 \m application/xml \n "xmlns:a=\"resume-schema\"" \t "file:///C:/display.xsl" \x a:Resume/a:Name \* MERGEFORMAT"#,
    );
    include_text.result = Cow::Borrowed("cached text");
    include_text.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    include_text.owner = FieldOwner::Body;
    include_text.position = 4;

    assert_eq!(include_text.field_type, FieldType::IncludeText);
    let text = include_text.external_include().unwrap();
    assert_eq!(text.kind(), IncludeFieldKind::Text);
    assert_eq!(text.source(), "missing source.xml");
    assert_eq!(text.bookmark(), Some("Summary"));
    assert_eq!(text.converter(), Some("Word8"));
    assert_eq!(
        text.options(),
        &[
            ExternalIncludeOption::Converter(Cow::Borrowed("Word8")),
            ExternalIncludeOption::Encoding(Cow::Borrowed("utf-8")),
            ExternalIncludeOption::MimeType(Cow::Borrowed("application/xml")),
            ExternalIncludeOption::NamespaceMapping(Cow::Borrowed("xmlns:a=\"resume-schema\"")),
            ExternalIncludeOption::Xslt(Cow::Borrowed("file:///C:/display.xsl")),
            ExternalIncludeOption::XPath(Cow::Borrowed("a:Resume/a:Name")),
        ]
    );
    assert!(text.suppresses_nested_field_updates());
    assert!(!text.omits_picture_data());
    assert_eq!(text.cached_result(), Some("cached text"));
    assert!(text.is_dirty());
    assert!(text.is_locked());
    assert_eq!(text.owner(), FieldOwner::Body);
    assert_eq!(text.position(), 4);
    assert_eq!(text.unknown_switches().len(), 1);
    assert_eq!(text.unknown_switches()[0].name, "*");
    assert_eq!(
        text.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    let unc_source = Field::parse_instruction(r#"INCLUDETEXT "\\server\\share\\source.docx""#);
    assert_eq!(
        unc_source.external_include().unwrap().source(),
        r"\server\share\source.docx"
    );

    let include_picture = Field::parse_instruction(
        r#"INCLUDEPICTURE "missing picture.gif" \c Pictim32 \d \* MERGEFORMAT"#,
    );
    assert_eq!(include_picture.field_type, FieldType::IncludePicture);
    let picture = include_picture.external_include().unwrap();
    assert_eq!(picture.kind(), IncludeFieldKind::Picture);
    assert_eq!(picture.source(), "missing picture.gif");
    assert_eq!(picture.bookmark(), None);
    assert_eq!(picture.converter(), Some("Pictim32"));
    assert_eq!(
        picture.options(),
        &[ExternalIncludeOption::Converter(Cow::Borrowed("Pictim32"))]
    );
    assert!(!picture.suppresses_nested_field_updates());
    assert!(picture.omits_picture_data());
    assert_eq!(picture.unknown_switches().len(), 1);
    assert_eq!(picture.unknown_switches()[0].name, "*");

    let legacy_text = Field::parse_instruction(r#"INCLUDE "missing legacy.docx" LegacySection \!"#);
    assert_eq!(legacy_text.field_type, FieldType::Include);
    let legacy_text = legacy_text.external_include().unwrap();
    assert_eq!(legacy_text.kind(), IncludeFieldKind::Text);
    assert_eq!(legacy_text.source(), "missing legacy.docx");
    assert_eq!(legacy_text.bookmark(), Some("LegacySection"));
    assert!(legacy_text.suppresses_nested_field_updates());

    let legacy_picture =
        Field::parse_instruction(r#"IMPORT "missing legacy.wmf" \c GraphicsFilter \d"#);
    assert_eq!(legacy_picture.field_type, FieldType::Import);
    let legacy_picture = legacy_picture.external_include().unwrap();
    assert_eq!(legacy_picture.kind(), IncludeFieldKind::Picture);
    assert_eq!(legacy_picture.source(), "missing legacy.wmf");
    assert_eq!(legacy_picture.converter(), Some("GraphicsFilter"));
    assert!(legacy_picture.omits_picture_data());

    assert!(
        Field::parse_instruction("INCLUDETEXT")
            .external_include()
            .is_none()
    );
    assert!(
        Field::parse_instruction("INCLUDETEXT \\c Word8")
            .external_include()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"INCLUDEPICTURE "picture.gif" Selector"#)
            .external_include()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"INCLUDEPICTURE "picture.gif" \d extra"#)
            .external_include()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"INCLUDETEXT source \e")
            .external_include()
            .is_none()
    );
    for instruction in [r#"INCLUDES "source.docx""#, r#"IMPORTS "picture.wmf""#] {
        let field = Field::parse_instruction(instruction);
        assert_eq!(field.field_type, FieldType::Unknown);
        assert!(field.external_include().is_none());
    }

    let mismatched = Field::new(
        FieldType::Include,
        Cow::Borrowed(r#"IMPORT "picture.wmf""#),
        Cow::Borrowed(""),
    );
    assert!(mismatched.external_include().is_none());
}

#[test]
fn table_of_contents_fields_preserve_stored_configuration_without_generation() {
    let mut field = Field::parse_instruction(
        r#"TOC \a Figure \b "Scope Bookmark" \c Table \d "/" \f A \h \l 1-3 \n "2-3" \o "1-4" \p " — " \s Figure \t "Custom,1,Appendix,2" \u \w \x \z \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached TOC");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Toc);
    let toc = field.table_of_contents().unwrap();
    assert_eq!(toc.instruction(), field.instruction);
    assert_eq!(
        toc.options(),
        &[
            TableOfContentsOption::CaptionWithoutLabel(Cow::Borrowed("Figure")),
            TableOfContentsOption::Bookmark(Cow::Borrowed("Scope Bookmark")),
            TableOfContentsOption::CaptionSequence(Cow::Borrowed("Table")),
            TableOfContentsOption::SequencePageSeparator(Cow::Borrowed("/")),
            TableOfContentsOption::TableEntryIdentifier(Cow::Borrowed("A")),
            TableOfContentsOption::Hyperlinks,
            TableOfContentsOption::TableEntryLevels(Cow::Borrowed("1-3")),
            TableOfContentsOption::OmitPageNumbers(Some(Cow::Borrowed("2-3"))),
            TableOfContentsOption::HeadingStyleRange(Some(Cow::Borrowed("1-4"))),
            TableOfContentsOption::EntryPageNumberSeparator(Cow::Borrowed(" — ")),
            TableOfContentsOption::SequenceIdentifier(Cow::Borrowed("Figure")),
            TableOfContentsOption::StyleMappings(Cow::Borrowed("Custom,1,Appendix,2")),
            TableOfContentsOption::OutlineLevels,
            TableOfContentsOption::PreserveTabs,
            TableOfContentsOption::PreserveNewlines,
            TableOfContentsOption::HidePageNumbersInWebView,
        ]
    );
    assert_eq!(toc.cached_result(), Some("cached TOC"));
    assert!(toc.is_dirty());
    assert!(toc.is_locked());
    assert_eq!(toc.owner(), FieldOwner::Body);
    assert_eq!(toc.position(), 4);
    assert_eq!(toc.unknown_switches().len(), 1);
    assert_eq!(toc.unknown_switches()[0].name, "*");
    assert_eq!(
        toc.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    let all_levels = Field::parse_instruction(r"TOC \n \o");
    assert_eq!(
        all_levels.table_of_contents().unwrap().options(),
        &[
            TableOfContentsOption::OmitPageNumbers(None),
            TableOfContentsOption::HeadingStyleRange(None),
        ]
    );

    assert!(
        Field::parse_instruction(r"TOC \a")
            .table_of_contents()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TOC \h unexpected")
            .table_of_contents()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TOC unexpected")
            .table_of_contents()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("TOCENTRIES").field_type,
        FieldType::Unknown
    );
}

#[test]
fn table_of_contents_entry_fields_preserve_stored_metadata_without_generation() {
    let mut field = Field::parse_instruction(r#"TC "Illustration 1" \f i \l 4 \n \* MERGEFORMAT"#);
    field.result = Cow::Borrowed("cached entry");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::TocEntry);
    let entry = field.table_of_contents_entry().unwrap();
    assert_eq!(entry.instruction(), field.instruction);
    assert_eq!(entry.entry(), "Illustration 1");
    assert_eq!(
        entry.options(),
        &[
            TableOfContentsEntryOption::ListIdentifier(Cow::Borrowed("i")),
            TableOfContentsEntryOption::Level(Cow::Borrowed("4")),
            TableOfContentsEntryOption::OmitPageNumber,
        ]
    );
    assert_eq!(entry.cached_result(), Some("cached entry"));
    assert!(entry.is_dirty());
    assert!(entry.is_locked());
    assert_eq!(entry.owner(), FieldOwner::Body);
    assert_eq!(entry.position(), 4);
    assert_eq!(entry.unknown_switches().len(), 1);
    assert_eq!(entry.unknown_switches()[0].name, "*");
    assert_eq!(
        entry.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(
        Field::parse_instruction("TC")
            .table_of_contents_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TC \f i")
            .table_of_contents_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TC entry unexpected")
            .table_of_contents_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TC entry \n unexpected")
            .table_of_contents_entry()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("TCC entry").field_type,
        FieldType::Unknown
    );
}

#[test]
fn table_of_authorities_entry_fields_preserve_stored_metadata_without_generation() {
    let mut field = Field::parse_instruction(
        r#"TA \l "Baldwin v. Alberti" \c 1 \s Baldwin \b \i \r PageRange \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached authority");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::TableOfAuthoritiesEntry);
    let entry = field.table_of_authorities_entry().unwrap();
    assert_eq!(entry.instruction(), field.instruction);
    assert_eq!(
        entry.options(),
        &[
            TableOfAuthoritiesEntryOption::LongCitation(Cow::Borrowed("Baldwin v. Alberti")),
            TableOfAuthoritiesEntryOption::Category(Cow::Borrowed("1")),
            TableOfAuthoritiesEntryOption::ShortCitation(Cow::Borrowed("Baldwin")),
            TableOfAuthoritiesEntryOption::BoldPageNumber,
            TableOfAuthoritiesEntryOption::ItalicPageNumber,
            TableOfAuthoritiesEntryOption::PageRangeBookmark(Cow::Borrowed("PageRange")),
        ]
    );
    assert_eq!(entry.cached_result(), Some("cached authority"));
    assert!(entry.is_dirty());
    assert!(entry.is_locked());
    assert_eq!(entry.owner(), FieldOwner::Body);
    assert_eq!(entry.position(), 4);
    assert_eq!(entry.unknown_switches().len(), 1);
    assert_eq!(entry.unknown_switches()[0].name, "*");
    assert_eq!(
        entry.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(
        Field::parse_instruction("TA")
            .table_of_authorities_entry()
            .is_some()
    );
    assert!(
        Field::parse_instruction(r"TA \c")
            .table_of_authorities_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TA \b unexpected")
            .table_of_authorities_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction("TA unexpected")
            .table_of_authorities_entry()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction(r"TAX \l Citation").field_type,
        FieldType::Unknown
    );
}

#[test]
fn table_of_authorities_fields_preserve_stored_configuration_without_generation() {
    let mut field = Field::parse_instruction(
        r#"TOA \b Authorities \c 2 \d "-" \e " — " \f \g "–" \h \l ", " \p \s Section \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached authorities");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::TableOfAuthorities);
    let toa = field.table_of_authorities().unwrap();
    assert_eq!(toa.instruction(), field.instruction);
    assert_eq!(
        toa.options(),
        &[
            TableOfAuthoritiesOption::Bookmark(Cow::Borrowed("Authorities")),
            TableOfAuthoritiesOption::Category(Cow::Borrowed("2")),
            TableOfAuthoritiesOption::SequencePageSeparator(Cow::Borrowed("-")),
            TableOfAuthoritiesOption::EntryPageNumberSeparator(Cow::Borrowed(" — ")),
            TableOfAuthoritiesOption::RemoveEntryFormatting,
            TableOfAuthoritiesOption::PageRangeSeparator(Cow::Borrowed("–")),
            TableOfAuthoritiesOption::CategoryHeadings,
            TableOfAuthoritiesOption::PageReferenceSeparator(Cow::Borrowed(", ")),
            TableOfAuthoritiesOption::UsePassim,
            TableOfAuthoritiesOption::SequenceIdentifier(Cow::Borrowed("Section")),
        ]
    );
    assert_eq!(toa.cached_result(), Some("cached authorities"));
    assert!(toa.is_dirty());
    assert!(toa.is_locked());
    assert_eq!(toa.owner(), FieldOwner::Body);
    assert_eq!(toa.position(), 4);
    assert_eq!(toa.unknown_switches().len(), 1);
    assert_eq!(toa.unknown_switches()[0].name, "*");
    assert_eq!(
        toa.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(
        Field::parse_instruction("TOA")
            .table_of_authorities()
            .is_some()
    );
    assert!(
        Field::parse_instruction(r"TOA \b")
            .table_of_authorities()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"TOA \f unexpected")
            .table_of_authorities()
            .is_none()
    );
    assert!(
        Field::parse_instruction("TOA unexpected")
            .table_of_authorities()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction(r"TOAX \c 2").field_type,
        FieldType::Unknown
    );
}

#[test]
fn index_fields_preserve_stored_configuration_without_generation() {
    let mut field = Field::parse_instruction(
        r#"INDEX \b Scope \c 2 \d "-" \e ", " \f Intro \g "–" \h "A Entries" \k ". " \l "; " \p A-C \r \s Figure \y \z 1033 \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached index");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Index);
    let index = field.index().unwrap();
    assert_eq!(index.instruction(), field.instruction);
    assert_eq!(
        index.options(),
        &[
            IndexOption::Bookmark(Cow::Borrowed("Scope")),
            IndexOption::Columns(Cow::Borrowed("2")),
            IndexOption::SequencePageSeparator(Cow::Borrowed("-")),
            IndexOption::EntryPageNumberSeparator(Cow::Borrowed(", ")),
            IndexOption::EntryType(Cow::Borrowed("Intro")),
            IndexOption::PageRangeSeparator(Cow::Borrowed("–")),
            IndexOption::Heading(Cow::Borrowed("A Entries")),
            IndexOption::CrossReferenceSeparator(Cow::Borrowed(". ")),
            IndexOption::PageNumberSeparator(Cow::Borrowed("; ")),
            IndexOption::LetterRange(Cow::Borrowed("A-C")),
            IndexOption::RunIn,
            IndexOption::SequenceIdentifier(Cow::Borrowed("Figure")),
            IndexOption::UseYomi,
            IndexOption::LanguageId(Cow::Borrowed("1033")),
        ]
    );
    assert_eq!(index.cached_result(), Some("cached index"));
    assert!(index.is_dirty());
    assert!(index.is_locked());
    assert_eq!(index.owner(), FieldOwner::Body);
    assert_eq!(index.position(), 4);
    assert_eq!(index.unknown_switches().len(), 1);
    assert_eq!(index.unknown_switches()[0].name, "*");
    assert_eq!(
        index.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(Field::parse_instruction("INDEX").index().is_some());
    assert!(Field::parse_instruction(r"INDEX \b").index().is_none());
    assert!(
        Field::parse_instruction(r"INDEX \r unexpected")
            .index()
            .is_none()
    );
    assert!(
        Field::parse_instruction("INDEX unexpected")
            .index()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction(r"INDEXES \c 2").field_type,
        FieldType::Unknown
    );
}

#[test]
fn index_entry_fields_preserve_stored_metadata_without_generation() {
    let mut field = Field::parse_instruction(
        r#"XE "Office Open XML:Syntax" \b \f Intro \i \r PageRange \t "See syntax" \y "Office" \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached entry");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::IndexEntry);
    let entry = field.index_entry().unwrap();
    assert_eq!(entry.instruction(), field.instruction);
    assert_eq!(entry.entry(), "Office Open XML:Syntax");
    assert_eq!(
        entry.options(),
        &[
            IndexEntryOption::BoldPageNumber,
            IndexEntryOption::EntryType(Cow::Borrowed("Intro")),
            IndexEntryOption::ItalicPageNumber,
            IndexEntryOption::PageRangeBookmark(Cow::Borrowed("PageRange")),
            IndexEntryOption::CrossReference(Cow::Borrowed("See syntax")),
            IndexEntryOption::Yomi(Cow::Borrowed("Office")),
        ]
    );
    assert_eq!(entry.cached_result(), Some("cached entry"));
    assert!(entry.is_dirty());
    assert!(entry.is_locked());
    assert_eq!(entry.owner(), FieldOwner::Body);
    assert_eq!(entry.position(), 4);
    assert_eq!(entry.unknown_switches().len(), 1);
    assert_eq!(entry.unknown_switches()[0].name, "*");
    assert_eq!(
        entry.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(Field::parse_instruction("XE").index_entry().is_none());
    assert!(
        Field::parse_instruction(r"XE \f Intro")
            .index_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"XE entry unexpected")
            .index_entry()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"XE entry \b unexpected")
            .index_entry()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction(r"XER entry").field_type,
        FieldType::Unknown
    );
}

#[test]
fn citation_fields_preserve_stored_metadata_without_resolving_sources() {
    let mut field = Field::parse_instruction(
        r#"CITATION Ecma01 \l 1033 \f "see " \s " (appendix)" \p 42 \v 2 \n \t \y \m Ecma02 \m Ecma03 \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached citation");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Citation);
    let citation = field.citation().unwrap();
    assert_eq!(citation.instruction(), field.instruction);
    assert_eq!(citation.source_tag(), "Ecma01");
    assert_eq!(
        citation.options(),
        &[
            CitationOption::LanguageId(Cow::Borrowed("1033")),
            CitationOption::Prefix(Cow::Borrowed("see ")),
            CitationOption::Suffix(Cow::Borrowed(" (appendix)")),
            CitationOption::PageNumber(Cow::Borrowed("42")),
            CitationOption::VolumeNumber(Cow::Borrowed("2")),
            CitationOption::SuppressAuthor,
            CitationOption::SuppressTitle,
            CitationOption::SuppressYear,
            CitationOption::AdditionalSourceTag(Cow::Borrowed("Ecma02")),
            CitationOption::AdditionalSourceTag(Cow::Borrowed("Ecma03")),
        ]
    );
    assert_eq!(citation.cached_result(), Some("cached citation"));
    assert!(citation.is_dirty());
    assert!(citation.is_locked());
    assert_eq!(citation.owner(), FieldOwner::Body);
    assert_eq!(citation.position(), 4);
    assert_eq!(citation.unknown_switches().len(), 1);
    assert_eq!(citation.unknown_switches()[0].name, "*");
    assert_eq!(
        citation.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(Field::parse_instruction("CITATION").citation().is_none());
    assert!(
        Field::parse_instruction(r"CITATION \l 1033")
            .citation()
            .is_none()
    );
    assert!(
        Field::parse_instruction("CITATION Ecma01 unexpected")
            .citation()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"CITATION Ecma01 \l")
            .citation()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"CITATION Ecma01 \n unexpected")
            .citation()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("CITATIONS Ecma01").field_type,
        FieldType::Unknown
    );
}

#[test]
fn bibliography_fields_preserve_stored_metadata_without_generation() {
    let mut field =
        Field::parse_instruction(r"BIBLIOGRAPHY \l 1033 \f en-US \m Ecma01 \* MERGEFORMAT");
    field.result = Cow::Borrowed("cached bibliography");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Bibliography);
    let bibliography = field.bibliography().unwrap();
    assert_eq!(bibliography.instruction(), field.instruction);
    assert_eq!(
        bibliography.options(),
        &[
            BibliographyOption::LanguageId(Cow::Borrowed("1033")),
            BibliographyOption::FilterLanguageId(Cow::Borrowed("en-US")),
            BibliographyOption::SourceTag(Cow::Borrowed("Ecma01")),
        ]
    );
    assert_eq!(bibliography.cached_result(), Some("cached bibliography"));
    assert!(bibliography.is_dirty());
    assert!(bibliography.is_locked());
    assert_eq!(bibliography.owner(), FieldOwner::Body);
    assert_eq!(bibliography.position(), 4);
    assert_eq!(bibliography.unknown_switches().len(), 1);
    assert_eq!(bibliography.unknown_switches()[0].name, "*");
    assert_eq!(
        bibliography.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(
        Field::parse_instruction("BIBLIOGRAPHY")
            .bibliography()
            .is_some()
    );
    assert!(
        Field::parse_instruction("BIBLIOGRAPHY unexpected")
            .bibliography()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"BIBLIOGRAPHY \f")
            .bibliography()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction(r"BIBLIOGRAPHIES \l 1033").field_type,
        FieldType::Unknown
    );
}

#[test]
fn document_variable_fields_preserve_names_without_resolution() {
    let mut field = Field::parse_instruction(r#"DOCVARIABLE "Customer Region" \* MERGEFORMAT"#);
    field.result = Cow::Borrowed("cached region");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::DocumentVariable);
    let variable = field.document_variable().unwrap();
    assert_eq!(variable.instruction(), field.instruction);
    assert_eq!(variable.variable_name(), "Customer Region");
    assert_eq!(variable.cached_result(), Some("cached region"));
    assert!(variable.is_dirty());
    assert!(variable.is_locked());
    assert_eq!(variable.owner(), FieldOwner::Body);
    assert_eq!(variable.position(), 4);
    assert_eq!(variable.unknown_switches().len(), 1);
    assert_eq!(variable.unknown_switches()[0].name, "*");
    assert_eq!(
        variable.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );

    assert!(
        Field::parse_instruction("DOCVARIABLE")
            .document_variable()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DOCVARIABLE \* MERGEFORMAT")
            .document_variable()
            .is_none()
    );
    assert!(
        Field::parse_instruction("DOCVARIABLE Customer unexpected")
            .document_variable()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("DOCVARIABLES Customer").field_type,
        FieldType::Unknown
    );
}

#[test]
fn document_property_fields_preserve_names_without_resolution() {
    let mut field =
        Field::parse_instruction(r#"DOCPROPERTY "Project Name" \* MERGEFORMAT \@ "MMMM d, yyyy""#);
    field.result = Cow::Borrowed("cached project");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::DocumentProperty);
    let property = field.document_property().unwrap();
    assert_eq!(property.instruction(), field.instruction);
    assert_eq!(property.property_name(), "Project Name");
    assert_eq!(property.cached_result(), Some("cached project"));
    assert!(property.is_dirty());
    assert!(property.is_locked());
    assert_eq!(property.owner(), FieldOwner::Body);
    assert_eq!(property.position(), 4);
    assert_eq!(property.switches().len(), 2);
    assert_eq!(property.switches()[0].name, "*");
    assert_eq!(property.switches()[0].value.as_deref(), Some("MERGEFORMAT"));
    assert_eq!(property.switches()[1].name, "@");
    assert_eq!(
        property.switches()[1].value.as_deref(),
        Some("MMMM d, yyyy")
    );

    assert!(
        Field::parse_instruction("DOCPROPERTY")
            .document_property()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DOCPROPERTY \")
            .document_property()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DOCPROPERTY \* MERGEFORMAT")
            .document_property()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"DOCPROPERTY """#)
            .document_property()
            .is_none()
    );
    assert!(
        Field::parse_instruction("DOCPROPERTY Project unexpected")
            .document_property()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DOCPROPERTY Project \")
            .document_property()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"DOCPROPERTY Project \* \")
            .document_property()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("DOCPROPERTYS Project").field_type,
        FieldType::Unknown
    );
    let too_long = Field::new(
        FieldType::DocumentProperty,
        Cow::Owned(format!("DOCPROPERTY {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.document_property().is_none());
}

#[test]
fn info_fields_preserve_stored_metadata_without_resolution_or_updates() {
    let mut field = Field::parse_instruction(
        r#"INFO TITLE "Stored title override" \* MERGEFORMAT \@ "opaque format""#,
    );
    field.result = Cow::Borrowed("cached title");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::Info);
    let information = field.info_field().unwrap();
    assert_eq!(information.instruction(), field.instruction);
    assert_eq!(information.information_type(), "TITLE");
    assert_eq!(information.new_value(), Some("Stored title override"));
    assert_eq!(information.cached_result(), Some("cached title"));
    assert!(information.is_dirty());
    assert!(information.is_locked());
    assert_eq!(information.owner(), FieldOwner::Body);
    assert_eq!(information.position(), 4);
    assert_eq!(information.switches().len(), 2);
    assert_eq!(information.switches()[0].name, "*");
    assert_eq!(
        information.switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(information.switches()[1].name, "@");
    assert_eq!(
        information.switches()[1].value.as_deref(),
        Some("opaque format")
    );

    let template = Field::parse_instruction("INFO TEMPLATE");
    assert_eq!(template.field_type, FieldType::Info);
    assert_eq!(
        template.info_field().unwrap().information_type(),
        "TEMPLATE"
    );
    assert_eq!(template.info_field().unwrap().new_value(), None);

    for instruction in [
        "INFO",
        r#"INFO "" "#,
        r#"INFO TITLE "Stored title" unexpected"#,
        r#"INFO TITLE "unterminated"#,
        r"INFO TITLE \",
    ] {
        assert!(
            Field::parse_instruction(instruction).info_field().is_none(),
            "{instruction}"
        );
    }

    assert_eq!(
        Field::parse_instruction("INFOS TITLE").field_type,
        FieldType::Unknown
    );
    assert!(
        Field::parse_instruction(r#"TITLE "Stored title override""#)
            .info_field()
            .is_none()
    );
    let too_long = Field::new(
        FieldType::Info,
        Cow::Owned(format!("INFO {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.info_field().is_none());
}

#[test]
fn document_information_fields_preserve_kinds_without_reading_or_calculating_values() {
    let mut field = Field::parse_instruction(r#"TITLE \* MERGEFORMAT \@ "opaque format""#);
    field.result = Cow::Borrowed("cached title");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::DocumentInformation);
    let information = field.document_information().unwrap();
    assert_eq!(information.instruction(), field.instruction);
    assert_eq!(information.kind(), DocumentInformationFieldKind::Title);
    assert_eq!(information.cached_result(), Some("cached title"));
    assert!(information.is_dirty());
    assert!(information.is_locked());
    assert_eq!(information.owner(), FieldOwner::Body);
    assert_eq!(information.position(), 4);
    assert_eq!(information.switches().len(), 2);
    assert_eq!(information.switches()[0].name, "*");
    assert_eq!(
        information.switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(information.switches()[1].name, "@");
    assert_eq!(
        information.switches()[1].value.as_deref(),
        Some("opaque format")
    );

    for (instruction, kind) in [
        ("TITLE", DocumentInformationFieldKind::Title),
        ("SUBJECT", DocumentInformationFieldKind::Subject),
        ("AUTHOR", DocumentInformationFieldKind::Author),
        ("KEYWORDS", DocumentInformationFieldKind::Keywords),
        ("COMMENTS", DocumentInformationFieldKind::Comments),
        ("LASTSAVEDBY", DocumentInformationFieldKind::LastSavedBy),
        ("CREATEDATE", DocumentInformationFieldKind::CreateDate),
        ("SAVEDATE", DocumentInformationFieldKind::SaveDate),
        ("PRINTDATE", DocumentInformationFieldKind::PrintDate),
        ("REVNUM", DocumentInformationFieldKind::RevisionNumber),
        ("EDITTIME", DocumentInformationFieldKind::EditTime),
        ("NUMPAGES", DocumentInformationFieldKind::NumberOfPages),
        ("NUMWORDS", DocumentInformationFieldKind::NumberOfWords),
        ("NUMCHARS", DocumentInformationFieldKind::NumberOfCharacters),
    ] {
        let field = Field::parse_instruction(instruction);
        assert_eq!(field.field_type, FieldType::DocumentInformation);
        let information = field.document_information().unwrap();
        assert_eq!(information.kind(), kind);
        assert_eq!(information.kind().field_keyword(), instruction);
        assert!(information.switches().is_empty());
    }

    assert!(
        Field::parse_instruction("TITLE unexpected")
            .document_information()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"AUTHOR "unterminated"#)
            .document_information()
            .is_none()
    );
    assert!(
        Field::parse_instruction("COMMENTS \\")
            .document_information()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"LASTSAVEDBY \* MERGEFORMAT unexpected")
            .document_information()
            .is_none()
    );
    assert!(
        Field::parse_instruction("NUMWORDS unexpected")
            .document_information()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("AUTHORS").field_type,
        FieldType::Unknown
    );
    let too_long = Field::new(
        FieldType::DocumentInformation,
        Cow::Owned(format!("TITLE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.document_information().is_none());
}

#[test]
fn document_context_fields_preserve_kinds_without_reading_or_calculating_values() {
    let mut field = Field::parse_instruction(r"FILENAME \p");
    field.result = Cow::Borrowed("cached file name");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::DocumentContext);
    let context = field.document_context().unwrap();
    assert_eq!(context.instruction(), field.instruction);
    assert_eq!(context.kind(), DocumentContextFieldKind::FileName);
    assert_eq!(context.cached_result(), Some("cached file name"));
    assert!(context.is_dirty());
    assert!(context.is_locked());
    assert_eq!(context.owner(), FieldOwner::Body);
    assert_eq!(context.position(), 4);
    assert_eq!(context.switches().len(), 1);
    assert_eq!(context.switches()[0].name, "p");
    assert_eq!(context.switches()[0].value, None);

    for (instruction, kind, field_type) in [
        (
            "FILENAME",
            DocumentContextFieldKind::FileName,
            FieldType::DocumentContext,
        ),
        (
            "TEMPLATE",
            DocumentContextFieldKind::Template,
            FieldType::DocumentContext,
        ),
        ("DATE", DocumentContextFieldKind::Date, FieldType::Date),
        ("TIME", DocumentContextFieldKind::Time, FieldType::Date),
        ("PAGE", DocumentContextFieldKind::Page, FieldType::Page),
        (
            "FILESIZE",
            DocumentContextFieldKind::FileSize,
            FieldType::DocumentContext,
        ),
        (
            "SECTION",
            DocumentContextFieldKind::Section,
            FieldType::DocumentContext,
        ),
        (
            "SECTIONPAGES",
            DocumentContextFieldKind::SectionPages,
            FieldType::DocumentContext,
        ),
    ] {
        let field = Field::parse_instruction(instruction);
        assert_eq!(field.field_type, field_type);
        let context = field.document_context().unwrap();
        assert_eq!(context.kind(), kind);
        assert_eq!(context.kind().field_keyword(), instruction);
        assert!(context.switches().is_empty());
    }

    assert!(
        Field::parse_instruction("FILENAME unexpected")
            .document_context()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"TEMPLATE "unterminated"#)
            .document_context()
            .is_none()
    );
    assert!(
        Field::parse_instruction("FILENAME \\")
            .document_context()
            .is_none()
    );
    assert!(
        Field::parse_instruction("PAGE unexpected")
            .document_context()
            .is_none()
    );
    assert!(
        Field::parse_instruction("SECTIONPAGES unexpected")
            .document_context()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("FILENAMES").field_type,
        FieldType::Unknown
    );
    assert_eq!(
        Field::parse_instruction("PAGES").field_type,
        FieldType::Unknown
    );
    assert_eq!(
        Field::parse_instruction("SECTIONPAGE").field_type,
        FieldType::Unknown
    );
    let too_long = Field::new(
        FieldType::DocumentContext,
        Cow::Owned(format!("FILENAME {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.document_context().is_none());
}

#[test]
fn merge_fields_preserve_names_without_merging() {
    let mut field = Field::parse_instruction(
        r#"MERGEFIELD "Customer Region" \b "Dear " \f "!" \* MERGEFORMAT"#,
    );
    field.result = Cow::Borrowed("cached customer");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::MergeField);
    let merge = field.merge_field().unwrap();
    assert_eq!(merge.instruction(), field.instruction);
    assert_eq!(merge.field_name(), "Customer Region");
    assert_eq!(merge.cached_result(), Some("cached customer"));
    assert!(merge.is_dirty());
    assert!(merge.is_locked());
    assert_eq!(merge.owner(), FieldOwner::Body);
    assert_eq!(merge.position(), 4);
    assert_eq!(merge.switches().len(), 3);
    assert_eq!(merge.switches()[0].name, "b");
    assert_eq!(merge.switches()[0].value.as_deref(), Some("Dear "));
    assert_eq!(merge.switches()[1].name, "f");
    assert_eq!(merge.switches()[1].value.as_deref(), Some("!"));
    assert_eq!(merge.switches()[2].name, "*");
    assert_eq!(merge.switches()[2].value.as_deref(), Some("MERGEFORMAT"));

    assert!(
        Field::parse_instruction("MERGEFIELD")
            .merge_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r#"MERGEFIELD \b "Dear ""#)
            .merge_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction("MERGEFIELD Customer unexpected")
            .merge_field()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("MERGEFIELDS Customer").field_type,
        FieldType::Unknown
    );
}

#[test]
fn mail_merge_data_fields_preserve_sources_without_connecting_or_merging() {
    let mut field = Field::parse_instruction(
        r#"DATA "recipients source.csv" "headers source.csv" \* MERGEFORMAT \q opaque"#,
    );
    field.result = Cow::Borrowed("cached mail-merge source");
    field.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    field.owner = FieldOwner::Body;
    field.position = 4;

    assert_eq!(field.field_type, FieldType::MailMergeData);
    let data = field.mail_merge_data().unwrap();
    assert_eq!(data.instruction(), field.instruction);
    assert_eq!(data.data_source(), "recipients source.csv");
    assert_eq!(data.header_source(), Some("headers source.csv"));
    assert_eq!(data.cached_result(), Some("cached mail-merge source"));
    assert!(data.is_dirty());
    assert!(data.is_locked());
    assert_eq!(data.owner(), FieldOwner::Body);
    assert_eq!(data.position(), 4);
    assert_eq!(data.switches().len(), 2);
    assert_eq!(data.switches()[0].name, "*");
    assert_eq!(data.switches()[0].value.as_deref(), Some("MERGEFORMAT"));
    assert_eq!(data.switches()[1].name, "q");
    assert_eq!(data.switches()[1].value.as_deref(), Some("opaque"));

    let without_header = Field::parse_instruction(r"data recipients.csv \q opaque");
    assert_eq!(without_header.field_type, FieldType::MailMergeData);
    let without_header = without_header.mail_merge_data().unwrap();
    assert_eq!(without_header.data_source(), "recipients.csv");
    assert_eq!(without_header.header_source(), None);
    assert_eq!(without_header.switches()[0].name, "q");
    assert_eq!(
        without_header.switches()[0].value.as_deref(),
        Some("opaque")
    );

    assert!(Field::parse_instruction("DATA").mail_merge_data().is_none());
    assert!(
        Field::parse_instruction(r"DATA \* MERGEFORMAT")
            .mail_merge_data()
            .is_none()
    );
    assert!(
        Field::parse_instruction("DATA recipients.csv headers.csv unexpected")
            .mail_merge_data()
            .is_none()
    );
    let database = Field::parse_instruction("DATABASE recipients.csv");
    assert_eq!(database.field_type, FieldType::Database);
    assert!(database.mail_merge_data().is_none());
    let too_long = Field::new(
        FieldType::MailMergeData,
        Cow::Owned(format!("DATA {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.mail_merge_data().is_none());
}

#[test]
fn database_fields_preserve_query_metadata_without_connecting_or_executing() {
    let mut database = Field::parse_instruction(
        r#"DATABASE \d "unavailable.csv" \c "DSN=NeverConnect" \s "SELECT * FROM Customers" \h"#,
    );
    database.result = Cow::Borrowed("cached database table");
    database.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    database.owner = FieldOwner::Body;
    database.position = 4;

    assert_eq!(database.field_type, FieldType::Database);
    let database_field = database.database_field().unwrap();
    assert_eq!(
        database_field.instruction(),
        r#"DATABASE \d "unavailable.csv" \c "DSN=NeverConnect" \s "SELECT * FROM Customers" \h"#
    );
    assert_eq!(
        database_field.opaque_instructions(),
        r#"\d "unavailable.csv" \c "DSN=NeverConnect" \s "SELECT * FROM Customers" \h"#
    );
    assert_eq!(
        database_field.cached_result(),
        Some("cached database table")
    );
    assert!(database_field.is_dirty());
    assert!(database_field.is_locked());
    assert_eq!(database_field.owner(), FieldOwner::Body);
    assert_eq!(database_field.position(), 4);

    let bare = Field::parse_instruction("database");
    assert_eq!(bare.field_type, FieldType::Database);
    assert_eq!(bare.database_field().unwrap().opaque_instructions(), "");

    let nonmatching = Field::parse_instruction("DATABASES");
    assert_eq!(nonmatching.field_type, FieldType::Unknown);
    assert!(nonmatching.database_field().is_none());

    let too_long = Field::new(
        FieldType::Database,
        Cow::Owned(format!("DATABASE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.database_field().is_none());
}

#[test]
fn referenced_document_fields_preserve_paths_without_opening_sources() {
    let mut reference = Field::parse_instruction(
        r#"RD "chapters/Chapter 1.doc" \f \* MERGEFORMAT \p "retained metadata""#,
    );
    reference.result = Cow::Borrowed("cached reference");
    reference.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    reference.owner = FieldOwner::Body;
    reference.position = 4;

    assert_eq!(reference.field_type, FieldType::ReferencedDocument);
    let reference = reference.referenced_document().unwrap();
    assert_eq!(reference.source(), "chapters/Chapter 1.doc");
    assert!(reference.uses_relative_path());
    assert_eq!(reference.cached_result(), Some("cached reference"));
    assert!(reference.is_dirty());
    assert!(reference.is_locked());
    assert_eq!(reference.owner(), FieldOwner::Body);
    assert_eq!(reference.position(), 4);
    assert_eq!(reference.switches().len(), 3);
    assert_eq!(reference.switches()[0].name, "f");
    assert_eq!(reference.switches()[0].value, None);
    assert_eq!(reference.switches()[1].name, "*");
    assert_eq!(
        reference.switches()[1].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(reference.switches()[2].name, "p");
    assert_eq!(
        reference.switches()[2].value.as_deref(),
        Some("retained metadata")
    );

    let absolute = Field::parse_instruction(r#"rd "archive.doc" \p"#);
    assert_eq!(absolute.field_type, FieldType::ReferencedDocument);
    let absolute = absolute.referenced_document().unwrap();
    assert_eq!(absolute.source(), "archive.doc");
    assert!(!absolute.uses_relative_path());
    assert_eq!(absolute.switches()[0].name, "p");

    for instruction in [
        "RD",
        r"RD \f",
        r#"RD "chapter.doc" \f unexpected"#,
        r#"RD "chapter.doc" \f \F"#,
        r#"RD "chapter.doc" unexpected"#,
        r#"RD "chapter.doc" \f """#,
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .referenced_document()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction("RDS chapter.doc").field_type,
        FieldType::Unknown
    );
    let wrong_type = Field::new(
        FieldType::Database,
        Cow::Borrowed(r#"RD "chapter.doc""#),
        Cow::Borrowed(""),
    );
    assert!(wrong_type.referenced_document().is_none());
    let too_long = Field::new(
        FieldType::ReferencedDocument,
        Cow::Owned(format!("RD {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.referenced_document().is_none());
}

#[test]
fn mail_merge_counters_preserve_cached_results_without_merging() {
    let mut record = Field::parse_instruction("MERGEREC");
    record.result = Cow::Borrowed("12");
    record.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    record.owner = FieldOwner::Body;
    record.position = 4;

    assert_eq!(record.field_type, FieldType::MergeRecord);
    let record_counter = record.mail_merge_counter().unwrap();
    assert_eq!(record_counter.instruction(), record.instruction);
    assert_eq!(record_counter.kind(), MailMergeCounterKind::Record);
    assert_eq!(record_counter.cached_result(), Some("12"));
    assert!(record_counter.is_dirty());
    assert!(record_counter.is_locked());
    assert_eq!(record_counter.owner(), FieldOwner::Body);
    assert_eq!(record_counter.position(), 4);

    let sequence = Field::parse_instruction("mergeSEQ");
    assert_eq!(sequence.field_type, FieldType::MergeSequence);
    assert_eq!(
        sequence.mail_merge_counter().unwrap().kind(),
        MailMergeCounterKind::Sequence
    );

    assert!(
        Field::parse_instruction("MERGEREC 12")
            .mail_merge_counter()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"MERGESEQ \* MERGEFORMAT")
            .mail_merge_counter()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("MERGERECORD").field_type,
        FieldType::Unknown
    );
}

#[test]
fn mail_merge_next_fields_preserve_cached_results_without_advancing_records() {
    let mut next = Field::parse_instruction("NEXT");
    next.result = Cow::Borrowed("cached next");
    next.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    next.owner = FieldOwner::Body;
    next.position = 4;

    assert_eq!(next.field_type, FieldType::MailMergeNext);
    let next_field = next.mail_merge_next().unwrap();
    assert_eq!(next_field.instruction(), next.instruction);
    assert_eq!(next_field.cached_result(), Some("cached next"));
    assert!(next_field.is_dirty());
    assert!(next_field.is_locked());
    assert_eq!(next_field.owner(), FieldOwner::Body);
    assert_eq!(next_field.position(), 4);

    assert!(
        Field::parse_instruction("NEXT 12")
            .mail_merge_next()
            .is_none()
    );
    assert!(
        Field::parse_instruction(r"NEXT \* MERGEFORMAT")
            .mail_merge_next()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("NEXTIF Customer = Ada").field_type,
        FieldType::MailMergeNextIf
    );
    assert!(
        Field::parse_instruction("NEXTIF Customer = Ada")
            .mail_merge_next()
            .is_none()
    );
}

#[test]
fn conditional_mail_merge_controls_preserve_cached_results_without_merging() {
    let mut next_if = Field::parse_instruction(r#"NEXTIF Customer = "Ada""#);
    next_if.result = Cow::Borrowed("cached nextif");
    next_if.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    next_if.owner = FieldOwner::Body;
    next_if.position = 4;

    assert_eq!(next_if.field_type, FieldType::MailMergeNextIf);
    let next_if_control = next_if.mail_merge_conditional_control().unwrap();
    assert_eq!(
        next_if_control.kind(),
        MailMergeConditionalControlKind::NextIf
    );
    assert_eq!(next_if_control.comparison(), r#"Customer = "Ada""#);
    assert_eq!(next_if_control.cached_result(), Some("cached nextif"));
    assert!(next_if_control.is_dirty());
    assert!(next_if_control.is_locked());
    assert_eq!(next_if_control.owner(), FieldOwner::Body);
    assert_eq!(next_if_control.position(), 4);

    let skip_if = Field::parse_instruction("skipif MERGEFIELD Order < 100");
    assert_eq!(skip_if.field_type, FieldType::MailMergeSkipIf);
    let skip_if_control = skip_if.mail_merge_conditional_control().unwrap();
    assert_eq!(
        skip_if_control.kind(),
        MailMergeConditionalControlKind::SkipIf
    );
    assert_eq!(skip_if_control.comparison(), "MERGEFIELD Order < 100");

    assert!(
        Field::parse_instruction("NEXTIF")
            .mail_merge_conditional_control()
            .is_none()
    );
    assert!(
        Field::parse_instruction("SKIPIF   ")
            .mail_merge_conditional_control()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("NEXTIFF Customer = Ada").field_type,
        FieldType::Unknown
    );
}

#[test]
fn if_fields_preserve_cached_results_without_evaluation() {
    let mut conditional = Field::parse_instruction(r#"IF "A" = "A" "yes" "no""#);
    conditional.result = Cow::Borrowed("yes");
    conditional.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    conditional.owner = FieldOwner::Body;
    conditional.position = 4;

    assert_eq!(conditional.field_type, FieldType::If);
    let if_field = conditional.if_field().unwrap();
    assert_eq!(if_field.instruction(), conditional.instruction);
    assert_eq!(if_field.expression(), r#""A" = "A" "yes" "no""#);
    assert_eq!(if_field.cached_result(), Some("yes"));
    assert!(if_field.is_dirty());
    assert!(if_field.is_locked());
    assert_eq!(if_field.owner(), FieldOwner::Body);
    assert_eq!(if_field.position(), 4);

    assert!(Field::parse_instruction("IF").if_field().is_none());
    assert_eq!(
        Field::parse_instruction(r#"IFF "A" = "A" "yes" "no""#).field_type,
        FieldType::Unknown
    );
}

#[test]
fn compare_fields_preserve_cached_results_without_evaluation() {
    let mut comparison = Field::parse_instruction(r#"COMPARE "CustomerNumber" >= 4"#);
    comparison.result = Cow::Borrowed("1");
    comparison.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    comparison.owner = FieldOwner::Body;
    comparison.position = 4;

    assert_eq!(comparison.field_type, FieldType::Compare);
    let compare_field = comparison.compare_field().unwrap();
    assert_eq!(compare_field.instruction(), comparison.instruction);
    assert_eq!(compare_field.comparison(), r#""CustomerNumber" >= 4"#);
    assert_eq!(compare_field.cached_result(), Some("1"));
    assert!(compare_field.is_dirty());
    assert!(compare_field.is_locked());
    assert_eq!(compare_field.owner(), FieldOwner::Body);
    assert_eq!(compare_field.position(), 4);

    let nested = Field::parse_instruction("compare MERGEFIELD CustomerRating <= 9");
    let compare_field = nested.compare_field().unwrap();
    assert_eq!(compare_field.comparison(), "MERGEFIELD CustomerRating <= 9");

    assert!(
        Field::parse_instruction("COMPARE")
            .compare_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction("COMPARE   ")
            .compare_field()
            .is_none()
    );
    assert_eq!(
        Field::parse_instruction("COMPARES Customer = 1").field_type,
        FieldType::Unknown
    );
}

#[test]
fn set_fields_preserve_cached_results_without_evaluation_or_state_changes() {
    let mut set =
        Field::parse_instruction(r#"SET "Customer Region" "North America" \* MERGEFORMAT"#);
    set.result = Cow::Borrowed("cached region");
    set.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    set.owner = FieldOwner::Body;
    set.position = 4;

    assert_eq!(set.field_type, FieldType::Set);
    let set_field = set.set_field().unwrap();
    assert_eq!(set_field.instruction(), set.instruction);
    assert_eq!(set_field.target_name(), "Customer Region");
    assert_eq!(set_field.expression(), r#""North America" \* MERGEFORMAT"#);
    assert_eq!(set_field.cached_result(), Some("cached region"));
    assert!(set_field.is_dirty());
    assert!(set_field.is_locked());
    assert_eq!(set_field.owner(), FieldOwner::Body);
    assert_eq!(set_field.position(), 4);

    let formula = Field::parse_instruction("set Total =SUM(ABOVE) + 1");
    assert_eq!(formula.field_type, FieldType::Set);
    let formula = formula.set_field().unwrap();
    assert_eq!(formula.target_name(), "Total");
    assert_eq!(formula.expression(), "=SUM(ABOVE) + 1");

    assert_eq!(
        Field::parse_instruction("SETTINGS value").field_type,
        FieldType::Unknown
    );
    for instruction in [
        "SET",
        r#"SET "" value"#,
        "SET Target",
        "SET Target   ",
        r"SET \* value",
        r#"SET "Target"expression"#,
    ] {
        assert!(
            Field::parse_instruction(instruction).set_field().is_none(),
            "{instruction}"
        );
    }
}

#[test]
fn sequence_fields_preserve_metadata_without_bookmark_lookup_or_numbering() {
    let mut sequence = Field::parse_instruction(r"SEQ Figure FigureChapter \r 3 \* ARABIC");
    sequence.result = Cow::Borrowed("3");
    sequence.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    sequence.owner = FieldOwner::Body;
    sequence.position = 4;

    assert_eq!(sequence.field_type, FieldType::Sequence);
    let sequence_field = sequence.sequence_field().unwrap();
    assert_eq!(sequence_field.instruction(), sequence.instruction);
    assert_eq!(sequence_field.identifier(), "Figure");
    assert_eq!(sequence_field.bookmark(), Some("FigureChapter"));
    assert_eq!(sequence_field.tail(), r"\r 3 \* ARABIC");
    assert_eq!(sequence_field.cached_result(), Some("3"));
    assert!(sequence_field.is_dirty());
    assert!(sequence_field.is_locked());
    assert_eq!(sequence_field.owner(), FieldOwner::Body);
    assert_eq!(sequence_field.position(), 4);

    let table = Field::parse_instruction(r"seq Table \s 1 \* ROMAN");
    assert_eq!(table.field_type, FieldType::Sequence);
    let table = table.sequence_field().unwrap();
    assert_eq!(table.identifier(), "Table");
    assert_eq!(table.bookmark(), None);
    assert_eq!(table.tail(), r"\s 1 \* ROMAN");

    let bare = Field::parse_instruction("SEQ Footnote");
    let bare = bare.sequence_field().unwrap();
    assert_eq!(bare.identifier(), "Footnote");
    assert_eq!(bare.bookmark(), None);
    assert_eq!(bare.tail(), "");

    assert_eq!(
        Field::parse_instruction("SEQUENCE Figure").field_type,
        FieldType::Unknown
    );
    for instruction in [
        "SEQ",
        r#"SEQ ""#,
        r#"SEQ Figure ""#,
        r#"SEQ "Figure"Bookmark"#,
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .sequence_field()
                .is_none(),
            "{instruction}"
        );
    }
}

#[test]
fn formula_fields_preserve_cached_results_without_evaluation() {
    let mut formula = Field::parse_instruction(r"=SUM(ABOVE) \* MERGEFORMAT");
    formula.result = Cow::Borrowed("42");
    formula.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    formula.owner = FieldOwner::Body;
    formula.position = 4;

    assert_eq!(formula.field_type, FieldType::Formula);
    let formula_field = formula.formula_field().unwrap();
    assert_eq!(formula_field.instruction(), formula.instruction);
    assert_eq!(formula_field.formula(), r"SUM(ABOVE) \* MERGEFORMAT");
    assert_eq!(formula_field.cached_result(), Some("42"));
    assert!(formula_field.is_dirty());
    assert!(formula_field.is_locked());
    assert_eq!(formula_field.owner(), FieldOwner::Body);
    assert_eq!(formula_field.position(), 4);

    let conditional = Field::parse_instruction(r#"= IF(1 = 1, "yes", "no")"#);
    assert_eq!(conditional.field_type, FieldType::Formula);
    let conditional = conditional.formula_field().unwrap();
    assert_eq!(conditional.formula(), r#"IF(1 = 1, "yes", "no")"#);

    let missing = Field::parse_instruction("=");
    assert_eq!(missing.field_type, FieldType::Formula);
    assert!(missing.formula_field().is_none());
    assert_eq!(
        Field::parse_instruction("EQUAL 1 + 1").field_type,
        FieldType::Unknown
    );
}

#[test]
fn quote_fields_preserve_cached_text_without_inserting_or_transforming_it() {
    let mut quote = Field::parse_instruction(r#"QUOTE "Stored literal" \* MERGEFORMAT \# "000""#);
    quote.result = Cow::Borrowed("cached literal");
    quote.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    quote.owner = FieldOwner::Body;
    quote.position = 4;

    assert_eq!(quote.field_type, FieldType::Quote);
    let quote_field = quote.quote_field().unwrap();
    assert_eq!(quote_field.instruction(), quote.instruction);
    assert_eq!(quote_field.text(), "Stored literal");
    assert_eq!(quote_field.cached_result(), Some("cached literal"));
    assert!(quote_field.is_dirty());
    assert!(quote_field.is_locked());
    assert_eq!(quote_field.owner(), FieldOwner::Body);
    assert_eq!(quote_field.position(), 4);
    assert_eq!(quote_field.switches().len(), 2);
    assert_eq!(quote_field.switches()[0].name, "*");
    assert_eq!(
        quote_field.switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(quote_field.switches()[1].name, "#");
    assert_eq!(quote_field.switches()[1].value.as_deref(), Some("000"));

    let unquoted = Field::parse_instruction(r#"quote CompatibilityText \@ "MMMM""#);
    assert_eq!(unquoted.field_type, FieldType::Quote);
    let unquoted = unquoted.quote_field().unwrap();
    assert_eq!(unquoted.text(), "CompatibilityText");
    assert_eq!(unquoted.switches()[0].name, "@");
    assert_eq!(unquoted.switches()[0].value.as_deref(), Some("MMMM"));

    for instruction in [
        "QUOTE",
        r"QUOTE \* MERGEFORMAT",
        r#"QUOTE "literal" unexpected"#,
        r#"QUOTE "unterminated"#,
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .quote_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction(r#"QUOTEY "not a quote field""#).field_type,
        FieldType::Unknown
    );
}

#[test]
fn symbol_fields_preserve_cached_metadata_without_mapping_codes_or_inserting_glyphs() {
    let mut symbol = Field::parse_instruction(r#"SYMBOL 0xA9 \f "Symbol" \s 12 \u"#);
    symbol.result = Cow::Borrowed("cached copyright");
    symbol.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    symbol.owner = FieldOwner::Body;
    symbol.position = 4;

    assert_eq!(symbol.field_type, FieldType::Symbol);
    let symbol_field = symbol.symbol_field().unwrap();
    assert_eq!(symbol_field.instruction(), symbol.instruction);
    assert_eq!(symbol_field.character_argument(), "0xA9");
    assert_eq!(symbol_field.cached_result(), Some("cached copyright"));
    assert!(symbol_field.is_dirty());
    assert!(symbol_field.is_locked());
    assert_eq!(symbol_field.owner(), FieldOwner::Body);
    assert_eq!(symbol_field.position(), 4);
    assert_eq!(symbol_field.switches().len(), 3);
    assert_eq!(symbol_field.switches()[0].name, "f");
    assert_eq!(symbol_field.switches()[0].value.as_deref(), Some("Symbol"));
    assert_eq!(symbol_field.switches()[1].name, "s");
    assert_eq!(symbol_field.switches()[1].value.as_deref(), Some("12"));
    assert_eq!(symbol_field.switches()[2].name, "u");
    assert_eq!(symbol_field.switches()[2].value, None);

    let symbol = Field::parse_instruction(r"symbol 163 \a \h \j");
    assert_eq!(symbol.field_type, FieldType::Symbol);
    let symbol = symbol.symbol_field().unwrap();
    assert_eq!(symbol.character_argument(), "163");
    assert_eq!(symbol.switches()[0].name, "a");
    assert_eq!(symbol.switches()[1].name, "h");
    assert_eq!(symbol.switches()[2].name, "j");

    for instruction in [
        "SYMBOL",
        r#"SYMBOL \f "Symbol""#,
        r"SYMBOL 0xA9 unexpected",
        r#"SYMBOL 0xA9 \f "unterminated"#,
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .symbol_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction("SYMBOLS 163").field_type,
        FieldType::Unknown
    );
}

#[test]
fn automatic_number_fields_preserve_cached_metadata_without_calculating_numbers_or_layout() {
    let mut automatic = Field::parse_instruction(r#"AUTONUM \s "." \* MERGEFORMAT"#);
    automatic.result = Cow::Borrowed("7.");
    automatic.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    automatic.owner = FieldOwner::Body;
    automatic.position = 4;

    assert_eq!(automatic.field_type, FieldType::AutoNumber);
    let automatic = automatic.auto_number_field().unwrap();
    assert_eq!(automatic.kind(), AutoNumberFieldKind::AutoNum);
    assert_eq!(automatic.kind().field_keyword(), "AUTONUM");
    assert_eq!(automatic.cached_result(), Some("7."));
    assert!(automatic.is_dirty());
    assert!(automatic.is_locked());
    assert_eq!(automatic.owner(), FieldOwner::Body);
    assert_eq!(automatic.position(), 4);
    assert_eq!(automatic.switches().len(), 2);
    assert_eq!(automatic.switches()[0].name, "s");
    assert_eq!(automatic.switches()[0].value.as_deref(), Some("."));
    assert_eq!(automatic.switches()[1].name, "*");
    assert_eq!(
        automatic.switches()[1].value.as_deref(),
        Some("MERGEFORMAT")
    );

    let legal = Field::parse_instruction(r#"autonumlgl \e \s ")" "#);
    assert_eq!(legal.field_type, FieldType::AutoNumber);
    let legal = legal.auto_number_field().unwrap();
    assert_eq!(legal.kind(), AutoNumberFieldKind::AutoNumLegal);
    assert_eq!(legal.kind().field_keyword(), "AUTONUMLGL");
    assert_eq!(legal.switches()[0].name, "e");
    assert_eq!(legal.switches()[0].value, None);
    assert_eq!(legal.switches()[1].name, "s");
    assert_eq!(legal.switches()[1].value.as_deref(), Some(")"));

    let outline = Field::parse_instruction("AUTONUMOUT");
    assert_eq!(outline.field_type, FieldType::AutoNumber);
    let outline = outline.auto_number_field().unwrap();
    assert_eq!(outline.kind(), AutoNumberFieldKind::AutoNumOutline);
    assert_eq!(outline.kind().field_keyword(), "AUTONUMOUT");
    assert!(outline.switches().is_empty());

    for instruction in [
        "AUTONUM unexpected",
        r#"AUTONUMLGL \s "unterminated"#,
        "AUTONUMOUT \\",
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .auto_number_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction("AUTONUMS").field_type,
        FieldType::Unknown
    );
    let too_long = Field::new(
        FieldType::AutoNumber,
        Cow::Owned(format!("AUTONUM {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.auto_number_field().is_none());
}

#[test]
fn list_number_fields_preserve_cached_metadata_without_reading_lists_or_calculating_numbers() {
    let mut numbered = Field::parse_instruction(r"LISTNUM NumberDefault \l 6 \s 3 \* MERGEFORMAT");
    numbered.result = Cow::Borrowed("(iii)");
    numbered.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    numbered.owner = FieldOwner::Body;
    numbered.position = 4;

    assert_eq!(numbered.field_type, FieldType::ListNumber);
    let numbered = numbered.list_number_field().unwrap();
    assert_eq!(numbered.list_name(), Some("NumberDefault"));
    assert_eq!(numbered.cached_result(), Some("(iii)"));
    assert!(numbered.is_dirty());
    assert!(numbered.is_locked());
    assert_eq!(numbered.owner(), FieldOwner::Body);
    assert_eq!(numbered.position(), 4);
    assert_eq!(numbered.switches().len(), 3);
    assert_eq!(numbered.switches()[0].name, "l");
    assert_eq!(numbered.switches()[0].value.as_deref(), Some("6"));
    assert_eq!(numbered.switches()[1].name, "s");
    assert_eq!(numbered.switches()[1].value.as_deref(), Some("3"));
    assert_eq!(numbered.switches()[2].name, "*");
    assert_eq!(numbered.switches()[2].value.as_deref(), Some("MERGEFORMAT"));

    let outline = Field::parse_instruction(r#"listnum "Outline Default" \l 4"#);
    assert_eq!(outline.field_type, FieldType::ListNumber);
    let outline = outline.list_number_field().unwrap();
    assert_eq!(outline.list_name(), Some("Outline Default"));
    assert_eq!(outline.switches()[0].name, "l");
    assert_eq!(outline.switches()[0].value.as_deref(), Some("4"));

    let unnamed = Field::parse_instruction(r"LISTNUM \l 2");
    assert_eq!(unnamed.field_type, FieldType::ListNumber);
    let unnamed = unnamed.list_number_field().unwrap();
    assert_eq!(unnamed.list_name(), None);
    assert_eq!(unnamed.switches()[0].name, "l");
    assert_eq!(unnamed.switches()[0].value.as_deref(), Some("2"));

    for instruction in [
        "LISTNUM NumberDefault unexpected",
        r#"LISTNUM "unterminated"#,
        "LISTNUM \\",
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .list_number_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction("LISTNUMBER NumberDefault").field_type,
        FieldType::Unknown
    );
    let too_long = Field::new(
        FieldType::ListNumber,
        Cow::Owned(format!("LISTNUM {}", "x".repeat(MAX_INSTRUCTION_LEN))),
        Cow::Borrowed(""),
    );
    assert!(too_long.list_number_field().is_none());
}

#[test]
fn style_reference_fields_preserve_metadata_without_style_or_layout_resolution() {
    let mut style_reference = Field::parse_instruction(
        r#"STYLEREF "Heading 1" \l \n \p \r \t \w \* MERGEFORMAT \q opaque"#,
    );
    style_reference.result = Cow::Borrowed("Cached heading");
    style_reference.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    style_reference.owner = FieldOwner::Body;
    style_reference.position = 4;

    assert_eq!(style_reference.field_type, FieldType::StyleReference);
    let style_reference = style_reference.style_reference_field().unwrap();
    assert_eq!(
        style_reference.instruction(),
        r#"STYLEREF "Heading 1" \l \n \p \r \t \w \* MERGEFORMAT \q opaque"#
    );
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
    assert_eq!(style_reference.unknown_switches().len(), 2);
    assert_eq!(style_reference.unknown_switches()[0].name, "*");
    assert_eq!(
        style_reference.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(style_reference.unknown_switches()[1].name, "q");
    assert_eq!(
        style_reference.unknown_switches()[1].value.as_deref(),
        Some("opaque")
    );
    assert_eq!(style_reference.cached_result(), Some("Cached heading"));
    assert!(style_reference.is_dirty());
    assert!(style_reference.is_locked());
    assert_eq!(style_reference.owner(), FieldOwner::Body);
    assert_eq!(style_reference.position(), 4);

    let title = Field::parse_instruction(r"styleref Title \n");
    assert_eq!(title.field_type, FieldType::StyleReference);
    let title = title.style_reference_field().unwrap();
    assert_eq!(title.style_name(), "Title");
    assert_eq!(
        title.options(),
        &[StyleReferenceFieldOption::ParagraphNumber]
    );
    assert!(title.unknown_switches().is_empty());
    assert_eq!(title.cached_result(), None);

    assert_eq!(
        Field::parse_instruction("STYLEREFS Heading").field_type,
        FieldType::Unknown
    );
    for instruction in [
        "STYLEREF",
        r#"STYLEREF ""#,
        r"STYLEREF Heading \l unexpected",
        "STYLEREF Heading unexpected",
        r"STYLEREF Heading \",
        r#"STYLEREF Heading "unterminated"#,
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .style_reference_field()
                .is_none(),
            "{instruction}"
        );
    }
}

#[test]
fn prompt_fields_preserve_metadata_without_displaying_prompts() {
    let mut ask =
        Field::parse_instruction(r#"ASK AskResponse "What is your first name?" \d "" \o"#);
    ask.result = Cow::Borrowed("cached ask response");
    ask.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    ask.owner = FieldOwner::Body;
    ask.position = 4;

    assert_eq!(ask.field_type, FieldType::Ask);
    let ask = ask.prompt_field().unwrap();
    assert_eq!(ask.kind(), PromptFieldKind::Ask);
    assert_eq!(ask.bookmark(), Some("AskResponse"));
    assert_eq!(ask.prompt(), Some("What is your first name?"));
    assert_eq!(ask.default_response(), Some(""));
    assert!(ask.prompts_once_per_mail_merge());
    assert_eq!(ask.cached_result(), Some("cached ask response"));
    assert!(ask.is_dirty());
    assert!(ask.is_locked());
    assert_eq!(ask.owner(), FieldOwner::Body);
    assert_eq!(ask.position(), 4);

    let fill_in = Field::parse_instruction(r#"fillin "Enter appointment time" \d "09:00""#);
    assert_eq!(fill_in.field_type, FieldType::FillIn);
    let fill_in = fill_in.prompt_field().unwrap();
    assert_eq!(fill_in.kind(), PromptFieldKind::FillIn);
    assert_eq!(fill_in.bookmark(), None);
    assert_eq!(fill_in.prompt(), Some("Enter appointment time"));
    assert_eq!(fill_in.default_response(), Some("09:00"));
    assert!(!fill_in.prompts_once_per_mail_merge());

    let default_only = Field::parse_instruction(r#"FILLIN \d "recent response" \o"#);
    let default_only = default_only.prompt_field().unwrap();
    assert_eq!(default_only.prompt(), None);
    assert_eq!(default_only.default_response(), Some("recent response"));
    assert!(default_only.prompts_once_per_mail_merge());

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
        assert!(
            Field::parse_instruction(instruction)
                .prompt_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction(r#"ASKER Answer "Question""#).field_type,
        FieldType::Unknown
    );
}

#[test]
fn user_identity_fields_preserve_metadata_without_reading_host_identity() {
    let mut address = Field::parse_instruction(r#"USERADDRESS "10 Top Secret Lane" \* Upper"#);
    address.result = Cow::Borrowed("10 TOP SECRET LANE");
    address.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    address.owner = FieldOwner::Body;
    address.position = 4;

    assert_eq!(address.field_type, FieldType::UserAddress);
    let address = address.user_identity_field().unwrap();
    assert_eq!(address.kind(), UserIdentityFieldKind::Address);
    assert_eq!(address.override_value(), Some("10 Top Secret Lane"));
    assert_eq!(address.formatting(), Some(UserIdentityFormatting::Upper));
    assert_eq!(address.cached_result(), Some("10 TOP SECRET LANE"));
    assert!(address.is_dirty());
    assert!(address.is_locked());
    assert_eq!(address.owner(), FieldOwner::Body);
    assert_eq!(address.position(), 4);

    let initials = Field::parse_instruction(r"userinitials \* Lower");
    assert_eq!(initials.field_type, FieldType::UserInitials);
    let initials = initials.user_identity_field().unwrap();
    assert_eq!(initials.kind(), UserIdentityFieldKind::Initials);
    assert_eq!(initials.override_value(), None);
    assert_eq!(initials.formatting(), Some(UserIdentityFormatting::Lower));

    let name = Field::parse_instruction(r#"USERNAME "Ada Lovelace" \* FirstCap"#);
    assert_eq!(name.field_type, FieldType::UserName);
    let name = name.user_identity_field().unwrap();
    assert_eq!(name.kind(), UserIdentityFieldKind::Name);
    assert_eq!(name.override_value(), Some("Ada Lovelace"));
    assert_eq!(name.formatting(), Some(UserIdentityFormatting::FirstCap));

    for instruction in [
        "USERADDRESS \\*",
        "USERINITIALS \\* Title",
        "USERNAME \\* Upper \\* Lower",
        "USERNAME Ada \\l 1033",
        "USERADDRESS Ada Lovelace",
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .user_identity_field()
                .is_none(),
            "{instruction}"
        );
    }

    let blank_override = Field::parse_instruction(r#"USERNAME "" \* Caps"#);
    let blank_override = blank_override.user_identity_field().unwrap();
    assert_eq!(blank_override.override_value(), Some(""));
    assert_eq!(
        blank_override.formatting(),
        Some(UserIdentityFormatting::Caps)
    );
    assert_eq!(
        Field::parse_instruction("USERNAMES Ada").field_type,
        FieldType::Unknown
    );
}

#[test]
fn advance_fields_preserve_placement_metadata_without_changing_layout() {
    let mut advance =
        Field::parse_instruction(r#"ADVANCE \u 6 \d 12 \l 20 \r -4 \x 150 \y "72" \d -3"#);
    advance.result = Cow::Borrowed("cached placement");
    advance.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    advance.owner = FieldOwner::Body;
    advance.position = 4;

    assert_eq!(advance.field_type, FieldType::Advance);
    let advance = advance.advance_field().unwrap();
    let adjustments = advance
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
    assert_eq!(advance.cached_result(), Some("cached placement"));
    assert!(advance.is_dirty());
    assert!(advance.is_locked());
    assert_eq!(advance.owner(), FieldOwner::Body);
    assert_eq!(advance.position(), 4);

    let no_adjustments = Field::parse_instruction("aDvAnCe");
    assert_eq!(no_adjustments.field_type, FieldType::Advance);
    assert!(
        no_adjustments
            .advance_field()
            .unwrap()
            .adjustments()
            .is_empty()
    );

    for instruction in [
        r"ADVANCE \d",
        r"ADVANCE \z 10",
        r"ADVANCE \x 1.5",
        r"ADVANCE \u 9223372036854775808",
        "ADVANCE 12",
        r"ADVANCE \d 6 trailing",
    ] {
        assert!(
            Field::parse_instruction(instruction)
                .advance_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction(r"ADVANCER \u 6").field_type,
        FieldType::Unknown
    );
}

#[test]
fn mail_merge_recipient_fields_preserve_layout_metadata_without_merging() {
    let mut address = Field::parse_instruction(
        r#"ADDRESSBLOCK \c 2 \d \e "United States" \e Canada \f "<<_FIRST0_>> <<_LAST0_>>" \l 1033 \* MERGEFORMAT"#,
    );
    address.result = Cow::Borrowed("cached address");
    address.status = FieldStatus {
        dirty: true,
        locked: true,
        ..FieldStatus::default()
    };
    address.owner = FieldOwner::Body;
    address.position = 4;

    assert_eq!(address.field_type, FieldType::AddressBlock);
    let address = address.mail_merge_recipient_field().unwrap();
    assert_eq!(address.kind(), MailMergeRecipientFieldKind::AddressBlock);
    assert_eq!(
        address.country_inclusion(),
        Some(AddressBlockCountryInclusion::UnlessExcluded)
    );
    assert!(address.formats_using_recipient_country());
    let excluded = address
        .excluded_countries()
        .iter()
        .map(Cow::as_ref)
        .collect::<Vec<_>>();
    assert_eq!(excluded, vec!["United States", "Canada"]);
    assert_eq!(address.format_template(), Some("<<_FIRST0_>> <<_LAST0_>>"));
    assert_eq!(address.language(), Some("1033"));
    assert_eq!(address.greeting_fallback_text(), None);
    assert_eq!(address.unknown_switches().len(), 1);
    assert_eq!(address.unknown_switches()[0].name, "*");
    assert_eq!(
        address.unknown_switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(address.cached_result(), Some("cached address"));
    assert!(address.is_dirty());
    assert!(address.is_locked());
    assert_eq!(address.owner(), FieldOwner::Body);
    assert_eq!(address.position(), 4);

    let greeting = Field::parse_instruction(
        r#"greetingline \f "Dear <<_FIRST0_>>," \e "To Whom It May Concern" \l en-US"#,
    );
    assert_eq!(greeting.field_type, FieldType::GreetingLine);
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
        assert!(
            Field::parse_instruction(instruction)
                .mail_merge_recipient_field()
                .is_none(),
            "{instruction}"
        );
    }
    assert_eq!(
        Field::parse_instruction(r"ADDRESSBLOCKING \c 1").field_type,
        FieldType::Unknown
    );
}

#[test]
fn document_discovers_document_variable_fields_without_resolving_them() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DOCVARIABLE CustomerName \\* MERGEFORMAT}{\fldrslt cached customer}}After}",
    )
    .unwrap();

    let fields = document.document_variable_fields();
    assert_eq!(document.document_variable_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].variable_name(), "CustomerName");
    assert_eq!(fields[0].cached_result(), Some("cached customer"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert!(document.document_variables().is_empty());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_document_property_fields_without_resolving_them() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DOCPROPERTY "Project Name" \\* MERGEFORMAT \\@ "MMMM d, yyyy"}{\fldrslt cached project}}After}"#,
    )
    .unwrap();

    let fields = document.document_property_fields();
    assert_eq!(document.document_property_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].property_name(), "Project Name");
    assert_eq!(fields[0].cached_result(), Some("cached project"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches().len(), 2);
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(fields[0].switches()[1].name, "@");
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_info_fields_without_reading_or_modifying_properties() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst INFO TITLE "Stored title override" \\* MERGEFORMAT}{\fldrslt cached title}}Middle {\field\flddirty\fldlock{\*\fldinst info COMMENTS "Stored comment" \\@ "opaque format"}{\fldrslt cached comment}}After}"#,
    )
    .unwrap();

    let fields = document.info_fields();
    assert_eq!(document.info_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].information_type(), "TITLE");
    assert_eq!(fields[0].new_value(), Some("Stored title override"));
    assert_eq!(fields[0].cached_result(), Some("cached title"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(
        fields[0].switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(fields[1].information_type(), "COMMENTS");
    assert_eq!(fields[1].new_value(), Some("Stored comment"));
    assert_eq!(fields[1].cached_result(), Some("cached comment"));
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(fields[1].switches()[0].name, "@");
    assert_eq!(
        fields[1].switches()[0].value.as_deref(),
        Some("opaque format")
    );
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_document_information_fields_without_reading_or_calculating_values() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TITLE \\* MERGEFORMAT}{\fldrslt cached title}}Middle {\field{\*\fldinst author \\@ "opaque format"}{\fldrslt cached author}}After}"#,
    )
    .unwrap();

    let fields = document.document_information_fields();
    assert_eq!(document.document_information_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), DocumentInformationFieldKind::Title);
    assert_eq!(fields[0].cached_result(), Some("cached title"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(fields[1].kind(), DocumentInformationFieldKind::Author);
    assert_eq!(fields[1].cached_result(), Some("cached author"));
    assert_eq!(fields[1].switches()[0].name, "@");
    assert_eq!(
        fields[1].switches()[0].value.as_deref(),
        Some("opaque format")
    );
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_document_information_statistics_without_calculating() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi {\field\flddirty\fldlock{\*\fldinst NUMWORDS \\* MERGEFORMAT}{\fldrslt cached words}}}",
    )
    .unwrap();

    let fields = document.document_information_fields();
    assert_eq!(document.document_information_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(
        fields[0].kind(),
        DocumentInformationFieldKind::NumberOfWords
    );
    assert_eq!(fields[0].cached_result(), Some("cached words"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches().len(), 1);
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(
        fields[0].switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
}

#[test]
fn document_discovers_document_context_fields_without_reading_or_calculating_values() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst FILENAME \\p}{\fldrslt cached file name}}Middle {\field{\*\fldinst TEMPLATE \\* MERGEFORMAT}{\fldrslt cached template}}After}",
    )
    .unwrap();

    let fields = document.document_context_fields();
    assert_eq!(document.document_context_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), DocumentContextFieldKind::FileName);
    assert_eq!(fields[0].cached_result(), Some("cached file name"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "p");
    assert_eq!(fields[1].kind(), DocumentContextFieldKind::Template);
    assert_eq!(fields[1].cached_result(), Some("cached template"));
    assert_eq!(fields[1].switches()[0].name, "*");
    assert_eq!(
        fields[1].switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_document_context_page_fields_without_calculation() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi {\field\flddirty\fldlock{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt cached page}}}",
    )
    .unwrap();

    let fields = document.document_context_fields();
    assert_eq!(document.document_context_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].kind(), DocumentContextFieldKind::Page);
    assert_eq!(fields[0].cached_result(), Some("cached page"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(
        fields[0].switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
}

#[test]
fn document_discovers_runtime_context_fields_without_file_or_layout_reads() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi {\field\flddirty\fldlock{\*\fldinst FILESIZE \\* MERGEFORMAT}{\fldrslt cached size}}{\field\flddirty\fldlock{\*\fldinst SECTION \\* MERGEFORMAT}{\fldrslt cached section}}{\field\flddirty\fldlock{\*\fldinst SECTIONPAGES \\* MERGEFORMAT}{\fldrslt cached section pages}}}",
    )
    .unwrap();

    let fields = document.document_context_fields();
    assert_eq!(document.document_context_field_count(), 3);
    assert_eq!(fields.len(), 3);
    assert_eq!(fields[0].kind(), DocumentContextFieldKind::FileSize);
    assert_eq!(fields[0].cached_result(), Some("cached size"));
    assert_eq!(fields[1].kind(), DocumentContextFieldKind::Section);
    assert_eq!(fields[1].cached_result(), Some("cached section"));
    assert_eq!(fields[2].kind(), DocumentContextFieldKind::SectionPages);
    assert_eq!(fields[2].cached_result(), Some("cached section pages"));
    for field in fields {
        assert!(field.is_dirty());
        assert!(field.is_locked());
        assert_eq!(field.switches()[0].name, "*");
        assert_eq!(field.switches()[0].value.as_deref(), Some("MERGEFORMAT"));
    }
}

#[test]
fn document_discovers_quote_fields_without_inserting_or_transforming_text() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst QUOTE "Stored literal" \\* MERGEFORMAT}{\fldrslt cached literal}}After}"#,
    )
    .unwrap();

    let fields = document.quote_fields();
    assert_eq!(document.quote_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].text(), "Stored literal");
    assert_eq!(fields[0].cached_result(), Some("cached literal"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(
        fields[0].switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_symbol_fields_without_mapping_codes_or_inserting_glyphs() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst SYMBOL 0xA9 \\f "Symbol" \\s 12 \\u}{\fldrslt cached copyright}}After}"#,
    )
    .unwrap();

    let fields = document.symbol_fields();
    assert_eq!(document.symbol_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].character_argument(), "0xA9");
    assert_eq!(fields[0].cached_result(), Some("cached copyright"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches().len(), 3);
    assert_eq!(fields[0].switches()[0].name, "f");
    assert_eq!(fields[0].switches()[0].value.as_deref(), Some("Symbol"));
    assert_eq!(fields[0].switches()[1].name, "s");
    assert_eq!(fields[0].switches()[1].value.as_deref(), Some("12"));
    assert_eq!(fields[0].switches()[2].name, "u");
    assert_eq!(fields[0].switches()[2].value, None);
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_print_fields_without_interpreting_or_sending_printer_commands() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst PRINT "ESC&l1O"}{\fldrslt cached printer}}{\field\flddirty\fldlock{\*\fldinst print \\p 2 "0 0 moveto"}{\fldrslt cached PostScript}}After}"#,
    )
    .unwrap();

    let fields = document.print_fields();
    assert_eq!(document.print_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].printer_instructions(), r#""ESC&l1O""#);
    assert_eq!(fields[0].cached_result(), Some("cached printer"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].printer_instructions(), r#"\p 2 "0 0 moveto""#);
    assert_eq!(fields[1].cached_result(), Some("cached PostScript"));
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_embed_fields_without_loading_or_activating_objects() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst EMBED Excel.Sheet.12 \\* MERGEFORMAT}{\fldrslt cached worksheet object}}{\field\flddirty\fldlock{\*\fldinst embed "Equation.DSMT4" \\d}{\fldrslt cached equation object}}After}"#,
    )
    .unwrap();

    let fields = document.embed_fields();
    assert_eq!(document.embed_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(
        fields[0].object_instructions(),
        r"Excel.Sheet.12 \* MERGEFORMAT"
    );
    assert_eq!(fields[0].cached_result(), Some("cached worksheet object"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].object_instructions(), r#""Equation.DSMT4" \d"#);
    assert_eq!(fields[1].cached_result(), Some("cached equation object"));
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_barcode_fields_without_decoding_or_rendering() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst BARCODE "4901234567894" EAN13 \\h 1440}{\fldrslt cached EAN13 barcode}}{\field\flddirty\fldlock{\*\fldinst barcode "ABC-123" CODE39 \\d}{\fldrslt cached Code39 barcode}}After}"#,
    )
    .unwrap();

    let fields = document.barcode_fields();
    assert_eq!(document.barcode_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(
        fields[0].barcode_instructions(),
        r#""4901234567894" EAN13 \h 1440"#
    );
    assert_eq!(fields[0].cached_result(), Some("cached EAN13 barcode"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].barcode_instructions(), r#""ABC-123" CODE39 \d"#);
    assert_eq!(fields[1].cached_result(), Some("cached Code39 barcode"));
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_bidi_outline_fields_without_resolving_numbering_or_layout() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst BIDIOUTLINE \\* MERGEFORMAT}{\fldrslt cached bidi outline number}}{\field\flddirty\fldlock{\*\fldinst bidioutline}{\fldrslt cached bare bidi outline}}After}",
    )
    .unwrap();

    let fields = document.bidi_outline_fields();
    assert_eq!(document.bidi_outline_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(
        fields[0].cached_result(),
        Some("cached bidi outline number")
    );
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].opaque_instructions(), "");
    assert_eq!(fields[1].cached_result(), Some("cached bare bidi outline"));
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_shape_fields_without_linking_or_rendering_drawings() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst SHAPE \\* MERGEFORMAT}{\fldrslt cached drawing anchor}}{\field\flddirty\fldlock{\*\fldinst shape}{\fldrslt cached bare drawing anchor}}After}",
    )
    .unwrap();

    let fields = document.shape_fields();
    assert_eq!(document.shape_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(fields[0].cached_result(), Some("cached drawing anchor"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].opaque_instructions(), "");
    assert_eq!(
        fields[1].cached_result(),
        Some("cached bare drawing anchor")
    );
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_legacy_form_fields_without_filling_or_executing() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst FORMTEXT \\* MERGEFORMAT}{\fldrslt cached text field}}{\field\flddirty\fldlock{\*\fldinst formcheckbox}{\fldrslt cached checkbox}}{\field\flddirty\fldlock{\*\fldinst FORMDROPDOWN \\* MERGEFORMAT}{\fldrslt cached drop-down selection}}After}",
    )
    .unwrap();

    let fields = document.legacy_form_fields();
    assert_eq!(document.legacy_form_field_count(), 3);
    assert_eq!(fields.len(), 3);
    assert_eq!(fields[0].kind(), LegacyFormFieldKind::Text);
    assert_eq!(fields[0].opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(fields[0].cached_result(), Some("cached text field"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].kind(), LegacyFormFieldKind::CheckBox);
    assert_eq!(fields[1].opaque_instructions(), "");
    assert_eq!(fields[1].cached_result(), Some("cached checkbox"));
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(fields[2].kind(), LegacyFormFieldKind::DropDown);
    assert_eq!(fields[2].opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(
        fields[2].cached_result(),
        Some("cached drop-down selection")
    );
    assert!(fields[2].is_dirty());
    assert!(fields[2].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_private_fields_without_conversion_or_layout() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst PRIVATE \\* MERGEFORMAT}{\fldrslt opaque converter payload}}{\field\flddirty\fldlock{\*\fldinst private}{\fldrslt cached bare private payload}}After}",
    )
    .unwrap();

    let fields = document.private_fields();
    assert_eq!(document.private_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].opaque_instructions(), r"\* MERGEFORMAT");
    assert_eq!(fields[0].cached_result(), Some("opaque converter payload"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].opaque_instructions(), "");
    assert_eq!(
        fields[1].cached_result(),
        Some("cached bare private payload")
    );
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_automatic_number_fields_without_calculating_numbers_or_layout() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst AUTONUM \\s "." \\* MERGEFORMAT}{\fldrslt 7.}}{\field\flddirty\fldlock{\*\fldinst AUTONUMLGL \\e \\s ")" }{\fldrslt 2.4}}{\field{\*\fldinst AUTONUMOUT}{\fldrslt III}}After}"#,
    )
    .unwrap();

    let fields = document.auto_number_fields();
    assert_eq!(document.auto_number_field_count(), 3);
    assert_eq!(fields.len(), 3);
    assert_eq!(fields[0].kind(), AutoNumberFieldKind::AutoNum);
    assert_eq!(fields[0].cached_result(), Some("7."));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "s");
    assert_eq!(fields[0].switches()[0].value.as_deref(), Some("."));
    assert_eq!(fields[1].kind(), AutoNumberFieldKind::AutoNumLegal);
    assert_eq!(fields[1].cached_result(), Some("2.4"));
    assert_eq!(fields[1].switches()[0].name, "e");
    assert_eq!(fields[1].switches()[1].name, "s");
    assert_eq!(fields[1].switches()[1].value.as_deref(), Some(")"));
    assert_eq!(fields[2].kind(), AutoNumberFieldKind::AutoNumOutline);
    assert_eq!(fields[2].cached_result(), Some("III"));
    assert!(fields[2].switches().is_empty());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_list_number_fields_without_reading_lists_or_calculating_numbers() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst LISTNUM NumberDefault \\l 6 \\s 3 \\* MERGEFORMAT}{\fldrslt (iii)}}{\field\flddirty\fldlock{\*\fldinst LISTNUM "Outline Default" \\l 4}{\fldrslt c}}{\field{\*\fldinst LISTNUM \\l 2}{\fldrslt i}}After}"#,
    )
    .unwrap();

    let fields = document.list_number_fields();
    assert_eq!(document.list_number_field_count(), 3);
    assert_eq!(fields.len(), 3);
    assert_eq!(fields[0].list_name(), Some("NumberDefault"));
    assert_eq!(fields[0].cached_result(), Some("(iii)"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "l");
    assert_eq!(fields[0].switches()[0].value.as_deref(), Some("6"));
    assert_eq!(fields[0].switches()[1].name, "s");
    assert_eq!(fields[0].switches()[1].value.as_deref(), Some("3"));
    assert_eq!(fields[1].list_name(), Some("Outline Default"));
    assert_eq!(fields[1].cached_result(), Some("c"));
    assert_eq!(fields[1].switches()[0].name, "l");
    assert_eq!(fields[1].switches()[0].value.as_deref(), Some("4"));
    assert_eq!(fields[2].list_name(), None);
    assert_eq!(fields[2].cached_result(), Some("i"));
    assert_eq!(fields[2].switches()[0].name, "l");
    assert_eq!(fields[2].switches()[0].value.as_deref(), Some("2"));
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_merge_fields_without_opening_data_sources() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst MERGEFIELD "Customer Region" \\b "Dear " \\* MERGEFORMAT}{\fldrslt cached customer}}After}"#,
    )
    .unwrap();

    let fields = document.merge_fields();
    assert_eq!(document.merge_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].field_name(), "Customer Region");
    assert_eq!(fields[0].cached_result(), Some("cached customer"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "b");
    assert_eq!(fields[0].switches()[0].value.as_deref(), Some("Dear "));
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_mail_merge_data_fields_without_opening_sources_or_merging() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DATA "recipients source.csv" "headers source.csv" \\* MERGEFORMAT}{\fldrslt cached mail-merge source}}{\field{\*\fldinst data recipients.csv}{\fldrslt cached bare source}}After}"#,
    )
    .unwrap();

    let fields = document.mail_merge_data_fields();
    assert_eq!(document.mail_merge_data_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].data_source(), "recipients source.csv");
    assert_eq!(fields[0].header_source(), Some("headers source.csv"));
    assert_eq!(fields[0].cached_result(), Some("cached mail-merge source"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[0].switches()[0].name, "*");
    assert_eq!(
        fields[0].switches()[0].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(fields[1].data_source(), "recipients.csv");
    assert_eq!(fields[1].header_source(), None);
    assert_eq!(fields[1].cached_result(), Some("cached bare source"));
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_database_fields_without_connecting_or_executing() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DATABASE \\d "unavailable.csv" \\c "DSN=NeverConnect" \\s "SELECT * FROM Customers" \\h}{\fldrslt cached database table}}{\field\flddirty\fldlock{\*\fldinst database}{\fldrslt cached bare database table}}After}"#,
    )
    .unwrap();

    let fields = document.database_fields();
    assert_eq!(document.database_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(
        fields[0].opaque_instructions(),
        r#"\d "unavailable.csv" \c "DSN=NeverConnect" \s "SELECT * FROM Customers" \h"#
    );
    assert_eq!(fields[0].cached_result(), Some("cached database table"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].opaque_instructions(), "");
    assert_eq!(
        fields[1].cached_result(),
        Some("cached bare database table")
    );
    assert!(fields[1].is_dirty());
    assert!(fields[1].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_referenced_documents_without_opening_sources() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst RD "chapters/Chapter 1.doc" \\f \\* MERGEFORMAT}{\fldrslt cached relative reference}}Middle {\field{\*\fldinst rd "archive.doc" \\p}{\fldrslt cached absolute reference}}After}"#,
    )
    .unwrap();

    let references = document.referenced_documents();
    assert_eq!(document.referenced_document_count(), 2);
    assert_eq!(references.len(), 2);
    assert_eq!(references[0].source(), "chapters/Chapter 1.doc");
    assert!(references[0].uses_relative_path());
    assert_eq!(
        references[0].cached_result(),
        Some("cached relative reference")
    );
    assert!(references[0].is_dirty());
    assert!(references[0].is_locked());
    assert_eq!(references[0].switches()[0].name, "f");
    assert_eq!(references[0].switches()[0].value, None);
    assert_eq!(references[0].switches()[1].name, "*");
    assert_eq!(
        references[0].switches()[1].value.as_deref(),
        Some("MERGEFORMAT")
    );
    assert_eq!(references[1].source(), "archive.doc");
    assert!(!references[1].uses_relative_path());
    assert_eq!(
        references[1].cached_result(),
        Some("cached absolute reference")
    );
    assert_eq!(references[1].switches()[0].name, "p");
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_mail_merge_counters_without_merging() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst MERGEREC}{\fldrslt 12}}Middle {\field{\*\fldinst mergeSEQ}{\fldrslt 3}}After}",
    )
    .unwrap();

    let counters = document.mail_merge_counters();
    assert_eq!(document.mail_merge_counter_count(), 2);
    assert_eq!(counters.len(), 2);
    assert_eq!(counters[0].kind(), MailMergeCounterKind::Record);
    assert_eq!(counters[0].cached_result(), Some("12"));
    assert!(counters[0].is_dirty());
    assert!(counters[0].is_locked());
    assert_eq!(counters[1].kind(), MailMergeCounterKind::Sequence);
    assert_eq!(counters[1].cached_result(), Some("3"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_mail_merge_next_fields_without_advancing_records() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst NEXT}{\fldrslt cached next}}After}",
    )
    .unwrap();

    let fields = document.mail_merge_next_fields();
    assert_eq!(document.mail_merge_next_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].cached_result(), Some("cached next"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_conditional_mail_merge_controls_without_merging() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst NEXTIF Customer = "Ada"}{\fldrslt cached nextif}}Middle {\field{\*\fldinst skipif MERGEFIELD Order < 100}{\fldrslt cached skipif}}After}"#,
    )
    .unwrap();

    let controls = document.mail_merge_conditional_controls();
    assert_eq!(document.mail_merge_conditional_control_count(), 2);
    assert_eq!(controls.len(), 2);
    assert_eq!(controls[0].kind(), MailMergeConditionalControlKind::NextIf);
    assert_eq!(controls[0].comparison(), r#"Customer = "Ada""#);
    assert_eq!(controls[0].cached_result(), Some("cached nextif"));
    assert!(controls[0].is_dirty());
    assert!(controls[0].is_locked());
    assert_eq!(controls[1].kind(), MailMergeConditionalControlKind::SkipIf);
    assert_eq!(controls[1].comparison(), "MERGEFIELD Order < 100");
    assert_eq!(controls[1].cached_result(), Some("cached skipif"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_if_fields_without_evaluation() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst IF "A" = "A" "yes" "no"}{\fldrslt yes}}After}"#,
    )
    .unwrap();

    let fields = document.if_fields();
    assert_eq!(document.if_field_count(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].expression(), r#""A" = "A" "yes" "no""#);
    assert_eq!(fields[0].cached_result(), Some("yes"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_compare_fields_without_evaluation() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst COMPARE "CustomerNumber" >= 4}{\fldrslt 1}}Middle {\field{\*\fldinst compare MERGEFIELD CustomerRating <= 9}{\fldrslt 0}}After}"#,
    )
    .unwrap();

    let fields = document.compare_fields();
    assert_eq!(document.compare_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].comparison(), r#""CustomerNumber" >= 4"#);
    assert_eq!(fields[0].cached_result(), Some("1"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].comparison(), "MERGEFIELD CustomerRating <= 9");
    assert_eq!(fields[1].cached_result(), Some("0"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_prompt_fields_without_displaying_prompts() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ASK AskResponse "What is your first name?" \\d "" \\o}{\fldrslt cached ask response}}Middle {\field{\*\fldinst FILLIN "Enter appointment time" \\d "09:00"}{\fldrslt 10:30}}After}"#,
    )
    .unwrap();

    let fields = document.prompt_fields();
    assert_eq!(document.prompt_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), PromptFieldKind::Ask);
    assert_eq!(fields[0].bookmark(), Some("AskResponse"));
    assert_eq!(fields[0].prompt(), Some("What is your first name?"));
    assert_eq!(fields[0].default_response(), Some(""));
    assert!(fields[0].prompts_once_per_mail_merge());
    assert_eq!(fields[0].cached_result(), Some("cached ask response"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].kind(), PromptFieldKind::FillIn);
    assert_eq!(fields[1].bookmark(), None);
    assert_eq!(fields[1].prompt(), Some("Enter appointment time"));
    assert_eq!(fields[1].default_response(), Some("09:00"));
    assert_eq!(fields[1].cached_result(), Some("10:30"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_user_identity_fields_without_reading_host_identity() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst USERADDRESS "10 Top Secret Lane" \\* Upper}{\fldrslt 10 TOP SECRET LANE}}Middle {\field{\*\fldinst USERINITIALS \\* Lower}{\fldrslt dw}}After {\field{\*\fldinst USERNAME "Ada Lovelace" \\* FirstCap}{\fldrslt Ada Lovelace}}}"#,
    )
    .unwrap();

    let fields = document.user_identity_fields();
    assert_eq!(document.user_identity_field_count(), 3);
    assert_eq!(fields.len(), 3);
    assert_eq!(fields[0].kind(), UserIdentityFieldKind::Address);
    assert_eq!(fields[0].override_value(), Some("10 Top Secret Lane"));
    assert_eq!(fields[0].formatting(), Some(UserIdentityFormatting::Upper));
    assert_eq!(fields[0].cached_result(), Some("10 TOP SECRET LANE"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].kind(), UserIdentityFieldKind::Initials);
    assert_eq!(fields[1].override_value(), None);
    assert_eq!(fields[1].formatting(), Some(UserIdentityFormatting::Lower));
    assert_eq!(fields[1].cached_result(), Some("dw"));
    assert_eq!(fields[2].kind(), UserIdentityFieldKind::Name);
    assert_eq!(fields[2].override_value(), Some("Ada Lovelace"));
    assert_eq!(
        fields[2].formatting(),
        Some(UserIdentityFormatting::FirstCap)
    );
    assert_eq!(fields[2].cached_result(), Some("Ada Lovelace"));
    assert_eq!(document.text(), "Before Middle After ");
}

#[test]
fn document_discovers_advance_fields_without_changing_layout() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ADVANCE \\u 6 \\d 12 \\l 20 \\r -4 \\x 150 \\y 72}{\fldrslt cached placement}}After}",
    )
    .unwrap();

    let fields = document.advance_fields();
    assert_eq!(document.advance_field_count(), 1);
    assert_eq!(fields.len(), 1);
    let adjustments = fields[0]
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
        ]
    );
    assert_eq!(fields[0].cached_result(), Some("cached placement"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_recipient_fields_without_merging() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ADDRESSBLOCK \\c 2 \\d \\e "United States" \\e Canada \\f "<<_FIRST0_>> <<_LAST0_>>" \\l 1033}{\fldrslt cached address}}Middle {\field{\*\fldinst GREETINGLINE \\f "Dear <<_FIRST0_>>," \\e "To Whom It May Concern" \\l en-US}{\fldrslt Dear Ada,}}After}"#,
    )
    .unwrap();

    let fields = document.mail_merge_recipient_fields();
    assert_eq!(document.mail_merge_recipient_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), MailMergeRecipientFieldKind::AddressBlock);
    assert_eq!(
        fields[0].country_inclusion(),
        Some(AddressBlockCountryInclusion::UnlessExcluded)
    );
    assert!(fields[0].formats_using_recipient_country());
    assert_eq!(fields[0].language(), Some("1033"));
    assert_eq!(fields[0].cached_result(), Some("cached address"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].kind(), MailMergeRecipientFieldKind::GreetingLine);
    assert_eq!(
        fields[1].greeting_fallback_text(),
        Some("To Whom It May Concern")
    );
    assert_eq!(fields[1].cached_result(), Some("Dear Ada,"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_bibliography_fields_without_loading_sources() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst CITATION Ecma01 \\l 1033 \\n \\m Ecma02}{\fldrslt cached citation}}Middle {\field{\*\fldinst BIBLIOGRAPHY \\l 1033 \\f en-US \\m Ecma01}{\fldrslt cached bibliography}}After}",
    )
    .unwrap();

    let citations = document.citations();
    assert_eq!(document.citation_count(), 1);
    assert_eq!(citations.len(), 1);
    assert_eq!(citations[0].source_tag(), "Ecma01");
    assert_eq!(
        citations[0].options(),
        &[
            CitationOption::LanguageId(Cow::Borrowed("1033")),
            CitationOption::SuppressAuthor,
            CitationOption::AdditionalSourceTag(Cow::Borrowed("Ecma02")),
        ]
    );
    assert_eq!(citations[0].cached_result(), Some("cached citation"));
    assert!(citations[0].is_dirty());
    assert!(citations[0].is_locked());

    let bibliographies = document.bibliographies();
    assert_eq!(document.bibliography_count(), 1);
    assert_eq!(bibliographies.len(), 1);
    assert_eq!(
        bibliographies[0].options(),
        &[
            BibliographyOption::LanguageId(Cow::Borrowed("1033")),
            BibliographyOption::FilterLanguageId(Cow::Borrowed("en-US")),
            BibliographyOption::SourceTag(Cow::Borrowed("Ecma01")),
        ]
    );
    assert_eq!(
        bibliographies[0].cached_result(),
        Some("cached bibliography")
    );
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_eq_fields_without_calculating_them() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field{\*\fldinst EQ \\f(1,2)}{\fldrslt }}After}",
    )
    .unwrap();

    let equations = document.equations();
    assert_eq!(document.equation_count(), 1);
    assert_eq!(equations.len(), 1);
    assert_eq!(equations[0].expression(), r"\f(1,2)");
    assert_eq!(equations[0].cached_result(), None);
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_dde_fields_without_starting_conversations() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty{\*\fldinst DDE Excel "missing.xlsx" "Sheet1!A1" \\a \\p}{\fldrslt cached DDE}}Middle {\field{\*\fldinst DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \\t}{\fldrslt cached auto}}After}"#,
    )
    .unwrap();

    let links = document.dde_links();
    assert_eq!(document.dde_link_count(), 2);
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].kind(), DdeFieldKind::Dde);
    assert_eq!(links[0].application(), "Excel");
    assert_eq!(links[0].source(), "missing.xlsx");
    assert_eq!(links[0].item(), Some("Sheet1!A1"));
    assert!(links[0].requests_automatic_updates());
    assert_eq!(links[0].representation(), Some(DdeRepresentation::Picture));
    assert_eq!(links[0].cached_result(), Some("cached DDE"));
    assert!(links[0].is_dirty());
    assert_eq!(links[1].kind(), DdeFieldKind::DdeAuto);
    assert_eq!(links[1].item(), Some("Sheet1!A2"));
    assert_eq!(links[1].representation(), Some(DdeRepresentation::Text));
    assert_eq!(links[1].cached_result(), Some("cached auto"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_link_fields_without_activating_ole_servers() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst LINK Excel.Sheet.8 "missing.xlsx" "Sheet1!A1" \\a \\f 4 \\p}{\fldrslt cached LINK}}Middle {\field{\*\fldinst LINK Word.Document.8 "missing.docx" Bookmark \\t}{\fldrslt cached text}}After}"#,
    )
    .unwrap();

    let links = document.link_fields();
    assert_eq!(document.link_field_count(), 2);
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].application_type(), "Excel.Sheet.8");
    assert_eq!(links[0].source(), "missing.xlsx");
    assert_eq!(links[0].item(), Some("Sheet1!A1"));
    assert!(links[0].requests_automatic_updates());
    assert_eq!(
        links[0].formatting_modes(),
        &[LinkFormatting::SpreadsheetSource]
    );
    assert_eq!(
        links[0].effective_result_option(),
        Some(LinkResultOption::Picture)
    );
    assert_eq!(links[0].cached_result(), Some("cached LINK"));
    assert!(links[0].is_dirty());
    assert!(links[0].is_locked());
    assert_eq!(links[1].application_type(), "Word.Document.8");
    assert_eq!(links[1].item(), Some("Bookmark"));
    assert_eq!(
        links[1].effective_result_option(),
        Some(LinkResultOption::Text)
    );
    assert_eq!(links[1].cached_result(), Some("cached text"));
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_external_includes_without_opening_sources() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty{\*\fldinst INCLUDETEXT "missing.docx" Summary \\!}{\fldrslt cached text}}Middle {\field{\*\fldinst INCLUDEPICTURE "missing.gif" \\d}{\fldrslt cached picture}}Legacy {\field{\*\fldinst INCLUDE "legacy.docx" LegacySection \\!}{\fldrslt cached legacy text}}Older {\field{\*\fldinst IMPORT "legacy.wmf" \\c GraphicsFilter \\d}{\fldrslt cached legacy picture}}After}"#,
    )
    .unwrap();

    let includes = document.external_includes();
    assert_eq!(document.external_include_count(), 4);
    assert_eq!(includes.len(), 4);
    assert_eq!(includes[0].kind(), IncludeFieldKind::Text);
    assert_eq!(includes[0].source(), "missing.docx");
    assert_eq!(includes[0].bookmark(), Some("Summary"));
    assert!(includes[0].suppresses_nested_field_updates());
    assert!(includes[0].is_dirty());
    assert_eq!(includes[1].kind(), IncludeFieldKind::Picture);
    assert_eq!(includes[1].source(), "missing.gif");
    assert!(includes[1].omits_picture_data());
    assert_eq!(includes[2].kind(), IncludeFieldKind::Text);
    assert_eq!(includes[2].source(), "legacy.docx");
    assert_eq!(includes[2].bookmark(), Some("LegacySection"));
    assert!(includes[2].suppresses_nested_field_updates());
    assert_eq!(includes[3].kind(), IncludeFieldKind::Picture);
    assert_eq!(includes[3].source(), "legacy.wmf");
    assert_eq!(includes[3].converter(), Some("GraphicsFilter"));
    assert!(includes[3].omits_picture_data());
    assert_eq!(document.text(), "Before Middle Legacy Older After");
}

#[test]
fn document_discovers_table_of_contents_without_regenerating_it() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TOC \\o "1-3" \\h \\z}{\fldrslt cached TOC}}After}"#,
    )
    .unwrap();

    let tables = document.table_of_contents();
    assert_eq!(document.table_of_contents_count(), 1);
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].cached_result(), Some("cached TOC"));
    assert!(tables[0].is_dirty());
    assert!(tables[0].is_locked());
    assert_eq!(
        tables[0].options(),
        &[
            TableOfContentsOption::HeadingStyleRange(Some(Cow::Borrowed("1-3"))),
            TableOfContentsOption::Hyperlinks,
            TableOfContentsOption::HidePageNumbersInWebView,
        ]
    );
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_table_of_contents_entries_without_generating_a_table() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TC "Illustration 1" \\f i \\l 4 \\n}{\fldrslt cached entry}}After}"#,
    )
    .unwrap();

    let entries = document.table_of_contents_entries();
    assert_eq!(document.table_of_contents_entry_count(), 1);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].entry(), "Illustration 1");
    assert_eq!(entries[0].cached_result(), Some("cached entry"));
    assert!(entries[0].is_dirty());
    assert!(entries[0].is_locked());
    assert_eq!(
        entries[0].options(),
        &[
            TableOfContentsEntryOption::ListIdentifier(Cow::Borrowed("i")),
            TableOfContentsEntryOption::Level(Cow::Borrowed("4")),
            TableOfContentsEntryOption::OmitPageNumber,
        ]
    );
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_table_of_authorities_entries_without_generating_a_table() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TA \\l "Baldwin v. Alberti" \\c 1 \\b}{\fldrslt cached authority}}After}"#,
    )
    .unwrap();

    let entries = document.table_of_authorities_entries();
    assert_eq!(document.table_of_authorities_entry_count(), 1);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].cached_result(), Some("cached authority"));
    assert!(entries[0].is_dirty());
    assert!(entries[0].is_locked());
    assert_eq!(
        entries[0].options(),
        &[
            TableOfAuthoritiesEntryOption::LongCitation(Cow::Borrowed("Baldwin v. Alberti")),
            TableOfAuthoritiesEntryOption::Category(Cow::Borrowed("1")),
            TableOfAuthoritiesEntryOption::BoldPageNumber,
        ]
    );
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_table_of_authorities_without_generating_it() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TOA \\b Authorities \\c 2 \\f \\h \\p}{\fldrslt cached authorities}}After}",
    )
    .unwrap();

    let tables = document.tables_of_authorities();
    assert_eq!(document.table_of_authorities_count(), 1);
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].cached_result(), Some("cached authorities"));
    assert!(tables[0].is_dirty());
    assert!(tables[0].is_locked());
    assert_eq!(
        tables[0].options(),
        &[
            TableOfAuthoritiesOption::Bookmark(Cow::Borrowed("Authorities")),
            TableOfAuthoritiesOption::Category(Cow::Borrowed("2")),
            TableOfAuthoritiesOption::RemoveEntryFormatting,
            TableOfAuthoritiesOption::CategoryHeadings,
            TableOfAuthoritiesOption::UsePassim,
        ]
    );
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_index_fields_without_generating_an_index() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst INDEX \\f Intro \\r}{\fldrslt cached index}}Middle {\field\flddirty{\*\fldinst XE "syntax:fields" \\f Intro \\t "See references"}{\fldrslt cached entry}}After}"#,
    )
    .unwrap();

    let indexes = document.indexes();
    assert_eq!(document.index_count(), 1);
    assert_eq!(indexes.len(), 1);
    assert_eq!(indexes[0].cached_result(), Some("cached index"));
    assert!(indexes[0].is_dirty());
    assert!(indexes[0].is_locked());
    assert_eq!(
        indexes[0].options(),
        &[
            IndexOption::EntryType(Cow::Borrowed("Intro")),
            IndexOption::RunIn,
        ]
    );

    let entries = document.index_entries();
    assert_eq!(document.index_entry_count(), 1);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].entry(), "syntax:fields");
    assert_eq!(entries[0].cached_result(), Some("cached entry"));
    assert!(entries[0].is_dirty());
    assert!(!entries[0].is_locked());
    assert_eq!(
        entries[0].options(),
        &[
            IndexEntryOption::EntryType(Cow::Borrowed("Intro")),
            IndexEntryOption::CrossReference(Cow::Borrowed("See references")),
        ]
    );
    assert_eq!(document.text(), "Before Middle After");
}

#[test]
fn document_discovers_macro_buttons_without_invoking_them() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst MACROBUTTON NoMacro Click here}{\fldrslt Click here}}After}",
    )
    .unwrap();

    let macro_buttons = document.macro_buttons();
    assert_eq!(document.macro_button_count(), 1);
    assert_eq!(macro_buttons.len(), 1);
    assert_eq!(macro_buttons[0].macro_name(), "NoMacro");
    assert_eq!(macro_buttons[0].display_text(), Some("Click here"));
    assert_eq!(macro_buttons[0].cached_result(), Some("Click here"));
    assert!(macro_buttons[0].is_dirty());
    assert!(macro_buttons[0].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_go_to_buttons_without_activating_them() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst GOTOBUTTON MyBookmark "Jump here"}{\fldrslt cached button}}After}"#,
    )
    .unwrap();

    let buttons = document.go_to_buttons();
    assert_eq!(document.go_to_button_count(), 1);
    assert_eq!(buttons.len(), 1);
    assert_eq!(buttons[0].target(), "MyBookmark");
    assert_eq!(buttons[0].button_text(), "Jump here");
    assert_eq!(buttons[0].cached_result(), Some("cached button"));
    assert!(buttons[0].is_dirty());
    assert!(buttons[0].is_locked());
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_active_content_fields_without_activation() {
    let document = crate::RtfDocument::parse(
        r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ADDIN opaque-add-in-data}{\fldrslt cached add-in result}}{\field{\*\fldinst CONTROL opaque-ocx-metadata}{\fldrslt cached control result}}{\field{\*\fldinst HTMLCONTROL opaque-html-control-metadata}{\fldrslt cached html result}}After}",
    )
    .unwrap();

    let fields = document.active_content_fields();
    assert_eq!(document.active_content_field_count(), 3);
    assert_eq!(fields.len(), 3);
    assert_eq!(fields[0].kind(), ActiveContentFieldKind::AddIn);
    assert_eq!(fields[0].cached_result(), Some("cached add-in result"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].kind(), ActiveContentFieldKind::OcxControl);
    assert_eq!(fields[1].cached_result(), Some("cached control result"));
    assert_eq!(fields[2].kind(), ActiveContentFieldKind::HtmlControl);
    assert_eq!(fields[2].cached_result(), Some("cached html result"));
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_auto_text_fields_without_lookup_or_insertion() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst GLOSSARY "Legacy Clause"}{\fldrslt cached glossary entry}}{\field{\*\fldinst AUTOTEXT ReusableClause}{\fldrslt cached auto text entry}}After}"#,
    )
    .unwrap();

    let fields = document.auto_text_fields();
    assert_eq!(document.auto_text_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), AutoTextFieldKind::Glossary);
    assert_eq!(fields[0].entry_name(), "Legacy Clause");
    assert_eq!(fields[0].cached_result(), Some("cached glossary entry"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].kind(), AutoTextFieldKind::AutoText);
    assert_eq!(fields[1].entry_name(), "ReusableClause");
    assert_eq!(fields[1].cached_result(), Some("cached auto text entry"));
    assert_eq!(document.text(), "Before After");
}

#[test]
fn document_discovers_auto_text_list_fields_without_selection_or_insertion() {
    let document = crate::RtfDocument::parse(
        r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst AUTOTEXTLIST "Choose a name" \\s "Name Style" \\t "Right-click to select" \\* MERGEFORMAT \\q opaque}{\fldrslt cached selection}}{\field{\*\fldinst AUTOTEXTLIST \\s NameStyle}{\fldrslt cached style-only selection}}After}"#,
    )
    .unwrap();

    let fields = document.auto_text_list_fields();
    assert_eq!(document.auto_text_list_field_count(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].display_text(), Some("Choose a name"));
    assert_eq!(
        fields[0].options(),
        &[
            AutoTextListOption::Style(Cow::Borrowed("Name Style")),
            AutoTextListOption::Tip(Cow::Borrowed("Right-click to select")),
        ]
    );
    assert_eq!(fields[0].unknown_switches().len(), 2);
    assert_eq!(fields[0].cached_result(), Some("cached selection"));
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].display_text(), None);
    assert_eq!(
        fields[1].options(),
        &[AutoTextListOption::Style(Cow::Borrowed("NameStyle"))]
    );
    assert_eq!(
        fields[1].cached_result(),
        Some("cached style-only selection")
    );
    assert_eq!(document.text(), "Before After");
}

#[test]
fn parses_libreoffice_internal_hyperlink_fixtures() {
    let fixture_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/rtf");
    for (fixture, expected) in [
        ("fdo86750.rtf", "anchor"),
        ("tdf134614_toc_indent.rtf", "_Toc1"),
    ] {
        let document =
            crate::RtfDocument::from_bytes(&std::fs::read(fixture_root.join(fixture)).unwrap())
                .unwrap();
        assert!(
            document.fields().iter().any(|field| {
                field.extract_bookmark().as_deref() == Some(expected)
                    && field.extract_url().as_deref() == Some(format!("#{expected}").as_str())
            }),
            "fixture {fixture} fields: {:?}",
            document.fields()
        );
    }

    let formatted =
        crate::RtfDocument::from_bytes(&std::fs::read(fixture_root.join("fdo82071.rtf")).unwrap())
            .unwrap();
    assert!(formatted.fields().iter().any(|field| matches!(
        field.parsed_code(),
        ParsedFieldCode::PageReference(ReferenceCode { ref bookmark, hyperlink: true, .. })
            if bookmark == "_Toc363816075"
    )));

    let backslashes = crate::RtfDocument::from_bytes(
        &std::fs::read(fixture_root.join("hyperlink-with-backslashes.rtf")).unwrap(),
    )
    .unwrap();
    assert_eq!(
        backslashes.fields()[0].extract_url().as_deref(),
        Some(r"c:\temp\doc1.doc")
    );

    let target = crate::RtfDocument::from_bytes(
        &std::fs::read(fixture_root.join("hyperlink-target.rtf")).unwrap(),
    )
    .unwrap();
    let ParsedFieldCode::Hyperlink(code) = target.fields()[0].parsed_code() else {
        panic!("expected target-frame hyperlink");
    };
    assert_eq!(code.target_frame.as_deref(), Some("_blank"));
}
