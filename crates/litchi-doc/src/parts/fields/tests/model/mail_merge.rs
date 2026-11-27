use super::super::super::model::*;

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
