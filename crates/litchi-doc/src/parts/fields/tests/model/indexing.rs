use super::super::super::codec::*;
use super::super::super::model::*;

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
