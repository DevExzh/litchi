use litchi_rtf::{DocumentWord2003Compatibility, RtfDocument, RtfWriter};

const CONTROLS: [&str; 14] = [
    "noafcnsttbl",
    "noindnmbrts",
    "felnbrelev",
    "indrlsweleven",
    "nocxsptable",
    "notcvasp",
    "notvatxbx",
    "spltpgpar",
    "hwelev",
    "afelev",
    "cachedcolbal",
    "utinl",
    "notbrkcnstfrctbl",
    "krnprsnet",
];

fn all_enabled() -> DocumentWord2003Compatibility {
    DocumentWord2003Compatibility {
        preserve_autofit_table_width_around_shapes: true,
        use_hanging_indent_as_numbering_tab: true,
        use_legacy_kinsoku_characters: true,
        use_legacy_floating_object_indentation: true,
        allow_contextual_spacing_in_tables: true,
        ignore_cell_vertical_alignment_with_floating_objects: true,
        ignore_text_box_vertical_alignment: true,
        split_page_break_paragraph: true,
        use_fixed_width_hangul: true,
        use_legacy_autofit_width_expansion: true,
        use_cached_column_balancing: true,
        underline_numbering_suffix: true,
        do_not_split_rows_around_floating_tables: true,
        use_ansi_kerning_pairs: true,
    }
}

fn write(doc: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(doc).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_complete_word_2003_matrix_as_passive_metadata() {
    let controls = CONTROLS
        .iter()
        .map(|name| format!("\\{name}"))
        .collect::<String>();
    let doc = RtfDocument::parse(&format!(r"{{\rtf1\ansi{controls} Body}}")).unwrap();
    assert_eq!(*doc.word_2003_compatibility(), all_enabled());
    assert_eq!(doc.text(), "Body");
    assert!(doc.tables().is_empty());
    assert!(doc.shapes().is_empty());
    assert!(doc.notes().is_empty());
}

#[test]
fn omission_is_empty_and_serializes_no_matrix_flags() {
    let doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    assert!(doc.word_2003_compatibility().is_empty());
    let serialized = write(&doc);
    for name in CONTROLS {
        assert!(!serialized.contains(&format!("\\{name}")));
    }
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_inertly() {
    let mut doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    doc.set_word_2003_compatibility(all_enabled());
    let serialized = write(&doc);
    let positions: Vec<_> = CONTROLS
        .iter()
        .map(|name| serialized.find(&format!("\\{name}")).unwrap())
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(*reparsed.word_2003_compatibility(), all_enabled());
    assert_eq!(reparsed.text(), "Body");

    doc.clear_word_2003_compatibility();
    assert!(doc.word_2003_compatibility().is_empty());
    assert_eq!(doc.text(), "Body");
}

#[test]
fn accepts_every_flag_independently_without_layout_side_effects() {
    for name in CONTROLS {
        let doc = RtfDocument::parse(&format!(r"{{\rtf1\ansi\{name} Body}}")).unwrap();
        assert_eq!(doc.text(), "Body");
        assert!(doc.tables().is_empty());
        assert!(doc.shapes().is_empty());
    }
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for name in CONTROLS {
        for input in [
            format!(r"{{\rtf1\ansi\{name}0 Body}}"),
            format!(r"{{\rtf1\ansi\{name}\{name} Body}}"),
            format!(r"{{\rtf1\ansi{{\*\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi{{\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi Body\{name}}}"),
        ] {
            assert!(RtfDocument::parse(&input).is_err(), "accepted {input}");
        }
    }
    assert!(RtfDocument::parse(r"{\rtf1\ansi\krnprsnet999999999999 Body}").is_err());
}

#[test]
fn parses_bundled_word_2003_producer_matrix_in_arbitrary_source_order() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/hidden-linebreaks.rtf"
    );
    let doc = RtfDocument::parse_bytes(&std::fs::read(path).unwrap()).unwrap();
    assert_eq!(*doc.word_2003_compatibility(), all_enabled());
    let serialized = write(&doc);
    let positions: Vec<_> = CONTROLS
        .iter()
        .map(|name| serialized.find(&format!("\\{name}")).unwrap())
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
}
