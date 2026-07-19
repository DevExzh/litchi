use litchi_rtf::{DocumentTableLayoutCompatibility, RtfDocument, RtfWriter};

const CONTROLS: [&str; 8] = [
    "otblrul",
    "alntblind",
    "lytcalctblwd",
    "lyttblrtgr",
    "nolnhtadjtbl",
    "nobrkwrptbl",
    "nogrowautofit",
    "newtblstyruls",
];

fn write(doc: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(doc).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_complete_cluster_as_passive_metadata() {
    let rtf = concat!(
        r"{\rtf1\ansi\otblrul\alntblind\lytcalctblwd\lyttblrtgr",
        r"\nolnhtadjtbl\nobrkwrptbl\nogrowautofit\newtblstyruls Body}"
    );
    let doc = RtfDocument::parse(rtf).unwrap();
    let compatibility = doc.table_layout_compatibility();
    assert!(compatibility.combine_borders_like_word_5);
    assert!(compatibility.do_not_align_rows_independently);
    assert!(compatibility.do_not_use_raw_table_width);
    assert!(compatibility.keep_rows_together);
    assert!(compatibility.do_not_adjust_line_height);
    assert!(compatibility.do_not_break_wrapped_tables_across_pages);
    assert!(compatibility.prevent_autofit_growth_into_margins);
    assert!(compatibility.use_word_2003_table_style_rules);
    assert_eq!(doc.text(), "Body");
    assert!(doc.tables().is_empty());
}

#[test]
fn omission_is_empty_and_serializes_no_table_policy() {
    let doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    assert!(doc.table_layout_compatibility().is_empty());
    let serialized = write(&doc);
    for name in CONTROLS {
        assert!(!serialized.contains(&format!("\\{name}")));
    }
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_inertly() {
    let mut doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    doc.set_table_layout_compatibility(DocumentTableLayoutCompatibility {
        combine_borders_like_word_5: true,
        do_not_align_rows_independently: true,
        do_not_use_raw_table_width: true,
        keep_rows_together: true,
        do_not_adjust_line_height: true,
        do_not_break_wrapped_tables_across_pages: true,
        prevent_autofit_growth_into_margins: true,
        use_word_2003_table_style_rules: true,
    });

    let serialized = write(&doc);
    let positions: Vec<_> = CONTROLS
        .iter()
        .map(|name| serialized.find(&format!("\\{name}")).unwrap())
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(
        *reparsed.table_layout_compatibility(),
        *doc.table_layout_compatibility()
    );
    assert_eq!(reparsed.text(), "Body");

    doc.clear_table_layout_compatibility();
    assert!(doc.table_layout_compatibility().is_empty());
    assert_eq!(doc.text(), "Body");
}

#[test]
fn accepts_independent_flags_without_synthesizing_table_layout() {
    for name in CONTROLS {
        let doc = RtfDocument::parse(&format!(r"{{\rtf1\ansi\{name} Body}}")).unwrap();
        assert_eq!(doc.text(), "Body");
        assert!(doc.tables().is_empty());
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
    assert!(RtfDocument::parse(r"{\rtf1\ansi\otblrul999999999999 Body}").is_err());
}

#[test]
fn parses_bundled_libreoffice_table_compatibility_preamble() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf161878.rtf"
    );
    let bytes = std::fs::read(path).unwrap();
    let doc = RtfDocument::parse_bytes(&bytes).unwrap();
    let compatibility = doc.table_layout_compatibility();
    assert!(!compatibility.combine_borders_like_word_5);
    assert!(compatibility.do_not_align_rows_independently);
    assert!(compatibility.do_not_use_raw_table_width);
    assert!(compatibility.keep_rows_together);
    assert!(compatibility.do_not_adjust_line_height);
    assert!(compatibility.do_not_break_wrapped_tables_across_pages);
    assert!(compatibility.prevent_autofit_growth_into_margins);
    assert!(compatibility.use_word_2003_table_style_rules);
}
