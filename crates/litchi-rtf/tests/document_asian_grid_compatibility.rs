use litchi_rtf::{DocumentAsianGridCompatibility, RtfDocument, RtfWriter};

const CONTROLS: [&str; 5] = [
    "ApplyBrkRules",
    "snaptogridincell",
    "wrppunct",
    "asianbrkrule",
    "toplinepunct",
];

fn write(doc: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(doc).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_complete_cluster_as_passive_metadata() {
    let doc = RtfDocument::parse(concat!(
        r"{\rtf1\ansi\ApplyBrkRules\snaptogridincell\wrppunct",
        r"\asianbrkrule\toplinepunct Body}"
    ))
    .unwrap();
    let compatibility = doc.asian_grid_compatibility();
    assert!(compatibility.apply_thai_line_breaking_rules);
    assert!(compatibility.snap_text_to_grid_inside_table);
    assert!(compatibility.allow_hanging_punctuation);
    assert!(compatibility.use_asian_line_breaking_rules);
    assert!(compatibility.compress_punctuation_at_line_start);
    assert_eq!(doc.text(), "Body");
    assert!(doc.tables().is_empty());
}

#[test]
fn omission_and_wrong_case_remain_empty() {
    let omitted = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    assert!(omitted.asian_grid_compatibility().is_empty());
    for wrong_case in ["applybrkrules", "APPLYBRKRULES", "Applybrkrules"] {
        let doc = RtfDocument::parse(&format!(r"{{\rtf1\ansi\{wrong_case} Body}}")).unwrap();
        assert!(doc.asian_grid_compatibility().is_empty());
        assert_eq!(doc.text(), "Body");
    }
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_inertly() {
    let mut doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    doc.set_asian_grid_compatibility(DocumentAsianGridCompatibility {
        apply_thai_line_breaking_rules: true,
        snap_text_to_grid_inside_table: true,
        allow_hanging_punctuation: true,
        use_asian_line_breaking_rules: true,
        compress_punctuation_at_line_start: true,
    });

    let serialized = write(&doc);
    let positions: Vec<_> = CONTROLS
        .iter()
        .map(|name| serialized.find(&format!("\\{name}")).unwrap())
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    assert!(!serialized.contains("\\tx"));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(
        *reparsed.asian_grid_compatibility(),
        *doc.asian_grid_compatibility()
    );
    assert_eq!(reparsed.text(), "Body");

    doc.clear_asian_grid_compatibility();
    assert!(doc.asian_grid_compatibility().is_empty());
    assert_eq!(doc.text(), "Body");
}

#[test]
fn accepts_each_flag_independently_without_layout_side_effects() {
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
    assert!(RtfDocument::parse(r"{\rtf1\ansi\toplinepunct999999999999 Body}").is_err());
}

#[test]
fn parses_bundled_libreoffice_asian_grid_flags() {
    let common_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf161878.rtf"
    );
    let common = RtfDocument::parse_bytes(&std::fs::read(common_path).unwrap()).unwrap();
    let compatibility = common.asian_grid_compatibility();
    assert!(compatibility.snap_text_to_grid_inside_table);
    assert!(compatibility.allow_hanging_punctuation);
    assert!(compatibility.use_asian_line_breaking_rules);

    let thai_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/all_gaps_word.rtf"
    );
    let thai = std::fs::read(thai_path).unwrap();
    let thai = String::from_utf8_lossy(&thai);
    assert!(thai.contains("\\ApplyBrkRules"));
    assert!(!thai.contains("\\applybrkrules"));
}
