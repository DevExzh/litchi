use litchi_rtf::{
    DocumentPrintLayoutSettings, DocumentThemeLanguages, LanguageId, RtfDocument, RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_flags_and_producer_style_theme_languages() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\gutterprl\twoonone\themelang1033"#,
        r#"\themelangfe2052\themelangcs1025 Body}"#,
    ))
    .unwrap();
    assert_eq!(
        *document.print_layout_settings(),
        DocumentPrintLayoutSettings {
            facing_pages: false,
            mirror_margins: false,
            document_gutter_twips: None,
            parallel_gutter: true,
            two_logical_pages_per_physical_page: true,
        }
    );
    assert_eq!(
        *document.theme_languages(),
        DocumentThemeLanguages {
            primary: Some(LanguageId::new(1033).unwrap()),
            east_asian: Some(LanguageId::new(2052).unwrap()),
            complex_script: Some(LanguageId::new(1025).unwrap()),
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_is_distinct_from_explicit_zero_languages() {
    let omitted = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(omitted.print_layout_settings().is_empty());
    assert!(omitted.theme_languages().is_empty());
    let serialized = String::from_utf8(write(&omitted)).unwrap();
    assert!(!serialized.contains("\\themelang"));

    let explicit =
        RtfDocument::parse(r#"{\rtf1\themelang0\themelangfe0\themelangcs0 Body}"#).unwrap();
    assert_eq!(explicit.theme_languages().primary.unwrap().value(), 0);
    assert_eq!(explicit.theme_languages().east_asian.unwrap().value(), 0);
    assert_eq!(
        explicit.theme_languages().complex_script.unwrap().value(),
        0
    );
    let serialized = String::from_utf8(write(&explicit)).unwrap();
    assert!(serialized.contains("\\themelang0"));
    assert!(serialized.contains("\\themelangfe0"));
    assert!(serialized.contains("\\themelangcs0"));
}

#[test]
fn typed_apis_round_trip_in_stable_order_and_clear_without_side_effects() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_print_layout_settings(DocumentPrintLayoutSettings {
        facing_pages: false,
        mirror_margins: false,
        document_gutter_twips: None,
        parallel_gutter: true,
        two_logical_pages_per_physical_page: true,
    })
    .unwrap();
    document.set_theme_languages(DocumentThemeLanguages {
        primary: Some(LanguageId::new(65535).unwrap()),
        east_asian: Some(LanguageId::new(1041).unwrap()),
        complex_script: Some(LanguageId::new(1037).unwrap()),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\gutterprl",
        "\\twoonone",
        "\\themelang65535",
        "\\themelangfe1041",
        "\\themelangcs1037",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.print_layout_settings(),
        document.print_layout_settings()
    );
    assert_eq!(reparsed.theme_languages(), document.theme_languages());
    assert_eq!(reparsed.text(), "Body");

    document.clear_print_layout_settings();
    document.clear_theme_languages();
    assert!(document.print_layout_settings().is_empty());
    assert!(document.theme_languages().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_theme_data_grid_and_rendering_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\themelangfe0\horzdoc\dgmargin\gutterprl"#,
        r#"{\*\themedata 0102}\themelang1031\jexpand\twoonone"#,
        r#"\themelangcs0 Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.print_layout_settings(),
        document.print_layout_settings()
    );
    assert_eq!(reparsed.theme_languages(), document.theme_languages());
    assert_eq!(reparsed.drawing_grid(), document.drawing_grid());
    assert_eq!(reparsed.rendering_settings(), document.rendering_settings());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_parameters_missing_values_ranges_overflow_and_duplicates() {
    for source in [
        r#"{\rtf1\gutterprl0 Body}"#,
        r#"{\rtf1\twoonone1 Body}"#,
        r#"{\rtf1\themelang Body}"#,
        r#"{\rtf1\themelang-1 Body}"#,
        r#"{\rtf1\themelang65536 Body}"#,
        r#"{\rtf1\themelangfe Body}"#,
        r#"{\rtf1\themelangfe-1 Body}"#,
        r#"{\rtf1\themelangfe65536 Body}"#,
        r#"{\rtf1\themelangcs Body}"#,
        r#"{\rtf1\themelangcs-1 Body}"#,
        r#"{\rtf1\themelangcs65536 Body}"#,
        r#"{\rtf1\themelang99999999999 Body}"#,
        r#"{\rtf1\gutterprl\gutterprl Body}"#,
        r#"{\rtf1\twoonone\twoonone Body}"#,
        r#"{\rtf1\themelang0\themelang1033 Body}"#,
        r#"{\rtf1\themelangfe0\themelangfe1041 Body}"#,
        r#"{\rtf1\themelangcs0\themelangcs1025 Body}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn rejects_every_starred_grouped_and_late_control() {
    for control in [
        r#"\gutterprl"#,
        r#"\twoonone"#,
        r#"\themelang1033"#,
        r#"\themelangfe2052"#,
        r#"\themelangcs1025"#,
    ] {
        for source in [
            format!(r#"{{\rtf1{{\*{control}}}Body}}"#),
            format!(r#"{{\rtf1{{{control}}}Body}}"#),
            format!(r#"{{\rtf1 Body{control}}}"#),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
