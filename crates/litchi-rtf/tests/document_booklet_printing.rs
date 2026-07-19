use litchi_rtf::{DocumentBookletPrinting, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> std::io::Result<Vec<u8>> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output).write_document(document)?;
    Ok(output)
}

#[test]
fn parses_complete_producer_style_booklet_metadata_passively() {
    let document =
        RtfDocument::parse(r#"{\rtf1\bookfold\bookfoldrev\bookfoldsheets16 Body}"#).unwrap();
    assert_eq!(
        *document.booklet_printing(),
        DocumentBookletPrinting {
            book_fold: true,
            reverse_book_fold: true,
            sheets_per_booklet: Some(16),
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn preserves_producer_observed_zero_separately_from_omission() {
    let explicit = RtfDocument::parse(r#"{\rtf1\bookfoldsheets0 Body}"#).unwrap();
    assert_eq!(explicit.booklet_printing().sheets_per_booklet, Some(0));
    assert_eq!(explicit.text(), "Body");

    let omitted = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(omitted.booklet_printing().is_empty());
    let serialized = String::from_utf8(write(&omitted).unwrap()).unwrap();
    assert!(!serialized.contains("bookfold"));
}

#[test]
fn accepts_independent_flags_without_synthesizing_layout_or_printing() {
    for (source, book_fold, reverse) in [
        (r#"{\rtf1\bookfold Body}"#, true, false),
        (r#"{\rtf1\bookfoldrev Body}"#, false, true),
    ] {
        let document = RtfDocument::parse(source).unwrap();
        assert_eq!(document.booklet_printing().book_fold, book_fold);
        assert_eq!(document.booklet_printing().reverse_book_fold, reverse);
        assert_eq!(document.text(), "Body");
        let serialized = String::from_utf8(write(&document).unwrap()).unwrap();
        assert!(!serialized.contains("\\landscape"));
        assert!(!serialized.contains("\\twoonone"));
    }
}

#[test]
fn typed_api_round_trips_boundaries_in_canonical_order_and_clears() {
    for sheets in [0, 4, 16, 2_147_483_644] {
        let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
        document.set_booklet_printing(DocumentBookletPrinting {
            book_fold: true,
            reverse_book_fold: true,
            sheets_per_booklet: Some(sheets),
        });
        let output = write(&document).unwrap();
        let serialized = String::from_utf8(output.clone()).unwrap();
        let controls = ["\\bookfold\\", "\\bookfoldrev\\", "\\bookfoldsheets"];
        for pair in controls.windows(2) {
            assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
        }
        let reparsed = RtfDocument::parse_bytes(&output).unwrap();
        assert_eq!(reparsed.booklet_printing(), document.booklet_printing());
        assert_eq!(reparsed.text(), "Body");
    }

    let mut document = RtfDocument::parse(r#"{\rtf1\bookfold Body}"#).unwrap();
    document.clear_booklet_printing();
    assert!(document.booklet_printing().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_independently_with_adjacent_print_layout_metadata() {
    let document = RtfDocument::parse(
        r#"{\rtf1\gutterprl\twoonone\bookfold\bookfoldsheets8\bookfoldrev Body}"#,
    )
    .unwrap();
    let output = write(&document).unwrap();
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.booklet_printing(), document.booklet_printing());
    assert_eq!(
        reparsed.print_layout_settings(),
        document.print_layout_settings()
    );
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_invalid_values_parameters_duplicates_starred_grouped_and_late_forms() {
    for source in [
        r#"{\rtf1\bookfoldsheets Body}"#,
        r#"{\rtf1\bookfoldsheets-4 Body}"#,
        r#"{\rtf1\bookfoldsheets1 Body}"#,
        r#"{\rtf1\bookfoldsheets2 Body}"#,
        r#"{\rtf1\bookfoldsheets5 Body}"#,
        r#"{\rtf1\bookfoldsheets2147483647 Body}"#,
        r#"{\rtf1\bookfoldsheets99999999999 Body}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
    for name in ["bookfold", "bookfoldrev"] {
        for suffix in ["0", "1", "2147483647", "99999999999"] {
            let source = format!(r#"{{\rtf1\{name}{suffix} Body}}"#);
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
    for control in ["bookfold", "bookfoldrev", "bookfoldsheets4"] {
        let name = control.trim_end_matches(|character: char| character.is_ascii_digit());
        for source in [
            format!(r#"{{\rtf1\{control}\{control} Body}}"#),
            format!(r#"{{\rtf1{{\*\{control}}}Body}}"#),
            format!(r#"{{\rtf1{{\{control}}}Body}}"#),
            format!(r#"{{\rtf1 Body\{control}}}"#),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {name}: {source}"
            );
        }
    }

    for invalid in [2, i32::MAX as u32 + 1] {
        let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
        document.set_booklet_printing(DocumentBookletPrinting {
            sheets_per_booklet: Some(invalid),
            ..DocumentBookletPrinting::default()
        });
        assert!(
            write(&document).is_err(),
            "serialized invalid value {invalid}"
        );
    }
}
