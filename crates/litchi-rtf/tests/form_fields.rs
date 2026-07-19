use litchi_rtf::{FormField, FormFieldType, FormTextType, RtfDocument, RtfWriter};
use std::borrow::Cow;
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_and_round_trips_unicode_dropdown_metadata_and_visible_result() {
    let rtf = concat!(
        r#"{\rtf1\ansi A{\field{\*\fldinst FORMDROPDOWN"#,
        r#"{\*\datafield 0a0b}{\*\formfield{"#,
        r#"\fftype2\fftypetxt0\ffhaslistbox\ffdefres0\ffres1"#,
        r#"{\*\ffname Choice}{\*\ffhelptext Pick}{\*\ffstattext Status}"#,
        r#"\ffownhelp\ffownstat{\*\ffentrymcr AutoOpen}{\*\ffexitmcr AutoClose}"#,
        r#"{\*\ffl One}{\*\ffl T\u20320?}}}}"#,
        r#"{\fldrslt Two}}Z}"#,
    );
    let document = RtfDocument::parse(rtf).unwrap();
    assert_eq!(document.text(), "ATwoZ");
    assert_eq!(document.form_fields().len(), 1);
    let field = &document.form_fields()[0];
    assert_eq!(field.field_type, FormFieldType::DropDown);
    assert_eq!(field.text_type, Some(FormTextType::Regular));
    assert_eq!(field.name.as_deref(), Some("Choice"));
    assert_eq!(field.list_entries, ["One", "T你"]);
    assert_eq!(field.selected_index(), Some(1));
    assert_eq!(field.selected_entry(), Some("T你"));
    assert_eq!(field.entry_macro.as_deref(), Some("AutoOpen"));
    assert_eq!(field.exit_macro.as_deref(), Some("AutoClose"));
    assert_eq!(field.data.as_ref(), [0x0a, 0x0b]);
    assert_eq!(field.position, 1);
    assert_eq!(field.range_end, 4);
    assert_eq!(field.result_text, "Two");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.form_fields(), document.form_fields());
}

#[test]
fn checkbox_undefined_result_and_mutation_are_typed_and_positional() {
    let source = concat!(
        r#"{\rtf1 X{\field{\*\fldinst FORMCHECKBOX"#,
        r#"{\*\formfield{\fftype1\fftypetxt0\ffhps20\ffdefres1\ffres25}}}"#,
        r#"{\fldrslt }}Y}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let checkbox = &document.form_fields()[0];
    assert_eq!(document.text(), "XY");
    assert_eq!(checkbox.position, 1);
    assert_eq!(checkbox.range_end, 1);
    assert_eq!(checkbox.default_checked(), Some(true));
    assert_eq!(checkbox.checked(), None);

    let mut mutated = RtfDocument::parse(r#"{\rtf1 AChoiceZ}"#).unwrap();
    mutated
        .push_form_field(FormField {
            field_type: FormFieldType::DropDown,
            text_type: Some(FormTextType::Regular),
            name: Some(Cow::Borrowed("Choice")),
            max_length: None,
            format: None,
            default_text: None,
            default_result: Some(0),
            result: Some(1),
            half_point_size: None,
            protected: false,
            calculate_on_exit: false,
            size_automatically: false,
            own_help: false,
            own_status: false,
            help_text: None,
            status_text: None,
            entry_macro: None,
            exit_macro: None,
            list_entries: vec![Cow::Borrowed("First"), Cow::Borrowed("Choice")],
            has_list_box: true,
            data: Cow::Borrowed(&[1, 2]),
            result_text: Cow::Borrowed("Choice"),
            position: 1,
            range_end: 7,
        })
        .unwrap();
    let output = write(&mutated);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap_or_else(|error| {
        panic!(
            "failed to parse writer output: {error}\n{}",
            String::from_utf8_lossy(&output)
        )
    });
    assert_eq!(reparsed.text(), "AChoiceZ");
    assert_eq!(reparsed.form_fields(), mutated.form_fields());
}

#[test]
fn rejects_malformed_conflicting_or_active_form_field_grammar() {
    let cases = [
        r#"{\rtf1{\*\formfield{\fftype0}}}"#,
        r#"{\rtf1{\*\datafield 00}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0\fftype0}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0\ffownhelp}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMDROPDOWN{\*\formfield{\fftype2\ffhaslistbox\ffres2{\*\ffl x}}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\datafield abc}{\*\formfield{\fftype0}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0{\object x}}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0{\*\ffentrymcr{\field x}}}}}{\fldrslt x}}}"#,
        r#"{\rtf1{\field{\*\fldinst FORMTEXT{\*\datafield\bin2 xx}{\*\formfield{\fftype0}}}{\fldrslt x}}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_bundled_libreoffice_form_field_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/extras/rtfexport/data/FORMDROPDOWN.rtf",
        "sw/qa/uibase/uiview/data/tdf152839_formtext.rtf",
        "sw/qa/extras/rtfimport/data/fdo44984.rtf",
        "sw/qa/extras/rtfimport/data/tdf96326.rtf",
        "sw/qa/extras/odfexport/data/tdf165315.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(
            !document.form_fields().is_empty(),
            "fixture exposed no form fields: {fixture}"
        );
    }
}
