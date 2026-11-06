#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{FormField, FormFieldType, FormTextType, write::Writer};
use std::borrow::Cow;
use std::fs;

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
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
    let document = Tree::parse(rtf, litchi_rtf::read::OpenOptions::default()).unwrap();
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

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
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
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let checkbox = &document.form_fields()[0];
    assert_eq!(document.text(), "XY");
    assert_eq!(checkbox.position, 1);
    assert_eq!(checkbox.range_end, 1);
    assert_eq!(checkbox.default_checked(), Some(true));
    assert_eq!(checkbox.checked(), None);

    let mut mutated = Tree::parse(
        r"{\rtf1 AChoiceZ}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
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
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default())
        .unwrap_or_else(|error| {
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
        r"{\rtf1{\*\formfield{\fftype0}}}",
        r"{\rtf1{\*\datafield 00}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0\fftype0}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0\ffownhelp}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMDROPDOWN{\*\formfield{\fftype2\ffhaslistbox\ffres2{\*\ffl x}}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\datafield abc}{\*\formfield{\fftype0}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0{\object x}}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\formfield{\fftype0{\*\ffentrymcr{\field x}}}}}{\fldrslt x}}}",
        r"{\rtf1{\field{\*\fldinst FORMTEXT{\*\datafield\bin2 xx}{\*\formfield{\fftype0}}}{\fldrslt x}}}",
    ];
    for rtf in cases {
        assert!(
            Tree::parse(rtf, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {rtf}"
        );
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
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        let document = Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default())
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(
            !document.form_fields().is_empty(),
            "fixture exposed no form fields: {fixture}"
        );
    }
}

#[test]
fn instruction_bookmarks_are_metadata_and_non_body_owners_remain_opaque() {
    let source = concat!(
        r"{\rtf1\ansi Before {\field{\*\fldinst {\*\bkmkstart Value} FORMTEXT",
        r"{\*\formfield{\fftype0{\*\ffname Value}}}}{\fldrslt kept}}{\*\bkmkend Value}After}"
    );
    let p = litchi_rtf::Package::parse(source).unwrap();
    assert_eq!(p.fields()[0].instruction.trim(), "FORMTEXT");
    assert_eq!(p.fields()[0].field_type, litchi_rtf::FieldType::FormText);
    let bookmark = p.raw().bookmarks().get("Value").unwrap();
    assert_eq!(bookmark.position, 7);
    assert_eq!(bookmark.content, "kept");
    let header = r"{\rtf1\ansi{\header{\field{\*\fldinst {\*\bkmkstart HeaderOnly} AUTHOR{\*\bkmkend HeaderOnly}}{\fldrslt cached}}}Body}";
    let p = litchi_rtf::Package::parse(header).unwrap();
    assert_eq!(p.fields()[0].instruction.trim(), "AUTHOR");
    assert!(p.raw().bookmarks().bookmarks().is_empty());
    assert_eq!(p.raw().opaque_nodes().len(), 2);
    assert!(p.to_bytes().unwrap() == header.as_bytes());
}

#[test]
fn adjacent_empty_forms_keep_their_own_modifiers_and_authored_forms_use_defaults() {
    let source = concat!(
        r"{\rtf1\ansi",
        r"{\field\flddirty\fldpriv{\*\fldinst FORMCHECKBOX{\*\formfield{\fftype1{\*\ffname One}}}}{\fldrslt }}",
        r"{\field\fldedit\fldlock{\*\fldinst FORMCHECKBOX{\*\formfield{\fftype1{\*\ffname Two}}}}{\fldrslt }}}"
    );
    let donor=Tree::parse(r"{\rtf1{\field{\*\fldinst FORMCHECKBOX{\*\formfield{\fftype1{\*\ffname Three}}}}{\fldrslt }}}",litchi_rtf::OpenOptions::default()).unwrap();
    let mut document = Tree::parse(source, litchi_rtf::OpenOptions::default()).unwrap();
    let expected = document
        .fields()
        .iter()
        .map(|f| f.status)
        .collect::<Vec<_>>();
    assert!(expected[0].dirty && expected[0].private && !expected[0].locked);
    assert!(expected[1].edited && expected[1].locked && !expected[1].private);
    document
        .push_form_field(donor.form_fields()[0].clone())
        .unwrap();
    let output = write(&document);
    let read = Tree::parse_bytes(&output, litchi_rtf::OpenOptions::default()).unwrap();
    assert_eq!(
        read.form_fields()
            .iter()
            .map(|f| f.name.as_deref())
            .collect::<Vec<_>>(),
        [Some("One"), Some("Two"), Some("Three")]
    );
    assert_eq!(
        read.fields().iter().map(|f| f.status).collect::<Vec<_>>(),
        [expected[0], expected[1], litchi_rtf::FieldStatus::default()]
    );
    // Clearing the separate generic projection must not discard modifiers
    // belonging to form owners that remain in the document.
    document.clear_fields();
    let output = write(&document);
    let read = Tree::parse_bytes(&output, litchi_rtf::OpenOptions::default()).unwrap();
    assert_eq!(
        read.fields().iter().map(|f| f.status).collect::<Vec<_>>(),
        [expected[0], expected[1], litchi_rtf::FieldStatus::default()]
    );
    document.clear_form_fields();
    document
        .push_form_field(donor.form_fields()[0].clone())
        .unwrap();
    let output = write(&document);
    let read = Tree::parse_bytes(&output, litchi_rtf::OpenOptions::default()).unwrap();
    assert_eq!(read.fields().len(), 1);
    assert_eq!(read.fields()[0].status, litchi_rtf::FieldStatus::default());
    assert_eq!(read.form_fields()[0].name.as_deref(), Some("Three"));
}
