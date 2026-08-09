#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{FormFieldType, RtfDocument, RtfWriter};

const SYNTHETIC: &str = concat!(
    r#"{\rtf1\ansi Start"#,
    r#"{\field{\*\fldinst FORMTEXT{\*\formfield{"#,
    r#"\fftype0\fftypetxt0\ffmaxlen12\ffprot\ffrecalc"#,
    r#"{\*\ffname Account}{\*\ffformat Uppercase}"#,
    r#"{\*\ffdeftext D\'8a\u20320?}}}}{\fldrslt value}}"#,
    r#"{\field{\*\fldinst FORMCHECKBOX{\*\formfield{"#,
    r#"\fftype1\fftypetxt0\ffhps20\ffsize}}}{\fldrslt x}}End}"#,
);

#[test]
fn parses_typed_properties_and_round_trips_inertly() {
    let document = RtfDocument::parse(SYNTHETIC).unwrap();
    assert_eq!(document.text(), "StartvaluexEnd");
    assert_eq!(document.form_fields().len(), 2);

    let text = &document.form_fields()[0];
    assert_eq!(text.field_type, FormFieldType::Text);
    assert_eq!(text.max_length, Some(12));
    assert_eq!(text.format.as_deref(), Some("Uppercase"));
    assert_eq!(text.default_text.as_deref(), Some("DŠ你"));
    assert!(text.protected);
    assert!(text.calculate_on_exit);
    assert!(!text.size_automatically);

    let checkbox = &document.form_fields()[1];
    assert_eq!(checkbox.field_type, FormFieldType::CheckBox);
    assert!(checkbox.size_automatically);

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.form_fields(), document.form_fields());
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

fn wrap_formfield(properties: &str) -> String {
    format!(
        r"{{\rtf1{{\field{{\*\fldinst FORMTEXT{{\*\formfield{{{properties}}}}}}}}}{{\fldrslt x}}}}}}"
    )
}

#[test]
fn rejects_wrong_order_kind_bounds_duplicates_and_active_content() {
    let malformed = [
        r"\fftype0\ffmaxlen-1",
        r"\fftype0\ffmaxlen65536",
        r"\fftype0\ffmaxlen1\ffmaxlen2",
        r"\fftype0\ffprot\ffprot0",
        r"\fftype0\ffrecalc\ffrecalc0",
        r"\fftype0\ffsize",
        r"\fftype1\ffmaxlen4",
        r"\fftype1{\*\ffformat x}",
        r"\fftype1{\*\ffdeftext x}",
        r"\fftype2\ffhaslistbox{\*\ffl x}\ffsize",
        r"\fftype0{\*\ffformat x}{\*\ffformat y}",
        r"\fftype0{\ffformat x}",
        r"\fftype0{\*\ffformat{\field danger}}",
        r"\fftype0{\*\ffdeftext{\object danger}}",
        r"\fftype0{\*\ffformat\bin2 xx}",
    ];
    for properties in malformed {
        let source = wrap_formfield(properties);
        assert!(
            RtfDocument::parse(&source).is_err(),
            "accepted malformed properties: {properties}"
        );
    }

    let unstarred = r"{\rtf1{\field{\*\fldinst FORMTEXT{\formfield{\fftype0}}}{\fldrslt x}}}";
    assert!(RtfDocument::parse(unstarred).is_err());
}

#[test]
fn parses_bundled_libreoffice_extended_formfield() {
    let fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-5.rtf"
    );
    let marker = br"{\*\formfield";
    let format_marker = br"\ffformat";
    let mut cursor = 0usize;
    let (start, end) = loop {
        let start = cursor
            + fixture[cursor..]
                .windows(marker.len())
                .position(|window| window == marker)
                .unwrap();
        let mut depth = 0usize;
        let mut end = None;
        for (offset, byte) in fixture[start..].iter().enumerate() {
            match byte {
                b'{' => depth += 1,
                b'}' => {
                    depth -= 1;
                    if depth == 0 {
                        end = Some(start + offset + 1);
                        break;
                    }
                },
                _ => {},
            }
        }
        let end = end.unwrap();
        if fixture[start..end]
            .windows(format_marker.len())
            .any(|window| window == format_marker)
        {
            break (start, end);
        }
        cursor = end;
    };
    let group = std::str::from_utf8(&fixture[start..end]).unwrap();
    let source = format!(r"{{\rtf1\ansi{{\field{{\*\fldinst FORMTEXT{group}}}{{\fldrslt x}}}}}}");
    let document = RtfDocument::parse(&source).unwrap();
    let field = &document.form_fields()[0];
    assert_eq!(field.field_type, FormFieldType::Text);
    assert_eq!(field.max_length, Some(50));
    assert!(field.format.is_some());
}
