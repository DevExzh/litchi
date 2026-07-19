use litchi_rtf::{Field, FieldStatus, FieldType, RtfDocument, RtfWriter};
use std::borrow::Cow;

fn write_document(document: &RtfDocument<'_>) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

fn field<'a>(document: &'a RtfDocument<'_>, instruction: &str) -> &'a Field<'a> {
    document
        .fields()
        .iter()
        .find(|field| field.instruction.trim() == instruction)
        .unwrap_or_else(|| panic!("missing field instruction {instruction}"))
}

#[test]
fn parses_presence_and_omission_without_changing_field_content() {
    let document = RtfDocument::parse(concat!(
        r"{\rtf1 A{\field\fldpriv\fldlock\fldedit\flddirty",
        r"{\*\fldinst MARKED}{\fldrslt cached}}",
        r"{\field{\*\fldinst OMITTED}{\fldrslt plain}}B}"
    ))
    .unwrap();

    assert_eq!(
        field(&document, "MARKED").status(),
        FieldStatus {
            dirty: true,
            edited: true,
            locked: true,
            private: true,
        }
    );
    assert_eq!(field(&document, "MARKED").result, "cached");
    assert_eq!(field(&document, "OMITTED").status(), FieldStatus::default());
    assert_eq!(field(&document, "OMITTED").result, "plain");
    assert_eq!(document.text(), "AB");
}

#[test]
fn nested_fields_have_independent_status() {
    let document = RtfDocument::parse(concat!(
        r"{\rtf1{\field\flddirty{\*\fldinst OUTER}{\fldrslt A",
        r"{\field\fldlock{\*\fldinst INNER}{\fldrslt B}}C}}}"
    ))
    .unwrap();

    assert_eq!(
        field(&document, "OUTER").status(),
        FieldStatus {
            dirty: true,
            ..FieldStatus::default()
        }
    );
    assert_eq!(
        field(&document, "INNER").status(),
        FieldStatus {
            locked: true,
            ..FieldStatus::default()
        }
    );
}

#[test]
fn rejects_parameters_duplicates_and_misplacement() {
    for control in ["flddirty", "fldedit", "fldlock", "fldpriv"] {
        let parameterized =
            format!("{{\\rtf1{{\\field\\{control}1{{\\*\\fldinst TEST}}{{\\fldrslt x}}}}}}");
        assert!(RtfDocument::parse(&parameterized).is_err(), "{control}");

        let duplicate = format!(
            "{{\\rtf1{{\\field\\{control}\\{control}{{\\*\\fldinst TEST}}{{\\fldrslt x}}}}}}"
        );
        assert!(RtfDocument::parse(&duplicate).is_err(), "{control}");

        let late = format!("{{\\rtf1{{\\field{{\\*\\fldinst TEST}}\\{control}{{\\fldrslt x}}}}}}");
        assert!(RtfDocument::parse(&late).is_err(), "{control}");

        let grouped =
            format!("{{\\rtf1{{\\field{{\\{control}}}{{\\*\\fldinst TEST}}{{\\fldrslt x}}}}}}");
        assert!(RtfDocument::parse(&grouped).is_err(), "{control}");
    }
}

#[test]
fn writer_uses_canonical_order_and_mutation_api() {
    let mut field = Field::new(
        FieldType::Unknown,
        Cow::Borrowed("TEST"),
        Cow::Borrowed("cached"),
    );
    field.set_status(FieldStatus {
        dirty: true,
        edited: true,
        locked: true,
        private: true,
    });

    let mut output = Vec::new();
    RtfWriter::new(&mut output).write_field(&field).unwrap();
    let serialized = String::from_utf8(output).unwrap();
    assert!(serialized.starts_with(r"{\field\flddirty\fldedit\fldlock\fldpriv{\*\fldinst{"));

    let document = RtfDocument::parse(&format!("{{\\rtf1{serialized}}}")).unwrap();
    assert_eq!(document.fields()[0].status(), field.status());
    let once = write_document(&document);
    let twice = write_document(&RtfDocument::parse(&once).unwrap());
    assert_eq!(once, twice);
}

#[test]
fn parses_libreoffice_field_status_fixtures() {
    let dirty = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf96326.rtf"
    )))
    .unwrap();
    assert!(dirty.fields().iter().any(|field| field.status.dirty));

    let edited = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/pgnlcrm.rtf"
    )))
    .unwrap();
    assert!(edited.fields().iter().any(|field| field.status.edited));

    let locked = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf100961_fixedDateTime.rtf"
    )))
    .unwrap();
    assert!(locked.fields().iter().any(|field| field.status.locked));

    let private = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf96326.rtf"
    )))
    .unwrap();
    assert!(private.fields().iter().any(|field| field.status.private));
}
