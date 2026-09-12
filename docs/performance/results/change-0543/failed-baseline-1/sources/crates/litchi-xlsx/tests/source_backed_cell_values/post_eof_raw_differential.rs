//! Public differential coverage for a raw materialization error after EOF.
//!
//! The follow-up shared traversal must preserve this public contract in both
//! the restored two-pass implementation and the candidate implementation:
//! validation accepts the complete XML stream, raw materialization rejects
//! the final typed boolean, and no partial snapshot is published.

use super::*;

const POST_EOF_RAW_ERROR: &str = "invalid worksheet boolean 'maybe'";
const SECOND_SHEET: &str = "/xl/worksheets/sheet2.xml";

#[derive(Clone, Copy, Debug)]
enum Shape {
    Medium,
    DenseSparse,
}

impl Shape {
    const fn dimensions(self) -> (usize, usize) {
        match self {
            Self::Medium => (96, 96),
            Self::DenseSparse => (128, 128),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Medium => "medium",
            Self::DenseSparse => "dense-sparse",
        }
    }
}

fn column_name(mut column: usize) -> String {
    let mut output = String::new();
    column += 1;
    while column != 0 {
        let remainder = (column - 1) % 26;
        output.push(char::from(
            b'A' + u8::try_from(remainder).expect("column remainder"),
        ));
        column = (column - 1) / 26;
    }
    output.chars().rev().collect()
}

fn post_eof_boolean_worksheet(shape: Shape) -> Vec<u8> {
    let (rows, columns) = shape.dimensions();
    let last_cell = format!("{}{}", column_name(columns - 1), rows);
    let mut xml = String::with_capacity(1_024 + rows * columns * 32);
    xml.push_str(&format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:{last_cell}"/><sheetData>"#
    ));
    for row in 0..rows {
        let row_number = row + 1;
        xml.push_str(&format!(r#"<row r="{row_number}">"#));
        for column in 0..columns {
            let address = format!("{}{}", column_name(column), row_number);
            let ordinal = row * columns + column + 1;
            if row + 1 == rows && column + 1 == columns {
                // The validator accepts the complete typed scalar. The raw
                // parser reaches this value only during post-EOF
                // materialization and reports its exact typed error.
                xml.push_str(&format!(r#"<c r="{address}" t="b"><v>maybe</v></c>"#));
            } else {
                xml.push_str(&format!(r#"<c r="{address}"><v>{ordinal}</v></c>"#));
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    xml.into_bytes()
}

fn source_for_shape(shape: Shape) -> (Vec<u8>, Vec<u8>) {
    let mut package = OpcPackage::from_bytes(&two_sheets()).unwrap();
    let second_sheet = package
        .get_part(&PackURI::new(SECOND_SHEET).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(post_eof_boolean_worksheet(shape));
    (PackageWriter::to_bytes(&package).unwrap(), second_sheet)
}

fn selected_error(editor: &SourceBackedEditor) -> Error {
    match editor.edit_sheets(["Sheet1".into()]) {
        Ok(_) => panic!("post-EOF raw error was unexpectedly accepted"),
        Err(error) => error,
    }
}

fn assert_post_eof_error(error: Error, shape: Shape, attempt: usize) {
    match error {
        Error::Invalid(actual) => assert_eq!(
            actual,
            POST_EOF_RAW_ERROR,
            "unexpected post-EOF raw message for {} attempt {attempt}",
            shape.name()
        ),
        other => panic!(
            "post-EOF raw error for {} attempt {attempt} had wrong type: {other:?}",
            shape.name()
        ),
    }
}

fn assert_shape_contract(shape: Shape) {
    let (bytes, second_sheet) = source_for_shape(shape);
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let original_version = source.version().unwrap();
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();

    // Run the same editor twice. This catches retained provisional parser or
    // validator state and requires the exact typed error on every retry.
    for attempt in 0..2 {
        assert_post_eof_error(selected_error(&editor), shape, attempt);
        assert_eq!(
            source.bytes.as_slice(),
            bytes.as_slice(),
            "source bytes changed after {} attempt {attempt}",
            shape.name()
        );
        assert_eq!(
            source.version().unwrap(),
            original_version,
            "source provenance changed after {} attempt {attempt}",
            shape.name()
        );
    }

    // A failed selected worksheet must leave the independent owner readable,
    // with its original source XML and value. No failed transaction has a
    // commit to publish; the later empty publication below must remain byte
    // exact as an additional atomicity oracle.
    let unaffected_source = editor.snapshot("Sheet2").unwrap();
    assert_eq!(unaffected_source.source_xml(), second_sheet.as_slice());
    assert_eq!(
        unaffected_source.value(address("A1")),
        Some(&Value::Number(Number::new("20").unwrap()))
    );

    let unaffected = editor
        .edit_sheets(["Sheet2".into()])
        .unwrap()
        .commit()
        .unwrap();
    assert!(!unaffected.changed(), "{} owner changed", shape.name());
    assert!(unaffected.patch().is_empty());
    assert_eq!(unaffected.snapshot().len(), 1);
    assert_eq!(
        unaffected.snapshot().value(0, address("A1")),
        Some(&Value::Number(Number::new("20").unwrap()))
    );

    let mut published = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut published, &unaffected)
        .unwrap();
    assert_eq!(
        published,
        bytes,
        "{} failure polluted publication",
        shape.name()
    );
    assert_eq!(source.bytes.as_slice(), bytes.as_slice());
    assert_eq!(source.version().unwrap(), original_version);
}

#[test]
fn post_eof_raw_boolean_error_is_differentially_stable_for_both_shapes() {
    for shape in [Shape::Medium, Shape::DenseSparse] {
        assert_shape_contract(shape);
    }
}
