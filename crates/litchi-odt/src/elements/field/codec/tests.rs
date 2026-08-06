use super::super::{MAX_FIELD_DEPTH, TEXT_DATABASE_NAMESPACE};
use super::validation::checked_field_depth;
use super::wire::{parse_database_fields, parse_meta_fields, parse_note_body_contents};

#[test]
fn field_depth_is_bounded() {
    assert!(checked_field_depth(MAX_FIELD_DEPTH).is_err());
}

#[test]
fn database_wire_parser_keeps_typed_fields() {
    let xml = format!(
        r#"<text:p xmlns:text="{TEXT_DATABASE_NAMESPACE}"><text:database-name text:table-name="Rows"/></text:p>"#
    );

    let fields = parse_database_fields(&xml).unwrap();
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].source.table_name, "Rows");
}

#[test]
fn metadata_wire_parser_requires_no_evaluation() {
    let xml = format!(
        r#"<text:p xmlns:text="{TEXT_DATABASE_NAMESPACE}" xmlns:xml="http://www.w3.org/XML/1998/namespace"><text:meta-field xml:id="field-1"/></text:p>"#
    );

    assert_eq!(parse_meta_fields(&xml).unwrap().len(), 1);
}

#[test]
fn note_body_wire_parser_preserves_mixed_content() {
    let xml = format!(
        r#"<text:note xmlns:text="{TEXT_DATABASE_NAMESPACE}"><text:note-body><text:p>note</text:p></text:note-body></text:note>"#
    );

    assert_eq!(parse_note_body_contents(&xml).unwrap().len(), 1);
}
