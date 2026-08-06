//! Core field identity and document extraction semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

#[test]
fn test_field_creation() {
    let field = Field::new("PAGE".to_string(), Some("1".to_string()), false);
    assert_eq!(field.instruction(), "PAGE");
    assert_eq!(field.result(), Some("1"));
    assert!(!field.is_dirty());
    assert_eq!(field.field_type(), "PAGE");
}

#[test]
fn test_field_type_extraction() {
    let field = Field::new("DATE \\@ \"MMMM d, yyyy\"".to_string(), None, false);
    assert_eq!(field.field_type(), "DATE");

    let field = Field::new(
        "REF bookmark1 \\h".to_string(),
        Some("See Section 1".to_string()),
        true,
    );
    assert_eq!(field.field_type(), "REF");
    assert!(field.is_dirty());
}

#[test]
fn extracts_decoded_field_instruction_and_result() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
            <w:r><w:instrText xml:space="preserve"> IF &quot;A&amp;B&quot; = &quot;A&amp;B&quot; </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t xml:space="preserve"> Yes &amp; no </w:t><w:tab/><w:br/></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].instruction(), r#"IF "A&B" = "A&B""#);
    assert_eq!(fields[0].result(), Some(" Yes & no \t\n"));
    assert!(fields[0].is_dirty());
}

#[test]
fn extracts_simple_fields_in_source_order_with_flags_and_nested_results() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MERGEFIELD &quot;Full Name&quot; " w:dirty="on" w:fldLock="1">
                <w:r><w:t xml:space="preserve"> Ada &amp; </w:t></w:r>
                <w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple>
                <w:r><w:t><![CDATA[ <Lovelace> ]]></w:t><w:tab/><w:br/><w:noBreakHyphen/><w:softHyphen/></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText> DATE </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>Today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr=" NUMPAGES "/>
        </w:p></w:body></w:document>"#;

    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert_eq!(fields[0].instruction(), r#"MERGEFIELD "Full Name""#);
    assert_eq!(
        fields[0].result(),
        Some(" Ada & 7 <Lovelace> \t\n‑\u{00ad}")
    );
    assert!(fields[0].is_dirty());
    assert!(fields[0].is_locked());
    assert_eq!(fields[1].instruction(), "PAGE");
    assert_eq!(fields[1].result(), Some("7"));
    assert_eq!(fields[2].instruction(), "DATE");
    assert_eq!(fields[2].result(), Some("Today"));
    assert!(fields[2].is_dirty());
    assert!(fields[2].is_locked());
    assert_eq!(fields[3].instruction(), "NUMPAGES");
    assert_eq!(fields[3].result(), None);
}

#[test]
fn rejects_simple_fields_without_instructions() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:fldSimple><w:r><w:t>result</w:t></w:r></w:fldSimple></w:p></w:body></w:document>"#;
    assert!(Field::extract_from_document(xml).is_err());
}
