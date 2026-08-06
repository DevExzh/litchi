//! Conditional, assignment, sequence, and formula field semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

#[test]
fn parses_inert_if_fields_without_evaluation() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" IF &quot;A&quot; = &quot;A&quot; &quot;yes&quot; &quot;no&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>yes</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>if MERGEFIELD Amount &gt; 100 "discount" "standard"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>discount</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="IFF 1 = 1 &quot;yes&quot; &quot;no&quot;"><w:r><w:t>not if</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_if_field());
    assert!(fields[1].is_if_field());
    assert!(!fields[2].is_if_field());

    let simple = fields[0].if_field().unwrap().unwrap();
    assert_eq!(simple.expression(), r#""A" = "A" "yes" "no""#);
    assert_eq!(simple.cached_result(), Some("yes"));
    assert!(simple.is_dirty());
    assert!(simple.is_locked());

    let complex = fields[1].if_field().unwrap().unwrap();
    assert_eq!(
        complex.expression(),
        r#"MERGEFIELD Amount > 100 "discount" "standard""#
    );
    assert_eq!(complex.cached_result(), Some("discount"));
    assert!(complex.is_dirty());
    assert!(complex.is_locked());
}

#[test]
fn rejects_if_fields_without_expressions() {
    let missing = Field::new("IF".to_string(), None, false);
    assert!(missing.if_field().is_err());
}

#[test]
fn parses_inert_compare_fields_without_evaluation() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" COMPARE &quot;CustomerNumber&quot; &gt;= 4 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>1</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>compare MERGEFIELD CustomerRating &lt;= 9</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>0</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="COMPARES Customer = 1"><w:r><w:t>not a comparison</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_compare_field());
    assert!(fields[1].is_compare_field());
    assert!(!fields[2].is_compare_field());

    let number = fields[0].compare_field().unwrap().unwrap();
    assert_eq!(number.comparison(), r#""CustomerNumber" >= 4"#);
    assert_eq!(number.cached_result(), Some("1"));
    assert!(number.is_dirty());
    assert!(number.is_locked());

    let rating = fields[1].compare_field().unwrap().unwrap();
    assert_eq!(rating.comparison(), "MERGEFIELD CustomerRating <= 9");
    assert_eq!(rating.cached_result(), Some("0"));
    assert!(rating.is_dirty());
    assert!(rating.is_locked());
    assert!(fields[2].compare_field().unwrap().is_none());
}

#[test]
fn rejects_compare_fields_without_comparisons() {
    let missing = Field::new("COMPARE".to_string(), None, false);
    assert!(missing.compare_field().is_err());
}

#[test]
fn parses_inert_set_fields_without_evaluation_or_state_changes() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SET RecipientName &quot;North America&quot; \* MERGEFORMAT" w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached recipient</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>set Total =SUM(ABOVE) + 1</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>125</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SETTINGS Value"><w:r><w:t>not set</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_set_field());
    assert!(fields[1].is_set_field());
    assert!(!fields[2].is_set_field());

    let recipient = fields[0].set_field().unwrap().unwrap();
    assert_eq!(recipient.target_name(), "RecipientName");
    assert_eq!(recipient.expression(), r#""North America" \* MERGEFORMAT"#);
    assert_eq!(recipient.cached_result(), Some("cached recipient"));
    assert!(recipient.is_dirty());
    assert!(recipient.is_locked());

    let total = fields[1].set_field().unwrap().unwrap();
    assert_eq!(total.target_name(), "Total");
    assert_eq!(total.expression(), "=SUM(ABOVE) + 1");
    assert_eq!(total.cached_result(), Some("125"));
    assert!(total.is_dirty());
    assert!(total.is_locked());
    assert!(fields[2].set_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_set_fields_without_evaluating_them() {
    for instruction in ["SET", "SET \"\" value", "SET Target", "SET Target   "] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.set_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!("SET Target {}", "x".repeat(MAX_SET_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.set_field().is_err());
}

#[test]
fn parses_inert_sequence_fields_without_bookmark_lookup_or_numbering() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SEQ Figure FigureChapter \r 3 \* ARABIC " w:dirty="true" w:fldLock="on">
                <w:r><w:t>3</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>seq Table \s 1 \* ROMAN</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>I</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SEQUENCE Figure"><w:r><w:t>not a sequence</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_sequence_field());
    assert!(fields[1].is_sequence_field());
    assert!(!fields[2].is_sequence_field());

    let figure = fields[0].sequence_field().unwrap().unwrap();
    assert_eq!(figure.identifier(), "Figure");
    assert_eq!(figure.bookmark(), Some("FigureChapter"));
    assert_eq!(figure.tail(), r"\r 3 \* ARABIC");
    assert_eq!(figure.cached_result(), Some("3"));
    assert!(figure.is_dirty());
    assert!(figure.is_locked());

    let table = fields[1].sequence_field().unwrap().unwrap();
    assert_eq!(table.identifier(), "Table");
    assert_eq!(table.bookmark(), None);
    assert_eq!(table.tail(), r"\s 1 \* ROMAN");
    assert_eq!(table.cached_result(), Some("I"));
    assert!(table.is_dirty());
    assert!(table.is_locked());
    assert!(fields[2].sequence_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_sequence_fields_without_numbering() {
    for instruction in ["SEQ", r#"SEQ ""#, r#"SEQ Figure ""#] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.sequence_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "SEQ Figure {}",
            "x".repeat(MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.sequence_field().is_err());
}

#[test]
fn parses_inert_formula_fields_without_evaluation() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" =SUM(ABOVE) \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>42</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>= IF(1 = 1, &quot;yes&quot;, &quot;no&quot;)</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>yes</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="EQUAL 1 + 1"><w:r><w:t>not a formula field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_formula_field());
    assert!(fields[1].is_formula_field());
    assert!(!fields[2].is_formula_field());

    let total = fields[0].formula_field().unwrap().unwrap();
    assert_eq!(total.formula(), r"SUM(ABOVE) \* MERGEFORMAT");
    assert_eq!(total.cached_result(), Some("42"));
    assert!(total.is_dirty());
    assert!(total.is_locked());

    let conditional = fields[1].formula_field().unwrap().unwrap();
    assert_eq!(conditional.formula(), r#"IF(1 = 1, "yes", "no")"#);
    assert_eq!(conditional.cached_result(), Some("yes"));
    assert!(conditional.is_dirty());
    assert!(conditional.is_locked());
    assert!(fields[2].formula_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_formula_fields_without_evaluating_them() {
    let missing = Field::new("=".to_string(), None, false);
    assert!(missing.is_formula_field());
    assert!(missing.formula_field().is_err());

    let too_long = Field::new(
        format!("={}", "x".repeat(MAX_FORMULA_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.formula_field().is_err());
}
