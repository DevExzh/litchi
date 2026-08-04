//! Focused field model, instruction, and XML regression tests.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::*;

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
fn extracts_mail_merge_field_names_without_switches() {
    let quoted = Field::new(
        r#"  MERGEFIELD "Full Name" \* MERGEFORMAT "#.to_string(),
        None,
        false,
    );
    assert!(quoted.is_merge_field());
    assert_eq!(quoted.merge_field_name(), Some("Full Name"));

    let unquoted = Field::new("mergefield CustomerId \\b prefix".to_string(), None, false);
    assert!(unquoted.is_merge_field());
    assert_eq!(unquoted.merge_field_name(), Some("CustomerId"));

    let missing = Field::new("MERGEFIELD \\* MERGEFORMAT".to_string(), None, false);
    assert_eq!(missing.merge_field_name(), None);
    let page = Field::new("PAGE".to_string(), None, false);
    assert!(!page.is_merge_field());
    assert_eq!(page.merge_field_name(), None);
}

#[test]
fn parses_inert_merge_fields_without_opening_data_sources() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MERGEFIELD &quot;Customer Region&quot; \b &quot;Dear &quot; \f &quot;!&quot; \m \v \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached region</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>mergefield CustomerName \b Prefix \f Suffix</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached customer</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="MERGEFIELDS CustomerName"><w:r><w:t>not a merge field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_merge_field());
    assert!(fields[1].is_merge_field());
    assert!(!fields[2].is_merge_field());

    let region = fields[0].merge_field().unwrap().unwrap();
    assert_eq!(region.field_name(), "Customer Region");
    assert_eq!(region.cached_result(), Some("cached region"));
    assert!(region.is_dirty());
    assert!(region.is_locked());
    assert_eq!(region.switches().len(), 5);
    assert_eq!(region.switches()[0].name(), 'b');
    assert_eq!(region.switches()[0].argument(), Some("Dear "));
    assert_eq!(region.switches()[1].name(), 'f');
    assert_eq!(region.switches()[1].argument(), Some("!"));
    assert!(region.has_switch('m'));
    assert!(region.has_switch('v'));
    assert!(region.has_switch('*'));
    assert_eq!(region.switches()[4].argument(), Some("MERGEFORMAT"));

    let customer = fields[1].merge_field().unwrap().unwrap();
    assert_eq!(customer.field_name(), "CustomerName");
    assert_eq!(customer.cached_result(), Some("cached customer"));
    assert!(customer.is_dirty());
    assert!(customer.is_locked());
    assert_eq!(customer.switches()[0].argument(), Some("Prefix"));
    assert_eq!(customer.switches()[1].argument(), Some("Suffix"));

    assert!(fields[2].merge_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_merge_field_semantics() {
    let missing_name = Field::new("MERGEFIELD \\* MERGEFORMAT".to_string(), None, false);
    assert!(missing_name.merge_field().is_err());

    let empty_name = Field::new(r#"MERGEFIELD "" "#.to_string(), None, false);
    assert!(empty_name.merge_field().is_err());

    let unexpected_operand = Field::new("MERGEFIELD Customer unexpected".to_string(), None, false);
    assert!(unexpected_operand.merge_field().is_err());
}

#[test]
fn parses_inert_mail_merge_data_fields_without_opening_sources_or_merging() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" DATA &quot;recipients source.csv&quot; &quot;headers source.csv&quot; \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached mail-merge source</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>data recipients.csv \q opaque</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached bare source</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="DATABASE recipients.csv"><w:r><w:t>not data metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_mail_merge_data());
    assert!(fields[1].is_mail_merge_data());
    assert!(!fields[2].is_mail_merge_data());

    let data = fields[0].mail_merge_data().unwrap().unwrap();
    assert_eq!(data.data_source(), "recipients source.csv");
    assert_eq!(data.header_source(), Some("headers source.csv"));
    assert_eq!(data.cached_result(), Some("cached mail-merge source"));
    assert!(data.is_dirty());
    assert!(data.is_locked());
    assert_eq!(data.switches().len(), 2);
    assert!(data.has_switch('*'));
    assert_eq!(data.switches()[0].argument(), Some("MERGEFORMAT"));
    assert!(data.has_switch('q'));
    assert_eq!(data.switches()[1].argument(), Some("opaque"));

    let bare = fields[1].mail_merge_data().unwrap().unwrap();
    assert_eq!(bare.data_source(), "recipients.csv");
    assert_eq!(bare.header_source(), None);
    assert_eq!(bare.cached_result(), Some("cached bare source"));
    assert!(bare.is_dirty());
    assert!(bare.is_locked());
    assert!(bare.has_switch('q'));
    assert_eq!(bare.switches()[0].argument(), Some("opaque"));

    assert!(fields[2].mail_merge_data().unwrap().is_none());

    let too_long = Field::new(
        format!(
            "DATA {}",
            "x".repeat(MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.mail_merge_data().is_err());
}

#[test]
fn rejects_invalid_mail_merge_data_field_semantics() {
    for instruction in [
        "DATA",
        r#"DATA \* MERGEFORMAT"#,
        r#"DATA """#,
        r#"DATA recipients.csv """#,
        "DATA recipients.csv headers.csv unexpected",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.is_mail_merge_data(), "{instruction}");
        assert!(field.mail_merge_data().is_err(), "{instruction}");
    }
}

#[test]
fn parses_inert_mail_merge_counter_fields_without_merging() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MERGEREC " w:dirty="true" w:fldLock="on">
                <w:r><w:t>12</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>mergeSEQ</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>3</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="MERGERECORD"><w:r><w:t>not a counter</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_merge_record());
    assert!(fields[0].is_mail_merge_counter());
    assert!(fields[1].is_merge_sequence());
    assert!(fields[1].is_mail_merge_counter());
    assert!(!fields[2].is_mail_merge_counter());

    let record = fields[0].mail_merge_counter().unwrap().unwrap();
    assert_eq!(record.kind(), MergeCounterKind::Record);
    assert_eq!(record.cached_result(), Some("12"));
    assert!(record.is_dirty());
    assert!(record.is_locked());

    let sequence = fields[1].mail_merge_counter().unwrap().unwrap();
    assert_eq!(sequence.kind(), MergeCounterKind::Sequence);
    assert_eq!(sequence.cached_result(), Some("3"));
    assert!(sequence.is_dirty());
    assert!(sequence.is_locked());

    assert!(fields[2].mail_merge_counter().unwrap().is_none());
}

#[test]
fn rejects_invalid_mail_merge_counter_field_semantics() {
    let record_argument = Field::new("MERGEREC 12".to_string(), None, false);
    assert!(record_argument.mail_merge_counter().is_err());

    let sequence_switch = Field::new("MERGESEQ \\* MERGEFORMAT".to_string(), None, false);
    assert!(sequence_switch.mail_merge_counter().is_err());
}

#[test]
fn parses_inert_mail_merge_next_fields_without_advancing_records() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" NEXT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached next</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>next</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached complex next</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="NEXTIF Customer = &quot;Ada&quot;"><w:r><w:t>not next</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_mail_merge_next());
    assert!(fields[1].is_mail_merge_next());
    assert!(!fields[2].is_mail_merge_next());

    let simple = fields[0].mail_merge_next().unwrap().unwrap();
    assert_eq!(simple.instruction(), "NEXT");
    assert_eq!(simple.cached_result(), Some("cached next"));
    assert!(simple.is_dirty());
    assert!(simple.is_locked());

    let complex = fields[1].mail_merge_next().unwrap().unwrap();
    assert_eq!(complex.cached_result(), Some("cached complex next"));
    assert!(complex.is_dirty());
    assert!(complex.is_locked());

    assert!(fields[2].mail_merge_next().unwrap().is_none());
}

#[test]
fn rejects_invalid_mail_merge_next_field_semantics() {
    let argument = Field::new("NEXT 12".to_string(), None, false);
    assert!(argument.mail_merge_next().is_err());

    let switch = Field::new("NEXT \\* MERGEFORMAT".to_string(), None, false);
    assert!(switch.mail_merge_next().is_err());
}

#[test]
fn parses_inert_conditional_mail_merge_controls_without_merging() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" NEXTIF Customer = &quot;Ada&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached nextif</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>skipif MERGEFIELD Order &lt; 100</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached skipif</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="NEXTIFF Customer = &quot;Ada&quot;"><w:r><w:t>not conditional</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_mail_merge_next_if());
    assert!(fields[0].is_mail_merge_conditional_control());
    assert!(fields[1].is_mail_merge_skip_if());
    assert!(fields[1].is_mail_merge_conditional_control());
    assert!(!fields[2].is_mail_merge_conditional_control());

    let next_if = fields[0].mail_merge_conditional_control().unwrap().unwrap();
    assert_eq!(next_if.kind(), MergeControlKind::NextIf);
    assert_eq!(next_if.comparison(), r#"Customer = "Ada""#);
    assert_eq!(next_if.cached_result(), Some("cached nextif"));
    assert!(next_if.is_dirty());
    assert!(next_if.is_locked());

    let skip_if = fields[1].mail_merge_conditional_control().unwrap().unwrap();
    assert_eq!(skip_if.kind(), MergeControlKind::SkipIf);
    assert_eq!(skip_if.comparison(), "MERGEFIELD Order < 100");
    assert_eq!(skip_if.cached_result(), Some("cached skipif"));
    assert!(skip_if.is_dirty());
    assert!(skip_if.is_locked());
}

#[test]
fn rejects_conditional_mail_merge_controls_without_comparisons() {
    let next_if = Field::new("NEXTIF".to_string(), None, false);
    assert!(next_if.mail_merge_conditional_control().is_err());

    let skip_if = Field::new("SKIPIF   ".to_string(), None, false);
    assert!(skip_if.mail_merge_conditional_control().is_err());
}

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
fn parses_inert_bookmark_reference_fields_without_resolution_or_navigation() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" REF &quot;Target Bookmark&quot; \d &quot;-&quot; \f \h \n \p \r \t \w \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached reference</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>pageref PageTarget \h \p</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>12 above</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr=" FTNREF FootnoteTarget \p \f ">
                <w:r><w:t>1 above</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr=" NOTEREF EndnoteTarget \p \f " w:dirty="true" w:fldLock="on">
                <w:r><w:t>i above</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr=" REFS Target "><w:r><w:t>not a reference</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 5);
    assert!(fields[0].is_reference_field());
    assert!(fields[1].is_reference_field());
    assert!(fields[2].is_reference_field());
    assert!(fields[3].is_reference_field());
    assert!(!fields[4].is_reference_field());

    let reference = fields[0].reference_field().unwrap().unwrap();
    assert_eq!(reference.kind(), ReferenceKind::Reference);
    assert_eq!(reference.bookmark(), "Target Bookmark");
    assert_eq!(
        reference.options(),
        &[
            ReferenceOption::SequencePageSeparator("-".to_string()),
            ReferenceOption::ReferencedNoteContent,
            ReferenceOption::Hyperlink,
            ReferenceOption::ParagraphNumberWithoutContext,
            ReferenceOption::RelativePosition,
            ReferenceOption::ParagraphNumberRelativeContext,
            ReferenceOption::SuppressNonNumberText,
            ReferenceOption::ParagraphNumberFullContext,
        ]
    );
    assert_eq!(reference.unknown_switches().len(), 2);
    assert_eq!(reference.unknown_switches()[0].name(), '*');
    assert_eq!(
        reference.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(reference.unknown_switches()[1].name(), 'q');
    assert_eq!(reference.unknown_switches()[1].argument(), Some("opaque"));
    assert_eq!(reference.cached_result(), Some("cached reference"));
    assert!(reference.is_dirty());
    assert!(reference.is_locked());

    let page_reference = fields[1].reference_field().unwrap().unwrap();
    assert_eq!(page_reference.kind(), ReferenceKind::PageReference);
    assert_eq!(page_reference.bookmark(), "PageTarget");
    assert_eq!(
        page_reference.options(),
        &[
            ReferenceOption::Hyperlink,
            ReferenceOption::RelativePosition,
        ]
    );
    assert_eq!(page_reference.cached_result(), Some("12 above"));
    assert!(page_reference.is_dirty());
    assert!(page_reference.is_locked());

    let footnote_reference = fields[2].reference_field().unwrap().unwrap();
    assert_eq!(footnote_reference.kind(), ReferenceKind::FootnoteReference);
    assert_eq!(footnote_reference.bookmark(), "FootnoteTarget");
    assert_eq!(
        footnote_reference.options(),
        &[
            ReferenceOption::RelativePosition,
            ReferenceOption::NoteMarkFormatting,
        ]
    );
    assert_eq!(footnote_reference.cached_result(), Some("1 above"));

    let note_reference = fields[3].reference_field().unwrap().unwrap();
    assert_eq!(note_reference.kind(), ReferenceKind::NoteReference);
    assert_eq!(note_reference.bookmark(), "EndnoteTarget");
    assert_eq!(
        note_reference.options(),
        &[
            ReferenceOption::RelativePosition,
            ReferenceOption::NoteMarkFormatting,
        ]
    );
    assert_eq!(note_reference.cached_result(), Some("i above"));
    assert!(note_reference.is_dirty());
    assert!(note_reference.is_locked());
    assert!(fields[4].reference_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_bookmark_reference_fields_without_resolution_or_navigation() {
    for instruction in [
        "REF",
        r#"REF ""#,
        r#"REF Bookmark \d"#,
        r#"REF Bookmark \f unexpected"#,
        r#"PAGEREF Bookmark \h unexpected"#,
        r#"NOTEREF Bookmark \f unexpected"#,
        r#"FTNREF Bookmark \p unexpected"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.reference_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!("REF {}", "x".repeat(MAX_REFERENCE_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.reference_field().is_err());

    let not_reference = Field::new(r#"REFS Bookmark"#.to_string(), None, false);
    assert!(!not_reference.is_reference_field());
    assert!(not_reference.reference_field().unwrap().is_none());
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

#[test]
fn parses_inert_equation_fields_without_calculation_or_rendering() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" EQ \o\ac(\fs24 Q,\fs16 R) " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached equation</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>eq \f(1,2)</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>1/2</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr=" EQ "/>
            <w:fldSimple w:instr="EQUAL 1 + 1"><w:r><w:t>not an equation</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_equation());
    assert!(fields[1].is_equation());
    assert!(fields[2].is_equation());
    assert!(!fields[3].is_equation());

    let stacked = fields[0].equation().unwrap().unwrap();
    assert_eq!(stacked.expression(), r#"\o\ac(\fs24 Q,\fs16 R)"#);
    assert_eq!(stacked.cached_result(), Some("cached equation"));
    assert!(stacked.is_dirty());
    assert!(stacked.is_locked());

    let fraction = fields[1].equation().unwrap().unwrap();
    assert_eq!(fraction.expression(), r#"\f(1,2)"#);
    assert_eq!(fraction.cached_result(), Some("1/2"));
    assert!(fraction.is_dirty());
    assert!(fraction.is_locked());

    let empty = fields[2].equation().unwrap().unwrap();
    assert_eq!(empty.expression(), "");
    assert!(fields[3].equation().unwrap().is_none());
}

#[test]
fn rejects_oversized_equation_fields_without_parsing_them() {
    let too_long = Field::new(
        format!("EQ {}", "x".repeat(MAX_EQUATION_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.equation().is_err());

    let not_equation = Field::new("EQUAL 1 + 1".to_string(), None, false);
    assert!(!not_equation.is_equation());
    assert!(not_equation.equation().unwrap().is_none());
}

#[test]
fn parses_inert_hyperlink_fields_without_opening_or_following_them() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" HYPERLINK &quot;https://example.test/a b&quot; \l &quot;_Toc1&quot; \o &quot;Stored tip&quot; \t &quot;_blank&quot; \m \n \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached external link</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>hyperlink \l &quot;JumpTarget&quot;</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached internal link</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="HYPERLINKER &quot;https://example.test&quot;"><w:r><w:t>not a link field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_hyperlink_field());
    assert!(fields[1].is_hyperlink_field());
    assert!(!fields[2].is_hyperlink_field());

    let external = fields[0].hyperlink_field().unwrap().unwrap();
    assert_eq!(external.external_target(), Some("https://example.test/a b"));
    assert_eq!(external.bookmark(), Some("_Toc1"));
    assert_eq!(external.screen_tip(), Some("Stored tip"));
    assert_eq!(external.target_frame(), Some("_blank"));
    assert!(external.appends_image_map_coordinates());
    assert!(external.opens_new_window());
    assert_eq!(external.unknown_switches().len(), 2);
    assert_eq!(external.unknown_switches()[0].name(), '*');
    assert_eq!(
        external.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(external.unknown_switches()[1].name(), 'q');
    assert_eq!(external.unknown_switches()[1].argument(), Some("opaque"));
    assert_eq!(external.cached_result(), Some("cached external link"));
    assert!(external.is_dirty());
    assert!(external.is_locked());

    let internal = fields[1].hyperlink_field().unwrap().unwrap();
    assert_eq!(internal.external_target(), None);
    assert_eq!(internal.bookmark(), Some("JumpTarget"));
    assert_eq!(internal.cached_result(), Some("cached internal link"));
    assert!(internal.is_dirty());
    assert!(internal.is_locked());
    assert!(fields[2].hyperlink_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_hyperlink_fields_without_resolving_targets() {
    for instruction in [
        "HYPERLINK",
        r#"HYPERLINK ""#,
        r#"HYPERLINK \l ""#,
        r#"HYPERLINK "https://example.test" \l First \l Second"#,
        r#"HYPERLINK "https://example.test" \o"#,
        r#"HYPERLINK "https://example.test" \m unexpected"#,
        r#"HYPERLINK "https://example.test" \m \m"#,
        r#"HYPERLINK "https://example.test" \n unexpected"#,
        r#"HYPERLINK "https://example.test" \n \n"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.hyperlink_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "HYPERLINK {}",
            "x".repeat(MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.hyperlink_field().is_err());

    let not_hyperlink = Field::new(
        r#"HYPERLINKER "https://example.test""#.to_string(),
        None,
        false,
    );
    assert!(!not_hyperlink.is_hyperlink_field());
    assert!(not_hyperlink.hyperlink_field().unwrap().is_none());
}

#[test]
fn parses_inert_quote_fields_without_inserting_or_transforming_text() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" QUOTE &quot;Stored literal&quot; \* MERGEFORMAT \# &quot;000&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached literal</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>quote "Complex literal" \@ "MMMM"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached complex literal</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="QUOTEY &quot;not a quote field&quot;"><w:r><w:t>not a quote field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_quote_field());
    assert!(fields[1].is_quote_field());
    assert!(!fields[2].is_quote_field());

    let literal = fields[0].quote_field().unwrap().unwrap();
    assert_eq!(literal.text(), "Stored literal");
    assert_eq!(literal.cached_result(), Some("cached literal"));
    assert!(literal.is_dirty());
    assert!(literal.is_locked());
    assert_eq!(literal.switches().len(), 2);
    assert_eq!(literal.switches()[0].name(), '*');
    assert_eq!(literal.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(literal.switches()[1].name(), '#');
    assert_eq!(literal.switches()[1].argument(), Some("000"));
    assert!(literal.has_switch('*'));

    let complex = fields[1].quote_field().unwrap().unwrap();
    assert_eq!(complex.text(), "Complex literal");
    assert_eq!(complex.cached_result(), Some("cached complex literal"));
    assert!(complex.is_dirty());
    assert!(complex.is_locked());
    assert_eq!(complex.switches()[0].name(), '@');
    assert_eq!(complex.switches()[0].argument(), Some("MMMM"));
    assert!(fields[2].quote_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_quote_fields_without_inserting_or_transforming_text() {
    for instruction in [
        "QUOTE",
        "QUOTE \\\\* MERGEFORMAT",
        r#"QUOTE "literal" unexpected"#,
        r#"QUOTE "unterminated"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.quote_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "QUOTE \"{}\"",
            "x".repeat(MAX_QUOTE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.quote_field().is_err());
}

#[test]
fn parses_inert_symbol_fields_without_mapping_codes_or_inserting_glyphs() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SYMBOL 0xA9 \f &quot;Symbol&quot; \s 12 \u " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached copyright</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>symbol 163 \a \h \j</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached pound</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SYMBOLS 163"><w:r><w:t>not a symbol field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_symbol_field());
    assert!(fields[1].is_symbol_field());
    assert!(!fields[2].is_symbol_field());

    let copyright = fields[0].symbol_field().unwrap().unwrap();
    assert_eq!(copyright.character_argument(), "0xA9");
    assert_eq!(copyright.cached_result(), Some("cached copyright"));
    assert!(copyright.is_dirty());
    assert!(copyright.is_locked());
    assert_eq!(copyright.switches().len(), 3);
    assert_eq!(copyright.switches()[0].name(), 'f');
    assert_eq!(copyright.switches()[0].argument(), Some("Symbol"));
    assert_eq!(copyright.switches()[1].name(), 's');
    assert_eq!(copyright.switches()[1].argument(), Some("12"));
    assert_eq!(copyright.switches()[2].name(), 'u');
    assert_eq!(copyright.switches()[2].argument(), None);
    assert!(copyright.has_switch('f'));

    let pound = fields[1].symbol_field().unwrap().unwrap();
    assert_eq!(pound.character_argument(), "163");
    assert_eq!(pound.cached_result(), Some("cached pound"));
    assert!(pound.is_dirty());
    assert!(pound.is_locked());
    assert_eq!(pound.switches().len(), 3);
    assert_eq!(pound.switches()[0].name(), 'a');
    assert_eq!(pound.switches()[1].name(), 'h');
    assert_eq!(pound.switches()[2].name(), 'j');
    assert!(fields[2].symbol_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_symbol_fields_without_mapping_codes_or_inserting_glyphs() {
    for instruction in [
        "SYMBOL",
        "SYMBOL \\f \"Symbol\"",
        "SYMBOL 0xA9 unexpected",
        "SYMBOL 0xA9 \\f \"unterminated",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.symbol_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!("SYMBOL {}", "x".repeat(MAX_SYMBOL_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.symbol_field().is_err());
}

#[test]
fn parses_inert_automatic_number_fields_without_calculating_numbers_or_layout() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" AUTONUM \s &quot;.&quot; \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>7.</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>autonumlgl \e \s ")"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>2.4</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTONUMOUT"><w:r><w:t>III</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="AUTONUMERIC"><w:r><w:t>not automatic numbering</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_auto_number_field());
    assert!(fields[1].is_auto_number_field());
    assert!(fields[2].is_auto_number_field());
    assert!(!fields[3].is_auto_number_field());

    let automatic = fields[0].auto_number_field().unwrap().unwrap();
    assert_eq!(automatic.kind(), AutoNumberKind::AutoNum);
    assert_eq!(automatic.kind().field_keyword(), "AUTONUM");
    assert_eq!(automatic.cached_result(), Some("7."));
    assert!(automatic.is_dirty());
    assert!(automatic.is_locked());
    assert_eq!(automatic.switches().len(), 2);
    assert_eq!(automatic.switches()[0].name(), 's');
    assert_eq!(automatic.switches()[0].argument(), Some("."));
    assert_eq!(automatic.switches()[1].name(), '*');
    assert_eq!(automatic.switches()[1].argument(), Some("MERGEFORMAT"));
    assert!(automatic.has_switch('s'));

    let legal = fields[1].auto_number_field().unwrap().unwrap();
    assert_eq!(legal.kind(), AutoNumberKind::AutoNumLegal);
    assert_eq!(legal.kind().field_keyword(), "AUTONUMLGL");
    assert_eq!(legal.cached_result(), Some("2.4"));
    assert!(legal.is_dirty());
    assert!(legal.is_locked());
    assert_eq!(legal.switches().len(), 2);
    assert_eq!(legal.switches()[0].name(), 'e');
    assert_eq!(legal.switches()[0].argument(), None);
    assert_eq!(legal.switches()[1].name(), 's');
    assert_eq!(legal.switches()[1].argument(), Some(")"));

    let outline = fields[2].auto_number_field().unwrap().unwrap();
    assert_eq!(outline.kind(), AutoNumberKind::AutoNumOutline);
    assert_eq!(outline.kind().field_keyword(), "AUTONUMOUT");
    assert_eq!(outline.cached_result(), Some("III"));
    assert!(!outline.is_dirty());
    assert!(!outline.is_locked());
    assert!(outline.switches().is_empty());
    assert!(fields[3].auto_number_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_automatic_number_fields_without_calculating_numbers_or_layout() {
    for instruction in [
        "AUTONUM unexpected",
        r#"AUTONUMLGL \s "unterminated"#,
        "AUTONUMOUT \\",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.auto_number_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "AUTONUM {}",
            "x".repeat(MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.auto_number_field().is_err());

    let other = Field::new("AUTONUMS".to_string(), None, false);
    assert!(!other.is_auto_number_field());
    assert!(other.auto_number_field().unwrap().is_none());
}

#[test]
fn parses_inert_list_number_fields_without_reading_lists_or_calculating_numbers() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" LISTNUM NumberDefault \l 6 \s 3 \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>(iii)</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>listnum "Outline Default" \l 4</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>c</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="LISTNUM \l 2"><w:r><w:t>i</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="LISTNUMBER NumberDefault"><w:r><w:t>not a list number field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_list_number_field());
    assert!(fields[1].is_list_number_field());
    assert!(fields[2].is_list_number_field());
    assert!(!fields[3].is_list_number_field());

    let numbered = fields[0].list_number_field().unwrap().unwrap();
    assert_eq!(numbered.list_name(), Some("NumberDefault"));
    assert_eq!(numbered.cached_result(), Some("(iii)"));
    assert!(numbered.is_dirty());
    assert!(numbered.is_locked());
    assert_eq!(numbered.switches().len(), 3);
    assert_eq!(numbered.switches()[0].name(), 'l');
    assert_eq!(numbered.switches()[0].argument(), Some("6"));
    assert_eq!(numbered.switches()[1].name(), 's');
    assert_eq!(numbered.switches()[1].argument(), Some("3"));
    assert_eq!(numbered.switches()[2].name(), '*');
    assert_eq!(numbered.switches()[2].argument(), Some("MERGEFORMAT"));
    assert!(numbered.has_switch('l'));

    let outline = fields[1].list_number_field().unwrap().unwrap();
    assert_eq!(outline.list_name(), Some("Outline Default"));
    assert_eq!(outline.cached_result(), Some("c"));
    assert!(outline.is_dirty());
    assert!(outline.is_locked());
    assert_eq!(outline.switches()[0].name(), 'l');
    assert_eq!(outline.switches()[0].argument(), Some("4"));

    let unnamed = fields[2].list_number_field().unwrap().unwrap();
    assert_eq!(unnamed.list_name(), None);
    assert_eq!(unnamed.cached_result(), Some("i"));
    assert_eq!(unnamed.switches()[0].name(), 'l');
    assert_eq!(unnamed.switches()[0].argument(), Some("2"));
    assert!(fields[3].list_number_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_list_number_fields_without_reading_lists_or_calculating_numbers() {
    for instruction in [
        "LISTNUM NumberDefault unexpected",
        r#"LISTNUM "unterminated"#,
        "LISTNUM \\",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.list_number_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "LISTNUM {}",
            "x".repeat(MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.list_number_field().is_err());

    let other = Field::new("LISTNUMBER NumberDefault".to_string(), None, false);
    assert!(!other.is_list_number_field());
    assert!(other.list_number_field().unwrap().is_none());
}

#[test]
fn parses_inert_style_reference_fields_without_style_or_layout_resolution() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" STYLEREF &quot;Heading 1&quot; \l \n \p \r \t \w \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Cached heading</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>styleref Title \n</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>1</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="STYLEREFS Heading 1"><w:r><w:t>not a style reference</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_style_reference_field());
    assert!(fields[1].is_style_reference_field());
    assert!(!fields[2].is_style_reference_field());

    let heading = fields[0].style_reference_field().unwrap().unwrap();
    assert_eq!(heading.style_name(), "Heading 1");
    assert_eq!(
        heading.options(),
        &[
            StyleOption::FollowingText,
            StyleOption::ParagraphNumber,
            StyleOption::RelativePosition,
            StyleOption::ParagraphNumberRelativeContext,
            StyleOption::SuppressNonNumberText,
            StyleOption::ParagraphNumberFullContext,
        ]
    );
    assert_eq!(heading.unknown_switches().len(), 2);
    assert_eq!(heading.unknown_switches()[0].name(), '*');
    assert_eq!(
        heading.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(heading.unknown_switches()[1].name(), 'q');
    assert_eq!(heading.unknown_switches()[1].argument(), Some("opaque"));
    assert_eq!(heading.cached_result(), Some("Cached heading"));
    assert!(heading.is_dirty());
    assert!(heading.is_locked());

    let title = fields[1].style_reference_field().unwrap().unwrap();
    assert_eq!(title.style_name(), "Title");
    assert_eq!(title.options(), &[StyleOption::ParagraphNumber]);
    assert_eq!(title.cached_result(), Some("1"));
    assert!(title.is_dirty());
    assert!(title.is_locked());
    assert!(fields[2].style_reference_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_style_reference_fields_without_style_or_layout_resolution() {
    for instruction in [
        "STYLEREF",
        r#"STYLEREF ""#,
        r#"STYLEREF Heading \l unexpected"#,
        "STYLEREF Heading unexpected",
        r#"STYLEREF Heading \"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.style_reference_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "STYLEREF Heading {}",
            "x".repeat(MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.style_reference_field().is_err());
}

#[test]
fn parses_inert_prompt_fields_without_displaying_prompts() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ASK AskResponse &quot;What is your first name?&quot; \d &quot;&quot; \o " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached ask response</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>fillin "Enter appointment time" \d "09:00"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>10:30</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="ASKER Answer &quot;not a prompt field&quot;"><w:r><w:t>not ask</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_ask_field());
    assert!(fields[0].is_prompt_field());
    assert!(fields[1].is_fill_in_field());
    assert!(fields[1].is_prompt_field());
    assert!(!fields[2].is_prompt_field());

    let ask = fields[0].prompt_field().unwrap().unwrap();
    assert_eq!(ask.kind(), PromptKind::Ask);
    assert_eq!(ask.bookmark(), Some("AskResponse"));
    assert_eq!(ask.prompt(), Some("What is your first name?"));
    assert_eq!(ask.default_response(), Some(""));
    assert!(ask.prompts_once_per_mail_merge());
    assert_eq!(ask.cached_result(), Some("cached ask response"));
    assert!(ask.is_dirty());
    assert!(ask.is_locked());

    let fill_in = fields[1].prompt_field().unwrap().unwrap();
    assert_eq!(fill_in.kind(), PromptKind::FillIn);
    assert_eq!(fill_in.bookmark(), None);
    assert_eq!(fill_in.prompt(), Some("Enter appointment time"));
    assert_eq!(fill_in.default_response(), Some("09:00"));
    assert!(!fill_in.prompts_once_per_mail_merge());
    assert_eq!(fill_in.cached_result(), Some("10:30"));
    assert!(fill_in.is_dirty());
    assert!(fill_in.is_locked());

    let default_only = Field::new(r#"FILLIN \d "recent response" \o"#.to_string(), None, false);
    let default_only = default_only.prompt_field().unwrap().unwrap();
    assert_eq!(default_only.kind(), PromptKind::FillIn);
    assert_eq!(default_only.bookmark(), None);
    assert_eq!(default_only.prompt(), None);
    assert_eq!(default_only.default_response(), Some("recent response"));
    assert!(default_only.prompts_once_per_mail_merge());

    assert!(fields[2].prompt_field().unwrap().is_none());
}

#[test]
fn rejects_malformed_prompt_field_metadata() {
    for instruction in [
        "ASK",
        r#"ASK "" "Question""#,
        "ASK Answer",
        r#"ASK Answer "Question" \d"#,
        r#"ASK Answer "Question" \o extra"#,
        r#"FILLIN "Question" \x"#,
        r#"FILLIN "Question" \d "first" \d "second""#,
        r#"FILLIN "Question" \o \o"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.prompt_field().is_err(), "{instruction}");
    }
}

#[test]
fn parses_inert_mail_merge_recipient_fields_without_merging() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ADDRESSBLOCK \c 2 \d \e &quot;United States&quot; \e Canada \f &quot;&lt;&lt;_FIRST0_&gt;&gt; &lt;&lt;_LAST0_&gt;&gt;&quot; \l 1033 \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached address</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>greetingline \f "Dear &lt;&lt;_FIRST0_&gt;&gt;," \e "To Whom It May Concern" \l en-US</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>Dear Ada,</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="ADDRESSBLOCKING \c 1"><w:r><w:t>not an address block</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_address_block());
    assert!(fields[0].is_mail_merge_recipient_field());
    assert!(fields[1].is_greeting_line());
    assert!(fields[1].is_mail_merge_recipient_field());
    assert!(!fields[2].is_mail_merge_recipient_field());

    let address = fields[0].mail_merge_recipient_field().unwrap().unwrap();
    assert_eq!(address.kind(), RecipientKind::AddressBlock);
    assert_eq!(
        address.country_inclusion(),
        Some(CountryInclusion::UnlessExcluded)
    );
    assert!(address.formats_using_recipient_country());
    let excluded = address
        .excluded_countries()
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(excluded, vec!["United States", "Canada"]);
    assert_eq!(address.format_template(), Some("<<_FIRST0_>> <<_LAST0_>>"));
    assert_eq!(address.language(), Some("1033"));
    assert_eq!(address.greeting_fallback_text(), None);
    assert_eq!(address.unknown_switches().len(), 1);
    assert_eq!(address.unknown_switches()[0].name(), '*');
    assert_eq!(
        address.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(address.cached_result(), Some("cached address"));
    assert!(address.is_dirty());
    assert!(address.is_locked());

    let greeting = fields[1].mail_merge_recipient_field().unwrap().unwrap();
    assert_eq!(greeting.kind(), RecipientKind::GreetingLine);
    assert_eq!(greeting.country_inclusion(), None);
    assert!(!greeting.formats_using_recipient_country());
    assert!(greeting.excluded_countries().is_empty());
    assert_eq!(greeting.format_template(), Some("Dear <<_FIRST0_>>,"));
    assert_eq!(greeting.language(), Some("en-US"));
    assert_eq!(
        greeting.greeting_fallback_text(),
        Some("To Whom It May Concern")
    );
    assert_eq!(greeting.cached_result(), Some("Dear Ada,"));
    assert!(greeting.is_dirty());
    assert!(greeting.is_locked());

    assert!(fields[2].mail_merge_recipient_field().unwrap().is_none());
}

#[test]
fn rejects_malformed_mail_merge_recipient_field_metadata() {
    for instruction in [
        "ADDRESSBLOCK text",
        "ADDRESSBLOCK \\c",
        "ADDRESSBLOCK \\c 3",
        "ADDRESSBLOCK \\d 1",
        "ADDRESSBLOCK \\d \\d",
        "ADDRESSBLOCK \\f",
        "GREETINGLINE \\f \"Dear\" \\f \"Hello\"",
        "GREETINGLINE \\l",
        "GREETINGLINE \\c \"First\" \\e \"Second\"",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.mail_merge_recipient_field().is_err(), "{instruction}");
    }
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
fn parses_inert_link_field_metadata_without_activating_sources() {
    let field = Field::new(
            r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#
                .to_string(),
            Some("cached LINK result".to_string()),
            true,
        );
    assert!(field.is_link());
    let link = field.link().unwrap().unwrap();
    assert_eq!(
        link.instruction(),
        r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#
    );
    assert_eq!(link.application_type(), "Excel.Sheet.8");
    assert_eq!(link.source(), r"C:\no-contact\source.xlsx");
    assert_eq!(link.item(), Some("Sheet1!R1C1:R4C4"));
    assert!(link.requests_automatic_updates());
    assert_eq!(
        link.result_options(),
        &[LinkResult::Picture, LinkResult::OmitGraphicData]
    );
    assert_eq!(
        link.effective_result_option(),
        Some(LinkResult::OmitGraphicData)
    );
    assert_eq!(link.formatting_modes(), &[LinkFormat::SpreadsheetSource]);
    assert_eq!(link.cached_result(), Some("cached LINK result"));
    assert!(link.is_dirty());
    assert!(!link.is_locked());
    assert_eq!(link.switches().len(), 5);
    assert_eq!(link.switches()[0].name(), 'a');
    assert_eq!(link.switches()[1].argument(), Some("4"));
    assert_eq!(link.switches()[4].name(), '*');
    assert_eq!(link.switches()[4].argument(), Some("MERGEFORMAT"));

    let multiple_formatting = Field::new(
        r"LINK Word.Document.8 source \f 0 \f 2 \t".to_string(),
        None,
        false,
    );
    let multiple_formatting = multiple_formatting.link().unwrap().unwrap();
    assert_eq!(
        multiple_formatting.formatting_modes(),
        &[LinkFormat::Source, LinkFormat::Destination]
    );
    assert_eq!(
        multiple_formatting.effective_result_option(),
        Some(LinkResult::Text)
    );

    let unsupported = Field::new(r"LINK Package source \f 1".to_string(), None, false);
    assert_eq!(
        unsupported.link().unwrap().unwrap().formatting_modes(),
        &[LinkFormat::Unsupported(1)]
    );

    let repeated_updates = Field::new(r"LINK Excel.Sheet.8 source \a \a".to_string(), None, false);
    assert!(
        repeated_updates
            .link()
            .unwrap()
            .unwrap()
            .requests_automatic_updates()
    );

    let not_link = Field::new("LINKAGE Excel.Sheet.8 source".to_string(), None, false);
    assert!(!not_link.is_link());
    assert!(not_link.link().unwrap().is_none());
    assert!(Field::new("LINK".to_string(), None, false).link().is_err());
    assert!(
        Field::new(
            r"LINK Excel.Sheet.8 source \f invalid".to_string(),
            None,
            false,
        )
        .link()
        .is_err()
    );
    assert!(
        Field::new(
            r"LINK Excel.Sheet.8 source \p unexpected".to_string(),
            None,
            false,
        )
        .link()
        .is_err()
    );
}

#[test]
fn parses_inert_dde_fields_without_starting_conversations() {
    let field = Field::new(
        r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#
            .to_string(),
        Some("cached DDE result".to_string()),
        true,
    );
    assert!(field.is_dde());
    assert!(!field.is_dde_auto());
    let dde = field.dde_link().unwrap().unwrap();
    assert_eq!(
        dde.instruction(),
        r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#
    );
    assert_eq!(dde.kind(), DdeKind::Dde);
    assert_eq!(dde.application(), "Excel");
    assert_eq!(dde.source(), r"C:\no-contact\source.xlsx");
    assert_eq!(dde.item(), Some("Sheet1!R1C1:R4C4"));
    assert!(dde.requests_automatic_updates());
    assert_eq!(dde.representation(), Some(DdeFormat::Picture));
    assert!(!dde.omits_graphic_data());
    assert_eq!(dde.cached_result(), Some("cached DDE result"));
    assert!(dde.is_dirty());
    assert!(!dde.is_locked());
    assert_eq!(dde.switches().len(), 3);
    assert_eq!(dde.switches()[0].name(), 'a');
    assert_eq!(dde.switches()[2].name(), '*');
    assert_eq!(dde.switches()[2].argument(), Some("MERGEFORMAT"));

    let automatic = Field::new(
        r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A1" \t"#.to_string(),
        None,
        false,
    );
    assert!(!automatic.is_dde());
    assert!(automatic.is_dde_auto());
    let automatic = automatic.dde_link().unwrap().unwrap();
    assert_eq!(automatic.kind(), DdeKind::DdeAuto);
    assert!(automatic.requests_automatic_updates());
    assert_eq!(automatic.representation(), Some(DdeFormat::Text));

    let omit_graphics = Field::new(r"DDE Excel source \a \d".to_string(), None, false)
        .dde_link()
        .unwrap()
        .unwrap();
    assert!(omit_graphics.requests_automatic_updates());
    assert!(omit_graphics.omits_graphic_data());
    assert_eq!(omit_graphics.representation(), None);

    assert!(
        Field::new("DDE".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    assert!(
        Field::new(r"DDE Excel \p".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    assert!(
        Field::new(r"DDE Excel source \p unexpected".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    assert!(
        Field::new(r"DDE Excel source \p \t".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    assert!(
        Field::new(r"DDEAUTO Excel source \p \t".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    assert!(
        Field::new(r"DDEAUTO Excel source \a".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    assert!(
        Field::new(r"DDE Excel source \a \a".to_string(), None, false)
            .dde_link()
            .is_err()
    );
    let not_dde = Field::new("DDEAUTOMATED Excel source".to_string(), None, false);
    assert!(!not_dde.is_dde());
    assert!(!not_dde.is_dde_auto());
    assert!(not_dde.dde_link().unwrap().is_none());
}

#[test]
fn parses_inert_referenced_document_fields_without_opening_sources() {
    let field = Field::with_flags(
        r#"RD "C:\\Manual\\Chapters\\Chapter 1.docx" \f \* MERGEFORMAT"#.to_string(),
        Some("cached RD result".to_string()),
        true,
        true,
    );
    assert!(field.is_referenced_document());
    let reference = field.referenced_document().unwrap().unwrap();
    assert_eq!(reference.source(), r"C:\Manual\Chapters\Chapter 1.docx");
    assert!(reference.uses_relative_path());
    assert_eq!(reference.cached_result(), Some("cached RD result"));
    assert!(reference.is_dirty());
    assert!(reference.is_locked());
    assert_eq!(reference.switches().len(), 2);
    assert_eq!(reference.switches()[0].name(), 'f');
    assert_eq!(reference.switches()[1].name(), '*');
    assert_eq!(reference.switches()[1].argument(), Some("MERGEFORMAT"));

    let absolute = Field::new(
        r#"RD "file:///no-contact/appendix.docx""#.to_string(),
        None,
        false,
    );
    let absolute = absolute.referenced_document().unwrap().unwrap();
    assert_eq!(absolute.source(), "file:///no-contact/appendix.docx");
    assert!(!absolute.uses_relative_path());

    assert!(
        Field::new("RD".to_string(), None, false)
            .referenced_document()
            .is_err()
    );
    assert!(
        Field::new(r#"RD "chapter.docx" \f relative"#.to_string(), None, false)
            .referenced_document()
            .is_err()
    );
    assert!(
        Field::new(r#"RD "chapter.docx" \f \f"#.to_string(), None, false)
            .referenced_document()
            .is_err()
    );
    let unknown = Field::new(r#"RD "chapter.docx" \p"#.to_string(), None, false)
        .referenced_document()
        .unwrap()
        .unwrap();
    assert!(!unknown.uses_relative_path());
    assert_eq!(unknown.switches()[0].name(), 'p');
    let not_rd = Field::new(r#"RDX "chapter.docx""#.to_string(), None, false);
    assert!(!not_rd.is_referenced_document());
    assert!(not_rd.referenced_document().unwrap().is_none());
}

#[test]
fn parses_inert_external_include_fields_without_resolving_sources() {
    let text_field = Field::new(
            r#"INCLUDETEXT "file:///C:/no-contact/source.xml" Summary \! \c Word8 \e utf-8 \m application/xml \n "xmlns:a=\"resume-schema\"" \t "file:///C:/display.xsl" \x a:Resume/a:Name \* MERGEFORMAT"#
                .to_string(),
            Some("cached included text".to_string()),
            true,
        );
    assert!(text_field.is_include_text());
    assert!(!text_field.is_include_picture());
    let text = text_field.external_include().unwrap().unwrap();
    assert_eq!(text.kind(), IncludeKind::Text);
    assert_eq!(text.source(), "file:///C:/no-contact/source.xml");
    assert_eq!(text.bookmark(), Some("Summary"));
    assert!(text.suppresses_nested_field_updates());
    assert!(!text.omits_picture_data());
    assert_eq!(
        text.options(),
        &[
            IncludeOption::Converter("Word8".to_string()),
            IncludeOption::Encoding("utf-8".to_string()),
            IncludeOption::MimeType("application/xml".to_string()),
            IncludeOption::NamespaceMapping("xmlns:a=\"resume-schema\"".to_string()),
            IncludeOption::Xslt("file:///C:/display.xsl".to_string()),
            IncludeOption::XPath("a:Resume/a:Name".to_string()),
        ]
    );
    assert_eq!(text.cached_result(), Some("cached included text"));
    assert!(text.is_dirty());
    assert!(!text.is_locked());
    assert_eq!(text.switches().len(), 8);
    assert_eq!(text.switches()[0].name(), '!');
    assert_eq!(text.switches()[7].name(), '*');

    let picture_field = Field::new(
        r#"INCLUDEPICTURE "file:///C:/no-contact/picture.gif" \c Pictim32 \d \* MERGEFORMAT"#
            .to_string(),
        Some("cached picture".to_string()),
        false,
    );
    assert!(!picture_field.is_include_text());
    assert!(picture_field.is_include_picture());
    let picture = picture_field.external_include().unwrap().unwrap();
    assert_eq!(picture.kind(), IncludeKind::Picture);
    assert_eq!(picture.source(), "file:///C:/no-contact/picture.gif");
    assert_eq!(picture.bookmark(), None);
    assert!(!picture.suppresses_nested_field_updates());
    assert!(picture.omits_picture_data());
    assert_eq!(
        picture.options(),
        &[IncludeOption::Converter("Pictim32".to_string())]
    );
    assert_eq!(picture.cached_result(), Some("cached picture"));
    assert_eq!(picture.switches()[2].name(), '*');

    let legacy_text_field = Field::new(
        r#"INCLUDE "file:///C:/no-contact/legacy.docx" LegacySection \!"#.to_string(),
        Some("cached legacy text".to_string()),
        true,
    );
    assert!(legacy_text_field.is_include_text());
    assert!(!legacy_text_field.is_include_picture());
    let legacy_text = legacy_text_field.external_include().unwrap().unwrap();
    assert_eq!(legacy_text.kind(), IncludeKind::Text);
    assert_eq!(legacy_text.source(), "file:///C:/no-contact/legacy.docx");
    assert_eq!(legacy_text.bookmark(), Some("LegacySection"));
    assert!(legacy_text.suppresses_nested_field_updates());
    assert_eq!(legacy_text.cached_result(), Some("cached legacy text"));
    assert!(legacy_text.is_dirty());

    let legacy_picture_field = Field::new(
        r#"IMPORT "file:///C:/no-contact/legacy.wmf" \c GraphicsFilter \d"#.to_string(),
        Some("cached legacy picture".to_string()),
        false,
    );
    assert!(!legacy_picture_field.is_include_text());
    assert!(legacy_picture_field.is_include_picture());
    let legacy_picture = legacy_picture_field.external_include().unwrap().unwrap();
    assert_eq!(legacy_picture.kind(), IncludeKind::Picture);
    assert_eq!(legacy_picture.source(), "file:///C:/no-contact/legacy.wmf");
    assert_eq!(legacy_picture.bookmark(), None);
    assert!(legacy_picture.omits_picture_data());
    assert_eq!(
        legacy_picture.options(),
        &[IncludeOption::Converter("GraphicsFilter".to_string())]
    );
    assert_eq!(
        legacy_picture.cached_result(),
        Some("cached legacy picture")
    );

    assert!(
        Field::new("INCLUDETEXT".to_string(), None, false)
            .external_include()
            .is_err()
    );
    assert!(
        Field::new(r"INCLUDETEXT \c Word8".to_string(), None, false)
            .external_include()
            .is_err()
    );
    assert!(
        Field::new(
            r#"INCLUDEPICTURE "picture.gif" Selector"#.to_string(),
            None,
            false,
        )
        .external_include()
        .is_err()
    );
    assert!(
        Field::new(
            r#"INCLUDEPICTURE "picture.gif" \d unexpected"#.to_string(),
            None,
            false,
        )
        .external_include()
        .is_err()
    );
    assert!(
        Field::new(r"INCLUDETEXT source \! unexpected".to_string(), None, false)
            .external_include()
            .is_err()
    );
    assert!(
        Field::new(r"INCLUDETEXT source \e".to_string(), None, false)
            .external_include()
            .is_err()
    );
    for instruction in [
        "INCLUDETEXTUAL missing.docx",
        r#"INCLUDES "source.docx""#,
        r#"IMPORTS "picture.wmf""#,
    ] {
        let not_include = Field::new(instruction.to_string(), None, false);
        assert!(!not_include.is_include_text());
        assert!(!not_include.is_include_picture());
        assert!(not_include.external_include().unwrap().is_none());
    }
}

#[test]
fn parses_toc_fields_and_standard_switches() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TOC \o &quot;1-3&quot; \h \z \b &quot;Main Bookmark&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Introduction</w:t><w:tab/><w:t>1</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText>TOC\o&quot;2-4&quot;\u \n &quot;2-2&quot; \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>Chapter</w:t><w:tab/><w:t>4</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="TOCENTRY \f ignored"><w:r><w:t>not a TOC</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_table_of_contents());
    assert!(fields[1].is_table_of_contents());
    assert!(!fields[2].is_table_of_contents());

    let first = fields[0].table_of_contents().unwrap().unwrap();
    assert_eq!(first.cached_result(), Some("Introduction\t1"));
    assert!(first.is_dirty());
    assert!(first.is_locked());
    assert!(first.includes_hyperlinks());
    assert!(first.hides_page_numbers_in_web_layout());
    assert!(!first.uses_outline_levels());
    assert_eq!(first.switches()[0].name(), 'o');
    assert_eq!(first.switches()[0].argument(), Some("1-3"));
    assert_eq!(first.switches()[3].argument(), Some("Main Bookmark"));
    assert_eq!(
        first.heading_style_levels().unwrap(),
        vec![TocLevelRange::new(1, 3).unwrap()]
    );

    let second = fields[1].table_of_contents().unwrap().unwrap();
    assert_eq!(second.cached_result(), Some("Chapter\t4"));
    assert!(second.uses_outline_levels());
    assert!(!second.includes_hyperlinks());
    assert_eq!(second.switches()[0].name(), 'o');
    assert_eq!(second.switches()[0].argument(), Some("2-4"));
    assert_eq!(second.switches()[3].name(), '*');
    assert_eq!(second.switches()[3].argument(), Some("MERGEFORMAT"));
    assert_eq!(
        second.heading_style_levels().unwrap(),
        vec![TocLevelRange::new(2, 4).unwrap()]
    );
}

#[test]
fn parses_table_of_contents_entry_fields_without_generating_contents() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TC &quot;Illustration 1&quot; \f i \l 4 \n \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached entry</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>tc&quot;Appendix A&quot;\l 2</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached appendix</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="TCC &quot;not an entry&quot;"><w:r><w:t>not a TC field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_table_of_contents_entry());
    assert!(fields[1].is_table_of_contents_entry());
    assert!(!fields[2].is_table_of_contents_entry());

    let illustration = fields[0].table_of_contents_entry().unwrap().unwrap();
    assert_eq!(illustration.entry(), "Illustration 1");
    assert_eq!(illustration.cached_result(), Some("cached entry"));
    assert!(illustration.is_dirty());
    assert!(illustration.is_locked());
    assert_eq!(illustration.list_identifier().unwrap(), Some("i"));
    assert_eq!(illustration.level().unwrap(), Some("4"));
    assert!(illustration.omits_page_number());
    assert_eq!(illustration.switches()[3].name(), '*');
    assert_eq!(illustration.switches()[3].argument(), Some("MERGEFORMAT"));

    let appendix = fields[1].table_of_contents_entry().unwrap().unwrap();
    assert_eq!(appendix.entry(), "Appendix A");
    assert_eq!(appendix.cached_result(), Some("cached appendix"));
    assert!(appendix.is_dirty());
    assert!(appendix.is_locked());
    assert_eq!(appendix.list_identifier().unwrap(), None);
    assert_eq!(appendix.level().unwrap(), Some("2"));
    assert!(!appendix.omits_page_number());

    assert!(fields[2].table_of_contents_entry().unwrap().is_none());
}

#[test]
fn rejects_invalid_table_of_contents_entry_field_semantics() {
    for instruction in [
        "TC",
        r#"TC """#,
        r#"TC "entry" unexpected"#,
        r#"TC "entry" \f"#,
        r#"TC "entry" \l"#,
        r#"TC "entry" \n unexpected"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.table_of_contents_entry().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "TC {}",
            "x".repeat(MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.table_of_contents_entry().is_err());

    let not_entry = Field::new(r#"TCC "entry""#.to_string(), None, false);
    assert!(!not_entry.is_table_of_contents_entry());
    assert!(not_entry.table_of_contents_entry().unwrap().is_none());
}

#[test]
fn parses_citation_and_bibliography_fields() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" CITATION Doe2024 \m &quot;Smith 2025&quot; \l 1033 \p &quot;14&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>(Doe, 2024; Smith, 2025, p. 14)</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>BIBLIOGRAPHY \l 1033 \f 1036 \m Doe2024 \m Smith2025</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>Doe. Example work.</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="CITATIONEXTRA ignored"><w:r><w:t>not a citation</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_citation());
    assert!(fields[1].is_bibliography());
    assert!(!fields[2].is_citation());
    assert!(!fields[2].is_bibliography());

    let citation = fields[0].citation().unwrap().unwrap();
    assert_eq!(
        citation.cached_result(),
        Some("(Doe, 2024; Smith, 2025, p. 14)")
    );
    assert!(citation.is_dirty());
    assert!(citation.is_locked());
    assert_eq!(citation.primary_source_tag(), "Doe2024");
    assert_eq!(citation.source_tags(), ["Doe2024", "Smith 2025"]);
    assert_eq!(citation.additional_source_tags(), ["Smith 2025"]);
    assert_eq!(citation.switches()[0].name(), 'm');
    assert_eq!(citation.switches()[0].argument(), Some("Smith 2025"));
    assert!(citation.has_switch('l'));
    assert!(citation.has_switch('p'));

    let documented_order = Field::new(
        r#"CITATION \l 1033 "Che 01" \v 3 \m Kra \v 2"#.to_string(),
        None,
        true,
    );
    let documented = documented_order.citation().unwrap().unwrap();
    assert_eq!(documented.source_tags(), ["Che 01", "Kra"]);
    assert_eq!(documented.switches()[0].name(), 'l');
    assert_eq!(documented.switches()[0].argument(), Some("1033"));
    assert!(documented.is_dirty());

    let bibliography = fields[1].bibliography().unwrap().unwrap();
    assert_eq!(bibliography.cached_result(), Some("Doe. Example work."));
    assert!(bibliography.is_dirty());
    assert!(bibliography.is_locked());
    assert_eq!(bibliography.switches()[0].name(), 'l');
    assert_eq!(bibliography.switches()[0].argument(), Some("1033"));
    assert!(bibliography.has_switch('f'));
    assert_eq!(bibliography.switches()[1].argument(), Some("1036"));
    assert_eq!(bibliography.switches()[2].argument(), Some("Doe2024"));
    assert_eq!(bibliography.switches()[3].argument(), Some("Smith2025"));
}

#[test]
fn parses_document_variable_fields_without_resolving_values() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" DOCVARIABLE &quot;Customer Region&quot; \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached region</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>DOCVARIABLE CustomerName</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached customer</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="DOCVARIABLES CustomerName"><w:r><w:t>not a variable</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_document_variable());
    assert!(fields[1].is_document_variable());
    assert!(!fields[2].is_document_variable());

    let region = fields[0].document_variable().unwrap().unwrap();
    assert_eq!(region.variable_name(), "Customer Region");
    assert_eq!(region.cached_result(), Some("cached region"));
    assert!(region.is_dirty());
    assert!(region.is_locked());
    assert!(region.has_switch('*'));
    assert_eq!(region.switches()[0].argument(), Some("MERGEFORMAT"));

    let customer = fields[1].document_variable().unwrap().unwrap();
    assert_eq!(customer.variable_name(), "CustomerName");
    assert_eq!(customer.cached_result(), Some("cached customer"));
    assert!(customer.is_dirty());
    assert!(customer.is_locked());
    assert!(customer.switches().is_empty());
    assert!(fields[2].document_variable().unwrap().is_none());
}

#[test]
fn rejects_invalid_document_variable_field_semantics() {
    let missing_name = Field::new("DOCVARIABLE \\* MERGEFORMAT".to_string(), None, false);
    assert!(missing_name.document_variable().is_err());

    let empty_name = Field::new(r#"DOCVARIABLE "" "#.to_string(), None, false);
    assert!(empty_name.document_variable().is_err());

    let unexpected_operand = Field::new("DOCVARIABLE Customer unexpected".to_string(), None, false);
    assert!(unexpected_operand.document_variable().is_err());
}

#[test]
fn parses_document_property_fields_without_resolving_values() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" DOCPROPERTY &quot;Project Name&quot; \* MERGEFORMAT \@ &quot;MMMM d, yyyy&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached project</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>docproperty Revision</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached revision</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="DOCPROPERTYS ProjectName"><w:r><w:t>not a property</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_document_property());
    assert!(fields[1].is_document_property());
    assert!(!fields[2].is_document_property());

    let project = fields[0].document_property().unwrap().unwrap();
    assert_eq!(project.property_name(), "Project Name");
    assert_eq!(project.cached_result(), Some("cached project"));
    assert!(project.is_dirty());
    assert!(project.is_locked());
    assert_eq!(project.switches().len(), 2);
    assert_eq!(project.switches()[0].name(), '*');
    assert_eq!(project.switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(project.switches()[1].name(), '@');
    assert_eq!(project.switches()[1].argument(), Some("MMMM d, yyyy"));
    assert!(project.has_switch('*'));
    assert!(project.has_switch('@'));

    let revision = fields[1].document_property().unwrap().unwrap();
    assert_eq!(revision.property_name(), "Revision");
    assert_eq!(revision.cached_result(), Some("cached revision"));
    assert!(revision.is_dirty());
    assert!(revision.is_locked());
    assert!(revision.switches().is_empty());
    assert!(fields[2].document_property().unwrap().is_none());
}

#[test]
fn rejects_invalid_document_property_field_semantics() {
    for instruction in [
        r#"DOCPROPERTY \* MERGEFORMAT"#,
        r#"DOCPROPERTY """#,
        "DOCPROPERTY Project unexpected",
        r#"DOCPROPERTY Project \"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.document_property().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "DOCPROPERTY {}",
            "x".repeat(MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.document_property().is_err());
}

#[test]
fn parses_inert_explicit_info_fields_without_reading_or_modifying_properties() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" INFO TITLE &quot;Stored title override&quot; \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached title</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>info COMMENTS "Stored comment" \@ "opaque format"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached comment</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="INFO TEMPLATE"><w:r><w:t>cached template</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="INFOS TITLE"><w:r><w:t>not an info field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_info_field());
    assert!(fields[1].is_info_field());
    assert!(fields[2].is_info_field());
    assert!(!fields[3].is_info_field());

    let title = fields[0].info_field().unwrap().unwrap();
    assert_eq!(title.information_type(), "TITLE");
    assert_eq!(title.new_value(), Some("Stored title override"));
    assert_eq!(title.cached_result(), Some("cached title"));
    assert!(title.is_dirty());
    assert!(title.is_locked());
    assert_eq!(title.switches().len(), 1);
    assert_eq!(title.switches()[0].name(), '*');
    assert_eq!(title.switches()[0].argument(), Some("MERGEFORMAT"));
    assert!(title.has_switch('*'));

    let comments = fields[1].info_field().unwrap().unwrap();
    assert_eq!(comments.information_type(), "COMMENTS");
    assert_eq!(comments.new_value(), Some("Stored comment"));
    assert_eq!(comments.cached_result(), Some("cached comment"));
    assert!(comments.is_dirty());
    assert!(comments.is_locked());
    assert_eq!(comments.switches().len(), 1);
    assert_eq!(comments.switches()[0].name(), '@');
    assert_eq!(comments.switches()[0].argument(), Some("opaque format"));

    let template = fields[2].info_field().unwrap().unwrap();
    assert_eq!(template.information_type(), "TEMPLATE");
    assert_eq!(template.new_value(), None);
    assert_eq!(template.cached_result(), Some("cached template"));
    assert!(!template.is_dirty());
    assert!(!template.is_locked());

    assert!(fields[3].info_field().unwrap().is_none());
    let ambiguous = Field::new(r#"TITLE "Stored title override""#.to_string(), None, false);
    assert!(!ambiguous.is_info_field());
    assert!(ambiguous.info_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_explicit_info_fields_without_reading_or_modifying_properties() {
    for instruction in [
        "INFO",
        r#"INFO "" "#,
        r#"INFO TITLE "Stored title" unexpected"#,
        r#"INFO TITLE "unterminated"#,
        r#"INFO TITLE \"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.info_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!("INFO {}", "x".repeat(MAX_INFO_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.info_field().is_err());
}

#[test]
fn parses_document_information_fields_without_reading_or_calculating_values() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TITLE \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached title</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>author \@ "opaque format"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached author</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTHORS"><w:r><w:t>not an author field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let extracted = Field::extract_from_document(xml).unwrap();
    assert_eq!(extracted.len(), 3);
    assert!(extracted[0].is_document_information());
    assert!(extracted[1].is_document_information());
    assert!(!extracted[2].is_document_information());

    let title = extracted[0].document_information().unwrap().unwrap();
    assert_eq!(title.kind(), InformationKind::Title);
    assert_eq!(title.cached_result(), Some("cached title"));
    assert!(title.is_dirty());
    assert!(title.is_locked());
    assert_eq!(title.switches()[0].name(), '*');
    assert_eq!(title.switches()[0].argument(), Some("MERGEFORMAT"));

    let author = extracted[1].document_information().unwrap().unwrap();
    assert_eq!(author.kind(), InformationKind::Author);
    assert_eq!(author.cached_result(), Some("cached author"));
    assert!(author.is_dirty());
    assert!(author.is_locked());
    assert!(author.has_switch('@'));
    assert_eq!(author.switches()[0].argument(), Some("opaque format"));
    assert!(extracted[2].document_information().unwrap().is_none());

    for (instruction, kind) in [
        (r"TITLE \* MERGEFORMAT", InformationKind::Title),
        (r"SUBJECT \* MERGEFORMAT", InformationKind::Subject),
        (r"AUTHOR \* MERGEFORMAT", InformationKind::Author),
        (r"KEYWORDS \* MERGEFORMAT", InformationKind::Keywords),
        (r"COMMENTS \* MERGEFORMAT", InformationKind::Comments),
        (r"LASTSAVEDBY \* MERGEFORMAT", InformationKind::LastSavedBy),
        (r"CREATEDATE \* MERGEFORMAT", InformationKind::CreateDate),
        (r"SAVEDATE \* MERGEFORMAT", InformationKind::SaveDate),
        (r"PRINTDATE \* MERGEFORMAT", InformationKind::PrintDate),
        (r"REVNUM \* MERGEFORMAT", InformationKind::RevisionNumber),
        (r"EDITTIME \* MERGEFORMAT", InformationKind::EditTime),
        (r"NUMPAGES \* MERGEFORMAT", InformationKind::NumberOfPages),
        (r"NUMWORDS \* MERGEFORMAT", InformationKind::NumberOfWords),
        (
            r"NUMCHARS \* MERGEFORMAT",
            InformationKind::NumberOfCharacters,
        ),
    ] {
        let cached_result = format!("cached {}", kind.field_keyword());
        let field = Field::with_flags(
            instruction.to_string(),
            Some(cached_result.clone()),
            true,
            true,
        );
        let information = field.document_information().unwrap().unwrap();
        assert_eq!(information.kind(), kind);
        assert_eq!(information.instruction(), instruction);
        assert_eq!(information.cached_result(), Some(cached_result.as_str()));
        assert!(information.is_dirty());
        assert!(information.is_locked());
        assert_eq!(information.switches()[0].name(), '*');
    }
}

#[test]
fn rejects_invalid_document_information_field_semantics() {
    for instruction in [
        "TITLE unexpected",
        r#"AUTHOR "unterminated"#,
        r"COMMENTS \",
        r"LASTSAVEDBY \* MERGEFORMAT unexpected",
        "NUMWORDS unexpected",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.document_information().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "TITLE \\* {}",
            "x".repeat(MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.document_information().is_err());
    assert_eq!(
        Field::new("SAVEDATE".to_string(), None, false)
            .document_information()
            .unwrap()
            .unwrap()
            .kind(),
        InformationKind::SaveDate
    );
    assert!(
        Field::new("SAVEDATES".to_string(), None, false)
            .document_information()
            .unwrap()
            .is_none()
    );
}

#[test]
fn parses_document_context_fields_without_reading_paths_files_or_layout() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" FILENAME \p " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached file name</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>template \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached template</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr=" SECTIONPAGES \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached section pages</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr="FILENAMES"><w:r><w:t>not a file-name field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let extracted = Field::extract_from_document(xml).unwrap();
    assert_eq!(extracted.len(), 4);
    assert!(extracted[0].is_document_context());
    assert!(extracted[1].is_document_context());
    assert!(extracted[2].is_document_context());
    assert!(!extracted[3].is_document_context());

    let file_name = extracted[0].document_context().unwrap().unwrap();
    assert_eq!(file_name.kind(), ContextKind::FileName);
    assert_eq!(file_name.cached_result(), Some("cached file name"));
    assert!(file_name.is_dirty());
    assert!(file_name.is_locked());
    assert!(file_name.has_switch('p'));

    let template = extracted[1].document_context().unwrap().unwrap();
    assert_eq!(template.kind(), ContextKind::Template);
    assert_eq!(template.cached_result(), Some("cached template"));
    assert!(template.is_dirty());
    assert!(template.is_locked());
    assert!(template.has_switch('*'));

    let section_pages = extracted[2].document_context().unwrap().unwrap();
    assert_eq!(section_pages.kind(), ContextKind::SectionPages);
    assert_eq!(section_pages.cached_result(), Some("cached section pages"));
    assert!(section_pages.is_dirty());
    assert!(section_pages.is_locked());
    assert!(section_pages.has_switch('*'));
    assert!(extracted[3].document_context().unwrap().is_none());

    for (instruction, kind, switch_name) in [
        (r"FILENAME \p", ContextKind::FileName, 'p'),
        (r"TEMPLATE \* MERGEFORMAT", ContextKind::Template, '*'),
        (r#"DATE \@ "opaque date format""#, ContextKind::Date, '@'),
        (r#"TIME \@ "opaque time format""#, ContextKind::Time, '@'),
        (r"PAGE \* MERGEFORMAT", ContextKind::Page, '*'),
        (r"FILESIZE \* MERGEFORMAT", ContextKind::FileSize, '*'),
        (r"SECTION \* MERGEFORMAT", ContextKind::Section, '*'),
        (
            r"SECTIONPAGES \* MERGEFORMAT",
            ContextKind::SectionPages,
            '*',
        ),
    ] {
        let cached_result = format!("cached {}", kind.field_keyword());
        let field = Field::with_flags(
            instruction.to_string(),
            Some(cached_result.clone()),
            true,
            true,
        );
        let context = field.document_context().unwrap().unwrap();
        assert_eq!(context.kind(), kind);
        assert_eq!(context.instruction(), instruction);
        assert_eq!(context.cached_result(), Some(cached_result.as_str()));
        assert!(context.is_dirty());
        assert!(context.is_locked());
        assert!(context.has_switch(switch_name));
    }
}

#[test]
fn rejects_invalid_document_context_field_semantics() {
    for instruction in [
        "FILENAME unexpected",
        r"TEMPLATE \",
        r"FILENAME \ ",
        "PAGE unexpected",
        "SECTIONPAGES unexpected",
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.document_context().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "FILENAME \\* {}",
            "x".repeat(MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.document_context().is_err());
    assert!(
        Field::new("FILENAMES".to_string(), None, false)
            .document_context()
            .unwrap()
            .is_none()
    );
    assert!(
        Field::new("PAGES".to_string(), None, false)
            .document_context()
            .unwrap()
            .is_none()
    );
    assert!(
        Field::new("SECTIONPAGE".to_string(), None, false)
            .document_context()
            .unwrap()
            .is_none()
    );
}

#[test]
fn parses_macro_button_fields_without_resolving_or_executing_targets() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MACROBUTTON &quot;Never Run&quot; &quot;Click here&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached button</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>MACROBUTTON NoMacro "Click again"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached second button</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="MACROBUTTONS NeverRun Button"><w:r><w:t>not a macro button</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_macro_button());
    assert!(fields[1].is_macro_button());
    assert!(!fields[2].is_macro_button());

    let first = fields[0].macro_button().unwrap().unwrap();
    assert_eq!(first.macro_name(), "Never Run");
    assert_eq!(first.display_text(), "Click here");
    assert_eq!(first.cached_result(), Some("cached button"));
    assert!(first.is_dirty());
    assert!(first.is_locked());

    let second = fields[1].macro_button().unwrap().unwrap();
    assert_eq!(second.macro_name(), "NoMacro");
    assert_eq!(second.display_text(), "Click again");
    assert_eq!(second.cached_result(), Some("cached second button"));
    assert!(second.is_dirty());
    assert!(second.is_locked());
    assert!(fields[2].macro_button().unwrap().is_none());
}

#[test]
fn rejects_invalid_macro_button_field_semantics() {
    let missing_name = Field::new("MACROBUTTON".to_string(), None, false);
    assert!(missing_name.macro_button().is_err());

    let empty_name = Field::new(r#"MACROBUTTON "" Button"#.to_string(), None, false);
    assert!(empty_name.macro_button().is_err());

    let missing_button = Field::new("MACROBUTTON NeverRun".to_string(), None, false);
    assert!(missing_button.macro_button().is_err());

    let empty_button = Field::new(r#"MACROBUTTON NeverRun """#.to_string(), None, false);
    assert!(empty_button.macro_button().is_err());

    let extra_argument = Field::new(
        "MACROBUTTON NeverRun Button unexpected".to_string(),
        None,
        false,
    );
    assert!(extra_argument.macro_button().is_err());

    let unsupported_switch = Field::new(
        r#"MACROBUTTON NeverRun Button \* MERGEFORMAT"#.to_string(),
        None,
        false,
    );
    assert!(unsupported_switch.macro_button().is_err());
}

#[test]
fn parses_go_to_button_fields_without_resolving_or_navigating_to_targets() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" GOTOBUTTON MyBookmark &quot;Jump to bookmark&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached bookmark button</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>GOTOBUTTON "f 2" Footnote</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached footnote button</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="GOTOBUTTONS MyBookmark Button"><w:r><w:t>not a button</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_go_to_button());
    assert!(fields[1].is_go_to_button());
    assert!(!fields[2].is_go_to_button());

    let first = fields[0].go_to_button().unwrap().unwrap();
    assert_eq!(first.target(), "MyBookmark");
    assert_eq!(first.button_text(), "Jump to bookmark");
    assert_eq!(first.cached_result(), Some("cached bookmark button"));
    assert!(first.is_dirty());
    assert!(first.is_locked());

    let second = fields[1].go_to_button().unwrap().unwrap();
    assert_eq!(second.target(), "f 2");
    assert_eq!(second.button_text(), "Footnote");
    assert_eq!(second.cached_result(), Some("cached footnote button"));
    assert!(second.is_dirty());
    assert!(second.is_locked());
    assert!(fields[2].go_to_button().unwrap().is_none());
}

#[test]
fn rejects_invalid_go_to_button_field_semantics() {
    let missing_target = Field::new("GOTOBUTTON".to_string(), None, false);
    assert!(missing_target.go_to_button().is_err());

    let empty_target = Field::new(r#"GOTOBUTTON "" Button"#.to_string(), None, false);
    assert!(empty_target.go_to_button().is_err());

    let missing_button = Field::new("GOTOBUTTON Destination".to_string(), None, false);
    assert!(missing_button.go_to_button().is_err());

    let empty_button = Field::new(r#"GOTOBUTTON Destination """#.to_string(), None, false);
    assert!(empty_button.go_to_button().is_err());

    let extra_argument = Field::new(
        "GOTOBUTTON Destination Button unexpected".to_string(),
        None,
        false,
    );
    assert!(extra_argument.go_to_button().is_err());

    let unsupported_switch = Field::new(
        r#"GOTOBUTTON Destination Button \* MERGEFORMAT"#.to_string(),
        None,
        false,
    );
    assert!(unsupported_switch.go_to_button().is_err());
}

#[test]
fn parses_active_content_fields_without_loading_or_activating_them() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ADDIN opaque-add-in-data " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached add-in result</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>control opaque-ocx-metadata</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached control result</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="HTMLCONTROL opaque-html-control-metadata">
                <w:r><w:t>cached html result</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr="ADDINS not-an-add-in"><w:r><w:t>not active content</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_add_in_field());
    assert!(fields[0].is_active_content_field());
    assert!(fields[1].is_control_field());
    assert!(fields[1].is_active_content_field());
    assert!(fields[2].is_html_control_field());
    assert!(fields[2].is_active_content_field());
    assert!(!fields[3].is_active_content_field());

    let add_in = fields[0].active_content_field().unwrap().unwrap();
    assert_eq!(add_in.kind(), ActiveContentKind::AddIn);
    assert_eq!(add_in.cached_result(), Some("cached add-in result"));
    assert!(add_in.is_dirty());
    assert!(add_in.is_locked());

    let ocx = fields[1].active_content_field().unwrap().unwrap();
    assert_eq!(ocx.kind(), ActiveContentKind::OcxControl);
    assert_eq!(ocx.cached_result(), Some("cached control result"));
    assert!(ocx.is_dirty());
    assert!(ocx.is_locked());

    let html = fields[2].active_content_field().unwrap().unwrap();
    assert_eq!(html.kind(), ActiveContentKind::HtmlControl);
    assert_eq!(html.cached_result(), Some("cached html result"));
    assert!(!html.is_dirty());
    assert!(!html.is_locked());
    assert!(fields[3].active_content_field().unwrap().is_none());
}

#[test]
fn parses_inert_print_fields_without_interpreting_or_sending_printer_commands() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" PRINT &quot;ESC&amp;l1O&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached printer result</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>print \p 2 "0 0 moveto"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached PostScript result</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="PRINTS not-a-print-field"><w:r><w:t>not printer metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_print_field());
    assert!(fields[1].is_print_field());
    assert!(!fields[2].is_print_field());

    let printer = fields[0].print_field().unwrap().unwrap();
    assert_eq!(printer.printer_instructions(), r#""ESC&l1O""#);
    assert_eq!(printer.cached_result(), Some("cached printer result"));
    assert!(printer.is_dirty());
    assert!(printer.is_locked());

    let postscript = fields[1].print_field().unwrap().unwrap();
    assert_eq!(postscript.printer_instructions(), r#"\p 2 "0 0 moveto""#);
    assert_eq!(postscript.cached_result(), Some("cached PostScript result"));
    assert!(postscript.is_dirty());
    assert!(postscript.is_locked());
    assert!(fields[2].print_field().unwrap().is_none());
}

#[test]
fn parses_inert_embed_fields_without_loading_or_activating_objects() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" EMBED Excel.Sheet.12 \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached worksheet object</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>embed "Equation.DSMT4" \d</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached equation object</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="EMBED"><w:r><w:t>cached bare object</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="EMBEDS Excel.Sheet.12"><w:r><w:t>not an embedded object field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_embed_field());
    assert!(fields[1].is_embed_field());
    assert!(fields[2].is_embed_field());
    assert!(!fields[3].is_embed_field());

    let worksheet = fields[0].embed_field().unwrap().unwrap();
    assert_eq!(
        worksheet.object_instructions(),
        r#"Excel.Sheet.12 \* MERGEFORMAT"#
    );
    assert_eq!(worksheet.cached_result(), Some("cached worksheet object"));
    assert!(worksheet.is_dirty());
    assert!(worksheet.is_locked());

    let equation = fields[1].embed_field().unwrap().unwrap();
    assert_eq!(equation.object_instructions(), r#""Equation.DSMT4" \d"#);
    assert_eq!(equation.cached_result(), Some("cached equation object"));
    assert!(equation.is_dirty());
    assert!(equation.is_locked());

    let bare = fields[2].embed_field().unwrap().unwrap();
    assert_eq!(bare.object_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare object"));
    assert!(!bare.is_dirty());
    assert!(!bare.is_locked());
    assert!(fields[3].embed_field().unwrap().is_none());

    let too_long = Field::new(
        format!("EMBED {}", "x".repeat(MAX_EMBED_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.embed_field().is_err());
}

#[test]
fn parses_inert_barcode_fields_without_decoding_or_rendering() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" BARCODE &quot;4901234567894&quot; EAN13 \h 1440 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached EAN13 barcode</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>barcode "ABC-123" CODE39 \d</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached Code39 barcode</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="BARCODE"><w:r><w:t>cached bare barcode</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="BARCODES 4901234567894"><w:r><w:t>not barcode metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_barcode_field());
    assert!(fields[1].is_barcode_field());
    assert!(fields[2].is_barcode_field());
    assert!(!fields[3].is_barcode_field());

    let ean13 = fields[0].barcode_field().unwrap().unwrap();
    assert_eq!(
        ean13.barcode_instructions(),
        r#""4901234567894" EAN13 \h 1440"#
    );
    assert_eq!(ean13.cached_result(), Some("cached EAN13 barcode"));
    assert!(ean13.is_dirty());
    assert!(ean13.is_locked());

    let code_39 = fields[1].barcode_field().unwrap().unwrap();
    assert_eq!(code_39.barcode_instructions(), r#""ABC-123" CODE39 \d"#);
    assert_eq!(code_39.cached_result(), Some("cached Code39 barcode"));
    assert!(code_39.is_dirty());
    assert!(code_39.is_locked());

    let bare = fields[2].barcode_field().unwrap().unwrap();
    assert_eq!(bare.barcode_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare barcode"));
    assert!(!bare.is_dirty());
    assert!(!bare.is_locked());
    assert!(fields[3].barcode_field().unwrap().is_none());

    let too_long = Field::new(
        format!(
            "BARCODE {}",
            "x".repeat(MAX_BARCODE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.barcode_field().is_err());
}

#[test]
fn parses_inert_bidi_outline_fields_without_resolving_numbering_or_layout() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" BIDIOUTLINE \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached bidi outline number</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>bidioutline</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached bare bidi outline</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="BIDIOUTLINES"><w:r><w:t>not bidi outline metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_bidi_outline_field());
    assert!(fields[1].is_bidi_outline_field());
    assert!(!fields[2].is_bidi_outline_field());

    let outline = fields[0].bidi_outline_field().unwrap().unwrap();
    assert_eq!(outline.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(outline.cached_result(), Some("cached bidi outline number"));
    assert!(outline.is_dirty());
    assert!(outline.is_locked());

    let bare = fields[1].bidi_outline_field().unwrap().unwrap();
    assert_eq!(bare.opaque_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare bidi outline"));
    assert!(bare.is_dirty());
    assert!(bare.is_locked());
    assert!(fields[2].bidi_outline_field().unwrap().is_none());

    let too_long = Field::new(
        format!(
            "BIDIOUTLINE {}",
            "x".repeat(MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.bidi_outline_field().is_err());
}

#[test]
fn parses_inert_shape_fields_without_linking_or_rendering_drawings() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SHAPE \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached drawing anchor</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>shape</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached bare drawing anchor</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SHAPES"><w:r><w:t>not shape metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_shape_field());
    assert!(fields[1].is_shape_field());
    assert!(!fields[2].is_shape_field());

    let shape = fields[0].shape_field().unwrap().unwrap();
    assert_eq!(shape.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(shape.cached_result(), Some("cached drawing anchor"));
    assert!(shape.is_dirty());
    assert!(shape.is_locked());

    let bare = fields[1].shape_field().unwrap().unwrap();
    assert_eq!(bare.opaque_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare drawing anchor"));
    assert!(bare.is_dirty());
    assert!(bare.is_locked());
    assert!(fields[2].shape_field().unwrap().is_none());

    let too_long = Field::new(
        format!("SHAPE {}", "x".repeat(MAX_SHAPE_FIELD_INSTRUCTION_BYTES)),
        None,
        false,
    );
    assert!(too_long.shape_field().is_err());
}

#[test]
fn parses_inert_legacy_form_fields_without_reading_or_filling_forms() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" FORMTEXT \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:ffData>
                    <w:name w:val="TextInput"/>
                    <w:entryMacro w:val="NeverRun"/>
                    <w:textInput><w:maxLength w:val="10"/></w:textInput>
                </w:ffData>
                <w:r><w:t>cached text field</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>formcheckbox</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached checkbox</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>FORMDROPDOWN \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached drop-down selection</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="FORMTEXTUAL"><w:r><w:t>not a form field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_legacy_form_field());
    assert!(fields[1].is_legacy_form_field());
    assert!(fields[2].is_legacy_form_field());
    assert!(!fields[3].is_legacy_form_field());

    let text = fields[0].legacy_form_field().unwrap().unwrap();
    assert_eq!(text.kind(), LegacyFormKind::Text);
    assert_eq!(text.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(text.cached_result(), Some("cached text field"));
    assert!(text.is_dirty());
    assert!(text.is_locked());

    let checkbox = fields[1].legacy_form_field().unwrap().unwrap();
    assert_eq!(checkbox.kind(), LegacyFormKind::CheckBox);
    assert_eq!(checkbox.opaque_instructions(), "");
    assert_eq!(checkbox.cached_result(), Some("cached checkbox"));
    assert!(checkbox.is_dirty());
    assert!(checkbox.is_locked());

    let drop_down = fields[2].legacy_form_field().unwrap().unwrap();
    assert_eq!(drop_down.kind(), LegacyFormKind::DropDown);
    assert_eq!(drop_down.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(
        drop_down.cached_result(),
        Some("cached drop-down selection")
    );
    assert!(drop_down.is_dirty());
    assert!(drop_down.is_locked());
    assert!(fields[3].legacy_form_field().unwrap().is_none());

    let too_long = Field::new(
        format!(
            "FORMTEXT {}",
            "x".repeat(MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.legacy_form_field().is_err());
}

#[test]
fn parses_inert_private_fields_without_conversion_or_layout() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" PRIVATE \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>opaque converter payload</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>private</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached bare private payload</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="PRIVATELY"><w:r><w:t>not private metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_private_field());
    assert!(fields[1].is_private_field());
    assert!(!fields[2].is_private_field());

    let private = fields[0].private_field().unwrap().unwrap();
    assert_eq!(private.opaque_instructions(), r#"\* MERGEFORMAT"#);
    assert_eq!(private.cached_result(), Some("opaque converter payload"));
    assert!(private.is_dirty());
    assert!(private.is_locked());

    let bare = fields[1].private_field().unwrap().unwrap();
    assert_eq!(bare.opaque_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare private payload"));
    assert!(bare.is_dirty());
    assert!(bare.is_locked());
    assert!(fields[2].private_field().unwrap().is_none());

    let too_long = Field::new(
        format!(
            "PRIVATE {}",
            "x".repeat(MAX_PRIVATE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.private_field().is_err());
}

#[test]
fn parses_inert_database_fields_without_connecting_or_executing() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" DATABASE \d &quot;unavailable.csv&quot; \c &quot;DSN=NeverConnect&quot; \s &quot;SELECT * FROM Customers&quot; \h " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached database table</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>database</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached bare database table</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="DATABASES"><w:r><w:t>not database metadata</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_database_field());
    assert!(fields[1].is_database_field());
    assert!(!fields[2].is_database_field());

    let database = fields[0].database_field().unwrap().unwrap();
    assert_eq!(
        database.opaque_instructions(),
        r#"\d "unavailable.csv" \c "DSN=NeverConnect" \s "SELECT * FROM Customers" \h"#
    );
    assert_eq!(database.cached_result(), Some("cached database table"));
    assert!(database.is_dirty());
    assert!(database.is_locked());

    let bare = fields[1].database_field().unwrap().unwrap();
    assert_eq!(bare.opaque_instructions(), "");
    assert_eq!(bare.cached_result(), Some("cached bare database table"));
    assert!(bare.is_dirty());
    assert!(bare.is_locked());
    assert!(fields[2].database_field().unwrap().is_none());

    let too_long = Field::new(
        format!(
            "DATABASE {}",
            "x".repeat(MAX_DATABASE_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.database_field().is_err());
}

#[test]
fn parses_auto_text_fields_without_lookup_or_insertion() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" GLOSSARY &quot;Legacy Clause&quot; \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached glossary entry</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>autotext "Reusable Clause" \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached auto text entry</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTOTEXTLIST display"><w:r><w:t>not an auto text field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_auto_text_field());
    assert!(fields[1].is_auto_text_field());
    assert!(!fields[2].is_auto_text_field());

    let glossary = fields[0].auto_text_field().unwrap().unwrap();
    assert_eq!(glossary.kind(), AutoTextKind::Glossary);
    assert_eq!(glossary.entry_name(), "Legacy Clause");
    assert_eq!(glossary.unknown_switches().len(), 2);
    assert_eq!(glossary.unknown_switches()[0].name(), '*');
    assert_eq!(
        glossary.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(glossary.unknown_switches()[1].name(), 'q');
    assert_eq!(glossary.unknown_switches()[1].argument(), Some("opaque"));
    assert_eq!(glossary.cached_result(), Some("cached glossary entry"));
    assert!(glossary.is_dirty());
    assert!(glossary.is_locked());

    let auto_text = fields[1].auto_text_field().unwrap().unwrap();
    assert_eq!(auto_text.kind(), AutoTextKind::AutoText);
    assert_eq!(auto_text.entry_name(), "Reusable Clause");
    assert_eq!(auto_text.unknown_switches().len(), 1);
    assert_eq!(auto_text.unknown_switches()[0].name(), '*');
    assert_eq!(
        auto_text.unknown_switches()[0].argument(),
        Some("MERGEFORMAT")
    );
    assert_eq!(auto_text.cached_result(), Some("cached auto text entry"));
    assert!(auto_text.is_dirty());
    assert!(auto_text.is_locked());
    assert!(fields[2].auto_text_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_auto_text_fields_without_lookup_or_insertion() {
    for instruction in [
        "GLOSSARY",
        r#"GLOSSARY ""#,
        "GLOSSARY Entry unexpected",
        r#"GLOSSARY Entry \"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.auto_text_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "AUTOTEXT Entry {}",
            "x".repeat(MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.auto_text_field().is_err());
}

#[test]
fn parses_auto_text_list_fields_without_selection_or_insertion() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" AUTOTEXTLIST &quot;Choose a name&quot; \s &quot;Name Style&quot; \t &quot;Right-click to select&quot; \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached selection</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>autotextlist \s NameStyle</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached style-only selection</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTOTEXTLISTS display"><w:r><w:t>not a list field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_auto_text_list_field());
    assert!(fields[1].is_auto_text_list_field());
    assert!(!fields[2].is_auto_text_list_field());

    let list = fields[0].auto_text_list_field().unwrap().unwrap();
    assert_eq!(list.display_text(), Some("Choose a name"));
    assert_eq!(
        list.options(),
        &[
            AutoTextListOption::Style("Name Style".to_string()),
            AutoTextListOption::Tip("Right-click to select".to_string()),
        ]
    );
    assert_eq!(list.unknown_switches().len(), 2);
    assert_eq!(list.unknown_switches()[0].name(), '*');
    assert_eq!(list.unknown_switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(list.unknown_switches()[1].name(), 'q');
    assert_eq!(list.unknown_switches()[1].argument(), Some("opaque"));
    assert_eq!(list.cached_result(), Some("cached selection"));
    assert!(list.is_dirty());
    assert!(list.is_locked());

    let style_only = fields[1].auto_text_list_field().unwrap().unwrap();
    assert_eq!(style_only.display_text(), None);
    assert_eq!(
        style_only.options(),
        &[AutoTextListOption::Style("NameStyle".to_string())]
    );
    assert_eq!(
        style_only.cached_result(),
        Some("cached style-only selection")
    );
    assert!(style_only.is_dirty());
    assert!(style_only.is_locked());
    assert!(fields[2].auto_text_list_field().unwrap().is_none());

    let empty_display = Field::new(r#"AUTOTEXTLIST "" \s NameStyle"#.to_string(), None, false)
        .auto_text_list_field()
        .unwrap()
        .unwrap();
    assert_eq!(empty_display.display_text(), Some(""));
}

#[test]
fn rejects_invalid_auto_text_list_fields_without_selection_or_insertion() {
    for instruction in [
        r#"AUTOTEXTLIST \s"#,
        r#"AUTOTEXTLIST \t"#,
        "AUTOTEXTLIST display unexpected",
        r#"AUTOTEXTLIST \"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.auto_text_list_field().is_err(), "{instruction}");
    }

    let too_long = Field::new(
        format!(
            "AUTOTEXTLIST {}",
            "x".repeat(MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES)
        ),
        None,
        false,
    );
    assert!(too_long.auto_text_list_field().is_err());
}

#[test]
fn parses_user_identity_fields_without_reading_host_identity() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" USERADDRESS &quot;10 Top Secret Lane&quot; \* Upper " w:dirty="true" w:fldLock="on">
                <w:r><w:t>10 TOP SECRET LANE</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>userinitials \* Lower</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>dw</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="USERNAME &quot;Ada Lovelace&quot; \* FirstCap"><w:r><w:t>Ada Lovelace</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="USERNAMES Ada"><w:r><w:t>not a user identity field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 4);
    assert!(fields[0].is_user_address());
    assert!(fields[0].is_user_identity_field());
    assert!(fields[1].is_user_initials());
    assert!(fields[1].is_user_identity_field());
    assert!(fields[2].is_user_name());
    assert!(fields[2].is_user_identity_field());
    assert!(!fields[3].is_user_identity_field());

    let address = fields[0].user_identity_field().unwrap().unwrap();
    assert_eq!(address.kind(), UserIdentityKind::Address);
    assert_eq!(address.override_value(), Some("10 Top Secret Lane"));
    assert_eq!(address.formatting(), Some(UserIdentityFormat::Upper));
    assert_eq!(address.cached_result(), Some("10 TOP SECRET LANE"));
    assert!(address.is_dirty());
    assert!(address.is_locked());

    let initials = fields[1].user_identity_field().unwrap().unwrap();
    assert_eq!(initials.kind(), UserIdentityKind::Initials);
    assert_eq!(initials.override_value(), None);
    assert_eq!(initials.formatting(), Some(UserIdentityFormat::Lower));
    assert_eq!(initials.cached_result(), Some("dw"));
    assert!(initials.is_dirty());
    assert!(initials.is_locked());

    let name = fields[2].user_identity_field().unwrap().unwrap();
    assert_eq!(name.kind(), UserIdentityKind::Name);
    assert_eq!(name.override_value(), Some("Ada Lovelace"));
    assert_eq!(name.formatting(), Some(UserIdentityFormat::FirstCap));
    assert_eq!(name.cached_result(), Some("Ada Lovelace"));
    assert!(!name.is_dirty());
    assert!(!name.is_locked());
    assert!(fields[3].user_identity_field().unwrap().is_none());
}

#[test]
fn rejects_invalid_user_identity_field_semantics() {
    let missing_format = Field::new("USERADDRESS \\*".to_string(), None, false);
    assert!(missing_format.user_identity_field().is_err());

    let unsupported_format = Field::new("USERINITIALS \\* Title".to_string(), None, false);
    assert!(unsupported_format.user_identity_field().is_err());

    let duplicate_format = Field::new("USERNAME \\* Upper \\* Lower".to_string(), None, false);
    assert!(duplicate_format.user_identity_field().is_err());

    let unsupported_switch = Field::new("USERNAME Ada \\l 1033".to_string(), None, false);
    assert!(unsupported_switch.user_identity_field().is_err());

    let unexpected_text = Field::new("USERADDRESS Ada Lovelace".to_string(), None, false);
    assert!(unexpected_text.user_identity_field().is_err());

    let blank_override = Field::new(r#"USERNAME "" \* Caps"#.to_string(), None, false);
    let blank_override = blank_override.user_identity_field().unwrap().unwrap();
    assert_eq!(blank_override.override_value(), Some(""));
    assert_eq!(blank_override.formatting(), Some(UserIdentityFormat::Caps));
}

#[test]
fn parses_inert_advance_fields_without_changing_layout() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ADVANCE \u 6 \d 12 \l 20 \r -4 \x 150 \y &quot;72&quot; \d -3 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached placement</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr="ADVANCER \u 6"><w:r><w:t>not an advance field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 2);
    assert!(fields[0].is_advance_field());
    assert!(!fields[1].is_advance_field());

    let advance = fields[0].advance_field().unwrap().unwrap();
    let adjustments = advance
        .adjustments()
        .iter()
        .map(|adjustment| (adjustment.operation(), adjustment.points()))
        .collect::<Vec<_>>();
    assert_eq!(
        adjustments,
        vec![
            (AdvanceOperation::Up, 6),
            (AdvanceOperation::Down, 12),
            (AdvanceOperation::Left, 20),
            (AdvanceOperation::Right, -4),
            (AdvanceOperation::HorizontalPosition, 150),
            (AdvanceOperation::VerticalPosition, 72),
            (AdvanceOperation::Down, -3),
        ]
    );
    assert_eq!(advance.cached_result(), Some("cached placement"));
    assert!(advance.is_dirty());
    assert!(advance.is_locked());
    assert!(fields[1].advance_field().unwrap().is_none());

    let no_adjustments = Field::new("aDvAnCe".to_string(), None, false);
    let no_adjustments = no_adjustments.advance_field().unwrap().unwrap();
    assert!(no_adjustments.adjustments().is_empty());
    assert_eq!(no_adjustments.cached_result(), None);
}

#[test]
fn rejects_invalid_advance_field_semantics() {
    for instruction in [
        r#"ADVANCE \d"#,
        r#"ADVANCE \z 10"#,
        r#"ADVANCE \x 1.5"#,
        r#"ADVANCE \u 9223372036854775808"#,
        "ADVANCE 12",
        r#"ADVANCE \d 6 trailing"#,
    ] {
        let field = Field::new(instruction.to_string(), None, false);
        assert!(field.advance_field().is_err(), "{instruction}");
    }
}

#[test]
fn parses_table_of_authorities_and_entry_fields() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TOA \c 0 \b &quot;Authorities&quot; \p \f \d &quot;-&quot; \s &quot;Chapter&quot; \e &quot;, &quot; \g &quot;&#x2013;&quot; \h \l &quot;, &quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Cases</w:t><w:tab/><w:t>1, 5</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>TA\l&quot;Long citation&quot;\s &quot;Short citation&quot; \c 1 \b \i</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>hidden citation marker</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="TABLE \c 1"><w:r><w:t>not an authority table</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_table_of_authorities());
    assert!(fields[1].is_table_of_authorities_entry());
    assert!(!fields[2].is_table_of_authorities());
    assert!(!fields[2].is_table_of_authorities_entry());

    let toa = fields[0].table_of_authorities().unwrap().unwrap();
    assert_eq!(toa.cached_result(), Some("Cases\t1, 5"));
    assert!(toa.is_dirty());
    assert!(toa.is_locked());
    assert_eq!(toa.category().unwrap(), Some(0));
    assert_eq!(toa.bookmark().unwrap(), Some("Authorities"));
    assert!(toa.uses_passim());
    assert!(toa.keeps_entry_formatting());
    assert_eq!(toa.sequence_page_separator().unwrap(), Some("-"));
    assert_eq!(toa.sequence_name().unwrap(), Some("Chapter"));
    assert_eq!(toa.entry_page_separator().unwrap(), Some(", "));
    assert_eq!(toa.page_range_separator().unwrap(), Some("–"));
    assert!(toa.includes_category_headers());
    assert_eq!(toa.page_number_separator().unwrap(), Some(", "));

    let entry = fields[1].table_of_authorities_entry().unwrap().unwrap();
    assert_eq!(entry.cached_result(), Some("hidden citation marker"));
    assert!(entry.is_dirty());
    assert!(entry.is_locked());
    assert_eq!(entry.long_citation().unwrap(), Some("Long citation"));
    assert_eq!(entry.short_citation().unwrap(), Some("Short citation"));
    assert_eq!(entry.category().unwrap(), Some(1));
    assert!(entry.is_bold());
    assert!(entry.is_italic());
}

#[test]
fn parses_index_and_index_entry_fields() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" INDEX \b Scope \c 2 \d &quot;.&quot; \e &quot;; &quot; \f &quot;topics&quot; \g &quot; to &quot; \h &quot;A&quot; \k &quot;: &quot; \l &quot; / &quot; \o &quot;P&quot; \p a-m \r \s Chapter \y \z 1033 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Rulers</w:t><w:tab/><w:t>4</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>XE&quot;Machiavelli: The Prince&quot;\b\i\f &quot;topics&quot; \r IndexRange \t &quot;See Rulers&quot; \y &quot;ma&quot;</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>hidden index marker</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="INDEXENTRY \f ignored"><w:r><w:t>not an index</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
    let fields = Field::extract_from_document(xml).unwrap();
    assert_eq!(fields.len(), 3);
    assert!(fields[0].is_index());
    assert!(fields[1].is_index_entry());
    assert!(!fields[2].is_index());
    assert!(!fields[2].is_index_entry());

    let index = fields[0].index().unwrap().unwrap();
    assert_eq!(index.cached_result(), Some("Rulers\t4"));
    assert!(index.is_dirty());
    assert!(index.is_locked());
    assert_eq!(index.bookmark().unwrap(), Some("Scope"));
    assert_eq!(index.columns().unwrap(), Some(2));
    assert_eq!(index.sequence_page_separator().unwrap(), Some("."));
    assert_eq!(index.entry_page_separator().unwrap(), Some("; "));
    assert_eq!(index.entry_identifier().unwrap(), Some("topics"));
    assert_eq!(index.page_range_separator().unwrap(), Some(" to "));
    assert_eq!(index.alphabetic_group_heading().unwrap(), Some("A"));
    assert_eq!(index.cross_reference_separator().unwrap(), Some(": "));
    assert_eq!(index.page_reference_separator().unwrap(), Some(" / "));
    assert_eq!(index.sort_order().unwrap(), Some(IndexOrder::Pronunciation));
    assert_eq!(index.letter_range().unwrap(), Some("a-m"));
    assert!(index.runs_subentries_inline());
    assert_eq!(index.sequence_name().unwrap(), Some("Chapter"));
    assert!(index.uses_yomi());
    assert_eq!(index.language_id().unwrap(), Some("1033"));

    let entry = fields[1].index_entry().unwrap().unwrap();
    assert_eq!(entry.cached_result(), Some("hidden index marker"));
    assert!(entry.is_dirty());
    assert!(entry.is_locked());
    assert_eq!(entry.entry(), "Machiavelli: The Prince");
    assert!(entry.is_bold());
    assert!(entry.is_italic());
    assert_eq!(entry.entry_identifier().unwrap(), Some("topics"));
    assert_eq!(entry.page_range_bookmark().unwrap(), Some("IndexRange"));
    assert_eq!(entry.cross_reference().unwrap(), Some("See Rulers"));
    assert_eq!(entry.yomi().unwrap(), Some("ma"));
}

#[test]
fn rejects_invalid_table_of_authorities_semantics() {
    let invalid_toa = Field::new(r#"TOA \c 17"#.to_string(), None, false);
    let toa = invalid_toa.table_of_authorities().unwrap().unwrap();
    assert!(toa.category().is_err());

    let invalid_entry = Field::new(r#"TA \c 0"#.to_string(), None, false);
    let entry = invalid_entry.table_of_authorities_entry().unwrap().unwrap();
    assert!(entry.category().is_err());

    let duplicate = Field::new(r#"TOA \b "a" \b "b""#.to_string(), None, false);
    let toa = duplicate.table_of_authorities().unwrap().unwrap();
    assert!(toa.bookmark().is_err());
}

#[test]
fn rejects_invalid_citation_and_bibliography_field_semantics() {
    let missing_source = Field::new("CITATION \\l 1033".to_string(), None, false);
    assert!(missing_source.citation().is_err());

    let empty_source = Field::new(r#"CITATION ""#.to_string(), None, false);
    assert!(empty_source.citation().is_err());

    let missing_multisource_tag =
        Field::new("CITATION Doe2024 \\m \\l 1033".to_string(), None, false);
    assert!(missing_multisource_tag.citation().is_err());

    let empty_multisource_tag = Field::new(r#"CITATION Doe2024 \m """#.to_string(), None, false);
    assert!(empty_multisource_tag.citation().is_err());

    let malformed_bibliography = Field::new("BIBLIOGRAPHY unexpected".to_string(), None, false);
    assert!(malformed_bibliography.bibliography().is_err());
}

#[test]
fn rejects_invalid_index_field_semantics() {
    let invalid_columns = Field::new(r#"INDEX \c 5"#.to_string(), None, false);
    let index = invalid_columns.index().unwrap().unwrap();
    assert!(index.columns().is_err());

    let invalid_sort = Field::new(r#"INDEX \o "radical""#.to_string(), None, false);
    let index = invalid_sort.index().unwrap().unwrap();
    assert!(index.sort_order().is_err());

    let missing_entry = Field::new(r#"XE \b"#.to_string(), None, false);
    assert!(missing_entry.index_entry().is_err());
    let empty_entry = Field::new(r#"XE """#.to_string(), None, false);
    assert!(empty_entry.index_entry().is_err());

    let duplicate_identifier = Field::new(
        r#"XE "topic" \f "first" \f "second""#.to_string(),
        None,
        false,
    );
    let entry = duplicate_identifier.index_entry().unwrap().unwrap();
    assert!(entry.entry_identifier().is_err());
}

#[test]
fn rejects_malformed_toc_switches_and_level_ranges() {
    let non_toc = Field::new("TOCENTRY \\f ignored".to_string(), None, false);
    assert!(!non_toc.is_table_of_contents());
    assert!(non_toc.table_of_contents().unwrap().is_none());

    let dangling = Field::new("TOC \\".to_string(), None, false);
    assert!(dangling.table_of_contents().is_err());
    let unterminated = Field::new(r#"TOC \o "1-3"#.to_string(), None, false);
    assert!(unterminated.table_of_contents().is_err());

    let invalid_levels = Field::new(r#"TOC \o "3-1""#.to_string(), None, false);
    let toc = invalid_levels.table_of_contents().unwrap().unwrap();
    assert!(toc.heading_style_levels().is_err());
    assert!(TocLevelRange::new(0, 1).is_err());
    assert!(TocLevelRange::new(1, 10).is_err());
}

#[test]
fn rejects_simple_fields_without_instructions() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:fldSimple><w:r><w:t>result</w:t></w:r></w:fldSimple></w:p></w:body></w:document>"#;
    assert!(Field::extract_from_document(xml).is_err());
}
