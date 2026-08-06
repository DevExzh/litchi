//! Mail-merge instruction families and inert metadata parsing.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
