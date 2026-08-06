//! Table-of-contents, bibliography, authority, and index field semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
