//! Layout, numbering, typography, and positioning field semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
