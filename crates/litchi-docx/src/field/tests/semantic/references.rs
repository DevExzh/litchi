//! Reference, link, DDE, and external-source field semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
