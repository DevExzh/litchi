use super::*;

#[test]
fn saves_and_reopens_inline_and_display_office_math() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let inline = crate::OfficeMath::text("x + y");
    let display =
        crate::OfficeMath::from_xml("<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        document
            .add_paragraph()
            .add_inline_office_math(inline.clone());
        document.add_display_office_math(display.clone());
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let document_xml =
        std::str::from_utf8(reopened.opc.get_part(&document_uri).unwrap().blob()).unwrap();
    let document_opening = &document_xml[..document_xml.find("><w:body>").unwrap()];
    assert!(
        document_opening
            .contains("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"")
    );
    let paragraphs = reopened.document().unwrap().paragraphs().unwrap();
    assert_eq!(paragraphs[0].inline_office_math().unwrap(), vec![inline]);
    assert_eq!(paragraphs[1].display_office_math().unwrap(), vec![display]);
}

#[test]
fn writes_and_rediscovers_distinct_watermarks() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    let mut watermark = crate::Watermark::text("INTERNAL");
    watermark.set_font("Aptos");
    watermark.set_color("808080");
    package.document_mut().unwrap().set_watermark(watermark);
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let watermarks = reopened.document().unwrap().watermarks().unwrap();

    assert_eq!(watermarks.len(), 1);
    assert_eq!(watermarks[0].get_text(), "INTERNAL");
    assert_eq!(watermarks[0].font(), "Aptos");
    assert_eq!(watermarks[0].color(), "#808080");
}

#[test]
fn writes_and_discovers_typed_table_of_contents_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        document.add_heading("Overview", 1).unwrap();
        document
            .add_toc(
                crate::TableOfContents::new()
                    .heading_levels(1, 4)
                    .hyperlinks(true),
            )
            .unwrap();
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let toc = document.table_of_contents().unwrap();
    assert_eq!(document.table_of_contents_count().unwrap(), 1);
    assert_eq!(toc.len(), 1);
    assert!(toc[0].includes_hyperlinks());
    assert!(toc[0].hides_page_numbers_in_web_layout());
    assert_eq!(
        toc[0].heading_style_levels().unwrap(),
        vec![crate::TocLevelRange::new(1, 4).unwrap()]
    );
}

#[test]
fn all_supported_headings_resolve_in_the_saved_style_catalog() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for level in 0..=9 {
            document
                .add_heading(&format!("Heading {level}"), level)
                .unwrap();
        }
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let ids = document
        .paragraphs()
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.style_id().unwrap().unwrap())
        .collect::<Vec<_>>();
    let expected = std::iter::once("Title".to_string())
        .chain((1..=9).map(|level| format!("Heading{level}")))
        .collect::<Vec<_>>();
    assert_eq!(ids, expected);

    let mut styles = document.styles().unwrap();
    let outlines = [
        None,
        Some(crate::styles::Outline::H1),
        Some(crate::styles::Outline::H2),
        Some(crate::styles::Outline::H3),
        Some(crate::styles::Outline::H4),
        Some(crate::styles::Outline::H5),
        Some(crate::styles::Outline::H6),
        Some(crate::styles::Outline::H7),
        Some(crate::styles::Outline::H8),
        Some(crate::styles::Outline::H9),
    ];
    for (id, outline) in expected.into_iter().zip(outlines) {
        let style = styles
            .get_by_id(&id)
            .unwrap()
            .unwrap_or_else(|| panic!("missing {id}"));
        assert_eq!(style.outline(), outline, "wrong outline for {id}");
    }
}

#[test]
fn writes_and_discovers_typed_table_of_contents_entry_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let entry = document.add_paragraph();
        entry.add_field(crate::writer::MutableField::with_result(
            r#"TC "Illustration 1" \f i \l 4 \n"#.to_string(),
            "cached entry".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let entries = document.table_of_contents_entries().unwrap();
    assert_eq!(document.table_of_contents_entry_count().unwrap(), 1);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].entry(), "Illustration 1");
    assert_eq!(entries[0].cached_result(), Some("cached entry"));
    assert_eq!(entries[0].list_identifier().unwrap(), Some("i"));
    assert_eq!(entries[0].level().unwrap(), Some("4"));
    assert!(entries[0].omits_page_number());
}

#[test]
fn writes_and_discovers_typed_table_of_authorities_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let authority_table = document.add_paragraph();
        authority_table.add_field(crate::writer::MutableField::with_result(
            r#"TOA \c 2 \b "Authorities" \p \h"#.to_string(),
            "Statutes\t3".to_string(),
        ));
        let entry = document.add_paragraph();
        entry.add_field(crate::writer::MutableField::with_result(
            r#"TA \l "Example Statute" \s "Example" \c 2 \b"#.to_string(),
            "hidden marker".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let authorities = document.tables_of_authorities().unwrap();
    assert_eq!(document.table_of_authorities_count().unwrap(), 1);
    assert_eq!(authorities.len(), 1);
    assert_eq!(authorities[0].category().unwrap(), Some(2));
    assert_eq!(authorities[0].bookmark().unwrap(), Some("Authorities"));
    assert!(authorities[0].uses_passim());
    assert!(authorities[0].includes_category_headers());

    let entries = document.table_of_authorities_entries().unwrap();
    assert_eq!(document.table_of_authorities_entry_count().unwrap(), 1);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].long_citation().unwrap(), Some("Example Statute"));
    assert_eq!(entries[0].short_citation().unwrap(), Some("Example"));
    assert_eq!(entries[0].category().unwrap(), Some(2));
    assert!(entries[0].is_bold());
}

#[test]
fn writes_and_discovers_typed_citation_and_bibliography_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let citation = document.add_paragraph();
        let mut primary = crate::CitationSource::new("Doe2024").unwrap();
        primary.set_prefix(Some("qtd. in".to_string())).unwrap();
        let mut citation_spec = crate::CitationFieldSpec::new(primary);
        citation_spec.set_locale(Some(1033));
        citation_spec
            .add_source(crate::CitationSource::new("Smith2025").unwrap())
            .unwrap();
        citation_spec
            .set_cached_result(Some("(Doe, 2024; Smith, 2025)".to_string()))
            .unwrap();
        citation_spec.set_dirty(false);
        citation.add_field(crate::writer::MutableField::citation(&citation_spec).unwrap());
        let bibliography = document.add_paragraph();
        let mut bibliography_spec = crate::BibliographyFieldSpec::new();
        bibliography_spec.set_locale(Some(1033));
        bibliography_spec.set_filter(Some(crate::BibliographyFilter::Locale(1036)));
        bibliography_spec.add_source_tag("Doe2024").unwrap();
        bibliography_spec.add_source_tag("Smith2025").unwrap();
        bibliography_spec
            .set_cached_result(Some("Doe. Example work.".to_string()))
            .unwrap();
        bibliography_spec.set_dirty(false);
        bibliography
            .add_field(crate::writer::MutableField::bibliography(&bibliography_spec).unwrap());
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let citations = document.citations().unwrap();
    assert_eq!(document.citation_count().unwrap(), 1);
    assert_eq!(citations.len(), 1);
    assert_eq!(citations[0].primary_source_tag(), "Doe2024");
    assert_eq!(citations[0].source_tags(), ["Doe2024", "Smith2025"]);
    assert!(citations[0].has_switch('l'));
    assert!(citations[0].has_switch('f'));
    assert!(!citations[0].is_dirty());
    assert_eq!(
        citations[0].cached_result(),
        Some("(Doe, 2024; Smith, 2025)")
    );

    let bibliographies = document.bibliographies().unwrap();
    assert_eq!(document.bibliography_count().unwrap(), 1);
    assert_eq!(bibliographies.len(), 1);
    assert_eq!(
        bibliographies[0].cached_result(),
        Some("Doe. Example work.")
    );
    assert!(bibliographies[0].has_switch('l'));
    assert!(bibliographies[0].has_switch('f'));
    assert!(bibliographies[0].has_switch('m'));
    assert!(!bibliographies[0].is_dirty());
    assert_eq!(bibliographies[0].switches()[0].argument(), Some("1033"));
    assert_eq!(bibliographies[0].switches()[1].argument(), Some("1036"));
}

#[test]
fn writes_and_discovers_inert_document_variable_fields_without_resolution() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_field(crate::writer::MutableField::with_result(
            r#"DOCVARIABLE CustomerName \* MERGEFORMAT"#.to_string(),
            "cached customer".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.document_variable_fields().unwrap();
    assert_eq!(document.document_variable_field_count().unwrap(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].variable_name(), "CustomerName");
    assert_eq!(fields[0].cached_result(), Some("cached customer"));
    assert!(fields[0].has_switch('*'));
    assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
    assert!(
        document
            .document_variables()
            .unwrap()
            .is_none_or(|variables| variables.get("CustomerName").is_none())
    );
}

#[test]
fn writes_and_discovers_inert_document_property_fields_without_resolution() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_field(crate::writer::MutableField::with_result(
            r#"DOCPROPERTY "Project Name" \* MERGEFORMAT \@ "MMMM d, yyyy""#.to_string(),
            "cached project".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.document_property_fields().unwrap();
    assert_eq!(document.document_property_field_count().unwrap(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].property_name(), "Project Name");
    assert_eq!(fields[0].cached_result(), Some("cached project"));
    assert!(fields[0].has_switch('*'));
    assert!(fields[0].has_switch('@'));
    assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(fields[0].switches()[1].argument(), Some("MMMM d, yyyy"));
}

#[test]
fn writes_and_discovers_inert_document_information_fields_without_resolution() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let title = document.add_paragraph();
        title.add_field(crate::writer::MutableField::with_result(
            r#"TITLE \* MERGEFORMAT"#.to_string(),
            "cached title".to_string(),
        ));
        let author = document.add_paragraph();
        author.add_field(crate::writer::MutableField::with_result(
            r#"AUTHOR \@ "opaque format""#.to_string(),
            "cached author".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.document_information_fields().unwrap();
    assert_eq!(document.document_information_field_count().unwrap(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), crate::InformationKind::Title);
    assert_eq!(fields[0].cached_result(), Some("cached title"));
    assert!(fields[0].has_switch('*'));
    assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
    assert_eq!(fields[1].kind(), crate::InformationKind::Author);
    assert_eq!(fields[1].cached_result(), Some("cached author"));
    assert!(fields[1].has_switch('@'));
    assert_eq!(fields[1].switches()[0].argument(), Some("opaque format"));
}

#[test]
fn writes_and_discovers_inert_document_context_fields_without_resolution() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let file_name = document.add_paragraph();
        file_name.add_field(crate::writer::MutableField::with_result(
            r#"FILENAME \p"#.to_string(),
            "cached file name".to_string(),
        ));
        let page = document.add_paragraph();
        page.add_field(crate::writer::MutableField::with_result(
            r#"PAGE \* MERGEFORMAT"#.to_string(),
            "cached page".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.document_context_fields().unwrap();
    assert_eq!(document.document_context_field_count().unwrap(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), crate::ContextKind::FileName);
    assert_eq!(fields[0].cached_result(), Some("cached file name"));
    assert!(fields[0].has_switch('p'));
    assert_eq!(fields[1].kind(), crate::ContextKind::Page);
    assert_eq!(fields[1].cached_result(), Some("cached page"));
    assert!(fields[1].has_switch('*'));
    assert_eq!(fields[1].switches()[0].argument(), Some("MERGEFORMAT"));
}

#[test]
fn writes_and_discovers_typed_inert_merge_fields_without_merging() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_field(crate::writer::MutableField::with_result(
            r#"MERGEFIELD "Customer Region" \b "Dear " \f "!" \m \v \* MERGEFORMAT"#.to_string(),
            "cached customer".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    assert!(reopened.mail_merge_settings().unwrap().is_none());
    let document = reopened.document().unwrap();
    let fields = document.typed_merge_fields().unwrap();
    assert_eq!(document.typed_merge_field_count().unwrap(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].field_name(), "Customer Region");
    assert_eq!(fields[0].cached_result(), Some("cached customer"));
    assert!(fields[0].has_switch('b'));
    assert!(fields[0].has_switch('f'));
    assert!(fields[0].has_switch('m'));
    assert!(fields[0].has_switch('v'));
    assert_eq!(fields[0].switches()[0].argument(), Some("Dear "));
    assert_eq!(fields[0].switches()[1].argument(), Some("!"));
}

#[test]
fn writes_and_discovers_typed_inert_mail_merge_counters_without_merging() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        document
            .add_paragraph()
            .add_field(crate::writer::MutableField::with_result(
                "MERGEREC".to_string(),
                "12".to_string(),
            ));
        document
            .add_paragraph()
            .add_field(crate::writer::MutableField::with_result(
                "MERGESEQ".to_string(),
                "3".to_string(),
            ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    assert!(reopened.mail_merge_settings().unwrap().is_none());
    let document = reopened.document().unwrap();
    let counters = document.mail_merge_counters().unwrap();
    assert_eq!(document.mail_merge_counter_count().unwrap(), 2);
    assert_eq!(counters.len(), 2);
    assert_eq!(counters[0].kind(), crate::MergeCounterKind::Record);
    assert_eq!(counters[0].cached_result(), Some("12"));
    assert_eq!(counters[1].kind(), crate::MergeCounterKind::Sequence);
    assert_eq!(counters[1].cached_result(), Some("3"));
}

#[test]
fn writes_and_discovers_inert_mail_merge_next_fields_without_advancing_records() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package.document_mut().unwrap().add_paragraph().add_field(
        crate::writer::MutableField::with_result("NEXT".to_string(), "cached next".to_string()),
    );
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    assert!(reopened.mail_merge_settings().unwrap().is_none());
    let document = reopened.document().unwrap();
    let fields = document.mail_merge_next_fields().unwrap();
    assert_eq!(document.mail_merge_next_field_count().unwrap(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].cached_result(), Some("cached next"));
}

#[test]
fn writes_and_discovers_inert_conditional_mail_merge_controls_without_merging() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package.document_mut().unwrap().add_paragraph().add_field(
        crate::writer::MutableField::with_result(
            r#"SKIPIF MERGEFIELD Order < 100"#.to_string(),
            "cached skipif".to_string(),
        ),
    );
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    assert!(reopened.mail_merge_settings().unwrap().is_none());
    let document = reopened.document().unwrap();
    let controls = document.mail_merge_conditional_controls().unwrap();
    assert_eq!(document.mail_merge_conditional_control_count().unwrap(), 1);
    assert_eq!(controls.len(), 1);
    assert_eq!(controls[0].kind(), crate::MergeControlKind::SkipIf);
    assert_eq!(controls[0].comparison(), "MERGEFIELD Order < 100");
    assert_eq!(controls[0].cached_result(), Some("cached skipif"));
}

#[test]
fn writes_and_discovers_inert_if_fields_without_evaluation() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package.document_mut().unwrap().add_paragraph().add_field(
        crate::writer::MutableField::with_result(
            r#"IF 1 = 1 "yes" "no""#.to_string(),
            "yes".to_string(),
        ),
    );
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.if_fields().unwrap();
    assert_eq!(document.if_field_count().unwrap(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].expression(), r#"1 = 1 "yes" "no""#);
    assert_eq!(fields[0].cached_result(), Some("yes"));
}

#[test]
fn writes_and_discovers_inert_document_state_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for (instruction, cached_result) in [
            (
                r#"SET RecipientName "North America" \* MERGEFORMAT"#,
                "cached recipient",
            ),
            (r#"SEQ Figure FigureChapter \r 3 \* ARABIC"#, "3"),
            (r#"=SUM(ABOVE) \* MERGEFORMAT"#, "42"),
            (r#"STYLEREF "Heading 1" \n \p"#, "1 above"),
        ] {
            document
                .add_paragraph()
                .add_field(crate::writer::MutableField::with_result(
                    instruction.to_string(),
                    cached_result.to_string(),
                ));
        }
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();

    let sets = document.set_fields().unwrap();
    assert_eq!(document.set_field_count().unwrap(), 1);
    assert_eq!(sets.len(), 1);
    assert_eq!(sets[0].target_name(), "RecipientName");
    assert_eq!(sets[0].expression(), r#""North America" \* MERGEFORMAT"#);
    assert_eq!(sets[0].cached_result(), Some("cached recipient"));

    let sequences = document.sequence_fields().unwrap();
    assert_eq!(document.sequence_field_count().unwrap(), 1);
    assert_eq!(sequences.len(), 1);
    assert_eq!(sequences[0].identifier(), "Figure");
    assert_eq!(sequences[0].bookmark(), Some("FigureChapter"));
    assert_eq!(sequences[0].tail(), r#"\r 3 \* ARABIC"#);
    assert_eq!(sequences[0].cached_result(), Some("3"));

    let formulas = document.formula_fields().unwrap();
    assert_eq!(document.formula_field_count().unwrap(), 1);
    assert_eq!(formulas.len(), 1);
    assert_eq!(formulas[0].formula(), r#"SUM(ABOVE) \* MERGEFORMAT"#);
    assert_eq!(formulas[0].cached_result(), Some("42"));

    let style_references = document.style_reference_fields().unwrap();
    assert_eq!(document.style_reference_field_count().unwrap(), 1);
    assert_eq!(style_references.len(), 1);
    assert_eq!(style_references[0].style_name(), "Heading 1");
    assert_eq!(
        style_references[0].options(),
        &[
            crate::StyleOption::ParagraphNumber,
            crate::StyleOption::RelativePosition,
        ]
    );
    assert_eq!(style_references[0].cached_result(), Some("1 above"));
}

#[test]
fn writes_and_discovers_inert_bookmark_reference_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for (instruction, cached_result) in [
            (
                r#"REF "Target Bookmark" \d "-" \f \h \n \p \r \t \w"#,
                "cached reference",
            ),
            (r#"PAGEREF PageTarget \h \p"#, "12 above"),
            (r#"FTNREF FootnoteTarget \p \f"#, "1 above"),
            (r#"NOTEREF EndnoteTarget \p \f"#, "i above"),
        ] {
            document
                .add_paragraph()
                .add_field(crate::writer::MutableField::with_result(
                    instruction.to_string(),
                    cached_result.to_string(),
                ));
        }
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let references = document.reference_fields().unwrap();
    assert_eq!(document.reference_field_count().unwrap(), 4);
    assert_eq!(references.len(), 4);

    assert_eq!(references[0].kind(), crate::ReferenceKind::Reference);
    assert_eq!(references[0].bookmark(), "Target Bookmark");
    assert_eq!(
        references[0].options(),
        &[
            crate::ReferenceOption::SequencePageSeparator("-".to_string()),
            crate::ReferenceOption::ReferencedNoteContent,
            crate::ReferenceOption::Hyperlink,
            crate::ReferenceOption::ParagraphNumberWithoutContext,
            crate::ReferenceOption::RelativePosition,
            crate::ReferenceOption::ParagraphNumberRelativeContext,
            crate::ReferenceOption::SuppressNonNumberText,
            crate::ReferenceOption::ParagraphNumberFullContext,
        ]
    );
    assert_eq!(references[0].cached_result(), Some("cached reference"));

    assert_eq!(references[1].kind(), crate::ReferenceKind::PageReference);
    assert_eq!(references[1].bookmark(), "PageTarget");
    assert_eq!(
        references[1].options(),
        &[
            crate::ReferenceOption::Hyperlink,
            crate::ReferenceOption::RelativePosition,
        ]
    );
    assert_eq!(references[1].cached_result(), Some("12 above"));

    assert_eq!(
        references[2].kind(),
        crate::ReferenceKind::FootnoteReference
    );
    assert_eq!(references[2].bookmark(), "FootnoteTarget");
    assert_eq!(
        references[2].options(),
        &[
            crate::ReferenceOption::RelativePosition,
            crate::ReferenceOption::NoteMarkFormatting,
        ]
    );
    assert_eq!(references[2].cached_result(), Some("1 above"));

    assert_eq!(references[3].kind(), crate::ReferenceKind::NoteReference);
    assert_eq!(references[3].bookmark(), "EndnoteTarget");
    assert_eq!(
        references[3].options(),
        &[
            crate::ReferenceOption::RelativePosition,
            crate::ReferenceOption::NoteMarkFormatting,
        ]
    );
    assert_eq!(references[3].cached_result(), Some("i above"));
}

#[test]
fn writes_and_discovers_inert_equation_fields_without_calculation_or_rendering() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for (instruction, cached_result) in [
            (r#"EQ \o\ac(\fs24 Q,\fs16 R)"#, "cached equation"),
            (r#"EQ \f(1,2)"#, "1/2"),
            ("EQ", ""),
        ] {
            document
                .add_paragraph()
                .add_field(crate::writer::MutableField::with_result(
                    instruction.to_string(),
                    cached_result.to_string(),
                ));
        }
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let equations = document.equations().unwrap();
    assert_eq!(document.equation_count().unwrap(), 3);
    assert_eq!(equations.len(), 3);
    assert_eq!(equations[0].expression(), r#"\o\ac(\fs24 Q,\fs16 R)"#);
    assert_eq!(equations[0].cached_result(), Some("cached equation"));
    assert_eq!(equations[1].expression(), r#"\f(1,2)"#);
    assert_eq!(equations[1].cached_result(), Some("1/2"));
    assert_eq!(equations[2].expression(), "");
}

#[test]
fn writes_and_discovers_inert_hyperlink_fields_without_opening_them() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for (instruction, cached_result) in [
            (
                r#"HYPERLINK "https://example.test/a b" \l "_Toc1" \o "Stored tip" \t "_blank" \m \n"#,
                "cached external link",
            ),
            (r#"HYPERLINK \l "JumpTarget""#, "cached internal link"),
        ] {
            document
                .add_paragraph()
                .add_field(crate::writer::MutableField::with_result(
                    instruction.to_string(),
                    cached_result.to_string(),
                ));
        }
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    assert_eq!(document.hyperlink_count().unwrap(), 0);
    let fields = document.hyperlink_fields().unwrap();
    assert_eq!(document.hyperlink_field_count().unwrap(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(
        fields[0].external_target(),
        Some("https://example.test/a b")
    );
    assert_eq!(fields[0].bookmark(), Some("_Toc1"));
    assert_eq!(fields[0].screen_tip(), Some("Stored tip"));
    assert_eq!(fields[0].target_frame(), Some("_blank"));
    assert!(fields[0].appends_image_map_coordinates());
    assert!(fields[0].opens_new_window());
    assert_eq!(fields[0].cached_result(), Some("cached external link"));
    assert_eq!(fields[1].external_target(), None);
    assert_eq!(fields[1].bookmark(), Some("JumpTarget"));
    assert_eq!(fields[1].cached_result(), Some("cached internal link"));
}

#[test]
fn writes_and_discovers_inert_prompt_fields_without_displaying_prompts() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let ask = document.add_paragraph();
        ask.add_field(crate::writer::MutableField::with_result(
            r#"ASK AskResponse "What is your first name?" \d "" \o"#.to_string(),
            "cached ask response".to_string(),
        ));
        let fill_in = document.add_paragraph();
        fill_in.add_field(crate::writer::MutableField::with_result(
            r#"FILLIN "Enter appointment time" \d "09:00""#.to_string(),
            "10:30".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.prompt_fields().unwrap();
    assert_eq!(document.prompt_field_count().unwrap(), 2);
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].kind(), crate::PromptKind::Ask);
    assert_eq!(fields[0].bookmark(), Some("AskResponse"));
    assert_eq!(fields[0].default_response(), Some(""));
    assert!(fields[0].prompts_once_per_mail_merge());
    assert_eq!(fields[0].cached_result(), Some("cached ask response"));
    assert_eq!(fields[1].kind(), crate::PromptKind::FillIn);
    assert_eq!(fields[1].bookmark(), None);
    assert_eq!(fields[1].prompt(), Some("Enter appointment time"));
    assert_eq!(fields[1].default_response(), Some("09:00"));
    assert_eq!(fields[1].cached_result(), Some("10:30"));
}

#[test]
fn writes_and_discovers_inert_macro_button_fields_without_execution() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_field(crate::writer::MutableField::with_result(
            r#"MACROBUTTON NeverRun "Click here""#.to_string(),
            "cached button".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let fields = document.macro_button_fields().unwrap();
    assert_eq!(document.macro_button_field_count().unwrap(), 1);
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].macro_name(), "NeverRun");
    assert_eq!(fields[0].display_text(), "Click here");
    assert_eq!(fields[0].cached_result(), Some("cached button"));
}

#[test]
fn writes_and_discovers_inert_active_and_building_block_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for (instruction, cached_result) in [
            ("ADDIN opaque-add-in-data", "cached add-in"),
            ("CONTROL opaque-control-data", "cached control"),
            ("HTMLCONTROL opaque-html-data", "cached HTML control"),
            (
                r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT"#,
                "cached glossary",
            ),
            (
                r#"AUTOTEXT "Reusable Clause" \q opaque"#,
                "cached auto text",
            ),
            (
                r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Select one""#,
                "cached auto text list",
            ),
        ] {
            document
                .add_paragraph()
                .add_field(crate::writer::MutableField::with_result(
                    instruction.to_string(),
                    cached_result.to_string(),
                ));
        }
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();

    let active_content = document.active_content_fields().unwrap();
    assert_eq!(document.active_content_field_count().unwrap(), 3);
    assert_eq!(active_content.len(), 3);
    assert_eq!(active_content[0].kind(), crate::ActiveContentKind::AddIn);
    assert_eq!(
        active_content[1].kind(),
        crate::ActiveContentKind::OcxControl
    );
    assert_eq!(
        active_content[2].kind(),
        crate::ActiveContentKind::HtmlControl
    );
    assert_eq!(
        active_content[2].cached_result(),
        Some("cached HTML control")
    );

    let auto_text = document.auto_text_fields().unwrap();
    assert_eq!(document.auto_text_field_count().unwrap(), 2);
    assert_eq!(auto_text.len(), 2);
    assert_eq!(auto_text[0].kind(), crate::AutoTextKind::Glossary);
    assert_eq!(auto_text[0].entry_name(), "Legacy Clause");
    assert_eq!(auto_text[1].kind(), crate::AutoTextKind::AutoText);
    assert_eq!(auto_text[1].entry_name(), "Reusable Clause");

    let auto_text_lists = document.auto_text_list_fields().unwrap();
    assert_eq!(document.auto_text_list_field_count().unwrap(), 1);
    assert_eq!(auto_text_lists.len(), 1);
    assert_eq!(auto_text_lists[0].display_text(), Some("Choose a name"));
    assert_eq!(
        auto_text_lists[0].cached_result(),
        Some("cached auto text list")
    );
}

#[test]
fn writes_and_discovers_typed_inert_dde_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let manual = document.add_paragraph();
        manual.add_field(crate::writer::MutableField::with_result(
            r#"DDE Excel "missing.xlsx" "Sheet1!A1" \a \p"#.to_string(),
            "cached DDE link".to_string(),
        ));
        let automatic = document.add_paragraph();
        automatic.add_field(crate::writer::MutableField::with_result(
            r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \t"#.to_string(),
            "cached DDE auto link".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let links = document.dde_links().unwrap();
    assert_eq!(document.dde_link_count().unwrap(), 2);
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].kind(), crate::DdeKind::Dde);
    assert_eq!(links[0].application(), "Excel");
    assert_eq!(links[0].source(), "missing.xlsx");
    assert_eq!(links[0].item(), Some("Sheet1!A1"));
    assert!(links[0].requests_automatic_updates());
    assert_eq!(links[0].representation(), Some(crate::DdeFormat::Picture));
    assert_eq!(links[0].cached_result(), Some("cached DDE link"));
    assert_eq!(links[1].kind(), crate::DdeKind::DdeAuto);
    assert_eq!(links[1].item(), Some("Sheet1!A2"));
    assert!(links[1].requests_automatic_updates());
    assert_eq!(links[1].representation(), Some(crate::DdeFormat::Text));
    assert_eq!(links[1].cached_result(), Some("cached DDE auto link"));
}

#[test]
fn writes_and_discovers_typed_inert_external_include_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let text = document.add_paragraph();
        text.add_field(crate::writer::MutableField::with_result(
            r#"INCLUDETEXT "file:///no-contact/source.docx" Summary \! \c Word8 \x /resume/name"#
                .to_string(),
            "cached included text".to_string(),
        ));
        let picture = document.add_paragraph();
        picture.add_field(crate::writer::MutableField::with_result(
            r#"INCLUDEPICTURE "file:///no-contact/picture.gif" \c Pictim32 \d"#.to_string(),
            "cached picture".to_string(),
        ));
        let legacy_text = document.add_paragraph();
        legacy_text.add_field(crate::writer::MutableField::with_result(
            r#"INCLUDE "file:///no-contact/legacy.docx" LegacySection \!"#.to_string(),
            "cached legacy text".to_string(),
        ));
        let legacy_picture = document.add_paragraph();
        legacy_picture.add_field(crate::writer::MutableField::with_result(
            r#"IMPORT "file:///no-contact/legacy.wmf" \c GraphicsFilter \d"#.to_string(),
            "cached legacy picture".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let includes = document.external_includes().unwrap();
    assert_eq!(document.external_include_count().unwrap(), 4);
    assert_eq!(includes.len(), 4);
    assert_eq!(includes[0].kind(), crate::IncludeKind::Text);
    assert_eq!(includes[0].source(), "file:///no-contact/source.docx");
    assert_eq!(includes[0].bookmark(), Some("Summary"));
    assert!(includes[0].suppresses_nested_field_updates());
    assert_eq!(
        includes[0].options(),
        &[
            crate::IncludeOption::Converter("Word8".to_string()),
            crate::IncludeOption::XPath("/resume/name".to_string()),
        ]
    );
    assert_eq!(includes[0].cached_result(), Some("cached included text"));
    assert_eq!(includes[1].kind(), crate::IncludeKind::Picture);
    assert_eq!(includes[1].source(), "file:///no-contact/picture.gif");
    assert!(includes[1].omits_picture_data());
    assert_eq!(
        includes[1].options(),
        &[crate::IncludeOption::Converter("Pictim32".to_string())]
    );
    assert_eq!(includes[1].cached_result(), Some("cached picture"));
    assert_eq!(includes[2].kind(), crate::IncludeKind::Text);
    assert_eq!(includes[2].source(), "file:///no-contact/legacy.docx");
    assert_eq!(includes[2].bookmark(), Some("LegacySection"));
    assert!(includes[2].suppresses_nested_field_updates());
    assert_eq!(includes[2].cached_result(), Some("cached legacy text"));
    assert_eq!(includes[3].kind(), crate::IncludeKind::Picture);
    assert_eq!(includes[3].source(), "file:///no-contact/legacy.wmf");
    assert!(includes[3].omits_picture_data());
    assert_eq!(
        includes[3].options(),
        &[crate::IncludeOption::Converter(
            "GraphicsFilter".to_string()
        )]
    );
    assert_eq!(includes[3].cached_result(), Some("cached legacy picture"));
}

#[test]
fn writes_and_discovers_typed_inert_referenced_document_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let relative = document.add_paragraph();
        relative.add_field(crate::writer::MutableField::with_result(
            r#"RD "C:\\Manual\\Chapters\\Chapter 1.docx" \f"#.to_string(),
            "cached relative reference".to_string(),
        ));
        let absolute = document.add_paragraph();
        absolute.add_field(crate::writer::MutableField::with_result(
            r#"RD "file:///no-contact/appendix.docx""#.to_string(),
            "cached absolute reference".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let references = document.referenced_documents().unwrap();
    assert_eq!(document.referenced_document_count().unwrap(), 2);
    assert_eq!(references.len(), 2);
    assert_eq!(references[0].source(), r"C:\Manual\Chapters\Chapter 1.docx");
    assert!(references[0].uses_relative_path());
    assert_eq!(
        references[0].cached_result(),
        Some("cached relative reference")
    );
    assert_eq!(references[1].source(), "file:///no-contact/appendix.docx");
    assert!(!references[1].uses_relative_path());
    assert_eq!(
        references[1].cached_result(),
        Some("cached absolute reference")
    );
}

#[test]
fn writes_and_discovers_typed_inert_link_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let spreadsheet = document.add_paragraph();
        spreadsheet.add_field(crate::writer::MutableField::with_result(
            r#"LINK Excel.Sheet.8 "missing.xlsx" "Sheet1!A1" \a \f 4 \p"#.to_string(),
            "cached spreadsheet link".to_string(),
        ));
        let text = document.add_paragraph();
        text.add_field(crate::writer::MutableField::with_result(
            r#"LINK Word.Document.8 "missing.docx" Bookmark \t"#.to_string(),
            "cached text link".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let links = document.link_fields().unwrap();
    assert_eq!(document.link_field_count().unwrap(), 2);
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].application_type(), "Excel.Sheet.8");
    assert_eq!(links[0].source(), "missing.xlsx");
    assert_eq!(links[0].item(), Some("Sheet1!A1"));
    assert!(links[0].requests_automatic_updates());
    assert_eq!(
        links[0].formatting_modes(),
        &[crate::LinkFormat::SpreadsheetSource]
    );
    assert_eq!(
        links[0].effective_result_option(),
        Some(crate::LinkResult::Picture)
    );
    assert_eq!(links[0].cached_result(), Some("cached spreadsheet link"));
    assert_eq!(links[1].application_type(), "Word.Document.8");
    assert_eq!(links[1].item(), Some("Bookmark"));
    assert_eq!(
        links[1].effective_result_option(),
        Some(crate::LinkResult::Text)
    );
    assert_eq!(links[1].cached_result(), Some("cached text link"));
}

#[test]
fn saves_and_discovers_typed_inert_bibliography_source_stores() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package
            .add_custom_xml(NewStore {
                xml: br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="/APA.XSL" StyleName="APA"><b:Source><b:Tag>Doe2024</b:Tag><b:SourceType>Book</b:SourceType><b:Title>Stored source</b:Title></b:Source></b:Sources>"#.to_vec(),
                content_type: "application/xml".to_string(),
                id: "{22222222-2222-2222-2222-222222222222}".to_string(),
                schemas: vec![
                    crate::OOXML_BIBLIOGRAPHY_NAMESPACE.to_string(),
                ],
                conformance: custom_xml::Conformance::Transitional,
            })
            .unwrap();
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let stores = reopened.bibliography_source_stores().unwrap();
    assert_eq!(stores.len(), 1);
    assert_eq!(
        stores[0].data_store_item_id(),
        Some("{22222222-2222-2222-2222-222222222222}")
    );
    assert_eq!(stores[0].selected_style(), Some("/APA.XSL"));
    assert_eq!(stores[0].style_name(), Some("APA"));
    assert_eq!(stores[0].source_count(), 1);
    assert_eq!(stores[0].sources()[0].tag(), Some("Doe2024"));
    assert_eq!(stores[0].sources()[0].source_type(), Some("Book"));
    assert_eq!(stores[0].sources()[0].title(), Some("Stored source"));

    let sources = reopened.bibliography_sources().unwrap();
    assert_eq!(reopened.bibliography_source_count().unwrap(), 1);
    assert_eq!(sources.len(), 1);
    assert_eq!(sources[0].tag(), Some("Doe2024"));
}

#[test]
fn writes_and_discovers_typed_index_fields() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let index = document.add_paragraph();
        index.add_field(crate::writer::MutableField::with_result(
            r#"INDEX \c 2 \f "topics" \r"#.to_string(),
            "Topic\t3".to_string(),
        ));
        let entry = document.add_paragraph();
        entry.add_field(crate::writer::MutableField::with_result(
            r#"XE "Topic" \f "topics" \r TopicRange \b"#.to_string(),
            "hidden marker".to_string(),
        ));
    }
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let document = reopened.document().unwrap();
    let indexes = document.indexes().unwrap();
    assert_eq!(document.index_count().unwrap(), 1);
    assert_eq!(indexes.len(), 1);
    assert_eq!(indexes[0].columns().unwrap(), Some(2));
    assert_eq!(indexes[0].entry_identifier().unwrap(), Some("topics"));
    assert!(indexes[0].runs_subentries_inline());

    let entries = document.index_entries().unwrap();
    assert_eq!(document.index_entry_count().unwrap(), 1);
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].entry(), "Topic");
    assert_eq!(entries[0].entry_identifier().unwrap(), Some("topics"));
    assert_eq!(
        entries[0].page_range_bookmark().unwrap(),
        Some("TopicRange")
    );
    assert!(entries[0].is_bold());
}
