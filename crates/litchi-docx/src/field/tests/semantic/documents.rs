//! Document variables, properties, information, and context semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
