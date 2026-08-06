//! Prompt, AutoText, and user-identity field semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
