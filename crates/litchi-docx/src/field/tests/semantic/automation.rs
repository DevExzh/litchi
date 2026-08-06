//! Automation, embedded-content, form, and inert action field semantics.

#[allow(
    clippy::wildcard_imports,
    reason = "tests exercise the complete public field vocabulary"
)]
use super::super::*;

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
