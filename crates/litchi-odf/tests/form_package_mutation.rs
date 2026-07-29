use litchi_odf::{
    Document, OdfAuthoredForm, OdfAuthoredFormControl, OdfButtonControl, OdfComboboxControl,
    OdfFileControl, OdfFixedTextControl, OdfFormControlKind, OdfFormProperty, OdfFrameControl,
    OdfGenericFormControl, OdfGridControl, OdfImageFrameControl, OdfInteractiveControl,
    OdfPasswordFileControl, OdfRadioControl, OdfSelectionControl, OdfTextControl,
    OdfTypedValueControl, OdfTypedValueControlKind, OdfValueRangeControl, OdfVisualControl,
    OwnedPackage, PackageWriter, Presentation, Spreadsheet, constants,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";

fn package(mimetype: &str, family: &str, inner: &str, signature: bool) -> Vec<u8> {
    let body = match family {
        "text" => format!("<office:text><!--unknown-preserved-->{inner}<text:p>body</text:p></office:text>"),
        "spreadsheet" => format!("<office:spreadsheet><!--unknown-preserved-->{inner}<table:table table:name=\"Sheet1\"/></office:spreadsheet>"),
        "presentation" => format!("<office:presentation><!--unknown-preserved-->{inner}<draw:page draw:name=\"Slide1\"/></office:presentation>"),
        _ => unreachable!(),
    };
    let xml = format!(
        "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"{OFFICE}\" xmlns:form=\"{FORM}\" xmlns:text=\"{TEXT}\" xmlns:table=\"{TABLE}\" xmlns:draw=\"{DRAW}\" xmlns:script=\"{SCRIPT}\" xmlns:xlink=\"{XLINK}\"><office:body>{body}</office:body></office:document-content>"
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer.add_file(constants::ODF_CONTENT, xml.as_bytes()).unwrap();
    if signature {
        writer.add_file_with_media_type("META-INF/documentsignatures.xml", b"<signatures/>", "text/xml").unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn exhaustive_form() -> OdfAuthoredForm {
    let mut form = OdfAuthoredForm::new("Main");
    form.xml_id = Some("main_form".to_string());
    form.datasource = Some("registered-database-name".to_string());
    form.command_type = Some("table".to_string());
    form.command = Some("Customers".to_string());
    form.href = Some("https://example.invalid/inert-datasource".to_string());
    form.properties.push(OdfFormProperty::number("Currency", "currency", "12.50", Some("USD".to_string())));

    let mut text = OdfTextControl::text("Text", "text_id");
    text.linked_cell = Some("Sheet1.A1".to_string());
    text.data_field = Some("customer_name".to_string());
    form.add_control(text).unwrap();

    let number = OdfTypedValueControl::new(OdfTypedValueControlKind::Number, "Number", "number_id");
    form.add_control(OdfAuthoredFormControl::TypedValue(number)).unwrap();

    let mut combo = OdfComboboxControl::new("Combo", "combo_id");
    combo.list_source = Some("SELECT name FROM customers".to_string());
    combo.source_cell_range = Some("Sheet1.A1:A20".to_string());
    form.add_control(OdfAuthoredFormControl::Selection(OdfSelectionControl::Combobox(combo))).unwrap();

    form.add_control(OdfAuthoredFormControl::Interactive(OdfInteractiveControl::Button(
        OdfButtonControl::new("Button", "button_id"),
    ))).unwrap();
    form.add_control(OdfAuthoredFormControl::Visual(OdfVisualControl::Radio(
        OdfRadioControl::new("Radio", "radio_id"),
    ))).unwrap();
    form.add_control(OdfAuthoredFormControl::PasswordFile(OdfPasswordFileControl::File(
        OdfFileControl::new("File", "file_id"),
    ))).unwrap();
    form.add_control(OdfAuthoredFormControl::Generic(OdfGenericFormControl::FixedText(
        OdfFixedTextControl::new("Label", "label_id"),
    ))).unwrap();
    form.add_control(OdfAuthoredFormControl::Grid(OdfGridControl::new("Grid", "grid_id"))).unwrap();
    form.add_control(OdfAuthoredFormControl::ImageFrame(OdfImageFrameControl::new(
        "ImageFrame", "image_frame_id",
    ))).unwrap();
    form.add_control(OdfAuthoredFormControl::ValueRange(OdfValueRangeControl::new(
        "Range", "range_id",
    ))).unwrap();

    let mut nested = OdfAuthoredForm::new("Nested");
    nested.form_id = Some("nested_form_id".to_string());
    nested.add_control(OdfAuthoredFormControl::Visual(OdfVisualControl::Frame(
        OdfFrameControl::new("Frame", "frame_id"),
    ))).unwrap();
    form.add_form(nested).unwrap();
    form
}

#[test]
fn odt_all_typed_controls_nested_forms_and_atomic_uniqueness() {
    let mut document = Document::from_bytes(package(constants::ODF_TEXT, "text", "", true)).unwrap();
    assert_eq!(document.add_form(0, &exhaustive_form()).unwrap(), 0);
    let parsed = document.forms().unwrap();
    let main = &parsed.groups[0].forms[0];
    assert_eq!(main.name.as_deref(), Some("Main"));
    assert_eq!(main.properties[0].name, "Currency");
    assert!(main.children.iter().any(|node| matches!(node, litchi_odf::OdfFormNode::Form(_))));
    let kinds: Vec<_> = main.children.iter().filter_map(|node| match node {
        litchi_odf::OdfFormNode::Control(control) => Some(control.kind.clone()),
        _ => None,
    }).collect();
    assert!(kinds.contains(&OdfFormControlKind::Text));
    assert!(kinds.contains(&OdfFormControlKind::Number));
    assert!(kinds.contains(&OdfFormControlKind::ComboBox));
    assert!(kinds.contains(&OdfFormControlKind::Button));
    assert!(kinds.contains(&OdfFormControlKind::Radio));
    assert!(kinds.contains(&OdfFormControlKind::File));
    assert!(kinds.contains(&OdfFormControlKind::FixedText));
    assert!(kinds.contains(&OdfFormControlKind::Grid));
    assert!(kinds.contains(&OdfFormControlKind::ImageFrame));
    assert!(kinds.contains(&OdfFormControlKind::ValueRange));

    document.move_form_control(9, 0).unwrap();
    let replacement = OdfAuthoredFormControl::Text(OdfTextControl::textarea("Changed", "changed_id"));
    document.replace_form_control(1, &replacement).unwrap();
    document.remove_form_control(2).unwrap();

    let before = document.to_bytes().unwrap();
    let duplicate = OdfAuthoredFormControl::Text(OdfTextControl::text("Duplicate", "changed_id"));
    assert!(document.add_form_control(0, &duplicate).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    assert!(document.move_form(0, 1).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);

    let archive = OwnedPackage::from_bytes(document.to_bytes().unwrap()).unwrap();
    let content = String::from_utf8(archive.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    assert!(content.contains("unknown-preserved"));
    assert!(!archive.has_file("META-INF/documentsignatures.xml").unwrap());
}

#[test]
fn control_reference_and_event_listener_are_inert_and_referentially_checked() {
    let inner = "<office:forms><form:form form:name=\"Main\"><office:event-listeners><script:event-listener script:event-name=\"dom:click\" script:language=\"ooo:script\" script:macro-name=\"Standard.Module1.Main\" xlink:type=\"simple\" xlink:href=\"vnd.sun.star.script:Standard.Module1.Main\"/></office:event-listeners><form:text form:name=\"Text\" xml:id=\"control_id\"/></form:form></office:forms><draw:control draw:control=\"control_id\"/>".to_string();
    let mut document = Document::from_bytes(package(constants::ODF_TEXT, "text", &inner, false)).unwrap();
    let forms = document.forms().unwrap();
    assert_eq!(forms.event_listeners.len(), 1);
    assert_eq!(forms.control_shapes.len(), 1);
    let before = document.to_bytes().unwrap();
    assert!(document.remove_form_control(0).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);

    document.add_form_control(
        0,
        &OdfAuthoredFormControl::Interactive(OdfInteractiveControl::Button(
            OdfButtonControl::new("SafeButton", "safe_button"),
        )),
    ).unwrap();
    assert_eq!(document.forms().unwrap().event_listeners.len(), 1);
}

#[test]
fn ods_and_odp_form_tree_crud_roundtrip() {
    let mut sheet = Spreadsheet::from_bytes(package(constants::ODF_SPREADSHEET, "spreadsheet", "", false)).unwrap();
    let root = OdfAuthoredForm::new("SheetForm");
    assert_eq!(sheet.add_form(0, &root).unwrap(), 0);
    assert_eq!(sheet.add_nested_form(0, &OdfAuthoredForm::new("Child")).unwrap(), 1);
    sheet.add_form_control(1, &OdfAuthoredFormControl::Text(OdfTextControl::text("Cell", "cell_id"))).unwrap();
    assert_eq!(sheet.forms().unwrap().groups[0].forms[0].children.len(), 1);
    sheet.remove_form(1).unwrap();
    assert!(sheet.forms().unwrap().groups[0].forms[0].children.is_empty());

    let mut slides = Presentation::from_bytes(package(constants::ODF_PRESENTATION, "presentation", "", false)).unwrap();
    slides.add_form(0, &OdfAuthoredForm::new("First")).unwrap();
    slides.add_form(0, &OdfAuthoredForm::new("Second")).unwrap();
    slides.move_form(1, 0).unwrap();
    let forms = slides.forms().unwrap();
    assert_eq!(forms.groups[0].forms[0].name.as_deref(), Some("Second"));
    slides.replace_form(0, &OdfAuthoredForm::new("Replacement")).unwrap();
    slides.remove_form(1).unwrap();
    assert_eq!(slides.forms().unwrap().groups[0].forms[0].name.as_deref(), Some("Replacement"));
}

#[test]
fn malformed_authored_forms_fail_without_package_changes() {
    let mut document = Document::from_bytes(package(constants::ODF_TEXT, "text", "", false)).unwrap();
    let before = document.to_bytes().unwrap();
    let mut malformed = OdfAuthoredForm::new("Bad");
    malformed.command_type = Some("execute".to_string());
    assert!(document.add_form(0, &malformed).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);

    let mut oversized = OdfAuthoredForm::new("Large");
    oversized.datasource = Some("x".repeat(1024 * 1024 + 1));
    assert!(document.add_form(0, &oversized).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}
