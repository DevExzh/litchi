//! XML value and extension-owner semantics for inert MS-OWEXML payloads.

use litchi_docx::web_extensions::{CustomFunctions, ExtList};

fn fragment(payload: &str) -> String {
    format!(
        concat!(
            "<we:extLst xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" ",
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
            "<a:ext uri=\"urn:test:custom-functions\">{}</a:ext></we:extLst>"
        ),
        payload
    )
}

fn read(payload: &str) -> CustomFunctions {
    ExtList::from_xml(fragment(payload).as_bytes())
        .unwrap()
        .custom_functions()
        .unwrap()
        .clone()
}

#[test]
fn function_ids_normalize_literal_line_endings_but_retain_character_references() {
    let value = read(concat!(
        "<we:customFunctionList>",
        "<we:customFunctionIds>left\r\nmiddle\rright&#13;end</we:customFunctionIds>",
        "<we:customFunctionIds><![CDATA[left\r\nmiddle\rright]]></we:customFunctionIds>",
        "<we:customFunctionIds> &amp;&lt;&gt;&#x1F34E; </we:customFunctionIds>",
        "<we:customFunctionIds/>",
        "</we:customFunctionList>"
    ));
    assert_eq!(
        value.custom_function_list().unwrap().ids(),
        [
            "left\nmiddle\nright\rend",
            "left\nmiddle\nright",
            " &<>🍎 ",
            ""
        ]
    );
}

#[test]
fn runtime_attributes_normalize_literal_whitespace_but_retain_character_references() {
    let value =
        read("<we:backgroundAppData state=\"-2147483648\" runtimeId=\"a\t\r\nb&#9;&#10;&#13;c\"/>");
    let runtime = value.background_app_data().unwrap();
    assert_eq!(runtime.state(), i32::MIN);
    assert_eq!(runtime.runtime_id(), "a  b\t\n\rc");
    let maximum = read("<we:backgroundAppData state=\"+2147483647\" runtimeId=\"\"/>");
    assert_eq!(maximum.background_app_data().unwrap().state(), i32::MAX);
    assert_eq!(maximum.background_app_data().unwrap().runtime_id(), "");
}

#[test]
fn boolean_and_integer_lexical_space_uses_xml_whitespace_only() {
    let value = read(concat!(
        "<we:containsCustomFunctions val=\" &#9;1&#13; \"/>",
        "<we:backgroundAppData state=\" &#10;+0007&#9; \" runtimeId=\"runtime\"/>"
    ));
    assert_eq!(
        value.contains_custom_functions().unwrap().explicit_value(),
        Some(true)
    );
    assert_eq!(value.background_app_data().unwrap().state(), 7);

    for payload in [
        "<we:containsCustomFunctions val=\"&#160;true&#160;\"/>",
        "<we:backgroundAppData state=\"&#160;7&#160;\" runtimeId=\"r\"/>",
        "<we:backgroundAppData state=\"2147483648\" runtimeId=\"r\"/>",
        "<we:backgroundAppData state=\"-2147483649\" runtimeId=\"r\"/>",
    ] {
        assert!(
            ExtList::from_xml(fragment(payload).as_bytes()).is_err(),
            "invalid schema value was accepted: {payload}"
        );
    }
}

#[test]
fn empty_list_and_empty_id_remain_distinct_during_an_unrelated_typed_edit() {
    for (payload, expected_ids) in [
        ("<we:customFunctionList/>", Vec::<String>::new()),
        (
            "<we:customFunctionList><we:customFunctionIds/></we:customFunctionList>",
            vec![String::new()],
        ),
    ] {
        let xml = fragment(payload);
        let mut extension = ExtList::from_xml(xml.as_bytes()).unwrap();
        let mut changed = extension.custom_functions().unwrap().clone();
        changed.set_background_app_data(Some(
            litchi_docx::web_extensions::BackgroundAppData::new(0, "").unwrap(),
        ));
        extension.set_custom_functions(Some(changed)).unwrap();
        let reopened = ExtList::from_xml(extension.as_xml()).unwrap();
        let metadata = reopened.custom_functions().unwrap();
        assert_eq!(metadata.custom_function_list().unwrap().ids(), expected_ids);
        assert!(metadata.contains_custom_functions().is_none());
        assert_eq!(metadata.background_app_data().unwrap().runtime_id(), "");
    }
}

#[test]
fn authored_xml_strings_retain_whitespace_and_delimiters_after_reopening() {
    use litchi_docx::web_extensions::{BackgroundAppData, CustomFunctionList, ExtKind};

    let runtime = "runtime\t\n\r<&>\"'🍎";
    let id = "function\t\n\r<&>\"'🍎";
    let mut ids = CustomFunctionList::new();
    ids.push_id(id).unwrap();
    let mut metadata = CustomFunctions::new();
    metadata
        .set_background_app_data(Some(BackgroundAppData::new(-1, runtime).unwrap()))
        .set_custom_function_list(Some(ids));
    let mut extension = ExtList::empty(ExtKind::AddIn).unwrap();
    extension
        .set_custom_functions(Some(metadata.clone()))
        .unwrap();
    let reopened = ExtList::from_xml(extension.as_xml()).unwrap();
    assert_eq!(reopened.custom_functions(), Some(&metadata));
}

#[test]
fn list_and_id_elements_refuse_undeclared_schema_attributes() {
    for payload in [
        "<we:customFunctionList bogus=\"value\"/>",
        "<we:customFunctionList bogus=\"value\"><we:customFunctionIds>id</we:customFunctionIds></we:customFunctionList>",
        "<we:customFunctionList><we:customFunctionIds bogus=\"value\"/></we:customFunctionList>",
        "<we:customFunctionList><we:customFunctionIds bogus=\"value\">id</we:customFunctionIds></we:customFunctionList>",
    ] {
        assert!(
            ExtList::from_xml(fragment(payload).as_bytes()).is_err(),
            "undeclared attribute was accepted: {payload}"
        );
    }
}

#[test]
fn newly_added_payloads_stay_in_the_existing_custom_function_extension_owner() {
    use litchi_docx::web_extensions::{BackgroundAppData, CustomFunctionList};

    let source = fragment("<we:containsCustomFunctions val=\"true\"/>").replace(
        "<a:ext uri=\"urn:test:custom-functions\">",
        concat!(
            "<a:ext uri=\"urn:test:vendor\"><v:opaque xmlns:v=\"urn:vendor\"/></a:ext>",
            "<a:ext uri=\"urn:test:custom-functions\">"
        ),
    );
    let mut extension = ExtList::from_xml(source.as_bytes()).unwrap();
    let mut changed = extension.custom_functions().unwrap().clone();
    changed
        .set_background_app_data(Some(BackgroundAppData::new(1, "runtime").unwrap()))
        .set_custom_function_list(Some(CustomFunctionList::new()));
    extension.set_custom_functions(Some(changed)).unwrap();
    let xml = extension.xml();
    let vendor_start = xml.find("uri=\"urn:test:vendor\"").unwrap();
    let vendor_end = vendor_start + xml[vendor_start..].find("</a:ext>").unwrap();
    let custom_start = xml.find("uri=\"urn:test:custom-functions\"").unwrap();
    let custom_end = custom_start + xml[custom_start..].find("</a:ext>").unwrap();
    for name in ["backgroundAppData", "customFunctionList"] {
        assert!(
            !xml[vendor_start..vendor_end].contains(name),
            "{name} moved into a vendor extension"
        );
        assert!(
            xml[custom_start..custom_end].contains(name),
            "{name} is absent from the existing custom extension"
        );
    }
}
