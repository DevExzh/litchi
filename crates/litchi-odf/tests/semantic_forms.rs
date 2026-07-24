use litchi_odf::{
    FlatOpenDocument, OdfFormControlKind, OdfFormNode, OdfFormPropertyValue, OdfFormScalarValue,
    OpenDocumentPackage,
};
use std::path::{Path, PathBuf};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XFORMS: &str = "http://www.w3.org/2002/xforms";

fn flat(body: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><o:document xmlns:o="{OFFICE}" xmlns:f="{FORM}" xmlns:d="{DRAW}" xmlns:t="{TEXT}" xmlns:s="{SCRIPT}" xmlns:x="{XLINK}" xmlns:xf="{XFORMS}" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:body><o:text>{body}</o:text></o:body></o:document>"#
    )
    .into_bytes()
}

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn collect<'a>(nodes: &'a [OdfFormNode], out: &mut Vec<&'a litchi_odf::OdfFormControl>) {
    for node in nodes {
        match node {
            OdfFormNode::Form(form) => collect(&form.children, out),
            OdfFormNode::Control(control) => {
                out.push(control);
                collect(&control.children, out);
            },
        }
    }
}

#[test]
fn parses_all_controls_nesting_typed_properties_and_links() {
    let names = [
        "text",
        "textarea",
        "password",
        "file",
        "formatted-text",
        "number",
        "date",
        "time",
        "fixed-text",
        "combobox",
        "item",
        "listbox",
        "option",
        "button",
        "image",
        "checkbox",
        "radio",
        "frame",
        "image-frame",
        "hidden",
        "grid",
        "column",
        "value-range",
        "generic-control",
    ];
    let controls = names
        .iter()
        .enumerate()
        .map(|(i, name)| format!(r#"<f:{name} xml:id="c{i}" f:id="c{i}"/>"#))
        .collect::<String>();
    let document = FlatOpenDocument::from_bytes(flat(&format!(
        r#"<o:forms f:automatic-focus="false" f:apply-design-mode="true"><f:form f:name="outer"><f:properties><f:property f:property-name="enabled" o:value-type="boolean" o:boolean-value="true"/><f:list-property f:property-name="choices" o:value-type="string"><f:list-value o:string-value="a"/><f:list-value o:string-value="b"/></f:list-property></f:properties><f:form f:name="nested">{controls}</f:form></f:form></o:forms><d:control d:control="c0" d:z-index="7"/>"#
    )))
    .unwrap();
    let forms = document.forms().unwrap();
    assert_eq!(forms.groups.len(), 1);
    assert_eq!(forms.groups[0].automatic_focus, Some(false));
    let mut parsed = Vec::new();
    collect(&forms.groups[0].forms[0].children, &mut parsed);
    assert_eq!(parsed.len(), names.len());
    assert!(matches!(parsed[0].kind, OdfFormControlKind::Text));
    assert!(matches!(
        parsed[23].kind,
        OdfFormControlKind::GenericControl
    ));
    assert!(matches!(
        forms.groups[0].forms[0].properties[0].value,
        OdfFormPropertyValue::Scalar(OdfFormScalarValue::Boolean(true))
    ));
    assert!(matches!(
        forms.groups[0].forms[0].properties[1].value,
        OdfFormPropertyValue::List { ref values, .. } if values.len() == 2
    ));
    assert!(forms.control_shapes[0].resolved_control.is_some());
}

#[test]
fn flags_behavior_but_preserves_external_values_inertly() {
    let xml = flat(
        r#"<o:forms><xf:model><xf:instance src="file:///never"/></xf:model><f:form f:datasource="https://never.test/db" f:command="DROP TABLE x"><f:image f:id="c" f:image-data="https://never.test/image"><o:event-listeners><s:event-listener s:macro-name="macro://never"/></o:event-listeners></f:image></f:form></o:forms><d:control d:control="c"/>"#,
    );
    let forms = FlatOpenDocument::from_bytes(xml).unwrap().forms().unwrap();
    assert!(forms.has_xforms && forms.has_event_listeners);
    assert!(
        forms.groups[0].forms[0]
            .attributes
            .iter()
            .any(|attribute| attribute.value == "DROP TABLE x")
    );
    let OdfFormNode::Control(control) = &forms.groups[0].forms[0].children[0] else {
        panic!()
    };
    assert_eq!(
        control.image_data.as_deref(),
        Some("https://never.test/image")
    );
}

#[test]
fn rejects_malformed_spoofed_unresolved_and_limited_inputs() {
    for xml in [
        flat(r#"<o:forms><f:form/></o:forms><d:control d:control="missing"/>"#),
        flat(
            r#"<o:forms><f:form><f:text f:id="x"/><f:button f:id="x"/></f:form></o:forms><d:control d:control="x"/>"#,
        ),
        flat(r#"<o:forms f:automatic-focus="yes"><f:form/></o:forms>"#),
    ] {
        assert!(FlatOpenDocument::from_bytes(xml).unwrap().forms().is_err());
    }
    let spoofed = format!(r#"<o:document xmlns:o="{OFFICE}" xmlns:f="urn:not-form" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text><o:forms><f:form/></o:forms></o:text></o:body></o:document>"#).into_bytes();
    assert!(
        FlatOpenDocument::from_bytes(spoofed)
            .unwrap()
            .forms()
            .unwrap()
            .groups[0]
            .forms
            .is_empty()
    );
    let oversized = "x".repeat(65_537);
    let limited = flat(&format!(
        r#"<o:forms><f:form f:name="{oversized}"/></o:forms>"#
    ));
    assert!(
        FlatOpenDocument::from_bytes(limited)
            .unwrap()
            .forms()
            .is_err()
    );
    let mut nested = String::from("<o:forms>");
    for _ in 0..129 {
        nested.push_str("<f:form>");
    }
    for _ in 0..129 {
        nested.push_str("</f:form>");
    }
    nested.push_str("</o:forms>");
    assert!(
        FlatOpenDocument::from_bytes(flat(&nested))
            .unwrap()
            .forms()
            .is_err()
    );
}

#[test]
fn parses_all_bundled_fixtures_without_rewriting_packages() {
    for relative in [
        "test-data/odfdo/tests/samples/forms.odt",
        "test-data/libreoffice-core/sw/qa/extras/odfexport/data/Formcontrol needs high z-index.odt",
        "test-data/libreoffice-core/xmloff/qa/unit/data/tdf156707_text_form_control_borders.odt",
        "test-data/libreoffice-core/xmloff/qa/unit/data/tdf167358_label_form_control_borders.odt",
    ] {
        let bytes = std::fs::read(fixture(relative)).unwrap();
        let package = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
        let forms = package.forms().unwrap();
        assert!(!forms.groups.is_empty(), "{relative}");
        assert!(!forms.control_shapes.is_empty(), "{relative}");
        assert_eq!(package.to_bytes(), bytes, "{relative}");
    }
    let bytes = std::fs::read(fixture(
        "test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/formcontrol.fodt",
    ))
    .unwrap();
    let forms = FlatOpenDocument::from_bytes(bytes)
        .unwrap()
        .forms()
        .unwrap();
    assert_eq!(forms.groups.len(), 1);
    assert_eq!(forms.control_shapes.len(), 1);
    assert!(forms.control_shapes[0].resolved_control.is_some());
}
