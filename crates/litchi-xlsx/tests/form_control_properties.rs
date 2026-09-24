//! Focused, standalone coverage for the SpreadsheetML 2009/9 form-control
//! properties leaf codec.
//!
//! These tests intentionally stop at the XML part boundary.  A worksheet
//! control, relationship graph, and package writer are not part of the leaf
//! owner's contract yet.
#![allow(clippy::unwrap_used, reason = "test assertions use panic-on-failure")]

use std::fs;
use std::path::{Path, PathBuf};

use litchi_core::BlobId;
use litchi_xlsx::form_control::{
    CONTROL_PROPERTIES_CONTENT_TYPE, CONTROL_PROPERTIES_RELATIONSHIP_TYPE, Checked, DropStyle,
    EditValidation, FORM_CONTROL_NAMESPACE, FormControlError, FormControlFormula, Item, ItemList,
    KnownOrUnknown, Limits, MAX_RETAINED_BYTES, ObjectType, OpaqueXml, Properties, ScalarField,
    ScalarValue, SelectionType, TextHAlign, TextVAlign, insert_item, inspect, inspect_with_limits,
    parse, parse_with_limits, remove_item, replace_items, replace_scalar, write, write_with_limits,
};
use serde_json::Value;
use soapberry_zip::office::ArchiveReader;

const CORE_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NATIVE_CORPUS: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json"
));

fn empty_root() -> Vec<u8> {
    format!(r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}"/>"#).into_bytes()
}

fn root_with_body(body: &str) -> Vec<u8> {
    format!(r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}">{body}</formControlPr>"#)
        .into_bytes()
}

fn root_with_attribute(name: &str, value: &str) -> Vec<u8> {
    format!(r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}" {name}="{value}"/>"#).into_bytes()
}

fn known_token<T: std::fmt::Display>(value: Option<&KnownOrUnknown<T>>) -> Option<String> {
    value.map(|value| match value {
        KnownOrUnknown::Known(value) => value.to_string(),
        KnownOrUnknown::Unknown(value) => value.clone(),
    })
}

fn assert_invalid(xml: &[u8], label: &str) {
    assert!(parse(xml).is_err(), "{label} unexpectedly parsed");
}

fn minimum_parse_retained_limit(source: &[u8]) -> usize {
    let mut low = 0usize;
    let mut high = MAX_RETAINED_BYTES;
    assert!(parse_with_limits(source, &Limits::new().with_max_retained_bytes(high)).is_ok());
    while low < high {
        let middle = low + (high - low) / 2;
        if parse_with_limits(source, &Limits::new().with_max_retained_bytes(middle)).is_ok() {
            high = middle;
        } else {
            low = middle + 1;
        }
    }
    low
}

fn minimum_write_retained_limit(properties: &Properties) -> usize {
    let mut low = 0usize;
    let mut high = MAX_RETAINED_BYTES;
    assert!(write_with_limits(properties, &Limits::new().with_max_retained_bytes(high)).is_ok());
    while low < high {
        let middle = low + (high - low) / 2;
        if write_with_limits(properties, &Limits::new().with_max_retained_bytes(middle)).is_ok() {
            high = middle;
        } else {
            low = middle + 1;
        }
    }
    low
}

fn assert_retained_limit<T>(result: litchi_xlsx::form_control::Result<T>, maximum: usize) {
    match result {
        Err(FormControlError::Limit {
            maximum: observed_maximum,
            ..
        }) => assert_eq!(observed_maximum, maximum),
        Err(error) => panic!("expected retained-byte limit at {maximum}, got {error:?}"),
        Ok(_) => panic!("expected retained-byte limit at {maximum}"),
    }
}

fn workspace_path(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn json_str<'a>(value: &'a Value, key: &str) -> &'a str {
    value[key]
        .as_str()
        .unwrap_or_else(|| panic!("missing string field {key}"))
}

fn native_package(path: &str) -> Vec<u8> {
    let path = workspace_path(path);
    fs::read(&path)
        .unwrap_or_else(|error| panic!("read native fixture {}: {error}", path.display()))
}

#[test]
fn native_corpus_parts_read_and_write_as_exact_no_ops() {
    let corpus: Value = serde_json::from_str(NATIVE_CORPUS).unwrap();
    let fixtures = corpus["fixtures"].as_array().unwrap();
    assert_eq!(fixtures.len(), 7);

    let mut part_count = 0usize;
    for fixture in fixtures {
        let package_path = json_str(fixture, "path");
        let retained_path = json_str(fixture, "retained_path");
        let package = native_package(retained_path);
        assert_eq!(
            BlobId::of(&package).as_hex(),
            json_str(fixture, "sha256"),
            "retained fixture hash {retained_path} (source {package_path})"
        );
        assert_eq!(package.len(), fixture["bytes"].as_u64().unwrap() as usize);
        let archive = ArchiveReader::new(&package).unwrap();
        let parts = fixture["form_control_parts"].as_array().unwrap();
        for part in parts {
            part_count += 1;
            let member = json_str(part, "member");
            let source = archive.read(member).unwrap_or_else(|error| {
                panic!("read native form-control member {package_path}!{member}: {error}")
            });
            assert_eq!(
                BlobId::of(&source).as_hex(),
                json_str(part, "sha256"),
                "native form-control member hash {package_path}!{member}"
            );
            assert_eq!(source.len(), part["bytes"].as_u64().unwrap() as usize);
            assert_eq!(
                json_str(part, "content_type"),
                CONTROL_PROPERTIES_CONTENT_TYPE
            );
            for relationship in part["incoming_relationships"].as_array().unwrap() {
                assert_eq!(
                    json_str(relationship, "type"),
                    CONTROL_PROPERTIES_RELATIONSHIP_TYPE
                );
            }

            let properties = parse(&source).unwrap();
            assert_eq!(
                write(&properties).unwrap(),
                source,
                "native no-op {package_path}!{member}"
            );
            assert_eq!(properties.source_bytes(), Some(source.as_slice()));
            assert_eq!(
                known_token(properties.object_type()),
                Some(json_str(&part["attributes"], "objectType").to_owned())
            );

            let attributes = part["attributes"].as_object().unwrap();
            for (name, expected) in attributes {
                let expected = expected.as_str().unwrap();
                match name.as_str() {
                    "objectType" => assert_eq!(
                        known_token(properties.object_type()),
                        Some(expected.to_owned())
                    ),
                    "checked" => {
                        assert_eq!(known_token(properties.checked()), Some(expected.to_owned()))
                    },
                    "fmlaLink" => assert_eq!(
                        properties.fmla_link().map(FormControlFormula::as_str),
                        Some(expected)
                    ),
                    "lockText" => assert_eq!(properties.lock_text(), Some(expected == "1")),
                    "noThreeD" => assert_eq!(properties.no_three_d(), Some(expected == "1")),
                    "firstButton" => assert_eq!(properties.first_button(), Some(expected == "1")),
                    other => panic!("uncovered native attribute {other}"),
                }
            }

            if attributes.get("fmlaLink").and_then(Value::as_str) == Some("#REF!") {
                let changed = replace_scalar(
                    &source,
                    ScalarField::LockText,
                    Some(ScalarValue::Boolean(false)),
                )
                .unwrap();
                let reopened = parse(&changed).unwrap();
                assert_eq!(
                    reopened.fmla_link().map(FormControlFormula::as_str),
                    Some("#REF!")
                );
                assert!(
                    changed
                        .windows(b"#REF!".len())
                        .any(|window| window == b"#REF!")
                );
            }
        }
    }
    assert_eq!(part_count, 10);
}

#[test]
fn empty_properties_keep_presence_states_and_effective_defaults() {
    let source = empty_root();
    let properties = parse(&source).unwrap();
    assert_eq!(write(&properties).unwrap(), source);
    assert!(properties.object_type().is_none());
    assert!(properties.checked().is_none());
    assert!(properties.colored().is_none());
    assert!(properties.drop_lines().is_none());
    assert!(properties.drop_style().is_none());
    assert!(properties.dx().is_none());
    assert!(properties.first_button().is_none());
    assert!(properties.fmla_group().is_none());
    assert!(properties.fmla_link().is_none());
    assert!(properties.fmla_range().is_none());
    assert!(properties.fmla_txbx().is_none());
    assert!(properties.horiz().is_none());
    assert!(properties.inc().is_none());
    assert!(properties.just_last_x().is_none());
    assert!(properties.lock_text().is_none());
    assert!(properties.max().is_none());
    assert!(properties.min().is_none());
    assert!(properties.multi_sel().is_none());
    assert!(properties.no_three_d().is_none());
    assert!(properties.no_three_d2().is_none());
    assert!(properties.page().is_none());
    assert!(properties.sel().is_none());
    assert!(properties.seltype().is_none());
    assert!(properties.text_h_align().is_none());
    assert!(properties.text_v_align().is_none());
    assert!(properties.val().is_none());
    assert!(properties.width_min().is_none());
    assert!(properties.edit_val().is_none());
    assert!(properties.multi_line().is_none());
    assert!(properties.vertical_bar().is_none());
    assert!(properties.password_edit().is_none());
    assert!(properties.item_list().is_none());
    assert!(properties.root_extension_list().is_none());

    assert!(!properties.effective_colored());
    assert_eq!(properties.effective_drop_lines(), 8);
    assert_eq!(properties.effective_dx(), 80);
    assert!(!properties.effective_first_button());
    assert!(!properties.effective_horiz());
    assert_eq!(properties.effective_inc(), 1);
    assert!(!properties.effective_just_last_x());
    assert!(!properties.effective_lock_text());
    assert!(!properties.effective_no_three_d());
    assert!(!properties.effective_no_three_d2());
    assert_eq!(
        properties.effective_text_h_align(),
        KnownOrUnknown::Known(TextHAlign::Left)
    );
    assert_eq!(
        properties.effective_text_v_align(),
        KnownOrUnknown::Known(TextVAlign::Top)
    );
    assert_eq!(properties.effective_val(), 0);
    assert_eq!(
        properties.effective_edit_val(),
        KnownOrUnknown::Known(EditValidation::Text)
    );
    assert!(!properties.effective_multi_line());
    assert!(!properties.effective_vertical_bar());
    assert!(!properties.effective_password_edit());
}

#[test]
fn all_thirty_one_scalars_retain_typed_presence_and_lexical_values() {
    let source = format!(
        r##"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}"
            objectType="CheckBox" checked="Mixed" colored="1" dropLines="30000"
            dropStyle="simple" dx="80" firstButton="true" fmlaGroup="Sheet1!$A$1"
            fmlaLink="#REF!" fmlaRange="A1:B2" fmlaTxbx="'Sheet 1'!A1" horiz="0"
            inc="2" justLastX="1" lockText="false" max="30000" min="0"
            multiSel=" 1,  3 " noThreeD="1" noThreeD2="0" page="4" sel="2"
            seltype="extended" textHAlign="distributed" textVAlign="center" val="5"
            widthMin="6" editVal="formula" multiLine="true" verticalBar="1"
            passwordEdit="false"/>"##
    );
    let properties = parse(source.as_bytes()).unwrap();

    assert_eq!(
        known_token(properties.object_type()),
        Some("CheckBox".into())
    );
    assert_eq!(known_token(properties.checked()), Some("Mixed".into()));
    assert_eq!(properties.colored(), Some(true));
    assert_eq!(properties.drop_lines(), Some(30_000));
    assert_eq!(known_token(properties.drop_style()), Some("simple".into()));
    assert_eq!(properties.dx(), Some(80));
    assert_eq!(properties.first_button(), Some(true));
    assert_eq!(
        properties.fmla_group().map(FormControlFormula::as_str),
        Some("Sheet1!$A$1")
    );
    assert_eq!(
        properties.fmla_link().map(FormControlFormula::as_str),
        Some("#REF!")
    );
    assert_eq!(
        properties.fmla_range().map(FormControlFormula::as_str),
        Some("A1:B2")
    );
    assert_eq!(
        properties.fmla_txbx().map(FormControlFormula::as_str),
        Some("'Sheet 1'!A1")
    );
    assert_eq!(properties.horiz(), Some(false));
    assert_eq!(properties.inc(), Some(2));
    assert_eq!(properties.just_last_x(), Some(true));
    assert_eq!(properties.lock_text(), Some(false));
    assert_eq!(properties.max(), Some(30_000));
    assert_eq!(properties.min(), Some(0));
    assert_eq!(properties.multi_sel(), Some(" 1,  3 "));
    assert_eq!(properties.no_three_d(), Some(true));
    assert_eq!(properties.no_three_d2(), Some(false));
    assert_eq!(properties.page(), Some(4));
    assert_eq!(properties.sel(), Some(2));
    assert_eq!(known_token(properties.seltype()), Some("extended".into()));
    assert_eq!(
        known_token(properties.text_h_align()),
        Some("distributed".into())
    );
    assert_eq!(
        known_token(properties.text_v_align()),
        Some("center".into())
    );
    assert_eq!(properties.val(), Some(5));
    assert_eq!(properties.width_min(), Some(6));
    assert_eq!(known_token(properties.edit_val()), Some("formula".into()));
    assert_eq!(properties.multi_line(), Some(true));
    assert_eq!(properties.vertical_bar(), Some(true));
    assert_eq!(properties.password_edit(), Some(false));

    let fields = [
        ScalarField::ObjectType,
        ScalarField::Checked,
        ScalarField::Colored,
        ScalarField::DropLines,
        ScalarField::DropStyle,
        ScalarField::Dx,
        ScalarField::FirstButton,
        ScalarField::FmlaGroup,
        ScalarField::FmlaLink,
        ScalarField::FmlaRange,
        ScalarField::FmlaTxbx,
        ScalarField::Horiz,
        ScalarField::Inc,
        ScalarField::JustLastX,
        ScalarField::LockText,
        ScalarField::Max,
        ScalarField::Min,
        ScalarField::MultiSel,
        ScalarField::NoThreeD,
        ScalarField::NoThreeD2,
        ScalarField::Page,
        ScalarField::Sel,
        ScalarField::SelType,
        ScalarField::TextHAlign,
        ScalarField::TextVAlign,
        ScalarField::Val,
        ScalarField::WidthMin,
        ScalarField::EditVal,
        ScalarField::MultiLine,
        ScalarField::VerticalBar,
        ScalarField::PasswordEdit,
    ];
    for field in fields {
        assert!(
            properties.scalar(field).is_some(),
            "missing scalar {field:?}"
        );
    }
}

#[test]
fn enum_families_accept_every_admitted_variant() {
    let source = empty_root();
    for (value, wire) in [
        (ObjectType::Button, "Button"),
        (ObjectType::CheckBox, "CheckBox"),
        (ObjectType::Drop, "Drop"),
        (ObjectType::GBox, "GBox"),
        (ObjectType::Label, "Label"),
        (ObjectType::List, "List"),
        (ObjectType::Radio, "Radio"),
        (ObjectType::Scroll, "Scroll"),
        (ObjectType::Spin, "Spin"),
        (ObjectType::EditBox, "EditBox"),
        (ObjectType::Dialog, "Dialog"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::ObjectType,
            Some(ScalarValue::ObjectType(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"objectType="{wire}""#))
        );
    }
    for (value, wire) in [
        (Checked::Unchecked, "Unchecked"),
        (Checked::Checked, "Checked"),
        (Checked::Mixed, "Mixed"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::Checked,
            Some(ScalarValue::Checked(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"checked="{wire}""#))
        );
    }
    for (value, wire) in [
        (DropStyle::Combo, "combo"),
        (DropStyle::ComboEdit, "comboedit"),
        (DropStyle::Simple, "simple"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::DropStyle,
            Some(ScalarValue::DropStyle(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"dropStyle="{wire}""#))
        );
    }
    for (value, wire) in [
        (SelectionType::Single, "single"),
        (SelectionType::Multi, "multi"),
        (SelectionType::Extended, "extended"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::SelType,
            Some(ScalarValue::SelectionType(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"seltype="{wire}""#))
        );
    }
    for (value, wire) in [
        (EditValidation::Text, "text"),
        (EditValidation::Integer, "integer"),
        (EditValidation::Number, "number"),
        (EditValidation::Reference, "reference"),
        (EditValidation::Formula, "formula"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::EditVal,
            Some(ScalarValue::EditValidation(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"editVal="{wire}""#))
        );
    }
    for (value, wire) in [
        (TextHAlign::Left, "left"),
        (TextHAlign::Center, "center"),
        (TextHAlign::Right, "right"),
        (TextHAlign::Justify, "justify"),
        (TextHAlign::Distributed, "distributed"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::TextHAlign,
            Some(ScalarValue::TextHAlign(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"textHAlign="{wire}""#))
        );
    }
    for (value, wire) in [
        (TextVAlign::Top, "top"),
        (TextVAlign::Center, "center"),
        (TextVAlign::Bottom, "bottom"),
        (TextVAlign::Justify, "justify"),
        (TextVAlign::Distributed, "distributed"),
    ] {
        let output = replace_scalar(
            &source,
            ScalarField::TextVAlign,
            Some(ScalarValue::TextVAlign(value)),
        )
        .unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"textVAlign="{wire}""#))
        );
    }
}

#[test]
fn booleans_ranges_and_unknown_enum_tokens_are_bounded() {
    let bool_attributes = [
        "colored",
        "firstButton",
        "horiz",
        "justLastX",
        "lockText",
        "noThreeD",
        "noThreeD2",
        "multiLine",
        "verticalBar",
        "passwordEdit",
    ];
    for attribute in bool_attributes {
        for value in ["0", "false", "1", "true"] {
            parse(&root_with_attribute(attribute, value)).unwrap();
        }
        assert_invalid(&root_with_attribute(attribute, "yes"), "invalid boolean");
    }

    for attribute in ["dropLines", "inc", "max", "min", "page"] {
        parse(&root_with_attribute(attribute, "0")).unwrap();
        parse(&root_with_attribute(attribute, "30000")).unwrap();
        assert_invalid(
            &root_with_attribute(attribute, "30001"),
            "out-of-range unsigned value",
        );
        assert_invalid(
            &root_with_attribute(attribute, "4294967296"),
            "overflowing unsigned value",
        );
        parse(&root_with_attribute(attribute, "  8 \n ")).unwrap();
    }

    let unknown = root_with_attribute("objectType", "FutureControl");
    let properties = parse(&unknown).unwrap();
    assert_eq!(
        known_token(properties.object_type()),
        Some("FutureControl".into())
    );
    assert_eq!(write(&properties).unwrap(), unknown);
    let changed = replace_scalar(
        &unknown,
        ScalarField::LockText,
        Some(ScalarValue::Boolean(true)),
    )
    .unwrap();
    assert!(
        changed
            .windows(b"objectType=\"FutureControl\"".len())
            .any(|window| { window == b"objectType=\"FutureControl\"" })
    );

    let alignment = parse(&root_with_attribute("textHAlign", " left ")).unwrap();
    assert_eq!(known_token(alignment.text_h_align()), Some(" left ".into()));
    let multi_sel = parse(&root_with_attribute("multiSel", " 1,  3 ")).unwrap();
    assert_eq!(multi_sel.multi_sel(), Some(" 1,  3 "));
}

#[test]
fn x14_expanded_names_and_item_string_values_are_exactly_preserved() {
    let source = format!(
        r#"<?xml version="1.0"?>
<x14:formControlPr xmlns:x14="{FORM_CONTROL_NAMESPACE}" xmlns:c="{CORE_NAMESPACE}"
    xmlns:u="urn:unknown" objectType="List" fmlaLink="  Sheet1!$A$1  "
    seltype="multi" multiSel=" 1,  3 " x14:foreign="keep">
  <?litchi-root keep?>
  <!-- root comment -->
  <x14:itemLst>
    <?litchi-item keep?>
    <x14:item val="  first &amp; &#10; second  "/>
    <x14:item val="second"/>
    <x14:extLst><!-- item extension --><c:ext uri="urn:item"><u:payload><![CDATA[item <payload>]]></u:payload></c:ext></x14:extLst>
  </x14:itemLst>
  <x14:extLst><!-- root extension --><c:ext><u:payload><![CDATA[root <payload>]]></u:payload></c:ext><x14:ext uri="urn:foreign"/></x14:extLst>
</x14:formControlPr>
"#
    );
    let source = source.into_bytes();
    let view = inspect(&source).unwrap();
    assert_eq!(view.source(), source.as_slice());
    assert!(
        view.properties().source_bytes().is_none(),
        "borrowed inspect must not retain a complete owned source copy"
    );
    assert_eq!(write(view.properties()).unwrap(), source);
    let properties = parse(&source).unwrap();
    assert_eq!(
        properties.object_type().and_then(KnownOrUnknown::known),
        Some(&ObjectType::List)
    );
    assert_eq!(properties.unknown_attributes().len(), 1);
    assert_eq!(properties.unknown_attributes()[0].name(), b"x14:foreign");
    assert_eq!(
        properties.item_list().unwrap().items()[0].value(),
        "  first & \n second  "
    );
    assert_eq!(properties.item_list().unwrap().items()[1].val(), "second");
    assert_eq!(
        properties.fmla_link().map(FormControlFormula::as_str),
        Some("  Sheet1!$A$1  ")
    );
    assert_eq!(properties.multi_sel(), Some(" 1,  3 "));
    assert_eq!(known_token(properties.seltype()), Some("multi".into()));
    assert_eq!(properties.item_list().unwrap().extension_list().unwrap().xml(),
        b"<x14:extLst><!-- item extension --><c:ext uri=\"urn:item\"><u:payload><![CDATA[item <payload>]]></u:payload></c:ext></x14:extLst>");
    assert_eq!(properties.root_extension_list().unwrap().xml(),
        b"<x14:extLst><!-- root extension --><c:ext><u:payload><![CDATA[root <payload>]]></u:payload></c:ext><x14:ext uri=\"urn:foreign\"/></x14:extLst>");
    assert_eq!(write(&properties).unwrap(), source);

    let changed = replace_scalar(
        &source,
        ScalarField::NoThreeD,
        Some(ScalarValue::Boolean(true)),
    )
    .unwrap();
    assert!(changed.starts_with(b"<?xml version=\"1.0\"?>\n<x14:formControlPr"));
    assert!(
        changed
            .windows(b"<!-- root comment -->".len())
            .any(|window| { window == b"<!-- root comment -->" })
    );
    assert!(
        changed
            .windows(b"<?litchi-root keep?>".len())
            .any(|window| { window == b"<?litchi-root keep?>" })
    );
    assert!(
        changed
            .windows(b"<?litchi-item keep?>".len())
            .any(|window| { window == b"<?litchi-item keep?>" })
    );
    assert!(
        changed
            .windows(b"objectType=\"List\" fmlaLink=\"  Sheet1!$A$1  \"".len())
            .any(|window| window == b"objectType=\"List\" fmlaLink=\"  Sheet1!$A$1  \"")
    );
    assert!(
        changed
            .windows(b"multiSel=\" 1,  3 \" x14:foreign=\"keep\"".len())
            .any(|window| window == b"multiSel=\" 1,  3 \" x14:foreign=\"keep\"")
    );
    assert!(
        changed
            .windows(b"<![CDATA[root <payload>]]>".len())
            .any(|window| { window == b"<![CDATA[root <payload>]]>" })
    );
    assert!(
        changed
            .windows(b"x14:foreign=\"keep\"".len())
            .any(|window| { window == b"x14:foreign=\"keep\"" })
    );
    assert_eq!(
        parse(&changed).unwrap().item_list().unwrap().items()[0].value(),
        "  first & \n second  "
    );
}

#[test]
fn opaque_empty_descendants_count_toward_exact_depth_cap() {
    let source = root_with_body(&format!(
        r#"<extLst><c:ext xmlns:c="{CORE_NAMESPACE}"><u:outer xmlns:u="urn:opaque"><u:inner/></u:outer></c:ext></extLst>"#
    ));
    parse_with_limits(&source, &Limits::new().with_max_depth(5)).unwrap();
    assert!(parse_with_limits(&source, &Limits::new().with_max_depth(4)).is_err());
}

#[test]
fn detached_rewrite_materializes_shadowed_default_namespaces_for_opaque_descendants() {
    let source = format!(
        r#"<x14:formControlPr xmlns:x14="{FORM_CONTROL_NAMESPACE}" xmlns="urn:root-default" xmlns:c="{CORE_NAMESPACE}" objectType="List">
  <x14:itemLst xmlns="urn:item-default">
    <x14:item val="one"/>
    <x14:extLst><c:ext uri="urn:item"><opaque><child/></opaque></c:ext></x14:extLst>
  </x14:itemLst>
  <x14:extLst><c:ext uri="urn:root"><opaque><child/></opaque></c:ext></x14:extLst>
</x14:formControlPr>"#
    )
    .into_bytes();
    let mut properties = parse(&source).unwrap();
    let root_extension = properties.root_extension_list().unwrap();
    assert!(
        root_extension
            .namespaces()
            .iter()
            .any(|binding| binding.prefix().is_empty() && binding.uri() == "urn:root-default")
    );
    let item_extension = properties.item_list().unwrap().extension_list().unwrap();
    assert!(
        item_extension
            .namespaces()
            .iter()
            .any(|binding| binding.prefix().is_empty() && binding.uri() == "urn:item-default")
    );

    properties.set_no_three_d(Some(true));
    let output = write(&properties).unwrap();
    let reopened = parse(&output).unwrap();
    let output_text = String::from_utf8_lossy(&output);
    assert!(
        output_text.contains(r#"xmlns="urn:root-default""#),
        "detached root opaque XML lost its default namespace context"
    );
    assert!(
        output_text.contains(r#"xmlns="urn:item-default""#),
        "detached item opaque XML lost its shadowed default namespace context"
    );
    assert!(
        reopened
            .item_list()
            .unwrap()
            .namespaces()
            .iter()
            .any(|binding| binding.prefix().is_empty() && binding.uri() == "urn:item-default")
    );
}

#[test]
fn missing_or_foreign_extension_uri_stays_opaque_and_unknown_item_markup_is_not_discarded() {
    let source = root_with_body(&format!(
        r#"<itemLst>
  <item val="one" extra="preserve"/>
  <extLst><c:ext xmlns:c="{CORE_NAMESPACE}"><u:payload xmlns:u="urn:payload"><![CDATA[item]]></u:payload></c:ext></extLst>
</itemLst>
<extLst><c:ext xmlns:c="{CORE_NAMESPACE}"><u:payload xmlns:u="urn:payload"><![CDATA[root]]></u:payload></c:ext><x14:ext xmlns:x14="{FORM_CONTROL_NAMESPACE}" uri="foreign"/></extLst>"#
    ));
    let properties = parse(&source).unwrap();
    assert_eq!(properties.item_list().unwrap().items()[0].value(), "one");
    assert!(properties.item_list().unwrap().items()[0].has_opaque_markup());
    assert_eq!(properties.item_list().unwrap().extension_list().unwrap().xml(),
        b"<extLst><c:ext xmlns:c=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><u:payload xmlns:u=\"urn:payload\"><![CDATA[item]]></u:payload></c:ext></extLst>");
    let root_extension = properties.root_extension_list().unwrap().xml();
    assert!(
        root_extension
            .windows(b"<c:ext".len())
            .any(|window| window == b"<c:ext")
    );
    assert!(
        root_extension
            .windows(b"<x14:ext".len())
            .any(|window| window == b"<x14:ext")
    );
    assert_eq!(write(&properties).unwrap(), source);

    let changed = replace_scalar(
        &source,
        ScalarField::LockText,
        Some(ScalarValue::Boolean(true)),
    )
    .unwrap();
    assert!(
        changed
            .windows(b"<![CDATA[item]]>".len())
            .any(|window| window == b"<![CDATA[item]]>")
    );
    assert!(
        changed
            .windows(b"<![CDATA[root]]>".len())
            .any(|window| window == b"<![CDATA[root]]>")
    );
    assert!(replace_items(&source, &[Item::new("changed").unwrap()]).is_err());
}

#[test]
fn item_replacements_insert_and_remove_preserve_both_extension_lists() {
    let source = root_with_body(&format!(
        r#"<itemLst>
  <?item-pi keep?>
  <!-- between items -->
  <item val="one"/><item val="two"/>
  <extLst><c:ext xmlns:c="{CORE_NAMESPACE}" uri="item"/></extLst>
</itemLst>
<extLst><c:ext xmlns:c="{CORE_NAMESPACE}" uri="root"/></extLst>"#
    ));
    let noop = replace_items(
        &source,
        &[Item::new("one").unwrap(), Item::new("two").unwrap()],
    )
    .unwrap();
    assert_eq!(noop, source);

    let replaced = replace_items(
        &source,
        &[
            Item::new("new & value").unwrap(),
            Item::new("last").unwrap(),
        ],
    )
    .unwrap();
    let properties = parse(&replaced).unwrap();
    let items = properties.item_list().unwrap().items();
    assert_eq!(
        items.iter().map(Item::value).collect::<Vec<_>>(),
        ["new & value", "last"]
    );
    assert!(
        String::from_utf8(replaced.clone())
            .unwrap()
            .contains("new &amp; value")
    );
    let replaced_text = String::from_utf8(replaced).unwrap();
    assert!(replaced_text.contains("<?item-pi keep?>"));
    assert!(replaced_text.contains("<!-- between items -->"));
    assert!(replaced_text.contains("uri=\"item\""));
    assert!(replaced_text.contains("uri=\"root\""));

    let inserted = insert_item(&source, 1, Item::new("middle").unwrap()).unwrap();
    assert_eq!(
        parse(&inserted).unwrap().item_list().unwrap().items()[1].value(),
        "middle"
    );
    let removed = remove_item(&inserted, 1).unwrap();
    assert_eq!(
        parse(&removed).unwrap().item_list().unwrap().items().len(),
        2
    );
    assert!(remove_item(&source, 2).is_err());
}

#[test]
fn item_val_is_required_unqualified_string_and_invalid_grammar_is_rejected() {
    assert_invalid(
        &root_with_body("<itemLst><item/></itemLst>"),
        "item without val",
    );
    assert_invalid(
        &root_with_body(&format!(
            r#"<itemLst><item xmlns:x14="{FORM_CONTROL_NAMESPACE}" x14:val="qualified"/></itemLst>"#
        )),
        "qualified item val",
    );
    assert_invalid(&root_with_body("<itemLst><item val="), "malformed item XML");
    assert_invalid(
        &root_with_body("<itemLst><item val=\"x\"><![CDATA[bad]]></item></itemLst>"),
        "item CDATA",
    );

    let invalid_sources = [
        format!(r#"<formControlPr xmlns="{CORE_NAMESPACE}"/>"#),
        r#"<formControlPr/>"#.to_owned(),
        format!(
            r#"<x14:formControlPr xmlns:x14="{FORM_CONTROL_NAMESPACE}"><itemLst/></x14:formControlPr>"#
        ),
        format!(
            r#"<x14:formControlPr xmlns:x14="{FORM_CONTROL_NAMESPACE}" xmlns:c="{CORE_NAMESPACE}"><c:itemLst/></x14:formControlPr>"#
        ),
        String::from_utf8(root_with_body("text")).unwrap(),
        String::from_utf8(root_with_body("<![CDATA[text]]>")).unwrap(),
        String::from_utf8(root_with_body("<itemLst/><itemLst/>")).unwrap(),
        String::from_utf8(root_with_body("<extLst/><extLst/>")).unwrap(),
        String::from_utf8(root_with_body("<extLst/><itemLst/>")).unwrap(),
        String::from_utf8(root_with_body(
            "<itemLst><item val=\"x\"/><extLst/><item val=\"y\"/></itemLst>",
        ))
        .unwrap(),
        String::from_utf8(root_with_body("<itemLst><unknown/></itemLst>")).unwrap(),
        String::from_utf8(root_with_body("<nested><child/></nested>")).unwrap(),
    ];
    for source in invalid_sources {
        assert_invalid(
            source.as_bytes(),
            "invalid form-control grammar or namespace",
        );
    }

    let standalone = format!(
        "\u{feff}<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n{}\n",
        String::from_utf8(empty_root()).unwrap()
    )
    .into_bytes();
    assert_eq!(write(parse(&standalone).unwrap()).unwrap(), standalone);
    let embedded = format!(
        "<wrapper>{}</wrapper>",
        String::from_utf8(empty_root()).unwrap()
    );
    assert_invalid(embedded.as_bytes(), "embedded form-control part");
}

#[test]
fn detached_writer_materializes_x14_namespace_for_opaque_extensions() {
    let source = format!(
        r#"<x14:formControlPr xmlns:x14="{FORM_CONTROL_NAMESPACE}" xmlns:c="{CORE_NAMESPACE}" objectType="Button"><x14:extLst><c:ext uri="urn:opaque"><u:payload xmlns:u="urn:u"><![CDATA[x]]></u:payload></c:ext></x14:extLst></x14:formControlPr>"#
    );
    let mut properties = parse(source.as_bytes()).unwrap();
    properties.set_lock_text(Some(true));
    let output = write(&properties).unwrap();
    let output_text = String::from_utf8(output.clone()).unwrap();
    assert!(output_text.starts_with(&format!(
        r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}""#
    )));
    assert!(output_text.contains(&format!(r#"xmlns:x14="{FORM_CONTROL_NAMESPACE}""#)));
    assert!(output_text.contains(&format!(r#"xmlns:c="{CORE_NAMESPACE}""#)));
    assert!(output_text.contains("<![CDATA[x]]>"));
    assert_eq!(parse(&output).unwrap().lock_text(), Some(true));
}

#[test]
fn limits_fail_at_one_byte_or_event_under_each_local_ceiling() {
    let empty = empty_root();
    assert!(
        parse_with_limits(&empty, &Limits::new().with_max_part_bytes(empty.len() - 1)).is_err()
    );
    parse_with_limits(&empty, &Limits::new().with_max_part_bytes(empty.len())).unwrap();

    let output = write(Properties::new()).unwrap();
    assert!(
        write_with_limits(
            Properties::new(),
            &Limits::new().with_max_output_bytes(output.len() - 1)
        )
        .is_err()
    );
    assert_eq!(
        write_with_limits(
            Properties::new(),
            &Limits::new().with_max_output_bytes(output.len())
        )
        .unwrap(),
        output
    );

    // An empty root emits exactly one XML event and one EOF event.
    assert!(parse_with_limits(&empty, &Limits::new().with_max_events(1)).is_err());
    parse_with_limits(&empty, &Limits::new().with_max_events(2)).unwrap();
    assert!(parse_with_limits(&empty, &Limits::new().with_max_depth(0)).is_err());
    parse_with_limits(&empty, &Limits::new().with_max_depth(1)).unwrap();

    let one_item = root_with_body("<itemLst><item val=\"abc\"/></itemLst>");
    assert!(parse_with_limits(&one_item, &Limits::new().with_max_items(0)).is_err());
    parse_with_limits(&one_item, &Limits::new().with_max_items(1)).unwrap();
    assert!(parse_with_limits(&one_item, &Limits::new().with_max_item_value_bytes(2)).is_err());
    parse_with_limits(&one_item, &Limits::new().with_max_item_value_bytes(3)).unwrap();
    assert!(parse_with_limits(&one_item, &Limits::new().with_max_depth(2)).is_err());
    parse_with_limits(&one_item, &Limits::new().with_max_depth(3)).unwrap();

    let one_attribute = empty_root();
    assert!(parse_with_limits(&one_attribute, &Limits::new().with_max_attributes(0)).is_err());
    parse_with_limits(&one_attribute, &Limits::new().with_max_attributes(1)).unwrap();

    let opaque = root_with_body(&format!(
        r#"<extLst><c:ext xmlns:c="{CORE_NAMESPACE}" uri="urn:opaque"><u:p xmlns:u="urn:p">payload larger than the namespace URI</u:p></c:ext></extLst>"#
    ));
    let opaque_len = parse(&opaque)
        .unwrap()
        .root_extension_list()
        .unwrap()
        .xml()
        .len();
    assert!(
        parse_with_limits(
            &opaque,
            &Limits::new().with_max_opaque_bytes(opaque_len - 1)
        )
        .is_err()
    );
    parse_with_limits(&opaque, &Limits::new().with_max_opaque_bytes(opaque_len)).unwrap();
}

#[test]
fn scalar_attribute_insertion_reserves_the_full_candidate() {
    let source = empty_root();
    let value = ScalarValue::Boolean(true);
    let expected = replace_scalar(&source, ScalarField::LockText, Some(value.clone())).unwrap();
    assert!(expected.len() > source.len());

    let exact =
        inspect_with_limits(&source, Limits::new().with_max_output_bytes(expected.len())).unwrap();
    assert_eq!(
        exact
            .replace_scalar(ScalarField::LockText, Some(value.clone()))
            .unwrap(),
        expected
    );

    let under = inspect_with_limits(
        &source,
        Limits::new().with_max_output_bytes(expected.len() - 1),
    )
    .unwrap();
    assert!(
        under
            .replace_scalar(ScalarField::LockText, Some(value))
            .is_err()
    );
}

#[test]
fn formula_multisel_and_generated_output_limits_have_exact_boundaries() {
    let source = format!(
        r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}" objectType="List" fmlaLink="Sheet1!$A$1" seltype="multi" multiSel="1, 3"/>"#
    )
    .into_bytes();
    let view = inspect(&source).unwrap();
    let source_len = source.len();
    assert_eq!(
        write_with_limits(
            view.properties(),
            &Limits::new().with_max_output_bytes(source_len)
        )
        .unwrap(),
        source
    );
    assert!(
        write_with_limits(
            view.properties(),
            &Limits::new().with_max_output_bytes(source_len - 1)
        )
        .is_err()
    );

    let mut properties = parse(&source).unwrap();
    properties.set_no_three_d(Some(true));
    let output = write(&properties).unwrap();
    assert!(String::from_utf8_lossy(&output).contains("Sheet1!$A$1"));
    assert!(String::from_utf8_lossy(&output).contains("multiSel=\"1, 3\""));
    assert_eq!(
        write_with_limits(
            &properties,
            &Limits::new().with_max_output_bytes(output.len())
        )
        .unwrap(),
        output
    );
    assert!(
        write_with_limits(
            &properties,
            &Limits::new().with_max_output_bytes(output.len() - 1)
        )
        .is_err()
    );
}

#[test]
fn retained_source_and_opaque_bytes_are_precharged_once_at_boundary() {
    let source = root_with_body("<extLst/>");
    // Discover the public boundary instead of duplicating private Vec/Arc
    // accounting (including target-dependent namespace-binding slot sizes).
    // The parser and dirty-model validator must expose the same retained
    // budget for this source-backed opaque extension.
    let parse_required = minimum_parse_retained_limit(&source);
    assert!(parse_required > 0);
    parse_with_limits(
        &source,
        &Limits::new().with_max_retained_bytes(parse_required),
    )
    .unwrap();
    assert_retained_limit(
        parse_with_limits(
            &source,
            &Limits::new().with_max_retained_bytes(parse_required - 1),
        ),
        parse_required - 1,
    );

    let mut dirty = parse(&source).unwrap();
    dirty.set_just_last_x(Some(true));
    let dirty_required = minimum_write_retained_limit(&dirty);
    assert_eq!(dirty_required, parse_required);
    write_with_limits(
        &dirty,
        &Limits::new().with_max_retained_bytes(parse_required),
    )
    .unwrap();
    assert_retained_limit(
        write_with_limits(
            &dirty,
            &Limits::new().with_max_retained_bytes(parse_required - 1),
        ),
        parse_required - 1,
    );

    let fragment = b"<extLst/>".to_vec();
    let mut detached = Properties::new();
    detached.set_root_extension_list(Some(OpaqueXml::new(fragment.clone()).unwrap()));
    let detached_required = minimum_write_retained_limit(&detached);
    assert!(detached_required > 0);
    write_with_limits(
        &detached,
        &Limits::new().with_max_retained_bytes(detached_required),
    )
    .unwrap();
    assert_retained_limit(
        write_with_limits(
            &detached,
            &Limits::new().with_max_retained_bytes(detached_required - 1),
        ),
        detached_required - 1,
    );
}

#[test]
fn xml10_character_boundaries_cover_items_scalar_strings_and_formulas() {
    let del = '\u{7f}';
    let item_value = format!("item{del}value");
    let item = Item::new(item_value.clone()).unwrap();
    let item_list = ItemList::new([item]).unwrap();
    let mut properties = Properties::new();
    properties.set_object_type(Some(ObjectType::List));
    properties.set_item_list(Some(item_list)).unwrap();
    let output = write(&properties).unwrap();
    assert_eq!(
        parse(&output).unwrap().item_list().unwrap().items()[0].value(),
        item_value
    );

    let scalar_value = format!("left{del}");
    let scalar_source =
        format!(r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}" textHAlign="{scalar_value}"/>"#)
            .into_bytes();
    let scalar_properties = parse(&scalar_source).unwrap();
    assert_eq!(
        known_token(scalar_properties.text_h_align()),
        Some(scalar_value.clone())
    );
    assert_eq!(write(&scalar_properties).unwrap(), scalar_source);
    let mut scalar_changed = scalar_properties.clone();
    scalar_changed.set_just_last_x(Some(true));
    let detached_scalar = write(&scalar_changed).unwrap();
    assert_eq!(
        known_token(parse(&detached_scalar).unwrap().text_h_align()),
        Some(scalar_value.clone())
    );

    let formula_value = format!("A1{del}");
    let formula_source =
        format!(r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}" fmlaLink="{formula_value}"/>"#)
            .into_bytes();
    let formula_properties = parse(&formula_source).unwrap();
    assert_eq!(
        formula_properties
            .fmla_link()
            .map(FormControlFormula::as_str),
        Some(formula_value.as_str())
    );
    assert_eq!(write(&formula_properties).unwrap(), formula_source);
    assert!(
        FormControlFormula::new(formula_value).is_err(),
        "formula grammar may reject DEL independently of XML character validity"
    );

    assert!(OpaqueXml::new(format!("<opaque>{del}</opaque>").into_bytes()).is_ok());
    for character in ['\u{0}', '\u{1}', '\u{b}', '\u{c}', '\u{fffe}', '\u{ffff}'] {
        assert!(
            Item::new(format!("item{character}")).is_err(),
            "item accepted XML-invalid U+{:04X}",
            character as u32
        );
        assert!(
            OpaqueXml::new(format!("<opaque>{character}</opaque>").into_bytes()).is_err(),
            "opaque XML accepted XML-invalid U+{:04X}",
            character as u32
        );
        let scalar = format!(
            r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}" textHAlign="left{character}"/>"#
        );
        assert!(
            parse(scalar.as_bytes()).is_err(),
            "scalar string accepted XML-invalid U+{:04X}",
            character as u32
        );
        let formula = format!(
            r#"<formControlPr xmlns="{FORM_CONTROL_NAMESPACE}" fmlaLink="A1{character}"/>"#
        );
        assert!(
            parse(formula.as_bytes()).is_err(),
            "formula source accepted XML-invalid U+{:04X}",
            character as u32
        );
    }
}

#[test]
fn typed_formula_setters_do_not_introduce_ref_errors() {
    assert!(FormControlFormula::new("#REF!").is_err());

    let source = root_with_attribute("fmlaLink", "#REF!");
    let changed = replace_scalar(
        &source,
        ScalarField::LockText,
        Some(ScalarValue::Boolean(true)),
    )
    .unwrap();
    assert_eq!(
        parse(&changed).unwrap().fmla_link().unwrap().as_str(),
        "#REF!"
    );

    let valid_formula = FormControlFormula::new("Sheet1!$A$1").unwrap();
    let output = replace_scalar(
        &source,
        ScalarField::FmlaLink,
        Some(ScalarValue::Formula(valid_formula)),
    )
    .unwrap();
    assert_eq!(
        parse(&output).unwrap().fmla_link().unwrap().as_str(),
        "Sheet1!$A$1"
    );
}
