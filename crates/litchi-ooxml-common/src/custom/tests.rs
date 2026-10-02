use super::Host;
use super::codec::{decode, encode};
use super::model::{Props, Value};
use super::package::custom_part_name;
use super::schema::{
    DOCUMENT_SUMMARY_FORMAT_ID, FORMAT_ID, MAX_NAME_CHARS, MAX_PROPERTIES, MAX_TEXT_BYTES,
    MAX_XML_BYTES, MAX_XML_NODES, PART_TARGET, SUMMARY_FORMAT_ID,
};
use crate::Error;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::rel::TargetMode;
use std::sync::Arc;

const PREFIX: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" "#,
    r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#
);
const HEADER_FONT: &str = "ClassificationContentMarkingHeaderFontProps";
const HEADER_SHAPES: &str = "ClassificationContentMarkingHeaderShapeIds";
const HEADER_TEXT: &str = "ClassificationContentMarkingHeaderText";
const HEADER_LOCATIONS: &str = "ClassificationContentMarkingHeaderLocations";

fn property(pid: i32, name: &str, value: &str) -> String {
    format!(r#"<property fmtid="{FORMAT_ID}" pid="{pid}" name="{name}">{value}</property>"#)
}

fn document(body: &str) -> String {
    format!("{PREFIX}{body}</Properties>")
}

#[test]
fn concise_crud_moves_values_and_orders_names() {
    let mut props = Props::new();
    assert_eq!(props.insert("Version", 1_i32).expect("insert"), None);
    props.insert("Author", "Ada").expect("insert");
    assert_eq!(props.names().collect::<Vec<_>>(), ["Author", "Version"]);
    assert_eq!(props.get("VERSION"), Some(&Value::I32(1)));
    assert!(props.contains("author"));
    assert_eq!(
        props.insert("Version", 2_i32).expect("replace"),
        Some(Value::I32(1))
    );
    assert_eq!(props.remove("vErSiOn"), Some(Value::I32(2)));
    assert!(!props.contains("Version"));
    props.clear();
    assert!(props.is_empty());
    props.insert("AfterClear", true).expect("PID reset insert");
    let xml = encode(&props).expect("encode");
    assert!(String::from_utf8_lossy(&xml).contains(r#"pid="2""#));
}

#[test]
fn case_insensitive_duplicate_names_are_rejected_on_insert_and_read() {
    let mut props = Props::new();
    props.insert("Project", "one").expect("first insert");
    assert!(props.insert("project", "two").is_err());

    let xml = document(&format!(
        "{}{}",
        property(2, "Project", "<vt:lpwstr>one</vt:lpwstr>"),
        property(3, "PROJECT", "<vt:lpwstr>two</vt:lpwstr>")
    ));
    assert!(decode(xml.as_bytes()).is_err());
}

#[test]
fn names_use_canonical_unicode_caseless_identity() {
    let mut props = Props::new();
    props.insert("Straße", "stored spelling").expect("insert");
    assert_eq!(
        props.get("STRASSE"),
        Some(&Value::Text("stored spelling".to_owned()))
    );
    assert!(props.contains("strasse"));
    assert!(props.insert("STRASSE", "duplicate").is_err());
    assert_eq!(
        props.remove("STRASSE"),
        Some(Value::Text("stored spelling".to_owned()))
    );
    assert!(props.is_empty());
}

#[test]
fn all_supported_values_round_trip_deterministically_in_pid_order() {
    let xml = document(&format!(
        "{}{}{}{}{}{}{}{}{}",
        property(9, "Last", "<vt:lpstr>narrow &amp; exact</vt:lpstr>"),
        property(2, "Empty", "<vt:empty/>"),
        property(3, "Text", "<vt:lpwstr>wide</vt:lpwstr>"),
        property(4, "I32", "<vt:i4>-7</vt:i4>"),
        property(5, "I64", "<vt:i8>9000000000</vt:i8>"),
        property(6, "F32", "<vt:r4>1.25</vt:r4>"),
        property(7, "F64", "<vt:r8>2.5</vt:r8>"),
        property(8, "Bool", "<vt:bool>1</vt:bool>"),
        property(
            10,
            "Time",
            "<vt:filetime>2024-05-06T07:08:09.123Z</vt:filetime>"
        )
    ));
    let props = decode(xml.as_bytes()).expect("decode");
    assert_eq!(props.get("Empty"), Some(&Value::Empty));
    assert_eq!(props.get("I64"), Some(&Value::I64(9_000_000_000)));
    assert_eq!(props.get("F32"), Some(&Value::F32(1.25)));
    let first = encode(&props).expect("first encode");
    let second = encode(&props).expect("second encode");
    assert_eq!(first, second);
    let output = String::from_utf8(first.clone()).expect("UTF-8 XML");
    assert!(output.contains("<vt:lpstr>narrow &amp; exact</vt:lpstr>"));
    assert!(output.find(r#"pid="2""#) < output.find(r#"pid="9""#));
    assert_eq!(
        encode(&decode(&first).expect("round-trip decode")).expect("encode"),
        first
    );
}

#[test]
fn filetime_is_rfc3339_not_a_numeric_windows_counter() {
    let valid = document(&property(
        2,
        "When",
        "<vt:filetime>2020-01-02T03:04:05+02:30</vt:filetime>",
    ));
    let props = decode(valid.as_bytes()).expect("RFC3339 filetime");
    let output = String::from_utf8(encode(&props).expect("encode")).expect("UTF-8");
    assert!(output.contains("2020-01-02T00:34:05Z"));

    let numeric = document(&property(
        2,
        "When",
        "<vt:filetime>132223104000000000</vt:filetime>",
    ));
    assert!(decode(numeric.as_bytes()).is_err());
}

#[test]
fn non_finite_floats_are_rejected_on_insert_and_read() {
    let mut props = Props::new();
    assert!(props.insert("NaN", f64::NAN).is_err());
    assert!(props.insert("Infinity", f32::INFINITY).is_err());
    let xml = document(&property(2, "NaN", "<vt:r8>NaN</vt:r8>"));
    assert!(decode(xml.as_bytes()).is_err());
}

#[test]
fn root_namespace_and_property_cardinality_are_strict() {
    let wrong_namespace = r#"<Properties xmlns="urn:wrong"/>"#;
    assert!(decode(wrong_namespace.as_bytes()).is_err());
    let missing = document(&property(2, "Missing", ""));
    assert!(decode(missing.as_bytes()).is_err());
    let duplicate = document(&property(
        2,
        "Duplicate",
        "<vt:i4>1</vt:i4><vt:i4>2</vt:i4>",
    ));
    assert!(decode(duplicate.as_bytes()).is_err());
    let wrong_value_namespace = document(&property(
        2,
        "Wrong",
        r#"<x:i4 xmlns:x="urn:wrong">1</x:i4>"#,
    ));
    assert!(decode(wrong_value_namespace.as_bytes()).is_err());
}

#[test]
fn malformed_and_duplicate_pids_are_rejected() {
    let below_minimum = document(&property(1, "Low", "<vt:i4>1</vt:i4>"));
    assert!(decode(below_minimum.as_bytes()).is_err());
    let duplicate = document(&format!(
        "{}{}",
        property(2, "One", "<vt:i4>1</vt:i4>"),
        property(2, "Two", "<vt:i4>2</vt:i4>")
    ));
    assert!(decode(duplicate.as_bytes()).is_err());
    let malformed = document(&property(2, "Bad", "<vt:i4>not-an-int</vt:i4>"));
    assert!(decode(malformed.as_bytes()).is_err());
}

#[test]
fn exhausted_pid_space_allows_replacement_but_not_allocation() {
    let xml = document(&property(i32::MAX, "Last", "<vt:lpwstr>value</vt:lpwstr>"));
    let mut props = decode(xml.as_bytes()).expect("maximum PID is valid");
    assert_eq!(
        props.insert("Last", "replacement").expect("replace"),
        Some(Value::Text("value".to_owned()))
    );
    assert!(props.insert("New", "cannot allocate").is_err());
}

#[test]
fn forbidden_and_malformed_format_ids_are_rejected() {
    for format_id in [SUMMARY_FORMAT_ID, DOCUMENT_SUMMARY_FORMAT_ID, "not-a-guid"] {
        let xml = document(&format!(
            r#"<property fmtid="{format_id}" pid="2" name="Bad"><vt:i4>1</vt:i4></property>"#
        ));
        assert!(decode(xml.as_bytes()).is_err());
    }
}

#[test]
fn dtd_unknown_entities_and_malformed_xml_are_rejected() {
    let dtd = format!(
        r#"<!DOCTYPE Properties [<!ENTITY x "expanded">]>{PREFIX}{} </Properties>"#,
        property(2, "X", "<vt:lpwstr>&x;</vt:lpwstr>")
    );
    assert!(decode(dtd.as_bytes()).is_err());
    let unknown = document(&property(2, "X", "<vt:lpwstr>&unknown;</vt:lpwstr>"));
    assert!(decode(unknown.as_bytes()).is_err());
    assert!(decode(b"<Properties><property></Properties>").is_err());
}

#[test]
fn byte_depth_node_name_and_text_limits_are_enforced() {
    let oversized = vec![b' '; MAX_XML_BYTES + 1];
    assert!(matches!(decode(&oversized), Err(Error::Limit { .. })));

    let deep = document(&property(
        2,
        "Deep",
        r#"<vt:lpwstr><x xmlns="urn:x"/></vt:lpwstr>"#,
    ));
    assert!(matches!(decode(deep.as_bytes()), Err(Error::Limit { .. })));

    let comments = "<!--x-->".repeat(MAX_XML_NODES + 1);
    let noisy = document(&comments);
    assert!(matches!(decode(noisy.as_bytes()), Err(Error::Limit { .. })));

    let mut props = Props::new();
    assert!(matches!(
        props.insert("n".repeat(MAX_NAME_CHARS + 1), "x"),
        Err(Error::Limit { .. })
    ));
    assert!(matches!(
        props.insert("Text", "x".repeat(MAX_TEXT_BYTES + 1)),
        Err(Error::Limit { .. })
    ));
}

#[test]
fn property_count_is_bounded() {
    let mut body = String::new();
    for index in 0..=MAX_PROPERTIES {
        body.push_str(&property(
            i32::try_from(index).expect("test PID") + 2,
            &format!("P{index}"),
            "<vt:empty/>",
        ));
    }
    assert!(matches!(
        decode(document(&body).as_bytes()),
        Err(Error::Limit { .. })
    ));
}

#[test]
fn absent_package_properties_are_empty_but_orphans_are_errors() {
    let package = OpcPackage::new();
    assert!(Props::read(&package).expect("absence is valid").is_empty());

    let mut orphan = OpcPackage::new();
    orphan.add_part(Box::new(BlobPart::new(
        custom_part_name().expect("URI"),
        ct::OFC_CUSTOM_PROPERTIES.to_owned(),
        document("").into_bytes(),
    )));
    assert!(matches!(Props::read(&orphan), Err(Error::Relationship(_))));
}

#[test]
fn package_write_read_and_clear_remove_the_complete_graph() {
    let mut package = OpcPackage::new();
    let mut props = Props::new();
    props.insert("Project", "Litchi").expect("insert");
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
    props.write(&mut package).expect("write");
    assert!(!package.is_signed());
    let first_blob = package
        .get_part(&custom_part_name().expect("URI"))
        .expect("custom part")
        .blob_arc();
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
    props.write(&mut package).expect("byte-identical no-op");
    assert!(package.is_signed(), "a true no-op preserves signatures");
    let second_blob = package
        .get_part(&custom_part_name().expect("URI"))
        .expect("custom part")
        .blob_arc();
    assert!(Arc::ptr_eq(&first_blob, &second_blob));
    assert_eq!(
        Props::read(&package).expect("read").get("Project"),
        props.get("Project")
    );
    assert!(package.get_part(&custom_part_name().expect("URI")).is_ok());
    assert_eq!(
        package
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
            .count(),
        1
    );

    Props::new().write(&mut package).expect("clear graph");
    assert!(!package.is_signed());
    assert!(package.get_part(&custom_part_name().expect("URI")).is_err());
    assert!(
        package
            .rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::CUSTOM_PROPERTIES)
    );
    assert!(Props::read(&package).expect("read cleared").is_empty());
}

#[test]
fn malformed_package_graphs_are_never_treated_as_absent() {
    let mut missing = OpcPackage::new();
    missing.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
    assert!(matches!(Props::read(&missing), Err(Error::Missing(_))));

    let mut wrong_type = OpcPackage::new();
    wrong_type.add_part(Box::new(BlobPart::new(
        custom_part_name().expect("URI"),
        "application/xml".to_owned(),
        document("").into_bytes(),
    )));
    wrong_type.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
    assert!(matches!(
        Props::read(&wrong_type),
        Err(Error::ContentType { .. })
    ));

    let mut duplicate = OpcPackage::new();
    duplicate.add_part(Box::new(BlobPart::new(
        custom_part_name().expect("URI"),
        ct::OFC_CUSTOM_PROPERTIES.to_owned(),
        document("").into_bytes(),
    )));
    duplicate.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
    duplicate
        .rels_mut()
        .try_add_relationship(
            rt::CUSTOM_PROPERTIES.to_owned(),
            PART_TARGET.to_owned(),
            "rId2".to_owned(),
            TargetMode::Internal,
        )
        .expect("second relationship");
    assert!(matches!(
        Props::read(&duplicate),
        Err(Error::Relationship(_))
    ));

    let mut external = OpcPackage::new();
    external.relate_to_external("https://example.invalid/custom.xml", rt::CUSTOM_PROPERTIES);
    assert!(matches!(
        Props::read(&external),
        Err(Error::Relationship(_))
    ));
}

#[test]
fn illegal_xml_characters_are_rejected_before_writing() {
    let mut props = Props::new();
    assert!(props.insert("Bad\0Name", "value").is_err());
    assert!(props.insert("BadText", "value\u{1}").is_err());
}

#[test]
fn reserved_sensitivity_properties_follow_the_explicit_specification() {
    let mut properties = Props::new();
    properties
        .insert("Sensitivity", "D9F23AE3-A239-45EA-BF23-0123456789AB")
        .expect("GUID label ID");
    properties
        .insert("MSIP_Label_not-a-guid_SetDate", Value::I32(7))
        .expect("opaque SDK metadata is preserved");
    properties
        .validate_for(Host::Excel)
        .expect("Excel sensitivity metadata");

    assert!(properties.insert("Sensitivity", "not-a-guid").is_err());
    assert!(Props::new().insert("Sensitivity", 7_i32).is_err());
    let mut mixed_case = Props::new();
    mixed_case
        .insert("sEnSiTiViTy", "d9f23ae3-a239-45ea-bf23-0123456789ab")
        .expect("case-insensitive reserved name");
    mixed_case
        .validate_for(Host::Excel)
        .expect("case-insensitive reserved validation");
    assert!(Props::new().insert("sensitivity", "not-a-guid").is_err());
    let malformed = document(&property(
        2,
        "Sensitivity",
        "<vt:lpwstr>not-a-guid</vt:lpwstr>",
    ));
    assert!(decode(malformed.as_bytes()).is_err());
}

#[test]
fn word_marking_properties_require_their_documented_text_grammar() {
    let mut properties = Props::new();
    properties
        .insert(HEADER_FONT, "#ffFF00,23,Calibri")
        .expect("font properties");
    properties
        .insert(HEADER_SHAPES, "1,A,f")
        .expect("base shape IDs");
    properties
        .insert("cLaSsIfIcAtIoNcOnTeNtMaRkInGhEaDeRsHaPeIdS-1", "2")
        .expect("first fragment");
    properties
        .insert("ClassificationContentMarkingHeaderShapeIds-2", "B")
        .expect("second fragment");
    properties
        .insert(HEADER_TEXT, "Header")
        .expect("header text");
    properties.validate_for(Host::Word).expect("Word markings");
    let output = String::from_utf8(encode(&properties).expect("encode"))
        .expect("custom-properties XML is UTF-8");
    assert!(output.contains(r#"name="cLaSsIfIcAtIoNcOnTeNtMaRkInGhEaDeRsHaPeIdS-1""#));

    let mut mixed_case = Props::new();
    mixed_case
        .insert(
            "cLaSsIfIcAtIoNcOnTeNtMaRkInGhEaDeRfOnTpRoPs",
            "#ffFF00,23,Calibri",
        )
        .expect("case-insensitive Word property");
    mixed_case
        .validate_for(Host::Word)
        .expect("case-insensitive Word validation");
    assert!(mixed_case.validate_for(Host::PowerPoint).is_err());
    assert!(
        Props::new()
            .insert(
                "classificationcontentmarkingheaderfontprops",
                "#12345G,11,Calibri",
            )
            .is_err()
    );

    assert!(properties.validate_for(Host::PowerPoint).is_err());
    assert!(
        Props::new()
            .insert(HEADER_FONT, "#12345G,11,Calibri")
            .is_err()
    );
    assert!(Props::new().insert(HEADER_SHAPES, "1,0x2").is_err());
    assert!(
        Props::new()
            .insert("ClassificationContentMarkingHeaderShapeIds-0x1", "1")
            .is_err()
    );

    let mut missing_first_fragment = Props::new();
    missing_first_fragment
        .insert(HEADER_SHAPES, "1")
        .expect("base shape IDs");
    missing_first_fragment
        .insert("ClassificationContentMarkingHeaderShapeIds-2", "2")
        .expect("fragment parses before collection validation");
    assert!(missing_first_fragment.validate_for(Host::Word).is_err());
    assert!(encode(&missing_first_fragment).is_err());
}

#[test]
fn powerpoint_locations_require_escaped_names_and_host_boundaries() {
    let mut properties = Props::new();
    properties
        .insert(HEADER_LOCATIONS, r"Office Theme:10\Second\:Theme:11")
        .expect("design master locations");
    properties
        .insert(HEADER_TEXT, "Header")
        .expect("header text");
    properties
        .validate_for(Host::PowerPoint)
        .expect("PowerPoint markings");
    assert!(properties.validate_for(Host::Word).is_err());

    assert!(
        Props::new()
            .insert(HEADER_LOCATIONS, r"Office\q:10")
            .is_err()
    );

    let mut package = OpcPackage::new();
    properties
        .write_for(&mut package, Host::PowerPoint)
        .expect("host-aware write");
    Props::read_for(&package, Host::PowerPoint).expect("host-aware read");
    // LENIENCY_AUDIT 8.4: on read the host scope is advisory.
    let word = Props::read_for(&package, Host::Word).expect("advisory on read");
    assert!(!word.contains(HEADER_LOCATIONS));
    assert!(word.contains(HEADER_TEXT));
    assert_eq!(word.tolerance_report().total(), 1);
}

/// Build a package whose custom-properties part holds `xml`.
fn package_with_custom(xml: &str) -> OpcPackage {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        custom_part_name().expect("URI"),
        ct::OFC_CUSTOM_PROPERTIES.to_owned(),
        xml.as_bytes().to_vec(),
    )));
    package.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
    package
}

/// `(defect, item)` pairs of a read collection's report.
fn recorded(props: &Props) -> Vec<(crate::leniency::CustomPropertiesDefect, Option<String>)> {
    props
        .tolerance_report()
        .iter()
        .map(|tolerated| {
            let crate::leniency::Defect::CustomProperties(defect) = tolerated.defect() else {
                panic!("unexpected defect family");
            };
            assert_eq!(tolerated.part(), "/docProps/custom.xml");
            (defect, tolerated.item().map(str::to_owned))
        })
        .collect()
}

/// LENIENCY_AUDIT 8.3: one bad property costs only that property. Unknown
/// variant types, unparseable values, missing/invalid identity, duplicate
/// PIDs and names, unknown root children, PIs and unknown entities are each
/// skipped and recorded; the good properties survive.
#[test]
fn package_reads_skip_bad_custom_properties_and_record_them() {
    use crate::leniency::CustomPropertiesDefect as D;
    let body = [
        property(2, "Good", "<vt:lpwstr>kept</vt:lpwstr>"),
        property(3, "Due", "<vt:date>2026-01-01</vt:date>"),
        property(4, "Flag", "<vt:bool>maybe</vt:bool>"),
        r#"<property pid="5" name="NoFmtid"><vt:i4>1</vt:i4></property>"#.to_owned(),
        format!(
            r#"<property fmtid="{FORMAT_ID}" pid="1" name="LowPid"><vt:i4>1</vt:i4></property>"#
        ),
        property(2, "SamePid", "<vt:i4>1</vt:i4>"),
        property(6, "good", "<vt:i4>1</vt:i4>"),
        r#"<x:ext xmlns:x="urn:x"><x:y/></x:ext>"#.to_owned(),
        "<?xml-stylesheet href=\"a\"?>".to_owned(),
        property(7, "Entity", "<vt:lpwstr>&nbsp;</vt:lpwstr>"),
        property(8, "Two", "<vt:i4>1</vt:i4><vt:i4>2</vt:i4>"),
        property(9, "Other", "<vt:i8>42</vt:i8>"),
    ]
    .concat();
    let props = Props::read(&package_with_custom(&document(&body))).expect("degrades");
    assert_eq!(props.len(), 2);
    assert_eq!(props.get("Good"), Some(&Value::Text("kept".to_owned())));
    assert_eq!(props.get("Other"), Some(&Value::I64(42)));
    let item = |name: &str| Some(name.to_owned());
    assert_eq!(
        recorded(&props),
        vec![
            (D::UnsupportedValueType, item("Due")),
            (D::InvalidValue, item("Flag")),
            (D::InvalidIdentity, item("NoFmtid")),
            (D::InvalidIdentity, item("LowPid")),
            (D::DuplicateProperty, item("SamePid")),
            (D::DuplicateProperty, item("good")),
            (D::UnknownElement, item("x:ext")),
            (D::ProcessingInstruction, None),
            (D::InvalidValue, item("Entity")),
            (D::MalformedProperty, item("Two")),
        ]
    );

    // The strict codec (used for authored round trips) still refuses.
    assert!(decode(document(&body).as_bytes()).is_err());

    // A conformant part reads clean, and well-formedness stays fatal.
    let clean = Props::read(&package_with_custom(&document(&property(
        2,
        "Good",
        "<vt:lpwstr>kept</vt:lpwstr>",
    ))))
    .expect("clean");
    assert!(clean.tolerance_report().is_clean());
    assert!(Props::read(&package_with_custom(&format!("{PREFIX}<property"))).is_err());
}

/// LENIENCY_AUDIT 8.4: reserved sensitivity-label checks are advisory on
/// read — a property scoped to another host or with a malformed value is
/// skipped and recorded instead of failing the document; writes stay strict.
#[test]
fn reserved_property_violations_are_skipped_on_read_for() {
    use crate::leniency::CustomPropertiesDefect as D;
    let body = [
        property(2, "Keep", "<vt:lpwstr>yes</vt:lpwstr>"),
        property(
            3,
            "ClassificationWatermarkLocations",
            "<vt:lpwstr>Office Theme:10</vt:lpwstr>",
        ),
        property(4, "Sensitivity", "<vt:lpwstr>not-a-guid</vt:lpwstr>"),
    ]
    .concat();
    let package = package_with_custom(&document(&body));
    let props = Props::read_for(&package, Host::Word).expect("advisory on read");
    assert_eq!(props.len(), 1);
    assert!(props.contains("Keep"));
    let reported: Vec<_> = props
        .tolerance_report()
        .iter()
        .map(|t| {
            (
                t.defect(),
                t.item().map(str::to_owned),
                t.value().map(str::to_owned),
            )
        })
        .collect();
    assert_eq!(
        reported,
        vec![
            (
                D::ReservedProperty.into(),
                Some("Sensitivity".to_owned()),
                None
            ),
            (
                D::ReservedProperty.into(),
                Some("ClassificationWatermarkLocations".to_owned()),
                Some("Word".to_owned())
            ),
        ]
    );

    // PowerPoint owns the watermark locations, so only the bad GUID is skipped.
    let props = Props::read_for(&package, Host::PowerPoint).expect("advisory on read");
    assert_eq!(props.len(), 2);

    // Authoring stays strict.
    let mut authored = Props::new();
    authored
        .insert("ClassificationWatermarkLocations", "Office Theme:10")
        .expect("host-independent insert");
    assert!(authored.validate_for(Host::Word).is_err());
}
