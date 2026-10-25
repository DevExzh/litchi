#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "unit tests panic on fixture failure"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

use super::*;
use crate::Error;

/// The `app.xml` Word 2016 writes, with its vector properties.
const WORD: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>3</TotalTime><Pages>1</Pages><Words>12</Words><Characters>70</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr></vt:lpstr></vt:vector></TitlesOfParts><Company>Old &amp; Co</Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>81</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>"#;

fn word() -> Props {
    Props {
        template: Some("Normal.dotm".into()),
        total_time: Some(3),
        pages: Some(1),
        words: Some(12),
        characters: Some(70),
        application: Some("Microsoft Office Word".into()),
        doc_security: Some(0),
        lines: Some(1),
        paragraphs: Some(1),
        scale_crop: Some(false),
        company: Some("Old & Co".into()),
        links_up_to_date: Some(false),
        characters_with_spaces: Some(81),
        shared_doc: Some(false),
        hyperlinks_changed: Some(false),
        app_version: Some("16.0000".into()),
        ..Props::default()
    }
}

#[test]
fn decodes_every_scalar_and_ignores_vectors() {
    assert_eq!(decode(WORD.as_bytes()).unwrap(), word());
}

#[test]
fn splice_changes_only_the_edited_elements() {
    let target = Props {
        company: Some("New <Co>".into()),
        ..word()
    };
    let output = splice(WORD.as_bytes(), &target).unwrap().unwrap();
    let expected = WORD.replace("Old &amp; Co", "New &lt;Co&gt;");
    assert_eq!(String::from_utf8(output).unwrap(), expected);
}

#[test]
fn splice_appends_removes_and_keeps_vectors() {
    let target = Props {
        manager: Some("Ada".into()),
        template: None,
        pages: Some(7),
        ..word()
    };
    let output = String::from_utf8(splice(WORD.as_bytes(), &target).unwrap().unwrap()).unwrap();
    let expected = WORD
        .replace("<Template>Normal.dotm</Template>", "")
        .replace("<Pages>1</Pages>", "<Pages>7</Pages>")
        .replace("</Properties>", "<Manager>Ada</Manager></Properties>");
    assert_eq!(output, expected);
    assert_eq!(decode(output.as_bytes()).unwrap(), target);
}

#[test]
fn equal_target_is_an_exact_noop() {
    assert!(splice(WORD.as_bytes(), &word()).unwrap().is_none());
}

#[test]
fn new_elements_use_the_root_prefix_and_expand_empty_roots() {
    let prefixed = r#"<ap:Properties xmlns:ap="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><ap:Company/></ap:Properties>"#;
    let target = Props::new().company("C").manager("M");
    let output = String::from_utf8(splice(prefixed.as_bytes(), &target).unwrap().unwrap()).unwrap();
    assert_eq!(
        output,
        r#"<ap:Properties xmlns:ap="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><ap:Company>C</ap:Company><ap:Manager>M</ap:Manager></ap:Properties>"#
    );

    let empty =
        r#"<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties" />"#;
    let output = String::from_utf8(splice(empty.as_bytes(), &target).unwrap().unwrap()).unwrap();
    assert_eq!(
        output,
        r#"<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties" ><Manager>M</Manager><Company>C</Company></Properties>"#
    );
    assert_eq!(decode(output.as_bytes()).unwrap(), target);
}

#[test]
fn text_round_trips_markup_and_carriage_returns() {
    let target = Props::new().company("a & b < c > d\r\ne\t\"'");
    let xml = encode_new(&target).unwrap();
    assert_eq!(decode(&xml).unwrap(), target);
    let edited = splice(WORD.as_bytes(), &Props { ..target.clone() })
        .unwrap()
        .unwrap();
    assert_eq!(decode(&edited).unwrap(), target);
}

#[test]
fn reads_are_strict() {
    let wrap = |body: &str| {
        format!(
            r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">{body}</Properties>"#
        )
    };
    for body in [
        "<Pages>many</Pages>",
        "<Pages>1</Pages><Pages>2</Pages>",
        "<ScaleCrop>yes</ScaleCrop>",
        "<Company><b/></Company>",
        "loose text",
    ] {
        assert!(
            matches!(decode(wrap(body).as_bytes()), Err(Error::Invalid(_))),
            "{body}"
        );
    }
    assert!(matches!(
        decode(br#"<Properties xmlns="urn:other"/>"#),
        Err(Error::Invalid(_))
    ));
    let mce = r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>"#;
    assert!(matches!(decode(mce.as_bytes()), Err(Error::Invalid(_))));
    // Whitespace around numbers and booleans collapses, as xsd defines.
    let lenient = wrap("<Pages> +4 </Pages><ScaleCrop>\n1\n</ScaleCrop><Unknown><x/></Unknown>");
    let props = decode(lenient.as_bytes()).unwrap();
    assert_eq!(props.pages, Some(4));
    assert_eq!(props.scale_crop, Some(true));
}

#[test]
fn values_xml_cannot_carry_are_refused() {
    let props = Props::new().company("bad\u{1}text");
    assert!(matches!(encode_new(&props), Err(Error::Invalid(_))));
    assert!(matches!(
        splice(WORD.as_bytes(), &props),
        Err(Error::Invalid(_))
    ));
}

#[test]
fn source_xml_controls_are_rejected_at_the_decode_boundary() {
    let sources = [
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" data="&#x1;"/>"#,
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:opaque="urn:opaque&#x1;"/>"#,
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Unknown><Nested>&#x1;</Nested></Unknown></Properties>"#,
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>&#x1;</Company></Properties>"#,
    ];
    for source in sources {
        assert!(
            matches!(decode(source.as_bytes()), Err(Error::Invalid(message)) if message.contains("U+0001")),
            "source control was not rejected: {source}"
        );
    }

    let cdata = format!(
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Unknown><![CDATA[{}]]></Unknown></Properties>"#,
        '\u{1}'
    );
    assert!(decode(cdata.as_bytes()).is_err());

    let comment = format!(
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Unknown><!--{}--></Unknown></Properties>"#,
        '\u{1}'
    );
    assert!(
        matches!(decode(comment.as_bytes()), Err(Error::Invalid(message)) if message.contains("U+0001")),
        "source control in a comment was not rejected"
    );

    let pi = format!(
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Unknown><?keep {}?></Unknown></Properties>"#,
        '\u{1}'
    );
    assert!(
        matches!(decode(pi.as_bytes()), Err(Error::Invalid(message)) if message.contains("U+0001")),
        "source control in a processing instruction was not rejected"
    );
}

#[test]
fn splice_preserves_valid_inert_comment_and_processing_instruction_bytes() {
    let source = r#"<?xml version="1.0" encoding="UTF-8"?><!--inert & opaque--><?keep data?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Old</Company><?tail data?></Properties>"#;
    let target = Props::new().company("New");
    let output = splice(source.as_bytes(), &target).unwrap().unwrap();
    let expected = source.replace("<Company>Old</Company>", "<Company>New</Company>");
    assert_eq!(String::from_utf8(output).unwrap(), expected);
}

fn package_with(xml: &[u8]) -> OpcPackage {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/docProps/app.xml").unwrap(),
        ct::OFC_EXTENDED_PROPERTIES.to_owned(),
        xml.to_vec(),
    )));
    package.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
    package
}

fn sign(package: &mut OpcPackage) {
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    )));
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
}

fn publish(
    package: &mut OpcPackage,
    props: Option<Props>,
) -> Result<bool, crate::properties::Error> {
    let mut edit = crate::properties::Edit::new(package, crate::custom::Host::Word)?;
    edit.set_app_props(props);
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    *package = commit.into_package();
    Ok(changed)
}

#[test]
fn package_write_creates_edits_and_clears_the_graph() {
    let mut package = OpcPackage::new();
    assert_eq!(read(&package).unwrap(), None);
    assert!(publish(&mut package, Some(Props::new().company("Litchi"))).unwrap());
    assert_eq!(
        read(&package).unwrap(),
        Some(Props::new().company("Litchi"))
    );
    let part = package
        .get_part(&PackURI::new("/docProps/app.xml").unwrap())
        .unwrap();
    assert_eq!(part.content_type(), ct::OFC_EXTENDED_PROPERTIES);
    assert!(!publish(&mut package, Some(Props::new().company("Litchi"))).unwrap());
    assert!(publish(&mut package, Some(Props::new().manager("Ada"))).unwrap());
    assert_eq!(read(&package).unwrap(), Some(Props::new().manager("Ada")));
    assert!(publish(&mut package, None).unwrap());
    assert!(!publish(&mut package, None).unwrap());
    assert_eq!(read(&package).unwrap(), None);
    assert!(
        package
            .rels()
            .iter()
            .all(|relationship| { relationship.reltype() != rt::EXTENDED_PROPERTIES })
    );
}

#[test]
fn signed_packages_keep_noops_and_refuse_changes() {
    let mut package = package_with(WORD.as_bytes());
    sign(&mut package);
    assert!(!publish(&mut package, Some(word())).unwrap());
    assert!(matches!(
        publish(&mut package, Some(Props::new())),
        Err(crate::properties::Error::Signed)
    ));
    assert!(matches!(
        publish(&mut package, None),
        Err(crate::properties::Error::Signed)
    ));
    assert_eq!(read(&package).unwrap(), Some(word()));
    assert!(package.is_signed());
    package.unsign();
    assert!(publish(&mut package, Some(Props::new())).unwrap());
    assert_eq!(read(&package).unwrap(), Some(Props::new()));
}

#[test]
fn graph_anomalies_are_errors() {
    let mut orphan = OpcPackage::new();
    orphan.add_part(Box::new(BlobPart::new(
        PackURI::new("/docProps/app.xml").unwrap(),
        ct::OFC_EXTENDED_PROPERTIES.to_owned(),
        WORD.as_bytes().to_vec(),
    )));
    assert!(matches!(read(&orphan), Err(Error::Relationship(_))));

    let mut shared = package_with(WORD.as_bytes());
    shared.relate_to("docProps/app.xml", "urn:test:other");
    assert!(matches!(
        publish(&mut shared, None),
        Err(crate::properties::Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
                && refusal.operation() == "remove_metadata"
                && refusal.reason() == "removing a shared OOXML metadata part is refused"
    ));
    assert_eq!(read(&shared).unwrap(), Some(word()));
}
