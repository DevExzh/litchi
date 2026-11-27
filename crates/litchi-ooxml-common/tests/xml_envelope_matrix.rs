//! What the strict XML readers accept before, around and after their root
//! element, and the exact error each one reports otherwise.
//!
//! Every reader gets the same rows, built around its own root element. Each
//! expected line is the row label followed by the reader's result.

use litchi_ooxml_common::{custom_data, custom_xml, mce::alternative};

const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

/// One reader's root element in the forms the rows are built from.
struct Root<'a> {
    /// The start tag.
    open: &'a str,
    /// Content that makes the element valid for its reader.
    body: &'a str,
    /// The end tag.
    close: &'a str,
    /// The element as one empty-element tag.
    empty: &'a str,
}

fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
    let Root {
        open,
        body,
        close,
        empty,
    } = *root;
    let document = format!("{open}{body}{close}");
    let d = document.as_str();
    vec![
        ("empty input", String::new()),
        ("whitespace only", " \n".to_owned()),
        ("declaration only", DECLARATION.to_owned()),
        ("empty-element root", empty.to_owned()),
        ("start and end root", format!("{open}{close}")),
        ("root then whitespace", format!("{d}\n ")),
        ("root then comment", format!("{d}<!--c-->")),
        ("root then PI", format!("{d}<?pi x?>")),
        ("root then text", format!("{d}text")),
        ("root then second root", format!("{d}{d}")),
        ("root then foreign start", format!("{d}<foreign></foreign>")),
        ("root then foreign empty", format!("{d}<foreign/>")),
        ("declaration after root", format!("{d}{DECLARATION}")),
        ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
        (
            "comment before declaration",
            format!("<!--c-->{DECLARATION}{d}"),
        ),
        ("text before root", format!("text{d}")),
        ("unterminated root", format!("{open}{body}")),
        ("stray end tag", format!("{d}</foreign>")),
        ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
        ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
        ("declaration then root", format!("{DECLARATION}{d}")),
        (
            "whitespace before declaration",
            format!(" {DECLARATION}{d}"),
        ),
        ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
        ("comment then root", format!("<!--c-->{d}")),
        ("PI then root", format!("<?pi x?>{d}")),
        ("root then text then second root", format!("{d}text{d}")),
        ("empty-element root then root", format!("{empty}{d}")),
        (
            "declaration after empty-element root",
            format!("{empty}{DECLARATION}"),
        ),
    ]
}

fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
    let actual = rows(root)
        .into_iter()
        .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
        .collect::<Vec<_>>();
    assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
}

fn outcome<T, E: std::fmt::Debug>(result: Result<T, E>, summary: impl Fn(T) -> String) -> String {
    match result {
        Ok(value) => format!("Ok({})", summary(value)),
        Err(error) => format!("{error:?}"),
    }
}

#[test]
fn alternate_content_fragment_envelope() {
    let root = Root {
        open: r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:a">"#,
        body: r#"<mc:Choice Requires="a"><a:x/></mc:Choice>"#,
        close: "</mc:AlternateContent>",
        empty: r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>"#,
    };
    assert_matrix(
        &root,
        &[
            "empty input: NonConformant(\"unterminated AlternateContent fragment\")",
            "whitespace only: NonConformant(\"unterminated AlternateContent fragment\")",
            "declaration only: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "empty-element root: NonConformant(\"AlternateContent cannot be empty\")",
            "start and end root: NonConformant(\"AlternateContent requires a Choice\")",
            "root then whitespace: Ok(1 branches)",
            "root then comment: Ok(1 branches)",
            "root then PI: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "root then text: NonConformant(\"AlternateContent may contain only whitespace outside branches\")",
            "root then second root: NonConformant(\"AlternateContent must have one root element\")",
            "root then foreign start: NonConformant(\"AlternateContent must have one root element\")",
            "root then foreign empty: NonConformant(\"AlternateContent must have one root element\")",
            "declaration after root: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "two declarations: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "comment before declaration: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "text before root: NonConformant(\"AlternateContent may contain only whitespace outside branches\")",
            "unterminated root: NonConformant(\"unterminated AlternateContent fragment\")",
            "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
            "DOCTYPE: NonConformant(\"DTD declarations are forbidden in AlternateContent\")",
            "top-level CDATA: NonConformant(\"CDATA is not allowed directly in AlternateContent\")",
            "declaration then root: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "whitespace before declaration: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "PI before declaration: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "comment then root: Ok(1 branches)",
            "PI then root: NonConformant(\"processing instructions and declarations are not valid in an AlternateContent fragment\")",
            "root then text then second root: NonConformant(\"AlternateContent may contain only whitespace outside branches\")",
            "empty-element root then root: NonConformant(\"AlternateContent cannot be empty\")",
            "declaration after empty-element root: NonConformant(\"AlternateContent cannot be empty\")",
        ],
        |xml| {
            outcome(
                alternative::read(xml, &alternative::Limits::default()),
                |alternatives| format!("{} branches", alternatives.len()),
            )
        },
    );
}

#[test]
fn custom_xml_payload_envelope() {
    let root = Root {
        open: r#"<r xmlns="urn:r">"#,
        body: "<c/>",
        close: "</r>",
        empty: r#"<r xmlns="urn:r"/>"#,
    };
    assert_matrix(
        &root,
        &[
            "empty input: Invalid(\"custom XML payload has no complete root element\")",
            "whitespace only: Invalid(\"custom XML payload has no complete root element\")",
            "declaration only: Invalid(\"custom XML payload has no complete root element\")",
            "empty-element root: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "start and end root: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "root then whitespace: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "root then comment: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "root then PI: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "root then text: Invalid(\"custom XML payload has text outside its root\")",
            "root then second root: Invalid(\"custom XML payload has multiple roots\")",
            "root then foreign start: Invalid(\"custom XML payload has multiple roots\")",
            "root then foreign empty: Invalid(\"custom XML payload has multiple roots\")",
            "declaration after root: Invalid(\"custom XML payload has a late XML declaration\")",
            "two declarations: Invalid(\"custom XML payload has a late XML declaration\")",
            "comment before declaration: Invalid(\"custom XML payload has a late XML declaration\")",
            "text before root: Invalid(\"custom XML payload has text outside its root\")",
            "unterminated root: Invalid(\"custom XML payload has no complete root element\")",
            "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
            "DOCTYPE: Invalid(\"DTD is forbidden in custom XML payloads\")",
            "top-level CDATA: Invalid(\"custom XML payload has CDATA outside its root\")",
            "declaration then root: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "whitespace before declaration: Invalid(\"custom XML payload has a late XML declaration\")",
            "PI before declaration: Invalid(\"custom XML payload has a late XML declaration\")",
            "comment then root: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "PI then root: Ok(Name { namespace: \"urn:r\", local_name: \"r\" })",
            "root then text then second root: Invalid(\"custom XML payload has text outside its root\")",
            "empty-element root then root: Invalid(\"custom XML payload has multiple roots\")",
            "declaration after empty-element root: Invalid(\"custom XML payload has a late XML declaration\")",
        ],
        |xml| {
            outcome(
                custom_xml::validate_payload(xml, &litchi_opc::SourceLimits::default()),
                |name| format!("{name:?}"),
            )
        },
    );
}

#[test]
fn custom_data_properties_envelope() {
    let root = Root {
        open: r#"<datastoreItem xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" id="s">"#,
        body: "<extLst/>",
        close: "</datastoreItem>",
        empty: r#"<datastoreItem xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" id="s"/>"#,
    };
    assert_matrix(
        &root,
        &[
            "empty input: Invalid(\"missing or unterminated Custom Data Properties XML root\")",
            "whitespace only: Invalid(\"missing or unterminated Custom Data Properties XML root\")",
            "declaration only: Invalid(\"missing or unterminated Custom Data Properties XML root\")",
            "empty-element root: Ok(id \"s\", extension list false)",
            "start and end root: Ok(id \"s\", extension list false)",
            "root then whitespace: Ok(id \"s\", extension list true)",
            "root then comment: Ok(id \"s\", extension list true)",
            "root then PI: Ok(id \"s\", extension list true)",
            "root then text: Invalid(\"text outside XML root\")",
            "root then second root: Invalid(\"multiple XML roots\")",
            "root then foreign start: Invalid(\"expected {http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}datastoreItem, got {}foreign\")",
            "root then foreign empty: Invalid(\"expected {http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}datastoreItem, got {}foreign\")",
            "declaration after root: Invalid(\"XML declaration is not at the document prolog\")",
            "two declarations: Invalid(\"XML declaration is not at the document prolog\")",
            "comment before declaration: Invalid(\"XML declaration is not at the document prolog\")",
            "text before root: Invalid(\"text outside XML root\")",
            "unterminated root: Invalid(\"unterminated Custom Data Properties XML\")",
            "stray end tag: Decode(Malformed(\"ill-formed document: close tag `</foreign>` does not match any open tag\"))",
            "DOCTYPE: Invalid(\"DTDs are rejected\")",
            "top-level CDATA: Invalid(\"CDATA outside XML root\")",
            "declaration then root: Ok(id \"s\", extension list true)",
            "whitespace before declaration: Invalid(\"XML declaration is not at the document prolog\")",
            "PI before declaration: Invalid(\"XML declaration is not at the document prolog\")",
            "comment then root: Ok(id \"s\", extension list true)",
            "PI then root: Ok(id \"s\", extension list true)",
            "root then text then second root: Invalid(\"text outside XML root\")",
            "empty-element root then root: Invalid(\"multiple XML roots\")",
            "declaration after empty-element root: Invalid(\"XML declaration is not at the document prolog\")",
        ],
        |xml| {
            outcome(
                custom_data::parse_properties(xml, &custom_data::Limits::standard()),
                |properties| {
                    format!(
                        "id {:?}, extension list {}",
                        properties.id,
                        properties.extension_list.is_some()
                    )
                },
            )
        },
    );
}

#[test]
fn custom_data_workbook_root_envelope() {
    let root = Root {
        open: r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        body: "<sheets/>",
        close: "</workbook>",
        empty: r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
    };
    assert_matrix(
        &root,
        &[
            "empty input: Invalid(\"missing XML root\")",
            "whitespace only: Invalid(\"missing XML root\")",
            "declaration only: Invalid(\"missing XML root\")",
            "empty-element root: Ok()",
            "start and end root: Ok()",
            "root then whitespace: Ok()",
            "root then comment: Ok()",
            "root then PI: Ok()",
            "root then text: Invalid(\"text outside XML root\")",
            "root then second root: Invalid(\"multiple XML roots\")",
            "root then foreign start: Invalid(\"multiple XML roots\")",
            "root then foreign empty: Invalid(\"multiple XML roots\")",
            "declaration after root: Invalid(\"XML declaration is not at the document prolog\")",
            "two declarations: Invalid(\"XML declaration is not at the document prolog\")",
            "comment before declaration: Invalid(\"XML declaration is not at the document prolog\")",
            "text before root: Invalid(\"text outside XML root\")",
            "unterminated root: Invalid(\"unterminated XML root\")",
            "stray end tag: Decode(Malformed(\"ill-formed document: close tag `</foreign>` does not match any open tag\"))",
            "DOCTYPE: Invalid(\"DTDs are rejected\")",
            "top-level CDATA: Invalid(\"CDATA outside XML root\")",
            "declaration then root: Ok()",
            "whitespace before declaration: Invalid(\"XML declaration is not at the document prolog\")",
            "PI before declaration: Invalid(\"XML declaration is not at the document prolog\")",
            "comment then root: Ok()",
            "PI then root: Ok()",
            "root then text then second root: Invalid(\"text outside XML root\")",
            "empty-element root then root: Invalid(\"multiple XML roots\")",
            "declaration after empty-element root: Invalid(\"XML declaration is not at the document prolog\")",
        ],
        |xml| {
            outcome(
                custom_data::validate_workbook_root(xml, &custom_data::Limits::standard()),
                |()| String::new(),
            )
        },
    );
}
