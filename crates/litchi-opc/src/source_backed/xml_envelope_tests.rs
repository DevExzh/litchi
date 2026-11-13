//! What the source-backed metadata readers accept before, around and after
//! their root element, and the exact error each one reports otherwise.
//!
//! Every reader gets the same rows, built around its own root element. Each
//! expected line is the row label followed by the reader's result. The
//! readers are private, so the rows are run from inside the module.

#![allow(
    clippy::expect_used,
    reason = "a fixture part name that fails to parse is a broken test"
)]

use super::content_types_plan::ContentTypesPlan;
use super::*;

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

const RELATIONSHIPS_ROOT: Root<'static> = Root {
    open: r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    body: "",
    close: "</Relationships>",
    empty: r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
};

const TYPES_ROOT: Root<'static> = Root {
    open: r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    body: r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    close: "</Types>",
    empty: r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
};

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

fn outcome<T>(result: Result<T>, summary: impl Fn(T) -> String) -> String {
    match result {
        Ok(value) => format!("Ok({})", summary(value)),
        Err(error) => format!("{error:?}"),
    }
}

#[test]
fn relationship_source_envelope() {
    let expected = Relationships::new("/".to_owned());
    assert_matrix(
        &RELATIONSHIPS_ROOT,
        &[
            "empty input: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: missing or unclosed Relationships root\")",
            "whitespace only: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: missing or unclosed Relationships root\")",
            "declaration only: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: missing or unclosed Relationships root\")",
            "empty-element root: Ok(2 events, root closes at 83, empty root ends at Some(85))",
            "start and end root: Ok(3 events, root closes at 84, empty root ends at None)",
            "root then whitespace: Ok(3 events, root closes at 84, empty root ends at None)",
            "root then comment: Ok(4 events, root closes at 84, empty root ends at None)",
            "root then PI: Ok(4 events, root closes at 84, empty root ends at None)",
            "root then text: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: non-whitespace text is not supported in a relationships member\")",
            "root then second root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: content appears after the Relationships root\")",
            "root then foreign start: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: content appears after the Relationships root\")",
            "root then foreign empty: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: content appears after the Relationships root\")",
            "declaration after root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: the XML declaration is out of place\")",
            "two declarations: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: the XML declaration is out of place\")",
            "comment before declaration: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: the XML declaration is out of place\")",
            "text before root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: non-whitespace text is not supported in a relationships member\")",
            "unterminated root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: missing or unclosed Relationships root\")",
            "stray end tag: QuickXmlError(IllFormed(UnmatchedEndTag(\"foreign\")))",
            "DOCTYPE: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: DTDs are not permitted in relationships members\")",
            "top-level CDATA: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: unsupported XML content is not safe to append\")",
            "declaration then root: Ok(4 events, root closes at 122, empty root ends at None)",
            "whitespace before declaration: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: the XML declaration is out of place\")",
            "PI before declaration: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: the XML declaration is out of place\")",
            "comment then root: Ok(4 events, root closes at 92, empty root ends at None)",
            "PI then root: Ok(4 events, root closes at 92, empty root ends at None)",
            "root then text then second root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: non-whitespace text is not supported in a relationships member\")",
            "empty-element root then root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: content appears after the Relationships root\")",
            "declaration after empty-element root: InvalidRelationshipsManifest(\"relationships member '_rels/.rels' cannot be appended safely: the XML declaration is out of place\")",
        ],
        |xml| {
            outcome(
                parse_noncanonical_relationship_source(
                    xml,
                    "_rels/.rels",
                    &expected,
                    SourceLimits::default(),
                    || Ok(()),
                ),
                |parsed| {
                    format!(
                        "{} events, root closes at {}, empty root ends at {:?}",
                        parsed.event_count, parsed.root_close_start, parsed.root_empty_end
                    )
                },
            )
        },
    );
}

#[test]
fn raw_content_types_envelope() {
    assert_matrix(
        &TYPES_ROOT,
        &[
            "empty input: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "whitespace only: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "declaration only: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "empty-element root: Ok(0 mappings)",
            "start and end root: Ok(0 mappings)",
            "root then whitespace: Ok(1 mappings)",
            "root then comment: Ok(1 mappings)",
            "root then PI: Ok(1 mappings)",
            "root then text: InvalidContentTypesManifest(\"non-whitespace text is not permitted in content-types metadata\")",
            "root then second root: InvalidContentTypesManifest(\"multiple root elements\")",
            "root then foreign start: InvalidContentTypesManifest(\"multiple root elements\")",
            "root then foreign empty: InvalidContentTypesManifest(\"multiple root elements\")",
            "declaration after root: InvalidContentTypesManifest(\"the XML declaration is out of place in content-types metadata\")",
            "two declarations: InvalidContentTypesManifest(\"the XML declaration is out of place in content-types metadata\")",
            "comment before declaration: InvalidContentTypesManifest(\"the XML declaration is out of place in content-types metadata\")",
            "text before root: InvalidContentTypesManifest(\"non-whitespace text is not permitted in content-types metadata\")",
            "unterminated root: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "stray end tag: QuickXmlError(IllFormed(UnmatchedEndTag(\"foreign\")))",
            "DOCTYPE: InvalidContentTypesManifest(\"DTDs are not permitted in the content types manifest\")",
            "top-level CDATA: InvalidContentTypesManifest(\"CDATA, DTDs, and general references are not permitted in raw content-types metadata\")",
            "declaration then root: Ok(1 mappings)",
            "whitespace before declaration: InvalidContentTypesManifest(\"the XML declaration is out of place in content-types metadata\")",
            "PI before declaration: InvalidContentTypesManifest(\"the XML declaration is out of place in content-types metadata\")",
            "comment then root: Ok(1 mappings)",
            "PI then root: Ok(1 mappings)",
            "root then text then second root: InvalidContentTypesManifest(\"multiple root elements\")",
            "empty-element root then root: InvalidContentTypesManifest(\"multiple root elements\")",
            "declaration after empty-element root: InvalidContentTypesManifest(\"the XML declaration is out of place in content-types metadata\")",
        ],
        |xml| {
            outcome(
                parse_strict_raw_content_types_xml(xml, SourceLimits::default()),
                |map| format!("{} mappings", map.mapping_count()),
            )
        },
    );
}

/// The reader every content-types manifest passes first: the raw reader
/// above only sees what this one accepts.
#[test]
fn content_type_map_envelope() {
    assert_matrix(
        &TYPES_ROOT,
        &[
            "empty input: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "whitespace only: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "declaration only: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "empty-element root: Ok(0 mappings)",
            "start and end root: Ok(0 mappings)",
            "root then whitespace: Ok(1 mappings)",
            "root then comment: Ok(1 mappings)",
            "root then PI: Ok(1 mappings)",
            "root then text: Ok(1 mappings)",
            "root then second root: InvalidContentTypesManifest(\"multiple root elements\")",
            "root then foreign start: InvalidContentTypesManifest(\"multiple root elements\")",
            "root then foreign empty: InvalidContentTypesManifest(\"multiple root elements\")",
            "declaration after root: Ok(1 mappings)",
            "two declarations: Ok(1 mappings)",
            "comment before declaration: Ok(1 mappings)",
            "text before root: Ok(1 mappings)",
            "unterminated root: InvalidContentTypesManifest(\"missing or unclosed Types root element\")",
            "stray end tag: QuickXmlError(IllFormed(UnmatchedEndTag(\"foreign\")))",
            "DOCTYPE: InvalidContentTypesManifest(\"DTDs are not permitted in the content types manifest\")",
            "top-level CDATA: Ok(1 mappings)",
            "declaration then root: Ok(1 mappings)",
            "whitespace before declaration: Ok(1 mappings)",
            "PI before declaration: Ok(1 mappings)",
            "comment then root: Ok(1 mappings)",
            "PI then root: Ok(1 mappings)",
            "root then text then second root: InvalidContentTypesManifest(\"multiple root elements\")",
            "empty-element root then root: InvalidContentTypesManifest(\"multiple root elements\")",
            "declaration after empty-element root: Ok(0 mappings)",
        ],
        |xml| {
            outcome(
                ContentTypeMap::from_xml(xml, SourceLimits::default()),
                |map| format!("{} mappings", map.mapping_count()),
            )
        },
    );
}

/// The plan adds one mapping, so the result shows where the scan placed the
/// end of the root. The output is shown with its long, fixed pieces
/// abbreviated: the root start tag as `<Types>`, the source mapping as
/// `<Default/>`, the added mapping as `<Override/>` and the declaration as
/// `<?xml?>`.
#[test]
fn content_types_plan_envelope() {
    let additions = [(
        PackURI::new("/word/new.bin").expect("fixture part name"),
        "application/octet-stream".to_owned(),
    )];
    assert_matrix(
        &TYPES_ROOT,
        &[
            "empty input: InvalidContentTypesManifest(\"missing or unclosed Types root\")",
            "whitespace only: InvalidContentTypesManifest(\"missing or unclosed Types root\")",
            "declaration only: InvalidContentTypesManifest(\"missing or unclosed Types root\")",
            "empty-element root: Ok(<Types><Override/></Types>)",
            "start and end root: Ok(<Types><Override/></Types>)",
            "root then whitespace: Ok(<Types><Default/><Override/></Types>\n )",
            "root then comment: Ok(<Types><Default/><Override/></Types><!--c-->)",
            "root then PI: Ok(<Types><Default/><Override/></Types><?pi x?>)",
            "root then text: InvalidContentTypesManifest(\"non-whitespace text is not permitted in the content types manifest\")",
            "root then second root: InvalidContentTypesManifest(\"content-types root must be Types in the OPC content-types namespace\")",
            "root then foreign start: InvalidContentTypesManifest(\"content-types root must be Types in the OPC content-types namespace\")",
            "root then foreign empty: InvalidContentTypesManifest(\"content-types root must be Types in the OPC content-types namespace\")",
            "declaration after root: Ok(<Types><Default/><Override/></Types><?xml?>)",
            "two declarations: Ok(<?xml?><?xml?><Types><Default/><Override/></Types>)",
            "comment before declaration: Ok(<!--c--><?xml?><Types><Default/><Override/></Types>)",
            "text before root: InvalidContentTypesManifest(\"non-whitespace text is not permitted in the content types manifest\")",
            "unterminated root: InvalidContentTypesManifest(\"missing or unclosed Types root\")",
            "stray end tag: QuickXmlError(IllFormed(UnmatchedEndTag(\"foreign\")))",
            "DOCTYPE: InvalidContentTypesManifest(\"CDATA and DTDs are not permitted in the content types manifest\")",
            "top-level CDATA: InvalidContentTypesManifest(\"CDATA and DTDs are not permitted in the content types manifest\")",
            "declaration then root: Ok(<?xml?><Types><Default/><Override/></Types>)",
            "whitespace before declaration: Ok( <?xml?><Types><Default/><Override/></Types>)",
            "PI before declaration: Ok(<?pi x?><?xml?><Types><Default/><Override/></Types>)",
            "comment then root: Ok(<!--c--><Types><Default/><Override/></Types>)",
            "PI then root: Ok(<?pi x?><Types><Default/><Override/></Types>)",
            "root then text then second root: InvalidContentTypesManifest(\"non-whitespace text is not permitted in the content types manifest\")",
            "empty-element root then root: InvalidContentTypesManifest(\"content-types root must be Types in the OPC content-types namespace\")",
            "declaration after empty-element root: Ok(<Types><Override/></Types><?xml?>)",
        ],
        |xml| {
            outcome(
                ContentTypesPlan::plan(xml, &additions, &[], SourceLimits::default(), None, None)
                    .and_then(|plan| plan.materialize(SourceLimits::default(), None, None)),
                |(output, _)| {
                    String::from_utf8_lossy(&output)
                    .replace(TYPES_ROOT.open, "<Types>")
                    .replace(TYPES_ROOT.body, "<Default/>")
                    .replace(
                        r#"<Override PartName="/word/new.bin" ContentType="application/octet-stream"/>"#,
                        "<Override/>",
                    )
                    .replace(DECLARATION, "<?xml?>")
                },
            )
        },
    );
}
