mod support;

use litchi_core::{Error, Result};
use litchi_ods::{
    Spreadsheet,
    document::{NumberStyleNode, Snapshot, StyleGraph},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const FOREIGN: &str = "urn:example:future-style";

fn package_with_style(style: &str) -> Vec<u8> {
    let content = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<office:document-content xmlns:office=\"{OFFICE}\" xmlns:table=\"{TABLE}\" xmlns:text=\"{TEXT}\" xmlns:s=\"{STYLE}\" xmlns:n=\"{NUMBER}\" xmlns:v=\"{FOREIGN}\" office:version=\"1.3\">\
<office:automatic-styles>{style}</office:automatic-styles>\
<office:body><office:spreadsheet><table:table table:name=\"Data\"><table:table-row><table:table-cell><text:p>seed</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>\
</office:document-content>"
    );
    support::raw_package(&[("content.xml", content.as_bytes(), "text/xml")])
}

fn decimal_source_style() -> &'static str {
    "<n:number-style s:name=\"Decimal\"><n:text>$</n:text><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/><n:text>USD</n:text></n:number-style>"
}

fn replacement_graph() -> StyleGraph {
    StyleGraph {
        number_styles: vec![NumberStyleNode {
            name: "Decimal".to_string(),
            decimal_places: 3,
            min_integer_digits: 1,
            prefix: None,
            suffix: Some(" changed".to_string()),
        }],
        ..StyleGraph::default()
    }
}

#[test]
fn decimal_replacement_accepts_namespace_aliases() -> Result<()> {
    let snapshot = Snapshot::from_bytes(package_with_style(decimal_source_style()))?;
    let mut edit = snapshot.edit();
    edit.replace_style_graph(&replacement_graph())?;
    let commit = edit.commit()?;
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    let content = reopened.content_xml();
    assert!(content.contains("<number:number-style"));
    assert!(content.contains("number:decimal-places=\"3\""));
    assert!(content.contains("<number:text> changed</number:text>"));
    Ok(())
}

#[test]
fn modeled_character_data_events_remain_supported() -> Result<()> {
    let sources = [
        "<n:number-style s:name=\"Decimal\"><n:text><![CDATA[$]]></n:text><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/><n:text>USD</n:text></n:number-style>",
        "<n:number-style s:name=\"Decimal\"><n:text>&amp;</n:text><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/><n:text>USD</n:text></n:number-style>",
    ];
    for source in sources {
        let snapshot = Snapshot::from_bytes(package_with_style(source))?;
        let mut edit = snapshot.edit();
        edit.replace_style_graph(&replacement_graph())?;
        let commit = edit.commit()?;
        let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        let content = reopened.content_xml();
        assert!(content.contains("<number:number-style"));
    }
    Ok(())
}

#[test]
fn unsupported_source_payloads_refuse_atomically_after_pending_edit() -> Result<()> {
    let cases = [
        (
            "fraction",
            "<n:number-style s:name=\"Decimal\"><n:fraction n:min-numerator-digits=\"1\"/></n:number-style>",
        ),
        (
            "scientific",
            "<n:number-style s:name=\"Decimal\"><n:scientific-number n:decimal-places=\"2\"/></n:number-style>",
        ),
        (
            "embedded-text",
            "<n:number-style s:name=\"Decimal\"><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"><n:embedded-text n:position=\"1\">x</n:embedded-text></n:number></n:number-style>",
        ),
        (
            "foreign-child",
            "<n:number-style s:name=\"Decimal\"><v:future/></n:number-style>",
        ),
        (
            "foreign-attribute",
            "<n:number-style s:name=\"Decimal\" v:future=\"keep\"><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/></n:number-style>",
        ),
        (
            "comment",
            "<n:number-style s:name=\"Decimal\"><!--opaque--><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/></n:number-style>",
        ),
        (
            "processing-instruction",
            "<n:number-style s:name=\"Decimal\"><?vendor keep?><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/></n:number-style>",
        ),
        (
            "number-text",
            "<n:number-style s:name=\"Decimal\"><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\">opaque</n:number></n:number-style>",
        ),
        (
            "empty-particle-whitespace",
            "<n:number-style s:name=\"Decimal\"><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"> </n:number></n:number-style>",
        ),
        (
            "cdata-in-number",
            "<n:number-style s:name=\"Decimal\"><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"><![CDATA[2]]></n:number></n:number-style>",
        ),
        (
            "general-reference-in-number",
            "<n:number-style s:name=\"Decimal\"><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\">&amp;</n:number></n:number-style>",
        ),
        (
            "cdata-in-style-root",
            "<n:number-style s:name=\"Decimal\"><![CDATA[opaque]]><n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/></n:number-style>",
        ),
        (
            "general-reference-in-style-root",
            "<n:number-style s:name=\"Decimal\">&amp;<n:number n:decimal-places=\"2\" n:min-integer-digits=\"1\"/></n:number-style>",
        ),
    ];

    for (label, style) in cases {
        let snapshot = Snapshot::from_bytes(package_with_style(style))?;
        let mut edit = snapshot.edit();
        edit.set_cell_style("Data", 0, 0, "Pending")?;
        let pending = edit.as_bytes().to_vec();
        let error = edit
            .replace_style_graph(&replacement_graph())
            .expect_err(label);
        assert!(
            matches!(&error, Error::Unsupported(message) if message.contains("unsupported source")),
            "{label}: {error}"
        );
        assert_eq!(edit.as_bytes(), pending.as_slice(), "{label}");
    }
    Ok(())
}

#[test]
fn empty_replacement_is_an_exact_noop_for_pending_work() -> Result<()> {
    let snapshot = Snapshot::from_bytes(package_with_style(decimal_source_style()))?;
    let mut edit = snapshot.edit();
    edit.set_cell_style("Data", 0, 0, "Pending")?;
    let pending = edit.as_bytes().to_vec();
    edit.replace_style_graph(&StyleGraph::default())?;
    assert_eq!(edit.as_bytes(), pending.as_slice());
    let commit = edit.commit()?;
    assert!(commit.changed());
    assert!(
        Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
            .content_xml()
            .contains("Pending")
    );
    Ok(())
}
