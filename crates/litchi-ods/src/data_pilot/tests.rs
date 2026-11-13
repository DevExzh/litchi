use super::{Field, Orientation, Table};
use crate::{Builder, Workbook, builder::Workbook as Draft};
use litchi_odf_common::XmlStreamLimits;

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:v="urn:example:vendor"><office:body><office:spreadsheet><t:table t:name="Input"/><t:data-pilot-tables><t:data-pilot-table t:name="Pivot" t:target-range-address="Output.A1:B4"><t:source-cell-range t:cell-range-address="Input.A1:B20"/><t:data-pilot-field t:source-field-name="Region" t:orientation="row"/><v:future v:flag="retain"><v:value>opaque</v:value></v:future></t:data-pilot-table></t:data-pilot-tables><t:shapes/></office:spreadsheet></office:body></office:document-content>"#;

const XML_WITHOUT_OWNER: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><t:table t:name="Input"/><t:shapes/></office:spreadsheet></office:body></office:document-content>"#;

fn package(xml: &str) -> Vec<u8> {
    Builder::new()
        .content_xml(xml)
        .build()
        .expect("test fixture or operation should succeed")
        .to_bytes()
        .expect("test fixture or operation should succeed")
}

fn table(name: &str) -> Table {
    let mut table = Table::new(name, "Output.A1:B4");
    table.fields.push(Field::new("Region", Orientation::Row));
    table
}

#[test]
fn catalog_reads_rich_metadata_and_no_op_is_byte_exact() {
    let bytes = package(XML);
    let spreadsheet =
        Workbook::from_bytes(bytes.clone()).expect("test fixture or operation should succeed");
    let catalog = spreadsheet
        .data_pilots()
        .expect("test fixture or operation should succeed");
    assert!(catalog.has_owner());
    assert_eq!(catalog.len(), 1);
    assert_eq!(
        catalog
            .named("Pivot")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .fields
            .len(),
        1
    );

    let commit = catalog
        .transaction()
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(!commit.changed());
    assert_eq!(commit.bytes(), bytes.as_slice());
}

#[test]
fn clone_staged_crud_rewrites_only_the_owner_and_reopens() {
    let typed_xml = XML
        .replace(
            "<v:future v:flag=\"retain\"><v:value>opaque</v:value></v:future>",
            "",
        )
        .replace(" xmlns:v=\"urn:example:vendor\"", "");
    let mut mutable = Workbook::from_bytes(package(&typed_xml))
        .map(Draft::from_workbook)
        .expect("test fixture or operation should succeed");
    mutable
        .edit_data_pilots(|editor| {
            editor.update("Pivot", |table| {
                table.name = "Renamed".to_string();
                Ok(())
            })
        })
        .expect("test fixture or operation should succeed");
    let content = mutable.workbook().content_xml().to_owned();
    assert!(content.contains("t:shapes"));

    let spreadsheet = Workbook::from_bytes(
        mutable
            .build()
            .expect("the draft builds")
            .as_bytes()
            .to_vec(),
    )
    .expect("test fixture or operation should succeed");
    let catalog = spreadsheet
        .data_pilots()
        .expect("test fixture or operation should succeed");
    assert_eq!(
        catalog
            .named("Renamed")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .name,
        "Renamed"
    );
}

#[test]
fn insertion_honors_spreadsheet_order_and_supports_removal() {
    let mut mutable = Workbook::from_bytes(package(XML_WITHOUT_OWNER))
        .map(Draft::from_workbook)
        .expect("test fixture or operation should succeed");
    mutable
        .edit_data_pilots(|editor| editor.add(table("Pivot")))
        .expect("test fixture or operation should succeed");
    let content = mutable.workbook().content_xml();
    assert!(
        content
            .find("data-pilot-tables")
            .expect("test fixture or operation should succeed")
            < content
                .find("t:shapes")
                .expect("test fixture or operation should succeed")
    );

    mutable
        .edit_data_pilots(|editor| {
            let removed = editor.remove("Pivot")?;
            assert_eq!(removed.name, "Pivot");
            Ok(())
        })
        .expect("test fixture or operation should succeed");
    let catalog = mutable
        .workbook()
        .data_pilots()
        .expect("test fixture or operation should succeed");
    assert!(!catalog.has_owner());
}

#[test]
fn opaque_markup_blocks_lossy_edits_and_leaves_source_unchanged() {
    let mut mutable = Workbook::from_bytes(package(XML))
        .map(Draft::from_workbook)
        .expect("test fixture or operation should succeed");
    let before = mutable.workbook().content_xml().to_owned();
    let result = mutable.edit_data_pilots(|editor| {
        editor.update("Pivot", |table| {
            table.name = "Rejected".to_string();
            Ok(())
        })
    });
    assert!(result.is_err());
    assert_eq!(mutable.workbook().content_xml(), before);
}

#[test]
fn failed_typed_update_does_not_change_the_staged_catalog() {
    let spreadsheet = Workbook::from_bytes(package(XML_WITHOUT_OWNER))
        .expect("test fixture or operation should succeed");
    let catalog = spreadsheet
        .data_pilots()
        .expect("test fixture or operation should succeed");
    let mut transaction = catalog.transaction();
    let result = transaction.editor().update(0, |table| {
        table.fields.push(Field::new("Region", Orientation::Page));
        Ok(())
    });
    assert!(result.is_err());
    assert!(transaction.tables().is_empty());
}

#[test]
fn owned_snapshot_commit_adds_the_owner() -> litchi_core::Result<()> {
    let snapshot = crate::data_pilot::Snapshot::from_bytes(
        package(XML_WITHOUT_OWNER),
        &XmlStreamLimits::owner_default_value(),
    )?;
    let mut edit = snapshot.edit();
    edit.editor().add(table("Pivot"))?;
    let commit = edit.commit()?;
    assert_ne!(commit.snapshot().as_bytes(), snapshot.as_bytes());
    assert!(commit.snapshot().has_owner());
    assert_eq!(commit.snapshot().tables()[0].name, "Pivot");
    Ok(())
}

#[test]
fn raw_data_pilot_limits_cover_nested_location_and_exact_output() {
    let nested = XmlStreamLimits::new(XML.len() as u64 + 64, 2, 100_000, 100_000, 64 * 1024)
        .expect("finite nested policy");
    assert!(super::codec::locate(XML, &nested).is_err());

    let default_limits = XmlStreamLimits::owner_default_value();
    let fragment =
        crate::model::data_pilot::write_data_pilot_table_fragment(&table("Pivot"), &default_limits)
            .expect("default output");
    let exact = XmlStreamLimits::new(fragment.len() as u64, 512, 100_000, 100_000, 64 * 1024)
        .expect("finite exact output policy");
    assert_eq!(
        crate::model::data_pilot::write_data_pilot_table_fragment(&table("Pivot"), &exact)
            .expect("exact output"),
        fragment
    );
    let tiny = XmlStreamLimits::new(1, 512, 100_000, 100_000, 64 * 1024)
        .expect("finite tiny output policy");
    assert!(
        crate::model::data_pilot::write_data_pilot_table_fragment(&table("Pivot"), &tiny).is_err()
    );
}

/// Runs `run` over the XML envelope orderings built around one root element
/// and compares each `label => outcome` line with the frozen expectation.
fn assert_envelope_matrix(
    open: &str,
    close: &str,
    empty: &str,
    expected: &[&str],
    run: impl Fn(&str) -> String,
) {
    const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
    let doc = format!("{open}{close}");
    let rows = [
        ("empty input", String::new()),
        ("whitespace only", " \n".to_owned()),
        ("declaration only", DECL.to_owned()),
        ("empty-element root", empty.to_owned()),
        ("start and end root", doc.clone()),
        ("declaration then root", format!("{DECL}{doc}")),
        ("declaration, whitespace, root", format!("{DECL}\n{doc}")),
        ("root then whitespace", format!("{doc}\n ")),
        ("root then comment", format!("{doc}<!--c-->")),
        ("root then PI", format!("{doc}<?pi x?>")),
        ("root then text", format!("{doc}text")),
        ("root then second root", format!("{doc}{doc}")),
        (
            "root then foreign start",
            format!("{doc}<foreign></foreign>"),
        ),
        ("root then foreign empty", format!("{doc}<foreign/>")),
        (
            "empty-element root then second root",
            format!("{empty}{doc}"),
        ),
        ("empty-element root then text", format!("{empty}text")),
        ("declaration after root", format!("{doc}{DECL}")),
        ("two declarations", format!("{DECL}{DECL}{doc}")),
        ("comment before declaration", format!("<!--c-->{DECL}{doc}")),
        ("comment before root", format!("<!--c-->{doc}")),
        ("PI before root", format!("<?pi x?>{doc}")),
        ("text before root", format!("text{doc}")),
        ("whitespace before declaration", format!(" {DECL}{doc}")),
        ("unterminated root", open.to_owned()),
        ("stray end tag", format!("{doc}</stray>")),
        ("DOCTYPE", format!("{DECL}<!DOCTYPE r>{doc}")),
        ("DOCTYPE after root", format!("{doc}<!DOCTYPE r>")),
        ("top-level CDATA", format!("<![CDATA[x]]>{doc}")),
        ("CDATA after root", format!("{doc}<![CDATA[x]]>")),
        ("entity reference before root", format!("&amp;{doc}")),
        ("entity reference after root", format!("{doc}&amp;")),
    ];
    let actual: Vec<String> = rows
        .iter()
        .map(|(label, xml)| format!("{label} => {}", run(xml)))
        .collect();
    let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
    assert!(
        actual
            .iter()
            .map(String::as_str)
            .eq(expected.iter().copied()),
        "actual envelope matrix:\n{}",
        listing.join("\n")
    );
}

const MATRIX_ROOT: &str = concat!(
    "office:document-content ",
    "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
);
const MATRIX_BODY: &str = "<office:body><office:spreadsheet></office:spreadsheet></office:body>";
const MATRIX_CLOSE: &str = "</office:document-content>";

#[test]
fn data_pilot_locate_envelope_matrix() {
    assert_envelope_matrix(
        &format!("<{MATRIX_ROOT}>{MATRIX_BODY}"),
        MATRIX_CLOSE,
        &format!("<{MATRIX_ROOT}/>"),
        &[
            "empty input => InvalidFormat(\"incomplete ODS content.xml document\")",
            "whitespace only => InvalidFormat(\"incomplete ODS content.xml document\")",
            "declaration only => InvalidFormat(\"incomplete ODS content.xml document\")",
            "empty-element root => InvalidFormat(\"ODS content.xml has no office:spreadsheet host\")",
            "start and end root => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "declaration then root => Ok(spreadsheet=202..243, container=false, opaque=false)",
            "declaration, whitespace, root => Ok(spreadsheet=203..244, container=false, opaque=false)",
            "root then whitespace => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "root then comment => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "root then PI => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "root then text => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "root then second root => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "root then foreign start => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "root then foreign empty => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "empty-element root then second root => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "empty-element root then text => InvalidFormat(\"ODS content.xml has no office:spreadsheet host\")",
            "declaration after root => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "two declarations => Ok(spreadsheet=240..281, container=false, opaque=false)",
            "comment before declaration => Ok(spreadsheet=210..251, container=false, opaque=false)",
            "comment before root => Ok(spreadsheet=172..213, container=false, opaque=false)",
            "PI before root => Ok(spreadsheet=172..213, container=false, opaque=false)",
            "text before root => Ok(spreadsheet=168..209, container=false, opaque=false)",
            "whitespace before declaration => Ok(spreadsheet=203..244, container=false, opaque=false)",
            "unterminated root => InvalidFormat(\"incomplete ODS content.xml document\")",
            "stray end tag => InvalidFormat(\"invalid ODS DataPilot XML: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => Ok(spreadsheet=214..255, container=false, opaque=false)",
            "DOCTYPE after root => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "top-level CDATA => Ok(spreadsheet=177..218, container=false, opaque=false)",
            "CDATA after root => Ok(spreadsheet=164..205, container=false, opaque=false)",
            "entity reference before root => Ok(spreadsheet=169..210, container=false, opaque=false)",
            "entity reference after root => Ok(spreadsheet=164..205, container=false, opaque=false)",
        ],
        |xml| match super::codec::locate(xml, &XmlStreamLimits::owner_default_value()) {
            Ok(value) => format!(
                "Ok(spreadsheet={}..{}, container={}, opaque={})",
                value.spreadsheet.start,
                value.spreadsheet.end,
                value.container.is_some(),
                value.opaque
            ),
            Err(error) => format!("{error:?}"),
        },
    );
}
