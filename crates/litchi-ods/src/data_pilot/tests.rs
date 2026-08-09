use super::{Field, Orientation, Table};
use crate::{Builder, MutableSpreadsheet, Spreadsheet};

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:v="urn:example:vendor"><office:body><office:spreadsheet><t:table t:name="Input"/><t:data-pilot-tables><t:data-pilot-table t:name="Pivot" t:target-range-address="Output.A1:B4"><t:source-cell-range t:cell-range-address="Input.A1:B20"/><t:data-pilot-field t:source-field-name="Region" t:orientation="row"/><v:future v:flag="retain"><v:value>opaque</v:value></v:future></t:data-pilot-table></t:data-pilot-tables><t:shapes/></office:spreadsheet></office:body></office:document-content>"#;

const XML_WITHOUT_OWNER: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><t:table t:name="Input"/><t:shapes/></office:spreadsheet></office:body></office:document-content>"#;

fn package(xml: &str) -> Vec<u8> {
    Builder::new()
        .content_xml(xml)
        .build()
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
        Spreadsheet::from_bytes(bytes.clone()).expect("test fixture or operation should succeed");
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
    let mut mutable = MutableSpreadsheet::from_bytes(package(&typed_xml))
        .expect("test fixture or operation should succeed");
    mutable
        .edit_data_pilots(|editor| {
            editor.update("Pivot", |table| {
                table.name = "Renamed".to_string();
                Ok(())
            })
        })
        .expect("test fixture or operation should succeed");
    let content = mutable.spreadsheet().content_xml().to_owned();
    assert!(content.contains("t:shapes"));

    let spreadsheet = Spreadsheet::from_bytes(mutable.to_bytes())
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
    let mut mutable = MutableSpreadsheet::from_bytes(package(XML_WITHOUT_OWNER))
        .expect("test fixture or operation should succeed");
    mutable
        .edit_data_pilots(|editor| editor.add(table("Pivot")))
        .expect("test fixture or operation should succeed");
    let content = mutable.spreadsheet().content_xml();
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
        .spreadsheet()
        .data_pilots()
        .expect("test fixture or operation should succeed");
    assert!(!catalog.has_owner());
}

#[test]
fn opaque_markup_blocks_lossy_edits_and_leaves_source_unchanged() {
    let mut mutable = MutableSpreadsheet::from_bytes(package(XML))
        .expect("test fixture or operation should succeed");
    let before = mutable.spreadsheet().content_xml().to_owned();
    let result = mutable.edit_data_pilots(|editor| {
        editor.update("Pivot", |table| {
            table.name = "Rejected".to_string();
            Ok(())
        })
    });
    assert!(result.is_err());
    assert_eq!(mutable.spreadsheet().content_xml(), before);
}

#[test]
fn failed_typed_update_does_not_change_the_staged_catalog() {
    let spreadsheet = Spreadsheet::from_bytes(package(XML_WITHOUT_OWNER))
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
fn owned_snapshot_commit_patch_inverse_and_conflict_are_atomic() -> litchi_core::Result<()> {
    let snapshot = crate::data_pilot::Snapshot::from_bytes(package(XML_WITHOUT_OWNER))?;
    let mut edit = snapshot.edit();
    edit.editor().add(table("Pivot"))?;
    let commit = edit.commit()?;
    assert!(commit.changed());
    assert!(commit.snapshot().has_owner());
    assert_eq!(commit.snapshot().tables()[0].name, "Pivot");

    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), snapshot.as_bytes());

    let other = crate::data_pilot::Snapshot::from_bytes(package(XML))?;
    assert!(commit.patch().apply(&other).is_err());
    Ok(())
}
