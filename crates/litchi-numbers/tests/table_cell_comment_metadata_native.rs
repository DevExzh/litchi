//! Native Numbers root-comment metadata survives save, close, and reopen.

use std::{fs, io::Read, path::PathBuf};

use litchi_numbers::{
    CellPosition, Package, SheetSelector, TableCellComment, TableSelector, cell::Value,
    table::cells::Storage,
};

const SOURCE: &[u8] = include_bytes!("fixtures/comment-edit-root.numbers");
const NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/comment-metadata-native-resaved.numbers");
const NATIVE_CREATED: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/comment-reader-native-created.numbers");
const OPTIONAL_NATIVE_PATH_ENV: &str = "LITCHI_NATIVE_COMMENT_METADATA_PATH";
const MAX_OPTIONAL_NATIVE_BYTES: u64 = 8 * 1024 * 1024;

fn assert_native_root_metadata(
    package: &Package,
) -> Result<TableCellComment, Box<dyn std::error::Error>> {
    let root = package
        .table_cell_comment_a1("Sheet 1", "Review", "$B$2")?
        .ok_or("native B2 comment missing")?;
    assert_eq!(root.text(), "Focused source comment");
    assert!(root.timestamp().is_some());
    assert_eq!(
        root.author().and_then(|author| author.display_name()),
        Some("litchi-iwa")
    );
    Ok(root)
}

fn assert_native_created_sentinel(package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    let state = package.table_cell_a1("Sheet 1", "Table 1", "B2")?;
    match state.storage() {
        Storage::Stored(Value::Text(value)) => assert_eq!(value, "Native comment parity"),
        storage => panic!("native-created B2 sentinel mismatch: {storage:?}"),
    }
    Ok(())
}

#[test]
fn native_saved_root_preserves_semantic_metadata() -> Result<(), Box<dyn std::error::Error>> {
    let source = Package::from_bytes(SOURCE)?;
    let native = Package::from_bytes(NATIVE)?;
    let source_root = source
        .table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 1),
        )?
        .ok_or("source B2 comment missing")?;
    let native_root = native
        .table_cell_comment_a1("Sheet 1", "Review", "$B$2")?
        .ok_or("native B2 comment missing")?;
    assert_native_root_metadata(&native)?;
    assert_eq!(native_root, source_root);
    assert!(
        native
            .table_cell_comment_a1("Sheet 1", "Review", "C2")?
            .is_none()
    );
    let mut unchanged = Vec::new();
    native.write_to(&mut unchanged)?;
    assert_eq!(unchanged, NATIVE);
    Ok(())
}

/// Set `LITCHI_NATIVE_COMMENT_METADATA_PATH` to a Numbers file saved by the
/// native application to run the same bounded projection check against it.
/// The opt-in path is deliberately capped so a mistaken environment variable
/// cannot turn this integration test into an unbounded file reader.
#[test]
fn optional_native_saved_root_readback_preserves_semantic_metadata()
-> Result<(), Box<dyn std::error::Error>> {
    let Some(path) = std::env::var_os(OPTIONAL_NATIVE_PATH_ENV) else {
        return Ok(());
    };
    let path = PathBuf::from(path);
    let max_read = usize::try_from(MAX_OPTIONAL_NATIVE_BYTES + 1)?;
    let mut file = fs::File::open(&path)?;
    let mut bytes = Vec::with_capacity(max_read);
    file.by_ref()
        .take(MAX_OPTIONAL_NATIVE_BYTES + 1)
        .read_to_end(&mut bytes)?;
    assert!(
        bytes.len() as u64 <= MAX_OPTIONAL_NATIVE_BYTES,
        "optional native Numbers fixture is too large: {} bytes",
        bytes.len()
    );

    let package = Package::from_bytes(&bytes)?;
    let expected = Package::from_bytes(SOURCE)?
        .table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 1),
        )?
        .ok_or("source B2 comment missing")?;
    let root = assert_native_root_metadata(&package)?;
    assert_eq!(root, expected);

    let mut saved = Vec::new();
    package.write_to(&mut saved)?;
    assert_eq!(saved, bytes);
    let reopened = Package::from_bytes(&saved)?;
    assert_eq!(assert_native_root_metadata(&reopened)?, expected);
    Ok(())
}

#[test]
fn native_saved_empty_reply_list_reads_without_field_info_header()
-> Result<(), Box<dyn std::error::Error>> {
    let native = Package::from_bytes(NATIVE)?;
    assert!(
        native
            .table_cell_comment_replies_a1("Sheet 1", "Review", "B2")?
            .is_empty()
    );
    let mut unchanged = Vec::new();
    native.write_to(&mut unchanged)?;
    assert_eq!(unchanged, NATIVE);
    Ok(())
}

#[test]
fn native_created_root_read_preserves_metadata_and_empty_replies()
-> Result<(), Box<dyn std::error::Error>> {
    let native = Package::from_bytes(NATIVE_CREATED)?;
    let root = native
        .table_cell_comment_a1("Sheet 1", "Table 1", "B2")?
        .ok_or("native-created B2 comment missing")?;
    assert_eq!(root.text(), "Native root comment");
    assert!(root.timestamp().is_some());
    assert_eq!(
        root.author().and_then(|author| author.display_name()),
        Some("Ryker Zhu")
    );
    assert_native_created_sentinel(&native)?;
    assert!(
        native
            .table_cell_comment_replies_a1("Sheet 1", "Table 1", "B2")?
            .is_empty()
    );
    let mut unchanged = Vec::new();
    native.write_to(&mut unchanged)?;
    assert_eq!(unchanged, NATIVE_CREATED);
    let reopened = Package::from_bytes(&unchanged)?;
    let reopened_root = reopened
        .table_cell_comment_a1("Sheet 1", "Table 1", "B2")?
        .ok_or("reopened native-created B2 comment missing")?;
    assert_eq!(reopened_root, root);
    assert_native_created_sentinel(&reopened)?;
    assert!(
        reopened
            .table_cell_comment_replies_a1("Sheet 1", "Table 1", "B2")?
            .is_empty()
    );
    Ok(())
}

#[test]
fn native_multiroot_comments_preserve_independent_metadata_and_values()
-> Result<(), Box<dyn std::error::Error>> {
    const BYTES: &[u8] =
        include_bytes!("../../../test-data/iwork/numbers/comment-compat-native-multiroot.numbers");
    let package = Package::from_bytes(BYTES)?;
    let original = Package::from_bytes(NATIVE_CREATED)?;
    assert_eq!(
        package.table_cell_comment_a1("Sheet 1", "Table 1", "B2")?,
        original.table_cell_comment_a1("Sheet 1", "Table 1", "B2")?,
    );
    for (row, address, text, sentinel) in [
        (1, "B2", "Native root comment", "Native comment parity"),
        (
            2,
            "B3",
            "Native second root",
            "Native compatibility sentinel",
        ),
    ] {
        let position = CellPosition::new(row, 1);
        let root = package
            .table_cell_comment("Sheet 1", "Table 1", position)?
            .ok_or("native root missing")?;
        assert_eq!(root.text(), text);
        assert!(root.timestamp().is_some());
        assert_eq!(
            root.author().and_then(|author| author.display_name()),
            Some("Ryker Zhu")
        );
        assert!(
            package
                .table_cell_comment_replies("Sheet 1", "Table 1", position)?
                .is_empty()
        );
        match package
            .table_cell_a1("Sheet 1", "Table 1", address)?
            .storage()
        {
            Storage::Stored(Value::Text(value)) => assert_eq!(value, sentinel),
            other => panic!("native sentinel mismatch: {other:?}"),
        }
        #[cfg(feature = "internal-iwork-source")]
        {
            let compatible = Package::__table_cell_comment_from_bytes_for_compatibility(
                BYTES, "Sheet 1", "Table 1", position,
            )?
            .ok_or("compatibility root missing")?;
            assert_eq!(compatible, root);
            assert!(
                Package::__table_cell_comment_replies_from_bytes_for_compatibility(
                    BYTES, "Sheet 1", "Table 1", position,
                )?
                .is_empty()
            );
        }
    }
    let mut unchanged = Vec::new();
    package.write_to(&mut unchanged)?;
    assert_eq!(unchanged, BYTES);
    Ok(())
}
