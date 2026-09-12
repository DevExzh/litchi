//! Regression gates for the Numbers host-to-focused merge-reader handoff.
//!
//! The host still receives a native model identifier for compatibility with
//! the existing editor API.  The focused package is the independent semantic
//! oracle used here to prove that the handoff preserves selector geometry,
//! exact source bytes, and lazy cell storage.

use std::error::Error as StdError;

use litchi_iwa::{
    Error as IwaError,
    numbers::{NumbersDocumentBuilder, NumbersEditor},
};
use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    iwa::{Archive, RawMessage, SnappyStream},
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::table::merge::Region;
use litchi_iwa_protos::tst;
use litchi_numbers::{MergeReader, Package};
use prost::Message as _;

const NATIVE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-merges-native.numbers"
));
const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "shared-model";
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn expected_merges() -> Vec<Region> {
    vec![Region::new(10, 1, 2, 2).expect("native merge geometry is valid")]
}

fn native_table_model_id(editor: &NumbersEditor) -> TestResult<u64> {
    let mut model_ids = Vec::new();
    for name in editor.package().iwa_entry_names() {
        for object in editor.package().archive(name)?.objects {
            if object
                .messages
                .iter()
                .any(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            {
                model_ids.push(
                    object
                        .archive_info
                        .identifier
                        .ok_or("native table model has no object identifier")?,
                );
            }
        }
    }
    match model_ids.as_slice() {
        [model_id] => Ok(*model_id),
        [] => Err("native Numbers merge table model is missing".into()),
        _ => Err("native Numbers merge table model ownership is ambiguous".into()),
    }
}

/// Replace the selected table's first referenced tile with an invalid BNC
/// payload while preserving the ZIP envelope and all merge metadata.
///
/// Full semantic construction must inspect the tile and reject this source;
/// metadata readers must stop after the table model and still return merges.
fn malformed_native_table_tile(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let tile_identifier = find_native_tile_identifier(&catalog)?;

    let mut replacement = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let Some(object) = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(tile_identifier))
        else {
            continue;
        };
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TILE_MESSAGE_TYPE)
            .ok_or("native merge tile payload is missing")?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: vec![0xff],
            },
        )?;
        replacement = Some((
            entry.name().to_owned(),
            SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
        ));
        break;
    }

    let Some((name, component)) = replacement else {
        return Err("native merge tile object is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn find_native_tile_identifier(catalog: &Catalog) -> TestResult<u64> {
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != TABLE_MODEL_MESSAGE_TYPE {
                    continue;
                }
                let Ok(model) = tst::TableModelArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if model.table_name != TABLE_NAME {
                    continue;
                }
                let tile = model
                    .base_data_store
                    .tiles
                    .tiles
                    .first()
                    .ok_or("native merge tile reference is missing")?;
                return Ok(tile.tile.identifier);
            }
        }
    }
    Err("native merge table model is missing".into())
}

/// Remove only the physical declaration for the selected table-info -> model
/// edge. The table-info payload and selected table model remain intact, so the
/// compatibility reader can still recover the merge owner while the focused
/// reader must reject the malformed rooted graph.
fn undeclared_model_reference_source() -> TestResult<(Vec<u8>, u64)> {
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Merges")
        .table_name("Original")
        .table_dimensions(4, 4)
        .build()?;
    let table_id = editor
        .tables()?
        .first()
        .ok_or("generated table missing")?
        .id();
    editor.merge_cells(table_id, Region::new(1, 0, 1, 2)?)?;
    let source = editor.to_bytes()?;
    let catalog = Catalog::from_bytes(&source)?;

    let mut replacement = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for (message_index, message) in object.messages.iter().enumerate() {
                if !matches!(message.type_, 6_000 | 6_003) {
                    continue;
                }
                let Ok(info) = tst::TableInfoArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if info.table_model.identifier != table_id {
                    continue;
                }
                object
                    .archive_info
                    .message_infos
                    .get_mut(message_index)
                    .ok_or("table-info header metadata missing")?
                    .object_references
                    .clear();
                changed = true;
            }
        }
        if changed {
            replacement = Some((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
            ));
            break;
        }
    }

    let Some((name, component)) = replacement else {
        return Err("generated table-info model edge is missing".into());
    };
    let malformed = catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?;
    Ok((malformed, table_id))
}

#[test]
fn host_merge_read_matches_focused_reader_and_preserves_native_bytes() -> TestResult {
    let expected = expected_merges();
    let focused = MergeReader::from_bytes(NATIVE)?;
    assert_eq!(focused.table_merges(SHEET_NAME, TABLE_NAME)?, expected);
    assert_eq!(
        Package::from_bytes(NATIVE)?.table_merges(SHEET_NAME, TABLE_NAME)?,
        expected
    );

    let host = NumbersEditor::from_bytes(NATIVE)?;
    let model_id = native_table_model_id(&host)?;
    assert_eq!(host.table_cell_merges(model_id)?, expected);
    assert_eq!(host.to_bytes()?, NATIVE);
    Ok(())
}

#[test]
fn host_merge_read_does_not_materialize_unneeded_bnc_tiles() -> TestResult {
    let malformed = malformed_native_table_tile(NATIVE)?;
    assert!(
        Package::from_bytes(&malformed).is_err(),
        "the full focused semantic package must inspect and reject malformed BNC"
    );
    let expected = expected_merges();
    assert_eq!(
        MergeReader::from_bytes(&malformed)?.table_merges(SHEET_NAME, TABLE_NAME)?,
        expected,
        "focused metadata reads must stop before the selected table tile"
    );

    let host = NumbersEditor::from_bytes(&malformed)?;
    let model_id = native_table_model_id(&host)?;
    assert_eq!(
        host.table_cell_merges(model_id)?,
        expected,
        "the host handoff must retain metadata-only merge reads"
    );
    assert_eq!(host.to_bytes()?, malformed);
    Ok(())
}

#[test]
fn host_merge_read_retains_source_after_input_path_is_removed() -> TestResult {
    let file = tempfile::NamedTempFile::new()?;
    std::fs::write(file.path(), NATIVE)?;
    let host = NumbersEditor::open(file.path())?;
    let model_id = native_table_model_id(&host)?;
    drop(file);

    assert_eq!(host.table_cell_merges(model_id)?, expected_merges());
    assert_eq!(host.to_bytes()?, NATIVE);
    Ok(())
}

#[test]
fn host_merge_read_does_not_mask_a_focused_root_edge_error() -> TestResult {
    let (malformed, table_id) = undeclared_model_reference_source()?;
    let focused = MergeReader::from_bytes(&malformed)?;
    assert!(matches!(
        focused.table_merges(0usize, 0usize),
        Err(litchi_numbers::TableMergesError::InvalidSource)
    ));

    let host = NumbersEditor::from_bytes(&malformed)?;
    let error = host
        .table_cell_merges(table_id)
        .expect_err("the focused handoff must reject an undeclared model edge");
    assert!(matches!(
        error,
        IwaError::NumbersTableMerges(litchi_numbers::TableMergesError::InvalidSource)
    ));
    assert_eq!(host.to_bytes()?, malformed);
    Ok(())
}
