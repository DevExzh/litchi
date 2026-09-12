//! Regression coverage for Pages and Keynote metadata-only merge handoffs.
//!
//! These tests exercise the public migration hosts together with their
//! format-owned readers.  The selected merge store is intentionally kept
//! separate from table-cell storage: corrupting the selected table's BNC tile
//! must fail a complete cell read while leaving the metadata-only merge read
//! usable and source preserving.

use std::error::Error as StdError;

use litchi_iwa::{keynote::KeynoteEditor, pages::PagesEditor};
use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    iwa::{Archive, RawMessage, SnappyStream},
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::table::merge::Region;
use litchi_iwa_protos::{tsce, tst};
use prost::Message as _;

const PAGES: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-merges-native.pages"
));
const KEYNOTE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/slide-table-merges-native.key"
));

const TILE_MESSAGE_TYPE: u32 = 6_002;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn pages_merges() -> [Region; 1] {
    [Region::new(3, 2, 1, 2).expect("native Pages merge geometry is valid")]
}

fn keynote_merges() -> [Region; 1] {
    [Region::new(3, 1, 2, 2).expect("native Keynote merge geometry is valid")]
}

fn model_with_merge_owner(message: &RawMessage) -> Option<tst::TableModelArchive> {
    let model = tst::TableModelArchive::decode(message.data.as_slice()).ok()?;
    model.merge_owner.as_ref()?;
    Some(model)
}

fn first_merge_tile_identifier(source: &[u8]) -> TestResult<u64> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            for message in &object.messages {
                let Some(model) = model_with_merge_owner(message) else {
                    continue;
                };
                let tile = model
                    .base_data_store
                    .tiles
                    .tiles
                    .first()
                    .ok_or("native merge table has no tile reference")?;
                return Ok(tile.tile.identifier);
            }
        }
    }
    Err("native merge table model is missing".into())
}

/// Replace the selected table's BNC payload with an invalid wire fragment.
/// The model, merge owner, ZIP envelope, and all metadata remain intact.
fn malformed_merge_tile(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let tile_identifier = first_merge_tile_identifier(source)?;
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

fn first_merge_formula_mut(
    model: &mut tst::TableModelArchive,
) -> TestResult<&mut tsce::FormulaArchive> {
    model
        .merge_owner
        .as_mut()
        .and_then(|owner| owner.formula_store.as_mut())
        .and_then(|store| store.formulas.first_mut())
        .map(|pair| &mut pair.formula)
        .ok_or_else(|| "native merge formula is missing".into())
}

/// Make the selected formula graph invalid while preserving its model and
/// object identity.  This is a terminal selected-graph error: a host adapter
/// must return it instead of recovering through its old raw archive route.
fn malformed_merge_formula(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;
    let mut mutate = true;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message_index in 0..object.messages.len() {
                let message = &object.messages[message_index];
                let Some(mut model) = model_with_merge_owner(message) else {
                    continue;
                };
                if !mutate {
                    continue;
                }
                let formula = first_merge_formula_mut(&mut model)?;
                let function = formula
                    .ast_node_array
                    .ast_node
                    .iter_mut()
                    .find(|node| {
                        node.ast_node_type
                            == tsce::ast_node_array_archive::AstNodeType::FunctionNode as i32
                    })
                    .ok_or("native merge formula has no function node")?;
                function.ast_function_node_index = Some(167);
                object.replace_message_preserving_header(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data: model.encode_to_vec(),
                    },
                )?;
                mutate = false;
                changed = true;
                break;
            }
            if changed {
                break;
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
        return Err("native merge model is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

#[test]
fn pages_merge_handoff_matches_focused_reader_and_preserves_native_bytes() -> TestResult {
    let expected = pages_merges();
    let reader = litchi_pages::MergeReader::from_bytes(PAGES)?;
    assert_eq!(reader.body_table_merges("Table 1")?, expected);
    let focused = litchi_pages::Package::from_bytes(PAGES)?;
    assert_eq!(focused.body_table_merges("Table 1")?, expected);
    assert_eq!(focused.body_table_merges(0usize)?, expected);

    let host = PagesEditor::from_bytes(PAGES)?;
    let table = host
        .tables()?
        .into_iter()
        .find(|table| table.name == "Table 1")
        .ok_or("native Pages merge table is missing")?;
    assert_eq!(host.table_cell_merges(table.model_object_id)?, expected);
    assert_eq!(host.to_bytes()?, PAGES);
    Ok(())
}

#[test]
fn keynote_merge_handoff_matches_focused_reader_and_preserves_native_bytes() -> TestResult {
    let expected = keynote_merges();
    let reader = litchi_keynote::MergeReader::from_bytes(KEYNOTE)?;
    assert_eq!(reader.slide_table_merges(0usize, 0usize)?, expected);
    let focused = litchi_keynote::Package::from_bytes(KEYNOTE)?;
    assert_eq!(focused.slide_table_merges(0usize, 0usize)?, expected);

    let host = KeynoteEditor::from_bytes(KEYNOTE)?;
    let table = host
        .slide_tables(0)?
        .into_iter()
        .next()
        .ok_or("native Keynote merge table is missing")?;
    assert_eq!(
        host.slide_table_cell_merges(0, table.model_object_id)?,
        expected
    );
    assert_eq!(host.to_bytes()?, KEYNOTE);
    Ok(())
}

#[test]
fn pages_merge_handoff_does_not_materialize_unneeded_bnc_tiles() -> TestResult {
    let malformed = malformed_merge_tile(PAGES)?;
    let reader = litchi_pages::MergeReader::from_bytes(&malformed)?;
    assert_eq!(
        reader.body_table_merges(litchi_pages::BodyTableSelector::index(0))?,
        pages_merges()
    );
    let focused = litchi_pages::Package::from_bytes(&malformed)?;
    assert!(
        focused
            .body_table_cells(litchi_pages::BodyTableSelector::index(0))
            .is_err(),
        "the complete Pages cell reader must inspect and reject malformed BNC"
    );
    assert_eq!(
        focused.body_table_merges(litchi_pages::BodyTableSelector::index(0))?,
        pages_merges()
    );

    let host = PagesEditor::from_bytes(&malformed)?;
    let table = host
        .tables()?
        .into_iter()
        .find(|table| table.name == "Table 1")
        .ok_or("native Pages merge table is missing")?;
    assert!(
        host.table(table.model_object_id).is_err(),
        "the host's complete table reader must inspect and reject malformed BNC"
    );
    assert_eq!(
        host.table_cell_merges(table.model_object_id)?,
        pages_merges()
    );
    assert_eq!(host.to_bytes()?, malformed);
    Ok(())
}

#[test]
fn keynote_merge_handoff_does_not_materialize_unneeded_bnc_tiles() -> TestResult {
    let malformed = malformed_merge_tile(KEYNOTE)?;
    let reader = litchi_keynote::MergeReader::from_bytes(&malformed)?;
    assert_eq!(reader.slide_table_merges(0usize, 0usize)?, keynote_merges());
    let focused = litchi_keynote::Package::from_bytes(&malformed)?;
    assert!(
        focused.slide_table_cells(0usize, 0usize).is_err(),
        "the complete Keynote cell reader must inspect and reject malformed BNC"
    );
    assert_eq!(
        focused.slide_table_merges(0usize, 0usize)?,
        keynote_merges()
    );

    let host = KeynoteEditor::from_bytes(&malformed)?;
    let table = host
        .slide_tables(0)?
        .into_iter()
        .next()
        .ok_or("native Keynote merge table is missing")?;
    assert!(
        host.slide_table(0, table.model_object_id).is_err(),
        "the host's complete table reader must inspect and reject malformed BNC"
    );
    assert_eq!(
        host.slide_table_cell_merges(0, table.model_object_id)?,
        keynote_merges()
    );
    assert_eq!(host.to_bytes()?, malformed);
    Ok(())
}

#[test]
fn pages_merge_handoff_retains_source_after_input_path_is_removed() -> TestResult {
    let directory = tempfile::tempdir()?;
    let path = directory.path().join("native.pages");
    std::fs::write(&path, PAGES)?;
    let host = PagesEditor::open(&path)?;
    let reader = litchi_pages::MergeReader::open(&path)?;
    std::fs::remove_file(&path)?;

    let table = host
        .tables()?
        .into_iter()
        .find(|table| table.name == "Table 1")
        .ok_or("native Pages merge table is missing")?;
    assert_eq!(
        host.table_cell_merges(table.model_object_id)?,
        pages_merges()
    );
    assert_eq!(reader.body_table_merges(0usize)?, pages_merges());
    assert_eq!(host.to_bytes()?, PAGES);
    Ok(())
}

#[test]
fn keynote_merge_handoff_retains_source_after_input_path_is_removed() -> TestResult {
    let directory = tempfile::tempdir()?;
    let path = directory.path().join("native.key");
    std::fs::write(&path, KEYNOTE)?;
    let host = KeynoteEditor::open(&path)?;
    let reader = litchi_keynote::MergeReader::open(&path)?;
    std::fs::remove_file(&path)?;

    let table = host
        .slide_tables(0)?
        .into_iter()
        .next()
        .ok_or("native Keynote merge table is missing")?;
    assert_eq!(
        host.slide_table_cell_merges(0, table.model_object_id)?,
        keynote_merges()
    );
    assert_eq!(reader.slide_table_merges(0usize, 0usize)?, keynote_merges());
    assert_eq!(host.to_bytes()?, KEYNOTE);
    Ok(())
}

#[test]
fn pages_selected_merge_graph_errors_are_terminal_and_source_preserving() -> TestResult {
    let malformed = malformed_merge_formula(PAGES)?;
    let focused = litchi_pages::Package::from_bytes(&malformed)?;
    assert!(
        matches!(
            focused.body_table_merges(litchi_pages::BodyTableSelector::index(0)),
            Err(litchi_pages::BodyTableMergesError::InvalidSource)
        ),
        "the focused Pages reader must reject the malformed selected graph"
    );

    let host = PagesEditor::from_bytes(&malformed)?;
    let table = host
        .tables()?
        .into_iter()
        .find(|table| table.name == "Table 1")
        .ok_or("native Pages merge table is missing")?;
    assert!(host.table_cell_merges(table.model_object_id).is_err());
    assert_eq!(host.to_bytes()?, malformed);
    Ok(())
}

#[test]
fn keynote_selected_merge_graph_errors_are_terminal_and_source_preserving() -> TestResult {
    let malformed = malformed_merge_formula(KEYNOTE)?;
    let focused = litchi_keynote::Package::from_bytes(&malformed)?;
    assert!(
        matches!(
            focused.slide_table_merges(0usize, 0usize),
            Err(litchi_keynote::SlideTableMergesError::InvalidSource)
        ),
        "the focused Keynote reader must reject the malformed selected graph"
    );

    let host = KeynoteEditor::from_bytes(&malformed)?;
    let table = host
        .slide_tables(0)?
        .into_iter()
        .next()
        .ok_or("native Keynote merge table is missing")?;
    assert!(
        host.slide_table_cell_merges(0, table.model_object_id)
            .is_err()
    );
    assert_eq!(host.to_bytes()?, malformed);
    Ok(())
}
