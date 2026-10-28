//! Keynote slide-table cell parity cases (wave W1-C, record 0372): cell
//! writes through `slide::drawable::Edit::set_table_cell`.

use litchi_iwa::keynote;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_protos::{tsd, tsp};
use litchi_keynote::slide::drawable::{Kind, TableCellValue};
use litchi_keynote::{DrawableSelector, Package, SlideSelector, TableSelector};
use prost::Message as _;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{leaf_position, leaf_run};
use crate::known_fields::{Decoded, top_level_fields};
use crate::leaf_bytes;

const TABLE: &str = "iwork/keynote/table-discovery.key";
const SLIDE: usize = 0;
const COMMENT_ROW: usize = 0;
const COMMENT_COLUMN: usize = 0;

fn masked_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tsd::CommentStorageArchive::decode(data)?;
    if message.creation_date.is_some() {
        message.creation_date = Some(tsp::Date { seconds: 0.0 });
    }
    if message.storage_uuid.is_some() {
        message.storage_uuid = Some(tsp::Uuid { lower: 1, upper: 1 });
    }
    let mut decoded =
        crate::known_fields::decode_as::<tsd::CommentStorageArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn comment_position() -> TestResult<At> {
    Ok(At::new(
        u32::try_from(COMMENT_ROW)?,
        u32::try_from(COMMENT_COLUMN)?,
    ))
}

fn host_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = match editor.slide_table_cell_comment(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN) {
        Ok(comment) => comment,
        Err(error) if error.to_string().contains("has no comment") => None,
        Err(error) => return Err(error.into()),
    };
    let replies = if comment.is_some() {
        editor.slide_table_cell_comment_replies(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?
    } else {
        Vec::new()
    };
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.comment.text.as_str()),
        replies
            .iter()
            .map(|reply| reply.comment.text.as_str())
            .collect::<Vec<_>>()
    ))
}

fn leaf_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let comment = package.slide_table_cell_comment(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        position,
    )?;
    let replies = package.slide_table_cell_comment_replies(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        position,
    )?;
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.text()),
        replies.iter().map(|reply| reply.text()).collect::<Vec<_>>()
    ))
}

fn comment_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft, u64) -> TestResult + 'static,
    leaf: impl FnOnce(Package, At) -> TestResult<LeafRun> + 'static,
) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(move |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            let bytes = editor.to_bytes()?;
            Ok(bytes)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            let result = leaf(package, comment_position()?)?;
            Ok(result)
        })
        .known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (creation date and storage UUID masked)",
                masked_comment_storage,
            )
        })
        .leaf_structural_metadata()
        .read("host slide-table comment thread", host_comment_thread)
        .read("leaf slide-table comment thread", leaf_comment_thread)
        .run()?;
    Ok(())
}

fn comment_read_case(name: &'static str, tier: Tier, evidence: Evidence) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host slide-table comment thread", host_comment_thread)
        .read("leaf slide-table comment thread", leaf_comment_thread)
        .run()?;
    Ok(())
}

fn host_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
        .first()
        .ok_or("no slide table")?
        .model_object_id)
}

/// Host read: every cell of the first slide table.
fn host_cells(bytes: &[u8]) -> TestResult<Vec<((usize, usize), String)>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let table = editor.slide_table(SLIDE, host_model(&editor)?)?;
    Ok(table
        .iter_cells()
        .map(|(at, value)| (at, format!("{value:?}")))
        .collect())
}

/// Leaf read: every cell of the first slide table.
fn leaf_cells(bytes: &[u8]) -> TestResult<Vec<((u32, u32), String)>> {
    let package = Package::from_bytes(bytes)?;
    let table = package.slide_table_cells(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    Ok(table
        .iter_cells()
        .map(|cell| {
            (
                (cell.position().row(), cell.position().column()),
                format!("{:?}", cell.value()),
            )
        })
        .collect())
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host cells", host_cells)
        .read("leaf cells", leaf_cells)
        .run()?;
    Ok(())
}

/// The first stored cell of the first slide table (a cleared absent cell
/// is not a clear).
fn first_stored(editor: &keynote::Draft) -> TestResult<(usize, usize)> {
    let table = editor.slide_table(SLIDE, host_model(editor)?)?;
    Ok(table
        .iter_cells()
        .map(|(position, _)| position)
        .next()
        .ok_or("table has no stored cell")?)
}

fn leaf_set(source: &[u8], cells: &[(u32, u32, TableCellValue)]) -> TestResult<LeafRun> {
    let table = leaf_position(source, Kind::Table)?;
    leaf_run(source, |edit| {
        for (row, column, value) in cells {
            edit.set_table_cell(
                SlideSelector::index(SLIDE),
                DrawableSelector::index(table),
                *row,
                *column,
                value.clone(),
            )?;
        }
        Ok(())
    })
}

#[test]
fn keynote_slide_table_comment_set_parity() -> TestResult {
    comment_case(
        "keynote_slide_table_comment_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.set_slide_table_cell_comment(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W1-G root",
            )?;
            Ok(())
        },
        |package, position| {
            let commit = package.set_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
                "W1-G root",
            )?;
            let restored = commit
                .package()
                .apply_slide_table_cell_comment(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

#[test]
fn keynote_slide_table_comment_clear_parity() -> TestResult {
    comment_case(
        "keynote_slide_table_comment_clear_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.set_slide_table_cell_comment(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W1-G root",
            )?;
            editor.clear_slide_table_cell_comment(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |package, position| {
            let set = package.set_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
                "W1-G root",
            )?;
            let clear = set.package().clear_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
            )?;
            let undo_clear = clear
                .package()
                .apply_slide_table_cell_comment(&clear.patch().inverse())?;
            let restored = undo_clear
                .package()
                .apply_slide_table_cell_comment(&set.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(clear.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

#[test]
fn keynote_slide_table_comment_read_parity() -> TestResult {
    comment_read_case(
        "keynote_slide_table_comment_read_parity",
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn keynote_slide_table_comment_replies_read_parity() -> TestResult {
    comment_read_case(
        "keynote_slide_table_comment_replies_read_parity",
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn keynote_slide_table_cell_set_parity() -> TestResult {
    run(
        "keynote_slide_table_cell_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            editor.set_slide_table_cell(
                SLIDE,
                model,
                1,
                1,
                keynote::TableCellValue::Text("Parity".into()),
            )?;
            Ok(())
        },
        |source| leaf_set(source, &[(1, 1, TableCellValue::Text("Parity".into()))]),
    )
}

#[test]
fn keynote_slide_table_cells_set_parity() -> TestResult {
    run(
        "keynote_slide_table_cells_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            editor.set_slide_table_cells(
                SLIDE,
                model,
                [
                    keynote::TableCellUpdate::new(
                        1,
                        1,
                        keynote::TableCellValue::Number(42.5.try_into()?),
                    ),
                    keynote::TableCellUpdate::new(2, 1, keynote::TableCellValue::Boolean(true)),
                ],
            )?;
            Ok(())
        },
        |source| {
            leaf_set(
                source,
                &[
                    (1, 1, TableCellValue::Number(42.5)),
                    (2, 1, TableCellValue::Boolean(true)),
                ],
            )
        },
    )
}

#[test]
fn keynote_slide_table_cell_clear_parity() -> TestResult {
    run(
        "keynote_slide_table_cell_clear_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            let (row, column) = first_stored(editor)?;
            editor.clear_slide_table_cell(SLIDE, model, row, column)?;
            Ok(())
        },
        |source| {
            let (row, column) = first_stored(&keynote::Draft::from_bytes(source)?)?;
            leaf_set(
                source,
                &[(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                    TableCellValue::Empty,
                )],
            )
        },
    )
}
