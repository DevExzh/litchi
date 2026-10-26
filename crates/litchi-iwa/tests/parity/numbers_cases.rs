//! Numbers parity cases.

use litchi_iwa::numbers;
use litchi_iwa::text::{TextPointSize, TextStyle};
use litchi_iwa_common::table::merge::Region;
use litchi_numbers::cell::text_style::CharacterProperties;
use litchi_numbers::table::At;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;

const MERGES: &str = "iwork/numbers/table-merges-native.numbers";
const SHEET: &str = "Sheet 1";
const TABLE: &str = "shared-model";
const NATIVE_MERGE: (u32, u32, u32, u32) = (10, 1, 2, 2);

/// The host's native table-model identifier of the fixture's only table.
///
/// The host route takes native identifiers; the harness resolves them from
/// the same selector the leaf route uses (here: the only `TST.TableModel`).
fn host_table_model_id(editor: &numbers::Draft) -> TestResult<u64> {
    let mut ids = Vec::new();
    for name in editor.package().iwa_entry_names() {
        for object in editor.package().archive(name)?.objects {
            if object.messages.iter().any(|message| message.type_ == 6_001) {
                ids.extend(object.archive_info.identifier);
            }
        }
    }
    match ids.as_slice() {
        [id] => Ok(*id),
        _ => Err(format!("expected one table model, found {ids:?}").into()),
    }
}

fn host_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let id = host_table_model_id(&editor)?;
    Ok(editor.table_cell_merges(id)?)
}

fn leaf_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    Ok(litchi_numbers::Package::from_bytes(bytes)?.table_merges(SHEET, TABLE)?)
}

enum MergeOp {
    Merge(Region),
    Unmerge(Region),
}

/// The operation expressed once, run through both routes.
fn merge_case(name: &'static str, tier: Tier, evidence: Evidence, op: MergeOp) -> TestResult {
    let input = fixture(MERGES)?;
    let op = &op;
    Case::new(name, Format::Numbers, tier, evidence, &input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let id = host_table_model_id(&editor)?;
            match *op {
                MergeOp::Merge(region) => editor.merge_cells(id, region)?,
                MergeOp::Unmerge(region) => {
                    if !editor.unmerge_cells(id, region)? {
                        return Err("host unmerge found no merge".into());
                    }
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let mut edit = package.edit_table_merges(SHEET, TABLE)?;
            match *op {
                MergeOp::Merge(region) => {
                    edit.merge(region)?;
                },
                MergeOp::Unmerge(region) => {
                    if !edit.unmerge(region)? {
                        return Err("leaf unmerge found no merge".into());
                    }
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_table_merges(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table_cell_merges", host_merges)
        .read("leaf table_merges", leaf_merges)
        .run()?;
    Ok(())
}

fn native_region() -> TestResult<Region> {
    let (row, column, rows, columns) = NATIVE_MERGE;
    Ok(Region::new(row, column, rows, columns)?)
}

#[test]
fn numbers_table_merges_read_parity() -> TestResult {
    let input = fixture(MERGES)?;
    Case::new(
        "numbers_table_merges_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Native("test-data/iwork/numbers/table-merges-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(leaf_bytes!(
            litchi_numbers::Package::from_bytes(source)?
        )))
    })
    .read("host table_cell_merges", host_merges)
    .read("leaf table_merges", leaf_merges)
    .read("expected native merge", |bytes| {
        Ok(leaf_merges(bytes)? == vec![native_region()?])
    })
    .run()?;
    Ok(())
}

#[test]
fn numbers_table_unmerge_parity() -> TestResult {
    merge_case(
        "numbers_table_unmerge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Unmerge(native_region()?),
    )
}

#[test]
fn numbers_table_merge_parity() -> TestResult {
    merge_case(
        "numbers_table_merge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Merge(Region::new(2, 0, 2, 2)?),
    )
}

/// `basic.numbers` (source-built by the host) through the four topology
/// routes. Record 0369 pinned a row-insertion divergence here; record 0370
/// resolved it (the leaf keeps the root previews and the drawable's layout
/// cache, as the host does; the host's tile `numrows` follows the native
/// row-count convention).
const TOPOLOGY: &str = "iwork/numbers/basic.numbers";

fn leaf_sheets(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    Ok(format!("{:?}", package.sheets()))
}

fn leaf_table_cells(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let mut cells = Vec::new();
    for sheet in package.sheets() {
        for table in sheet.tables() {
            for cell in table.iter_cells() {
                cells.push(format!("{cell:?}"));
            }
        }
    }
    Ok(cells)
}

#[derive(Clone, Copy)]
enum TopologyOp {
    InsertRow(litchi_numbers::table::topology::RowInsertion),
    InsertColumn(litchi_numbers::table::topology::ColumnInsertion),
    RemoveRow(litchi_numbers::table::topology::RowDeletion),
    RemoveColumn(litchi_numbers::table::topology::ColumnDeletion),
}

fn topology_case(name: &'static str, tier: Tier, evidence: Evidence, op: TopologyOp) -> TestResult {
    let input = fixture(TOPOLOGY)?;
    Case::new(name, Format::Numbers, tier, evidence, &input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let table = litchi_numbers::TableSelector::index(0);
            match op {
                TopologyOp::InsertRow(at) => editor.insert_table_row(table, at)?,
                TopologyOp::InsertColumn(at) => editor.insert_table_column(table, at)?,
                TopologyOp::RemoveRow(at) => editor.remove_table_row(table, at)?,
                TopologyOp::RemoveColumn(at) => editor.remove_table_column(table, at)?,
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let edit = package.edit_table_topology(0usize, 0usize)?;
            let edit = match op {
                TopologyOp::InsertRow(at) => edit.insert_row(at)?,
                TopologyOp::InsertColumn(at) => edit.insert_column(at)?,
                TopologyOp::RemoveRow(at) => edit.remove_row(at)?,
                TopologyOp::RemoveColumn(at) => edit.remove_column(at)?,
            };
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_table_topology(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("leaf sheets", leaf_sheets)
        .read("leaf table cells", leaf_table_cells)
        .run()?;
    Ok(())
}

#[test]
fn numbers_table_insert_row_parity() -> TestResult {
    topology_case(
        "numbers_table_insert_row_parity",
        Tier::A,
        Evidence::Queued,
        TopologyOp::InsertRow(litchi_numbers::table::topology::RowInsertion::body(1)),
    )
}

#[test]
fn numbers_table_insert_column_parity() -> TestResult {
    topology_case(
        "numbers_table_insert_column_parity",
        Tier::A,
        Evidence::Queued,
        TopologyOp::InsertColumn(litchi_numbers::table::topology::ColumnInsertion::body(1)),
    )
}

/// Row and column removal (record 0370b): on a table no formula touches,
/// the leaf drops the removed cells in its structural pass, as the host
/// does; the host now releases a dropped cell's style and format entries.
#[test]
fn numbers_table_remove_row_parity() -> TestResult {
    topology_case(
        "numbers_table_remove_row_parity",
        Tier::A,
        Evidence::Queued,
        TopologyOp::RemoveRow(litchi_numbers::table::topology::RowDeletion::body(0)),
    )
}

#[test]
fn numbers_table_remove_column_parity() -> TestResult {
    topology_case(
        "numbers_table_remove_column_parity",
        Tier::A,
        Evidence::Queued,
        TopologyOp::RemoveColumn(litchi_numbers::table::topology::ColumnDeletion::body(0)),
    )
}

/// Every tracked Numbers fixture the harness sweeps.
pub const NUMBERS_CORPUS: &[&str] = &[
    "iwork/numbers/audio-playback-native-resaved.numbers",
    "iwork/numbers/audio-playback-native.numbers",
    "iwork/numbers/basic.numbers",
    "iwork/numbers/cell-value-native.numbers",
    "iwork/numbers/chart-arrangement-native.numbers",
    "iwork/numbers/chart-arrangement-retirement-resaved.numbers",
    "iwork/numbers/chart-data-edited-native.numbers",
    "iwork/numbers/chart-data-native.numbers",
    "iwork/numbers/chart-grid-authored-native.numbers",
    "iwork/numbers/chart-grid-replaced-native.numbers",
    "iwork/numbers/comment-compat-native-multiroot.numbers",
    "iwork/numbers/comment-metadata-native-resaved.numbers",
    "iwork/numbers/comment-reader-native-created.numbers",
    "iwork/numbers/custom-datetime-native-resaved.numbers",
    "iwork/numbers/custom-datetime-native.numbers",
    "iwork/numbers/custom-number-native-resaved.numbers",
    "iwork/numbers/custom-number-native.numbers",
    "iwork/numbers/custom-retirement-cleared.numbers",
    "iwork/numbers/custom-retirement-resaved.numbers",
    "iwork/numbers/custom-text-native-resaved.numbers",
    "iwork/numbers/custom-text-native-unicode.numbers",
    "iwork/numbers/custom-text-native.numbers",
    "iwork/numbers/duration-native-resaved.numbers",
    "iwork/numbers/duration-native.numbers",
    "iwork/numbers/formula-arena-native.numbers",
    "iwork/numbers/formula-budget-native.numbers",
    "iwork/numbers/formula-envelope-native.numbers",
    "iwork/numbers/formula-events-native.numbers",
    "iwork/numbers/image-adjustments-direct-resaved.numbers",
    "iwork/numbers/image-adjustments-native-resaved.numbers",
    "iwork/numbers/image-adjustments-native.numbers",
    "iwork/numbers/image-adjustments-retirement-resaved.numbers",
    "iwork/numbers/shared-model-native.numbers",
    "iwork/numbers/table-data-list-native.numbers",
    "iwork/numbers/table-merges-native.numbers",
    "iwork/numbers/table-relocation-native-resaved.numbers",
    "iwork/numbers/table-relocation-native.numbers",
];

/// The two insertion routes over the whole corpus (first table of the
/// first sheet), in every section. Where the host edits, the leaf
/// output must pass tier B (most cases are byte-identical; the others
/// differ only in the Snappy framing of a changed member); where the host
/// refuses, the leaf may refuse too or edit (it never loses a capability
/// the host has).
#[test]
fn numbers_table_topology_corpus_parity() -> TestResult {
    use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};
    let (_tier, _evidence) = (Tier::B, Evidence::Queued);
    topology_corpus(&[
        TopologyOp::InsertRow(RowInsertion::header(0)),
        TopologyOp::InsertRow(RowInsertion::body(0)),
        TopologyOp::InsertRow(RowInsertion::body(1)),
        TopologyOp::InsertRow(RowInsertion::footer(0)),
        TopologyOp::InsertColumn(ColumnInsertion::header(0)),
        TopologyOp::InsertColumn(ColumnInsertion::body(0)),
        TopologyOp::InsertColumn(ColumnInsertion::body(2)),
    ])
}

/// Row and column removal at seven positions on the whole corpus (record
/// 0370b), compared as the insertion sweep is.
#[test]
fn numbers_table_removal_corpus_parity() -> TestResult {
    use litchi_numbers::table::topology::{ColumnDeletion, RowDeletion};
    let (_tier, _evidence) = (Tier::B, Evidence::Queued);
    topology_corpus(&[
        TopologyOp::RemoveRow(RowDeletion::header(0)),
        TopologyOp::RemoveRow(RowDeletion::body(0)),
        TopologyOp::RemoveRow(RowDeletion::body(1)),
        TopologyOp::RemoveRow(RowDeletion::footer(0)),
        TopologyOp::RemoveColumn(ColumnDeletion::header(0)),
        TopologyOp::RemoveColumn(ColumnDeletion::body(0)),
        TopologyOp::RemoveColumn(ColumnDeletion::body(1)),
    ])
}

fn topology_corpus(ops: &[TopologyOp]) -> TestResult {
    let mut failures = Vec::new();
    let mut compared = 0_usize;
    let mut identical = 0_usize;
    for path in NUMBERS_CORPUS {
        let input = fixture(path)?;
        for &op in ops {
            let host = (|| -> TestResult<Vec<u8>> {
                let mut editor = numbers::Draft::from_bytes(&input.bytes)?;
                let table = litchi_numbers::TableSelector::index(0);
                match op {
                    TopologyOp::InsertRow(at) => editor.insert_table_row(table, at)?,
                    TopologyOp::InsertColumn(at) => editor.insert_table_column(table, at)?,
                    TopologyOp::RemoveRow(at) => editor.remove_table_row(table, at)?,
                    TopologyOp::RemoveColumn(at) => editor.remove_table_column(table, at)?,
                }
                Ok(editor.to_bytes()?)
            })();
            let leaf = (|| -> TestResult<Vec<u8>> {
                let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
                let edit = package.edit_table_topology(0usize, 0usize)?;
                let edit = match op {
                    TopologyOp::InsertRow(at) => edit.insert_row(at)?,
                    TopologyOp::InsertColumn(at) => edit.insert_column(at)?,
                    TopologyOp::RemoveRow(at) => edit.remove_row(at)?,
                    TopologyOp::RemoveColumn(at) => edit.remove_column(at)?,
                };
                Ok(leaf_bytes!(edit.commit()?.package()))
            })();
            let label = format!("{path} {} {:?}", op.label(), op.position());
            match (host, leaf) {
                (Ok(host), Ok(leaf)) => {
                    compared += 1;
                    if host == leaf {
                        identical += 1;
                        continue;
                    }
                    let diff = crate::archive_diff::compare(
                        &input.bytes,
                        &host,
                        &leaf,
                        &crate::known_fields::KnownFields::for_format(Format::Numbers),
                    )?;
                    if !diff.violations.is_empty() {
                        failures.push(format!("{label}: tier B violations\n{}", diff.render()));
                    }
                },
                (Ok(_), Err(error)) => {
                    failures.push(format!(
                        "{label}: the leaf refuses what the host edits: {error}"
                    ));
                },
                (Err(host), Ok(_)) => eprintln!("{label}: host refuses ({host}); the leaf edits"),
                (Err(host), Err(leaf)) => {
                    eprintln!("{label}: both refuse (host {host}; leaf {leaf})")
                },
            }
        }
    }
    eprintln!(
        "topology corpus: {compared} routes compared, {identical} byte-identical (tier A), the rest tier B"
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

impl TopologyOp {
    fn position(self) -> String {
        match self {
            Self::InsertRow(at) => format!("{at:?}"),
            Self::InsertColumn(at) => format!("{at:?}"),
            Self::RemoveRow(at) => format!("{at:?}"),
            Self::RemoveColumn(at) => format!("{at:?}"),
        }
    }

    fn label(self) -> &'static str {
        match self {
            Self::InsertRow(_) => "insert row",
            Self::InsertColumn(_) => "insert column",
            Self::RemoveRow(_) => "remove row",
            Self::RemoveColumn(_) => "remove column",
        }
    }
}

/// Whole-cell character style (record 0370): the host's
/// `set_table_cell_text_style` (point size, bold, italic) against the leaf's
/// `edit_table_cell_text_style`, ported from the host's
/// `cell_paragraph_style` new-style path (record 0309).
const TEXT_STYLE_FIXTURE: &str = "iwork/numbers/basic.numbers";

fn host_text_style(bytes: &[u8], row: u32, column: u32) -> TestResult<TextStyle> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let id = host_table_model_id(&editor)?;
    Ok(editor.table_cell_text_style(id, row as usize, column as usize)?)
}

fn leaf_text_style(bytes: &[u8], row: u32, column: u32) -> TestResult<CharacterProperties> {
    Ok(
        litchi_numbers::Package::from_bytes(bytes)?.table_cell_text_style(
            0usize,
            0usize,
            At::new(row, column),
        )?,
    )
}

fn text_style_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    cells: &[(u32, u32)],
    style: TextStyle,
) -> TestResult {
    let input = fixture(TEXT_STYLE_FIXTURE)?;
    let mut properties = CharacterProperties::new();
    properties.size = Some(style.point_size.points());
    properties.bold = Some(style.bold);
    properties.italic = Some(style.italic);
    let properties = &properties;
    let mut case = Case::new(name, Format::Numbers, tier, evidence, &input)
        .leaf_structural_metadata()
        .known_fields(|known| {
            known.with(
                2_022,
                "TSWP.ParagraphStyleArchive",
                decode_as::<litchi_iwa_protos::tswp::ParagraphStyleArchive>,
            )
        })
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let id = host_table_model_id(&editor)?;
            for (row, column) in cells {
                editor.set_table_cell_text_style(id, *row as usize, *column as usize, style)?;
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let mut edit = package.edit_table_cell_text_style(0usize, 0usize)?;
            for (row, column) in cells {
                edit = edit.set(At::new(*row, *column), properties.clone())?;
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_table_cell_text_style(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("leaf table cells", leaf_table_cells);
    for (row, column) in cells {
        let (row, column) = (*row, *column);
        case = case
            .read("host table_cell_text_style", move |bytes| {
                host_text_style(bytes, row, column)
            })
            .read("leaf table_cell_text_style", move |bytes| {
                leaf_text_style(bytes, row, column)
            });
    }
    case.run()?;
    Ok(())
}

/// The host's whole-cell point size, bold and italic read equals the
/// leaf's effective character properties on every corpus cell B2 and on
/// the cells both routes styled.
#[test]
fn numbers_cell_text_style_read_parity() -> TestResult {
    // A read-only comparison of the same bytes; no native receipt.
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let mut compared = 0_usize;
    for path in NUMBERS_CORPUS {
        let input = fixture(path)?;
        let mut inputs = vec![input.bytes.clone()];
        if let Ok(mut editor) = numbers::Draft::from_bytes(&input.bytes)
            && let Ok(id) = host_first_table_model_id(&editor)
            && editor
                .set_table_cell_text_style(
                    id,
                    1,
                    1,
                    TextStyle::new(TextPointSize::from_points(21.0)?).with_italic(true),
                )
                .is_ok()
        {
            inputs.push(editor.to_bytes()?);
        }
        for bytes in inputs {
            let Ok(editor) = numbers::Draft::from_bytes(&bytes) else {
                continue;
            };
            let Ok(id) = host_first_table_model_id(&editor) else {
                continue;
            };
            let Ok(host) = editor.table_cell_text_style(id, 1, 1) else {
                continue;
            };
            let leaf = leaf_text_style(&bytes, 1, 1)?;
            // An unset size reads as the host's 12 pt default.
            assert_eq!(
                leaf.size.unwrap_or(TextPointSize::TWELVE.points()),
                host.point_size.points(),
                "{path}"
            );
            assert_eq!(leaf.bold.unwrap_or(false), host.bold, "{path}");
            assert_eq!(leaf.italic.unwrap_or(false), host.italic, "{path}");
            compared += 1;
        }
    }
    eprintln!("text style reads compared: {compared}");
    assert!(compared >= 30, "only {compared} reads compared");
    Ok(())
}

/// Two cells of one table take the same style: each takes its own
/// variation on both routes.
#[test]
fn numbers_cell_text_style_shared_parity() -> TestResult {
    text_style_case(
        "numbers_cell_text_style_shared_parity",
        Tier::B,
        Evidence::Queued,
        &[(1, 1), (2, 1)],
        TextStyle::new(TextPointSize::from_points(9.5)?)
            .with_bold(false)
            .with_italic(true),
    )
}

/// The style route over every tracked fixture whose first table stores
/// cell B2: where the host edits, the leaf passes tier B.
#[test]
fn numbers_cell_text_style_corpus_parity() -> TestResult {
    let style = TextStyle::new(TextPointSize::from_points(14.0)?).with_bold(true);
    let mut properties = CharacterProperties::new();
    properties.size = Some(14.0);
    properties.bold = Some(true);
    properties.italic = Some(false);
    let known = crate::known_fields::KnownFields::for_format(Format::Numbers).with(
        2_022,
        "TSWP.ParagraphStyleArchive",
        decode_as::<litchi_iwa_protos::tswp::ParagraphStyleArchive>,
    );
    let mut failures = Vec::new();
    let mut compared = 0_usize;
    for path in NUMBERS_CORPUS {
        let input = fixture(path)?;
        let host = (|| -> TestResult<Vec<u8>> {
            let mut editor = numbers::Draft::from_bytes(&input.bytes)?;
            let id = host_first_table_model_id(&editor)?;
            editor.set_table_cell_text_style(id, 1, 1, style)?;
            Ok(editor.to_bytes()?)
        })();
        let leaf = (|| -> TestResult<Vec<u8>> {
            let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
            let commit = package
                .edit_table_cell_text_style(0usize, 0usize)?
                .set(At::new(1, 1), properties.clone())?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        })();
        match (host, leaf) {
            (Ok(host), Ok(leaf)) => {
                compared += 1;
                let diff = crate::archive_diff::compare_with_policies(
                    &input.bytes,
                    &host,
                    &leaf,
                    &known,
                    &[],
                    crate::archive_diff::MetadataPolicy {
                        save_token: true,
                        fresh_uuids: true,
                    },
                )?;
                if !diff.violations.is_empty() {
                    failures.push(format!("{path}: tier B violations\n{}", diff.render()));
                }
            },
            (Ok(_), Err(error)) => {
                failures.push(format!(
                    "{path}: the leaf refuses what the host edits: {error}"
                ));
            },
            (Err(host), Ok(_)) => eprintln!("{path}: host refuses ({host}); the leaf edits"),
            (Err(host), Err(leaf)) => eprintln!("{path}: both refuse (host {host}; leaf {leaf})"),
        }
    }
    eprintln!("text style corpus: {compared} routes compared");
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

#[test]
fn numbers_cell_text_style_parity() -> TestResult {
    text_style_case(
        "numbers_cell_text_style_parity",
        Tier::B,
        Evidence::Queued,
        &[(1, 1)],
        TextStyle::new(TextPointSize::from_points(18.0)?)
            .with_bold(true)
            .with_italic(false),
    )
}
/// The host's native identifier of the first sheet's first table model.
pub fn host_first_table_model_id(editor: &numbers::Draft) -> TestResult<u64> {
    let package = editor.package();
    let document = package
        .iwa_entry_names()
        .find_map(|name| {
            package
                .archive(name)
                .ok()?
                .objects
                .into_iter()
                .find_map(|object| {
                    object
                        .messages
                        .iter()
                        .find(|message| message.type_ == 1)
                        .map(|message| message.data.clone())
                })
        })
        .ok_or("no document")?;
    let document =
        <litchi_iwa_protos::tn::DocumentArchive as prost::Message>::decode(document.as_slice())?;
    let sheet = document.sheets.first().ok_or("no sheet")?.identifier;
    let find = |identifier: u64, types: &[u32]| -> TestResult<Vec<u8>> {
        for name in package.iwa_entry_names() {
            for object in package.archive(name)?.objects {
                if object.archive_info.identifier == Some(identifier)
                    && let Some(message) = object.messages.iter().find(|m| types.contains(&m.type_))
                {
                    return Ok(message.data.clone());
                }
            }
        }
        Err(format!("object {identifier} not found").into())
    };
    let sheet = <litchi_iwa_protos::tn::SheetArchive as prost::Message>::decode(
        find(sheet, &[2])?.as_slice(),
    )?;
    for drawable in &sheet.drawable_infos {
        if let Ok(info) = find(drawable.identifier, &[6_000]) {
            let info = <litchi_iwa_protos::tst::TableInfoArchive as prost::Message>::decode(
                info.as_slice(),
            )?;
            return Ok(info.table_model.identifier);
        }
    }
    Err("the first sheet has no table".into())
}

/// The host's workbook catalog reads (`sheets`, `tables`) against the
/// leaf's `Package::sheets` and each sheet's tables with
/// `Package::table_appearance`, on every tracked fixture the host opens
/// (record 0370).
#[test]
fn numbers_workbook_catalog_read_parity() -> TestResult {
    // A read-only comparison of the same bytes; no native receipt needed
    // beyond the fixtures' own.
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let mut compared = 0_usize;
    for path in NUMBERS_CORPUS {
        let input = fixture(path)?;
        let editor = numbers::Draft::from_bytes(&input.bytes)?;
        let leaf = litchi_numbers::Package::from_bytes(&input.bytes)?;
        let host_sheets: Vec<String> = editor
            .sheets()?
            .into_iter()
            .map(|sheet| sheet.name)
            .collect();
        let leaf_sheets: Vec<String> = leaf
            .sheets()
            .iter()
            .map(|sheet| sheet.name().to_owned())
            .collect();
        assert_eq!(host_sheets, leaf_sheets, "{path}: sheets");
        let Ok(host_tables) = editor.tables() else {
            // The host's catalog refuses tables its appearance read rejects.
            continue;
        };
        let mut leaf_tables = Vec::new();
        for (sheet_index, sheet) in leaf.sheets().iter().enumerate() {
            for (table_index, table) in sheet.tables().enumerate() {
                let appearance = leaf.table_appearance(sheet_index, table_index)?;
                leaf_tables.push(format!(
                    "{} {}x{} {appearance:?}",
                    table.name(),
                    table.row_count(),
                    table.column_count()
                ));
            }
        }
        let host_tables: Vec<String> = host_tables
            .into_iter()
            .map(|table| {
                format!(
                    "{} {}x{} {:?}",
                    table.name, table.rows, table.columns, table.appearance
                )
            })
            .collect();
        assert_eq!(host_tables, leaf_tables, "{path}: tables");
        compared += 1;
    }
    eprintln!("workbook catalogs compared: {compared}");
    assert!(compared >= 20, "only {compared} catalogs compared");
    Ok(())
}
