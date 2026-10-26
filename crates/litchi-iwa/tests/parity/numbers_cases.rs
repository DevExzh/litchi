//! Numbers parity cases.

use litchi_iwa::numbers;
use litchi_iwa_common::table::merge::Region;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
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

/// Row and column removal are not at parity yet (record 0370 §2): the leaf
/// clears the removed cells through the cell writer first, which deletes
/// the root previews, releases the cells' format keys and lowers the
/// perpendicular header counts, where the host keeps the previews, leaks
/// the format keys, keeps stale header counts and leaves the pre-BNC
/// offset tables naming the removed column. Until a wave resolves it, this
/// test pins the divergence so the routes stay listed, not shipped.
#[test]
fn numbers_table_removal_divergence_is_reported() -> TestResult {
    use litchi_numbers::table::topology::{ColumnDeletion, RowDeletion};
    for (name, op) in [
        ("remove row", TopologyOp::RemoveRow(RowDeletion::body(0))),
        (
            "remove column",
            TopologyOp::RemoveColumn(ColumnDeletion::body(0)),
        ),
    ] {
        let Err(error) = topology_case(
            "numbers_table_removal_divergence_is_reported",
            Tier::B,
            Evidence::Queued,
            op,
        ) else {
            return Err(
                format!("the {name} routes now agree: turn this into a parity case").into(),
            );
        };
        let text = error.to_string();
        eprintln!("{text}");
        assert!(text.contains("ZIP member sets differ"), "{text}");
    }
    Ok(())
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
    let ops = [
        TopologyOp::InsertRow(RowInsertion::header(0)),
        TopologyOp::InsertRow(RowInsertion::body(0)),
        TopologyOp::InsertRow(RowInsertion::body(1)),
        TopologyOp::InsertRow(RowInsertion::footer(0)),
        TopologyOp::InsertColumn(ColumnInsertion::header(0)),
        TopologyOp::InsertColumn(ColumnInsertion::body(0)),
        TopologyOp::InsertColumn(ColumnInsertion::body(2)),
    ];
    let mut failures = Vec::new();
    let mut compared = 0_usize;
    let mut identical = 0_usize;
    for path in NUMBERS_CORPUS {
        let input = fixture(path)?;
        for op in ops {
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
