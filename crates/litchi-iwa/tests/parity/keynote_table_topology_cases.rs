//! Keynote slide-table row and column insertion and removal (record 0412e)
//! through `litchi_keynote::slide::table::topology`, over the shared
//! table-topology engine the Numbers family runs
//! (`litchi_iwa_table::topology`).

use litchi_iwa::keynote;
use litchi_iwa_common::table::topology::{
    ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
};
use litchi_keynote::{Package, SlideSelector, TableSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_table_cases::{SLIDE, TABLE, host_cells, host_model, leaf_cells};
use crate::leaf_bytes;

#[derive(Clone, Copy)]
enum Op {
    InsertRow(RowInsertion),
    InsertColumn(ColumnInsertion),
    RemoveRow(RowDeletion),
    RemoveColumn(ColumnDeletion),
}

fn host(editor: &mut keynote::Draft, op: Op) -> TestResult {
    let model = host_model(editor)?;
    match op {
        Op::InsertRow(at) => editor.insert_slide_table_row(SLIDE, model, at)?,
        Op::InsertColumn(at) => editor.insert_slide_table_column(SLIDE, model, at)?,
        Op::RemoveRow(at) => editor.remove_slide_table_row(SLIDE, model, at)?,
        Op::RemoveColumn(at) => editor.remove_slide_table_column(SLIDE, model, at)?,
    }
    Ok(())
}

fn leaf(package: &Package, op: Op) -> TestResult<litchi_keynote::slide::table::topology::Commit> {
    let edit =
        package.edit_slide_table_topology(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    let edit = match op {
        Op::InsertRow(at) => edit.insert_row(at)?,
        Op::InsertColumn(at) => edit.insert_column(at)?,
        Op::RemoveRow(at) => edit.remove_row(at)?,
        Op::RemoveColumn(at) => edit.remove_column(at)?,
    };
    Ok(edit.commit()?)
}

fn topology_case(name: &'static str, tier: Tier, evidence: Evidence, op: Op) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(move |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, op)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            let commit = leaf(&package, op)?;
            let restored = commit
                .package()
                .apply_slide_table_topology(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host slide-table cells", host_cells)
        .read("leaf slide-table cells", leaf_cells)
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_table_insert_row_parity() -> TestResult {
    topology_case(
        "keynote_slide_table_insert_row_parity",
        Tier::A,
        Evidence::Queued,
        Op::InsertRow(RowInsertion::body(1)),
    )
}

#[test]
fn keynote_slide_table_insert_column_parity() -> TestResult {
    topology_case(
        "keynote_slide_table_insert_column_parity",
        Tier::A,
        Evidence::Queued,
        Op::InsertColumn(ColumnInsertion::body(1)),
    )
}

#[test]
fn keynote_slide_table_remove_row_parity() -> TestResult {
    topology_case(
        "keynote_slide_table_remove_row_parity",
        Tier::A,
        Evidence::Queued,
        Op::RemoveRow(RowDeletion::body(1)),
    )
}

#[test]
fn keynote_slide_table_remove_column_parity() -> TestResult {
    topology_case(
        "keynote_slide_table_remove_column_parity",
        Tier::A,
        Evidence::Queued,
        Op::RemoveColumn(ColumnDeletion::body(1)),
    )
}

/// Out-of-range positions and the last axis are typed refusals.
#[test]
fn keynote_slide_table_topology_refusals_are_typed() -> TestResult {
    use litchi_keynote::slide::table::topology::{Error, Refusal};
    let input = fixture(TABLE)?;
    let package = Package::from_bytes(&input.bytes)?;
    let refused = package
        .edit_slide_table_topology(SlideSelector::index(SLIDE), TableSelector::index(0))?
        .remove_row(RowDeletion::body(10_000));
    assert!(
        matches!(
            refused,
            Err(Error::Refused(Refusal::PositionOutOfRange { .. }))
        ),
        "an out-of-range row must be refused: {refused:?}"
    );
    let empty = package
        .edit_slide_table_topology(SlideSelector::index(SLIDE), TableSelector::index(0))?
        .commit();
    assert!(
        matches!(empty, Err(Error::Refused(Refusal::NoStagedOperation))),
        "an empty edit must be refused: {empty:?}"
    );
    Ok(())
}

fn probe(op: Op) -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLE)?,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, op)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            Ok(leaf_bytes!(leaf(&package, op)?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_insert_row_cost_abba() -> TestResult {
    probe(Op::InsertRow(RowInsertion::body(1)))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_insert_column_cost_abba() -> TestResult {
    probe(Op::InsertColumn(ColumnInsertion::body(1)))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_remove_row_cost_abba() -> TestResult {
    probe(Op::RemoveRow(RowDeletion::body(1)))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_remove_column_cost_abba() -> TestResult {
    probe(Op::RemoveColumn(ColumnDeletion::body(1)))
}
