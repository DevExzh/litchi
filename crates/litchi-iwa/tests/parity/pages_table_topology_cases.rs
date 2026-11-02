//! Pages body-table row and column insertion and removal (record 0412e)
//! through `litchi_pages::table::topology`, over the shared table-topology
//! engine the Numbers family runs (`litchi_iwa_table::topology`).

use litchi_iwa::pages;
use litchi_iwa_common::table::topology::{
    ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
};
use litchi_pages::Package;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::pages_table_cases::{
    CATALOG, CELLS, COMMENT_ROW, COMMENTS, TABLE, host_cells, host_model, leaf_cells,
};

#[derive(Clone, Copy)]
enum Op {
    InsertRow(RowInsertion),
    InsertColumn(ColumnInsertion),
    RemoveRow(RowDeletion),
    RemoveColumn(ColumnDeletion),
}

fn host(editor: &mut pages::Draft, op: Op) -> TestResult {
    let model = host_model(editor)?;
    match op {
        Op::InsertRow(at) => editor.insert_table_row(model, at)?,
        Op::InsertColumn(at) => editor.insert_table_column(model, at)?,
        Op::RemoveRow(at) => editor.remove_table_row(model, at)?,
        Op::RemoveColumn(at) => editor.remove_table_column(model, at)?,
    }
    Ok(())
}

fn leaf(package: &Package, op: Op) -> TestResult<litchi_pages::table::topology::Commit> {
    let edit = package.edit_body_table_topology(TABLE)?;
    let edit = match op {
        Op::InsertRow(at) => edit.insert_row(at)?,
        Op::InsertColumn(at) => edit.insert_column(at)?,
        Op::RemoveRow(at) => edit.remove_row(at)?,
        Op::RemoveColumn(at) => edit.remove_column(at)?,
    };
    Ok(edit.commit()?)
}

fn topology_case(name: &'static str, tier: Tier, evidence: Evidence, op: Op) -> TestResult {
    topology_case_on(CATALOG, name, tier, evidence, op)
}

fn topology_case_on(
    path: &'static str,
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    op: Op,
) -> TestResult {
    let input = fixture(path)?;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(move |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            host(&mut editor, op)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            let commit = leaf(&package, op)?;
            let restored = commit
                .package()
                .apply_body_table_topology(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host body-table cells", host_cells)
        .read("leaf body-table cells", leaf_cells)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_insert_row_parity() -> TestResult {
    topology_case(
        "pages_body_table_insert_row_parity",
        Tier::A,
        Evidence::Queued,
        Op::InsertRow(RowInsertion::body(1)),
    )
}

#[test]
fn pages_body_table_insert_column_parity() -> TestResult {
    topology_case(
        "pages_body_table_insert_column_parity",
        Tier::A,
        Evidence::Queued,
        Op::InsertColumn(ColumnInsertion::body(1)),
    )
}

#[test]
fn pages_body_table_remove_row_parity() -> TestResult {
    topology_case(
        "pages_body_table_remove_row_parity",
        Tier::A,
        Evidence::Queued,
        Op::RemoveRow(RowDeletion::body(1)),
    )
}

#[test]
fn pages_body_table_remove_column_parity() -> TestResult {
    topology_case(
        "pages_body_table_remove_column_parity",
        Tier::A,
        Evidence::Queued,
        Op::RemoveColumn(ColumnDeletion::body(1)),
    )
}

/// Removing the row of a commented cell clears the comment (and its reply
/// thread) through the comment transaction before the row goes.
#[test]
fn pages_body_table_remove_commented_row_parity() -> TestResult {
    topology_case_on(
        COMMENTS,
        "pages_body_table_remove_commented_row_parity",
        Tier::A,
        Evidence::Queued,
        Op::RemoveRow(RowDeletion::body(COMMENT_ROW - 1)),
    )
}

/// Calculation state the coordinate shift does not rewrite (here a native
/// table with volatile or spanning dependencies) is refused by both routes
/// before anything is published: the host with a parse error, the leaf with
/// a typed [`Unsupported::Formulas`](litchi_pages::table::topology::Unsupported).
#[test]
fn pages_body_table_topology_formula_state_refused_like_host() -> TestResult {
    use litchi_pages::table::topology::{Error, Refusal, Unsupported};
    let input = fixture(CELLS)?;
    let package = Package::from_bytes(&input.bytes)?;
    for op in [
        Op::InsertRow(RowInsertion::body(0)),
        Op::RemoveRow(RowDeletion::body(0)),
    ] {
        let mut editor = pages::Draft::from_bytes(&input.bytes)?;
        assert!(host(&mut editor, op).is_err(), "the host must refuse");
        let refused = leaf(&package, op);
        assert!(
            matches!(
                refused
                    .as_ref()
                    .err()
                    .and_then(|error| error.downcast_ref::<Error>()),
                Some(Error::Refused(Refusal::Unsupported {
                    reason: Unsupported::Formulas
                }))
            ),
            "the leaf must refuse with a typed formula reason: {refused:?}"
        );
    }
    Ok(())
}

/// Out-of-range positions and the last axis are typed refusals.
#[test]
fn pages_body_table_topology_refusals_are_typed() -> TestResult {
    use litchi_pages::table::topology::{Error, Refusal};
    let input = fixture(CATALOG)?;
    let package = Package::from_bytes(&input.bytes)?;
    let refused = package
        .edit_body_table_topology(TABLE)?
        .remove_row(RowDeletion::body(10_000));
    assert!(
        matches!(
            refused,
            Err(Error::Refused(Refusal::PositionOutOfRange { .. }))
        ),
        "an out-of-range row must be refused: {refused:?}"
    );
    let empty = package.edit_body_table_topology(TABLE)?.commit();
    assert!(
        matches!(empty, Err(Error::Refused(Refusal::NoStagedOperation))),
        "an empty edit must be refused: {empty:?}"
    );
    Ok(())
}

fn probe(op: Op) -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(CATALOG)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
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
fn pages_body_table_insert_row_cost_abba() -> TestResult {
    probe(Op::InsertRow(RowInsertion::body(1)))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_insert_column_cost_abba() -> TestResult {
    probe(Op::InsertColumn(ColumnInsertion::body(1)))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_remove_row_cost_abba() -> TestResult {
    probe(Op::RemoveRow(RowDeletion::body(1)))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_remove_column_cost_abba() -> TestResult {
    probe(Op::RemoveColumn(ColumnDeletion::body(1)))
}
