//! Callgrind ABBA cost probes: each migrated route against its host route
//! (record 0363 §5, record 0370).
//!
//! Each probe runs the host route and the leaf route of one operation in
//! ABBA order on the same fixture bytes: open, edit, commit and serialize.
//! The probes are ignored tests; measure one route at a time from a release
//! build with debug information:
//!
//! ```sh
//! CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
//! taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*cost_probes::host_topology*' \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   cost_probes::numbers_topology_insert_cost_abba --test-threads=1
//! # likewise with '--toggle-collect=*cost_probes::leaf_topology*'
//! ```

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};

use crate::harness::{TestResult, fixture};
use crate::leaf_bytes;

#[derive(Clone, Copy)]
enum Insert {
    Row(RowInsertion),
    Column(ColumnInsertion),
}

#[inline(never)]
fn host_topology(source: &[u8], op: Insert) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = litchi_numbers::TableSelector::index(0);
    match op {
        Insert::Row(at) => editor.insert_table_row(table, at)?,
        Insert::Column(at) => editor.insert_table_column(table, at)?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_topology(source: &[u8], op: Insert) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let edit = package.edit_table_topology(0usize, 0usize)?;
    let edit = match op {
        Insert::Row(at) => edit.insert_row(at)?,
        Insert::Column(at) => edit.insert_column(at)?,
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Row and column insertion on `basic.numbers` and a native fixture.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_topology_insert_cost_abba() -> TestResult {
    let inputs = [
        fixture("iwork/numbers/basic.numbers")?,
        fixture("iwork/numbers/custom-number-native.numbers")?,
    ];
    let ops = [
        Insert::Row(RowInsertion::body(1)),
        Insert::Column(ColumnInsertion::body(0)),
    ];
    for input in &inputs {
        for op in ops {
            black_box(host_topology(&input.bytes, op)?);
            black_box(leaf_topology(&input.bytes, op)?);
            black_box(leaf_topology(&input.bytes, op)?);
            black_box(host_topology(&input.bytes, op)?);
        }
    }
    Ok(())
}
