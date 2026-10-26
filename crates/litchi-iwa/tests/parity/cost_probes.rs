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

#[inline(never)]
fn host_text_style(source: &[u8], cells: &[(u32, u32)]) -> TestResult<Vec<u8>> {
    use litchi_iwa::text::{TextPointSize, TextStyle};
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = editor
        .package()
        .iwa_entry_names()
        .find_map(|name| {
            editor
                .package()
                .archive(name)
                .ok()?
                .objects
                .into_iter()
                .find_map(|object| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == 6_001)
                        .then_some(object.archive_info.identifier)
                        .flatten()
                })
        })
        .ok_or("no table model")?;
    let style = TextStyle::new(TextPointSize::from_points(18.0)?).with_bold(true);
    for (row, column) in cells {
        editor.set_table_cell_text_style(id, *row as usize, *column as usize, style)?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_text_style(source: &[u8], cells: &[(u32, u32)]) -> TestResult<Vec<u8>> {
    use litchi_numbers::cell::text_style::CharacterProperties;
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut properties = CharacterProperties::new();
    properties.size = Some(18.0);
    properties.bold = Some(true);
    let mut edit = package.edit_table_cell_text_style(0usize, 0usize)?;
    for (row, column) in cells {
        edit = edit.set(
            litchi_numbers::table::At::new(*row, *column),
            properties.clone(),
        )?;
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Whole-cell text style on `basic.numbers`: one cell, then two.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_text_style_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    for cells in [&[(1, 1)][..], &[(1, 1), (2, 1)][..]] {
        black_box(host_text_style(&input.bytes, cells)?);
        black_box(leaf_text_style(&input.bytes, cells)?);
        black_box(leaf_text_style(&input.bytes, cells)?);
        black_box(host_text_style(&input.bytes, cells)?);
    }
    Ok(())
}

#[inline(never)]
fn host_cell_text(source: &[u8], reset: bool) -> TestResult<Vec<u8>> {
    use litchi_iwa::text::Alignment;
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = editor
        .package()
        .iwa_entry_names()
        .find_map(|name| {
            editor
                .package()
                .archive(name)
                .ok()?
                .objects
                .into_iter()
                .find_map(|object| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == 6_001)
                        .then_some(object.archive_info.identifier)
                        .flatten()
                })
        })
        .ok_or("no table model")?;
    if reset {
        editor.reset_table_cell_text_alignment(id, 1, 1)?;
    } else {
        editor.set_table_cell_text_alignment(id, 1, 1, Alignment::Justified)?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_cell_text(source: &[u8], reset: bool) -> TestResult<Vec<u8>> {
    use litchi_iwa::text::Alignment;
    use litchi_numbers::cell::text::{Property, PropertyKind};
    let package = litchi_numbers::Package::from_bytes(source)?;
    let at = litchi_numbers::table::At::new(1, 1);
    let edit = package.edit_table_cell_text(0usize, 0usize)?;
    let edit = if reset {
        edit.reset(at, PropertyKind::Alignment)?
    } else {
        edit.set(at, Property::Alignment(Alignment::Justified))?
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Cell paragraph alignment on `basic.numbers` B2: set, then reset on the
/// host's output.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_cell_text_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let set = host_cell_text(&input.bytes, false)?;
    for (bytes, reset) in [(&input.bytes, false), (&set, true)] {
        black_box(host_cell_text(bytes, reset)?);
        black_box(leaf_cell_text(bytes, reset)?);
        black_box(leaf_cell_text(bytes, reset)?);
        black_box(host_cell_text(bytes, reset)?);
    }
    Ok(())
}

#[inline(never)]
fn host_cell_list(source: &[u8], numbered: bool) -> TestResult<Vec<u8>> {
    use litchi_iwa_text::paragraph::list::ParagraphList;
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let list = if numbered {
        ParagraphList::Numbered
    } else {
        ParagraphList::Bullet
    };
    editor.set_table_cell_paragraph_list(id, 1, 1, list)?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_cell_list(source: &[u8], numbered: bool) -> TestResult<Vec<u8>> {
    use litchi_numbers::cell::list::ParagraphList;
    let package = litchi_numbers::Package::from_bytes(source)?;
    let list = if numbered {
        ParagraphList::Numbered
    } else {
        ParagraphList::Bullet
    };
    let commit = package
        .edit_table_cell_lists(0usize, 0usize)?
        .set_list(litchi_numbers::table::At::new(1, 1), list)?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// Cell paragraph list on `basic.numbers` B2: a bullet list on the plain
/// cell (promotion), then a numbered list on the host's output.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_cell_list_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let listed = host_cell_list(&input.bytes, false)?;
    for (bytes, numbered) in [(&input.bytes, false), (&listed, true)] {
        black_box(host_cell_list(bytes, numbered)?);
        black_box(leaf_cell_list(bytes, numbered)?);
        black_box(leaf_cell_list(bytes, numbered)?);
        black_box(host_cell_list(bytes, numbered)?);
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum Remove {
    Row(litchi_numbers::table::topology::RowDeletion),
    Column(litchi_numbers::table::topology::ColumnDeletion),
}

#[inline(never)]
fn host_removal(source: &[u8], op: Remove) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = litchi_numbers::TableSelector::index(0);
    match op {
        Remove::Row(at) => editor.remove_table_row(table, at)?,
        Remove::Column(at) => editor.remove_table_column(table, at)?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_removal(source: &[u8], op: Remove) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let edit = package.edit_table_topology(0usize, 0usize)?;
    let edit = match op {
        Remove::Row(at) => edit.remove_row(at)?,
        Remove::Column(at) => edit.remove_column(at)?,
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Row and column removal on `basic.numbers` and a native fixture.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_topology_remove_cost_abba() -> TestResult {
    use litchi_numbers::table::topology::{ColumnDeletion, RowDeletion};
    let inputs = [
        fixture("iwork/numbers/basic.numbers")?,
        fixture("iwork/numbers/custom-number-native.numbers")?,
    ];
    let ops = [
        Remove::Row(RowDeletion::body(0)),
        Remove::Column(ColumnDeletion::body(0)),
    ];
    for input in &inputs {
        for op in ops {
            black_box(host_removal(&input.bytes, op)?);
            black_box(leaf_removal(&input.bytes, op)?);
            black_box(leaf_removal(&input.bytes, op)?);
            black_box(host_removal(&input.bytes, op)?);
        }
    }
    Ok(())
}
