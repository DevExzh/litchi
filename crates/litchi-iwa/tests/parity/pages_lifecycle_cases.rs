//! Pages section and body-table lifecycle parity cases of wave W1-B
//! (record 0371): the host's `{insert,append,remove}_section`, `sections`,
//! `add_table` and `remove_table` against the leaf `section::lifecycle`,
//! `table::creation` and `table::deletion` families.

use litchi_core::Position;
use litchi_iwa::pages;
use litchi_pages::table::creation::NewTable;
use litchi_pages::{BodyTableSelector, Package, SectionSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const SECTIONS: &str = "iwork/pages/body-sections-unicode.pages";
const TABLES: &str = "iwork/pages/body-table-catalog-native.pages";

fn host_sections(bytes: &[u8]) -> TestResult<Vec<(u32, Option<String>)>> {
    Ok(pages::Draft::from_bytes(bytes)?
        .sections()
        .iter()
        .map(|section| (section.character_index, section.name.clone()))
        .collect())
}

fn leaf_sections(bytes: &[u8]) -> TestResult<Vec<Option<String>>> {
    Ok(Package::from_bytes(bytes)?
        .sections()
        .iter()
        .map(|section| section.name().map(str::to_owned))
        .collect())
}

fn body_text(bytes: &[u8]) -> TestResult<String> {
    Ok(pages::Draft::from_bytes(bytes)?.body_text()?)
}

fn section_case(name: &'static str, tier: Tier, evidence: Evidence, remove: bool) -> TestResult {
    let input = fixture(SECTIONS)?;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            if remove {
                let id = editor
                    .sections()
                    .get(1)
                    .ok_or("no second section")?
                    .object_id;
                drop(editor.remove_section(id)?);
            } else {
                let id = editor.sections().last().ok_or("no section")?.object_id;
                drop(editor.append_section(id, "Appendix")?);
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let edit = package.edit_section_lifecycle();
            let edit = if remove {
                edit.remove(SectionSelector::index(1))?
            } else {
                let last = package.sections().len() - 1;
                edit.add(SectionSelector::index(last), Some("Appendix"))?
            };
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_section_lifecycle(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host sections", host_sections)
        .read("leaf sections", leaf_sections)
        .read("host body text", body_text)
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

#[test]
fn pages_section_append_parity() -> TestResult {
    section_case(
        "pages_section_append_parity",
        Tier::B,
        Evidence::Queued,
        false,
    )
}

#[test]
fn pages_section_remove_parity() -> TestResult {
    section_case(
        "pages_section_remove_parity",
        Tier::B,
        Evidence::Queued,
        true,
    )
}

fn host_tables(bytes: &[u8]) -> TestResult<Vec<(String, usize, usize, usize)>> {
    Ok(pages::Draft::from_bytes(bytes)?
        .tables()?
        .into_iter()
        .map(|table| {
            (
                table.name,
                table.anchor_character_index,
                table.rows,
                table.columns,
            )
        })
        .collect())
}

fn leaf_tables(bytes: &[u8]) -> TestResult<Vec<Option<String>>> {
    let package = Package::from_bytes(bytes)?;
    let catalog = package.body_tables()?;
    Ok((0..catalog.len())
        .map(|index| catalog.get(index).map(|entry| entry.name().to_owned()))
        .collect())
}

/// Table creation is not at parity: the leaf `table::creation` family
/// (record 0329) writes a different graph from the host's `add_table` (a
/// random table UUID, its own identifier order, a formula owner of type
/// 6267 where the host writes 4008). `add_table` stays host-only (record
/// 0371).
#[test]
fn pages_body_table_add_divergence_is_reported() -> TestResult {
    let input = fixture(TABLES)?;
    let outcome = Case::new(
        "pages_body_table_add_divergence_is_reported",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        drop(editor.add_table(0, "Litchi table", 3, 4)?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let commit = package
            .edit_body_table_creation()
            .insert(
                Position::new(0),
                NewTable::new(3, 4).with_name("Litchi table"),
            )?
            .commit()?;
        let restored = commit
            .package()
            .apply_body_table_creation(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host tables", host_tables)
    .read("leaf body_tables", leaf_tables)
    .run();
    let Err(error) = outcome else {
        return Err("body-table creation is now at parity; migrate add_table".into());
    };
    let report = error.to_string();
    assert!(report.contains("leaf [(6267"), "{report}");
    Ok(())
}

#[test]
fn pages_body_table_remove_parity() -> TestResult {
    let input = fixture(TABLES)?;
    Case::new(
        "pages_body_table_remove_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = editor
            .tables()?
            .get(1)
            .ok_or("no second table")?
            .model_object_id;
        drop(editor.remove_table(model)?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let commit = package
            .edit_body_table_deletion()
            .remove(BodyTableSelector::index(1))?
            .commit()?;
        let restored = commit
            .package()
            .apply_body_table_deletion(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host tables", host_tables)
    .read("leaf body_tables", leaf_tables)
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn section_append_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(SECTIONS)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor.sections().last().ok_or("no section")?.object_id;
            drop(editor.append_section(id, "Appendix")?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let last = package.sections().len() - 1;
            let commit = package
                .edit_section_lifecycle()
                .add(SectionSelector::index(last), Some("Appendix"))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn section_remove_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(SECTIONS)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor
                .sections()
                .get(1)
                .ok_or("no second section")?
                .object_id;
            drop(editor.remove_section(id)?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_section_lifecycle()
                .remove(SectionSelector::index(1))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_remove_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLES)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor
                .tables()?
                .get(1)
                .ok_or("no second table")?
                .model_object_id;
            drop(editor.remove_table(model)?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_deletion()
                .remove(BodyTableSelector::index(1))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// `insert_section` against the leaf section split: section 0 is copied and
/// the body breaks inside the first section's text.
#[test]
fn pages_section_insert_parity() -> TestResult {
    let input = fixture(SECTIONS)?;
    Case::new(
        "pages_section_insert_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let id = editor.sections().first().ok_or("no section")?.object_id;
        drop(editor.insert_section(id, 3, "Split")?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let commit = package
            .edit_section_lifecycle()
            .insert(
                SectionSelector::index(0),
                litchi_pages::TextPosition::from_utf16_index(3)?,
                Some("Split"),
            )?
            .commit()?;
        let restored = commit
            .package()
            .apply_section_lifecycle(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host sections", host_sections)
    .read("leaf sections", leaf_sections)
    .read("host body text", body_text)
    .leaf_structural_metadata()
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn section_insert_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(SECTIONS)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor.sections().first().ok_or("no section")?.object_id;
            drop(editor.insert_section(id, 3, "Split")?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_section_lifecycle()
                .insert(
                    SectionSelector::index(0),
                    litchi_pages::TextPosition::from_utf16_index(3)?,
                    Some("Split"),
                )?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// `body-table-catalog-native.pages` table 0 with column 0 numbers in
/// descending and column 1 text, and a persisted sort order (set through the
/// leaf sort-order family) of `scope` on column `column`.
fn sortable_fixture(
    scope: litchi_iwa_common::table::sort::Scope,
    path: &'static str,
    column: usize,
    direction: litchi_iwa_common::table::sort::Direction,
) -> TestResult<crate::harness::Fixture> {
    use litchi_iwa_common::table::sort::{ColumnIndex, Order, Rule};
    use litchi_pages::table::values::Value;
    let input = fixture(path)?;
    let package = Package::from_bytes(&input.bytes)?;
    let mut edit = package.edit_body_table_values(0)?;
    for (row, number, text) in [
        (0, 3.0, "kiwi"),
        (1, 7.5, "pear"),
        (2, -2.0, "apple"),
        (3, 42.0, "fig"),
        (4, 7.5, "date"),
    ] {
        edit.set(row, 0, Value::Number(number))?;
        edit.set(row, 1, Value::Text(text.to_owned()))?;
    }
    let filled = edit.commit()?;
    let order = Order::with_scope(scope, [Rule::new(ColumnIndex::new(column)?, direction)])?;
    let sorted = filled
        .package()
        .edit_body_table_sort_order(BodyTableSelector::index(0))?
        .set(order)
        .commit()?;
    Ok(crate::harness::Fixture {
        path: "native body table 0 with keys and a sort order",
        bytes: leaf_bytes!(sorted.package()),
    })
}

fn sort_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &crate::harness::Fixture,
    rows: Option<litchi_iwa_common::table::sort::RowRange>,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor.tables()?[0].model_object_id;
            let _moved = match rows {
                None => editor.apply_table_sort_order(model)?,
                Some(rows) => editor.apply_table_sort_order_to_rows(model, rows)?,
            };
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_sort_execution(0)?;
            match rows {
                None => drop(edit.apply()?),
                Some(rows) => drop(edit.apply_to_rows(rows)?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_table_sort_execution(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host tables", host_tables)
        .read("leaf body_table_cells", |bytes: &[u8]| {
            Ok(format!(
                "{:?}",
                Package::from_bytes(bytes)?.body_table_cells(0)?
            ))
        })
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_sort_apply_parity() -> TestResult {
    use litchi_iwa_common::table::sort::{Direction, Scope};
    for (column, direction) in [(0, Direction::Ascending), (1, Direction::Descending)] {
        sort_case(
            "pages_body_table_sort_apply_parity",
            Tier::A,
            Evidence::Queued,
            &sortable_fixture(Scope::EntireTable, TABLES, column, direction)?,
            None,
        )?;
    }
    // Explicit cell borders: a row layer moves with its row and a column
    // layer's runs split with the rows.
    let base = sortable_fixture(Scope::EntireTable, TABLES, 0, Direction::Ascending)?;
    let mut editor = pages::Draft::from_bytes(&base.bytes)?;
    let model = editor.tables()?[0].model_object_id;
    let stroke = litchi_iwa_common::shape::stroke::Stroke::new(
        litchi_iwa_common::color::Rgba::new(
            0.2,
            0.4,
            0.6,
            1.0,
            litchi_iwa_common::color::RgbColorSpace::Srgb,
        )?,
        litchi_iwa_common::shape::stroke::Width::new(2.0)?,
        litchi_iwa_common::shape::stroke::Pattern::Solid,
    );
    for (row, column, side) in [
        (1, 0, litchi_iwa_common::table::cell::BorderSide::Top),
        (2, 1, litchi_iwa_common::table::cell::BorderSide::Left),
        (3, 1, litchi_iwa_common::table::cell::BorderSide::Left),
        (4, 2, litchi_iwa_common::table::cell::BorderSide::Bottom),
    ] {
        editor.set_table_cell_border(model, row, column, side, stroke)?;
    }
    let bordered = crate::harness::Fixture {
        path: "native body table 0 with keys, a sort order and host-set cell borders",
        bytes: editor.to_bytes()?,
    };
    sort_case(
        "pages_body_table_sort_apply_parity",
        Tier::A,
        Evidence::Queued,
        &bordered,
        None,
    )?;
    // A table with user-hidden rows and columns: their positions stay.
    sort_case(
        "pages_body_table_sort_apply_parity",
        Tier::A,
        Evidence::Queued,
        &sortable_fixture(
            Scope::EntireTable,
            "iwork/pages/body-table-hidden-axes-native.pages",
            0,
            Direction::Descending,
        )?,
        None,
    )?;
    Ok(())
}

#[test]
fn pages_body_table_sort_apply_to_rows_parity() -> TestResult {
    use litchi_iwa_common::table::sort::{Direction, RowRange, Scope};
    sort_case(
        "pages_body_table_sort_apply_to_rows_parity",
        Tier::A,
        Evidence::Queued,
        &sortable_fixture(Scope::SelectedRows, TABLES, 0, Direction::Descending)?,
        Some(RowRange::new(1, 4)?),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_sort_apply_cost_abba() -> TestResult {
    use litchi_iwa_common::table::sort::{Direction, Scope};
    crate::pages_cost_probes::abba(
        &sortable_fixture(Scope::EntireTable, TABLES, 0, Direction::Ascending)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor.tables()?[0].model_object_id;
            let _moved = editor.apply_table_sort_order(model)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_sort_execution(0)?;
            edit.apply()?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
