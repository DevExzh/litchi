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
