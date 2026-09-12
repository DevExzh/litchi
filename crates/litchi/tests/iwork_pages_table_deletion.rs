//! Public Pages table deletion parity and exact-source transactions.
#![cfg(feature = "pages")]

use litchi::pages::{BodyTableDeletionError, BodyTableSelector, Package};

type TestResult = Result<(), Box<dyn std::error::Error>>;

const FIXTURES: &[(&str, &[u8])] = &[
    (
        "catalog",
        include_bytes!("../../../test-data/iwork/pages/body-table-catalog-native.pages"),
    ),
    (
        "cells",
        include_bytes!("../../../test-data/iwork/pages/body-table-cells-native.pages"),
    ),
    (
        "merges",
        include_bytes!("../../../test-data/iwork/pages/body-table-merges-native.pages"),
    ),
    (
        "comments",
        include_bytes!(
            "../../../test-data/iwork/pages/body-table-comment-hidden-native-saved.pages"
        ),
    ),
    (
        "cross-reference",
        include_bytes!("../../../test-data/iwork/pages/body-table-cross-reference-native.pages"),
    ),
];

#[test]
fn pages_deletion_preserves_survivors_and_supports_inverse() -> TestResult {
    for &(name, bytes) in FIXTURES {
        let source = Package::from_bytes(bytes)?;
        let catalog = source.body_tables()?;
        let cells = (0..catalog.len())
            .map(|position| source.body_table_cells(BodyTableSelector::index(position)))
            .collect::<Result<Vec<_>, _>>()?;
        for position in 0..catalog.len() {
            if name == "cross-reference" && position == 0 {
                continue;
            }
            let committed = source
                .remove_body_table(BodyTableSelector::index(position))
                .map_err(|error| format!("{name}, deletion of table {position}: {error}"))?;
            assert_eq!(
                committed.removed_table(),
                catalog.get(position).expect("selected table")
            );
            let after = committed.package().body_tables()?;
            assert_eq!(after.len(), catalog.len() - 1, "{name}, table {position}");
            for (new_position, original_position) in (0..catalog.len())
                .filter(|candidate| *candidate != position)
                .enumerate()
            {
                assert_eq!(
                    after.get(new_position).expect("surviving table").name(),
                    catalog
                        .get(original_position)
                        .expect("original table")
                        .name()
                );
                assert_eq!(
                    committed
                        .package()
                        .body_table_cells(BodyTableSelector::index(new_position))?,
                    cells[original_position],
                    "{name}, survivor {original_position}"
                );
                assert_eq!(
                    committed
                        .package()
                        .body_table_merges(BodyTableSelector::index(new_position))?,
                    source.body_table_merges(BodyTableSelector::index(original_position))?,
                    "{name}, surviving merges {original_position}"
                );
                assert_eq!(
                    committed
                        .package()
                        .body_table_hidden_axes(BodyTableSelector::index(new_position)),
                    source.body_table_hidden_axes(BodyTableSelector::index(original_position)),
                    "{name}, surviving hidden axes {original_position}"
                );
            }
            assert!(
                committed
                    .package()
                    .apply_body_table_deletion(committed.patch())
                    .is_err()
            );
            let restored = committed
                .package()
                .apply_body_table_deletion(&committed.patch().inverse())?;
            let mut restored_bytes = Vec::new();
            restored.package().write_to(&mut restored_bytes)?;
            assert_eq!(restored_bytes, bytes, "{name}, exact inverse");
            let mut original_bytes = Vec::new();
            source.write_to(&mut original_bytes)?;
            assert_eq!(original_bytes, bytes, "{name}, immutable source");
        }
    }
    Ok(())
}

#[test]
fn pages_deletion_refuses_missing_tables_and_incoming_formula_dependencies() -> TestResult {
    let bytes = FIXTURES.last().expect("cross-reference fixture").1;
    let source = Package::from_bytes(bytes)?;
    assert!(matches!(
        source.remove_body_table(BodyTableSelector::index(0)),
        Err(BodyTableDeletionError::UnsupportedDependency)
    ));
    assert!(matches!(
        source.remove_body_table(BodyTableSelector::index(usize::MAX)),
        Err(BodyTableDeletionError::TableNotFound)
    ));
    let mut unchanged = Vec::new();
    source.write_to(&mut unchanged)?;
    assert_eq!(unchanged, bytes);
    Ok(())
}

#[test]
fn pages_deletion_resolves_names_and_can_remove_the_last_table() -> TestResult {
    let bytes = FIXTURES[0].1;
    let source = Package::from_bytes(bytes)?;
    let catalog = source.body_tables()?;
    let selected = catalog.get(1).expect("second native table");
    let first = source.remove_body_table(selected.name_selector())?;
    assert_eq!(first.removed_table(), selected);
    let second = first
        .package()
        .remove_body_table(BodyTableSelector::index(0))?;
    assert!(second.package().body_tables()?.is_empty());
    assert!(
        second
            .package()
            .remove_body_table(BodyTableSelector::index(0))
            .is_err()
    );

    let restored_first = second
        .package()
        .apply_body_table_deletion(&second.patch().inverse())?;
    let restored_source = restored_first
        .package()
        .apply_body_table_deletion(&first.patch().inverse())?;
    let mut restored = Vec::new();
    restored_source.package().write_to(&mut restored)?;
    assert_eq!(restored, bytes);
    Ok(())
}
