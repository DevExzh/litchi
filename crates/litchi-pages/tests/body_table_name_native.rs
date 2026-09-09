//! Native and source-built regression coverage for body-table name ownership.
//!
//! These fixtures exercise the focused name owner against the two native
//! topology profiles currently retained by the Pages test corpus and the
//! compact source-built table.  The metadata mutations below intentionally
//! remove one unrelated UUID registration or the selected model registration;
//! they verify that partial native registries remain readable while selected
//! ownership remains mandatory.

use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::fs;
use std::path::Path;

use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::tsp;
use litchi_pages::table::dimension::Dimension;
use litchi_pages::{
    BodyTableDimensionError, BodyTableNameError, BodyTableNameLimitKind, BodyTableSelector, Package,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const VISIBLE_TABLE_COMPONENT: &str = "Index/CalculationEngine-1732611.iwa";
const SOURCE_TABLE_COMPONENT: &str = "Index/Document.iwa";
const SAVED_TABLE_COMPONENT: &str = "Index/CalculationEngine-176.iwa";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const NATIVE_VISIBLE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-name-visible-native.pages"
));
const SOURCE_BUILT_BEFORE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-before.pages"
));
const NATIVE_SAVED: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-native-saved.pages"
));
const FOCUSED_RUST: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-name-focused.pages"
));
const FOCUSED_NATIVE_SAVED: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-name-focused-native-saved.pages"
));

const FOCUSED_NAME: &str = "Native renamed table";
const BODY_MARKER: &str = "Pages hidden-axis native oracle — 2026-09-05";

#[derive(Clone, Copy)]
struct Fixture {
    label: &'static str,
    source: &'static [u8],
    expected_name: &'static str,
    renamed_name: &'static str,
    selected_model_identifier: u64,
    selected_component: &'static str,
}

const FIXTURES: [Fixture; 3] = [
    Fixture {
        label: "native-visible",
        source: NATIVE_VISIBLE,
        expected_name: "Table 1",
        renamed_name: "Native renamed table",
        selected_model_identifier: 1_733_258,
        selected_component: VISIBLE_TABLE_COMPONENT,
    },
    Fixture {
        label: "source-built-before",
        source: SOURCE_BUILT_BEFORE,
        expected_name: "Cities",
        renamed_name: "Source renamed table",
        selected_model_identifier: 10,
        selected_component: SOURCE_TABLE_COMPONENT,
    },
    Fixture {
        label: "native-saved",
        source: NATIVE_SAVED,
        expected_name: "Cities",
        renamed_name: "Source renamed table",
        selected_model_identifier: 152,
        selected_component: SAVED_TABLE_COMPONENT,
    },
];

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_focused_body_and_dimensions(package: &Package, label: &str) -> TestResult {
    let text = package.text()?;
    assert!(
        text.contains(BODY_MARKER),
        "{label}: the native body marker was lost"
    );

    // The retained UI receipt records a five-by-four table.  The public
    // dimension reader exposes the checked axis bounds even when the focused
    // name owner deliberately leaves the table payload untouched.
    for row in 0..5 {
        package
            .body_table_dimension_size(BodyTableSelector::index(0), Dimension::Row(row))
            .map_err(|error| format!("{label}: row {row} dimension read failed: {error:?}"))?;
    }
    for column in 0..4 {
        package
            .body_table_dimension_size(BodyTableSelector::index(0), Dimension::Column(column))
            .map_err(|error| {
                format!("{label}: column {column} dimension read failed: {error:?}")
            })?;
    }
    assert!(matches!(
        package.body_table_dimension_size(BodyTableSelector::index(0), Dimension::Row(5)),
        Err(BodyTableDimensionError::InvalidSource)
    ));
    assert!(matches!(
        package.body_table_dimension_size(BodyTableSelector::index(0), Dimension::Column(4)),
        Err(BodyTableDimensionError::InvalidSource)
    ));
    Ok(())
}

fn assert_focused_name_readback(source: &[u8], label: &str) -> TestResult {
    let package = Package::from_bytes(source)?;
    assert_eq!(
        package
            .body_table_name(BodyTableSelector::index(0))?
            .as_str(),
        FOCUSED_NAME,
        "{label}: index selector"
    );
    assert_eq!(
        package
            .body_table_name(BodyTableSelector::name(FOCUSED_NAME))?
            .as_str(),
        FOCUSED_NAME,
        "{label}: name selector"
    );
    assert_focused_body_and_dimensions(&package, label)?;
    assert_source_roundtrip(&package, source, label)?;

    let no_op = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name(FOCUSED_NAME)?
        .commit()?;
    assert!(no_op.patch().is_noop(), "{label}: no-op patch");
    assert_eq!(
        exact_bytes(no_op.package())?,
        source,
        "{label}: no-op bytes"
    );

    let reopened = Package::from_bytes(&exact_bytes(&package)?)?;
    assert_eq!(
        reopened
            .body_table_name(BodyTableSelector::name(FOCUSED_NAME))?
            .as_str(),
        FOCUSED_NAME,
        "{label}: reopened name selector"
    );
    assert_focused_body_and_dimensions(&reopened, label)?;
    Ok(())
}

fn member_bytes(source: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

fn assert_source_roundtrip(package: &Package, source: &[u8], label: &str) -> TestResult {
    let bytes = exact_bytes(package)?;
    assert_eq!(bytes, source, "{label}: exact source changed during read");
    Ok(())
}

fn assert_locality(
    source: &[u8],
    target: &[u8],
    selected_component: &str,
    label: &str,
) -> TestResult {
    let before = member_bytes(source)?;
    let after = member_bytes(target)?;
    assert_ne!(
        before.get(selected_component),
        after.get(selected_component),
        "{label}: selected table component did not change"
    );
    for (name, payload) in &before {
        if name == selected_component || PREVIEWS.contains(&name.as_str()) {
            continue;
        }
        assert_eq!(
            after.get(name),
            Some(payload),
            "{label}: unrelated member {name} changed"
        );
    }
    for name in PREVIEWS {
        if before.contains_key(name) {
            assert!(
                !after.contains_key(name),
                "{label}: preview {name} survived"
            );
        }
    }
    for name in after.keys() {
        assert!(
            before.contains_key(name) || name == selected_component,
            "{label}: unexpected new member {name}"
        );
    }
    Ok(())
}

fn maybe_export(label: &str, bytes: &[u8]) -> TestResult {
    let Some(directory) = std::env::var_os("LITCHI_PAGES_NAME_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    fs::write(directory.join(format!("{label}-renamed.pages")), bytes)?;
    Ok(())
}

fn rewrite_metadata(
    source: &[u8],
    mutate: impl FnOnce(&mut tsp::PackageMetadata) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or("missing Metadata.iwa")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == METADATA_MESSAGE_TYPE)
        })
        .ok_or("missing PackageMetadata object")?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or("missing PackageMetadata message")?;
    let mut metadata = tsp::PackageMetadata::decode(object.messages[index].data.as_slice())?;
    mutate(&mut metadata)?;
    object.replace_message_preserving_header(
        index,
        RawMessage {
            type_: METADATA_MESSAGE_TYPE,
            data: metadata.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(METADATA_MEMBER, compressed.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn remove_uuid_registration(source: &[u8], identifier: u64) -> TestResult<(Vec<u8>, bool)> {
    let mut removed = false;
    let bytes = rewrite_metadata(source, |metadata| {
        for component in &mut metadata.components {
            component.object_uuid_map_entries.retain(|entry| {
                let keep = entry.identifier != identifier;
                if !keep {
                    removed = true;
                }
                keep
            });
        }
        Ok(())
    })?;
    Ok((bytes, removed))
}

fn first_unrelated_uuid(source: &[u8], selected: u64) -> TestResult<u64> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or("missing Metadata.iwa")?;
    let archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .objects
        .iter()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == METADATA_MESSAGE_TYPE)
        })
        .ok_or("missing PackageMetadata object")?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or("missing PackageMetadata message")?;
    let metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
    metadata
        .components
        .iter()
        .flat_map(|component| component.object_uuid_map_entries.iter())
        .map(|entry| entry.identifier)
        .find(|identifier| *identifier != selected)
        .ok_or_else(|| "missing unrelated UUID registration".into())
}

fn next_component_identifier(metadata: &tsp::PackageMetadata) -> TestResult<u64> {
    metadata
        .components
        .iter()
        .chain(metadata.versioned_components.iter())
        .map(|component| component.identifier)
        .max()
        .and_then(|identifier| identifier.checked_add(1))
        .ok_or_else(|| "component identifier space is exhausted".into())
}

fn assert_exact_transactions(fixture: Fixture) -> TestResult {
    let package = Package::from_bytes(fixture.source)?;
    let index_name = package
        .body_table_name(BodyTableSelector::index(0))
        .map_err(|error| format!("{}: index read failed: {error:?}", fixture.label))?;
    assert_eq!(
        index_name.as_str(),
        fixture.expected_name,
        "{}: index read",
        fixture.label
    );
    let named_name = package
        .body_table_name(BodyTableSelector::name(fixture.expected_name))
        .map_err(|error| format!("{}: name read failed: {error:?}", fixture.label))?;
    assert_eq!(
        named_name.as_str(),
        fixture.expected_name,
        "{}: name read",
        fixture.label
    );
    assert_source_roundtrip(&package, fixture.source, fixture.label)?;

    let noop = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name(fixture.expected_name)?
        .commit()?;
    assert!(noop.patch().is_noop(), "{}: no-op patch", fixture.label);
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_source_roundtrip(noop.package(), fixture.source, fixture.label)?;

    let changed = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name(fixture.renamed_name)?
        .commit()?;
    assert!(
        changed.diagnostics().changed(),
        "{}: changed",
        fixture.label
    );
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(changed.patch().before().as_str(), fixture.expected_name);
    assert_eq!(changed.patch().after().as_str(), fixture.renamed_name);
    assert_eq!(
        changed
            .package()
            .body_table_name(BodyTableSelector::index(0))?
            .as_str(),
        fixture.renamed_name,
        "{}: changed read",
        fixture.label
    );
    assert_eq!(
        changed
            .package()
            .body_table_name(BodyTableSelector::name(fixture.renamed_name))?
            .as_str(),
        fixture.renamed_name,
        "{}: changed name selector",
        fixture.label
    );
    let changed_bytes = exact_bytes(changed.package())?;
    assert_locality(
        fixture.source,
        &changed_bytes,
        fixture.selected_component,
        fixture.label,
    )?;
    maybe_export(fixture.label, &changed_bytes)?;

    let reopened = Package::from_bytes(&changed_bytes)?;
    assert_eq!(
        reopened
            .body_table_name(BodyTableSelector::index(0))?
            .as_str(),
        fixture.renamed_name,
        "{}: reopened read",
        fixture.label
    );
    let restored = changed
        .package()
        .apply_body_table_name(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, fixture.source);
    assert_eq!(
        restored
            .package()
            .body_table_name(BodyTableSelector::index(0))?
            .as_str(),
        fixture.expected_name,
        "{}: inverse read",
        fixture.label
    );
    Ok(())
}

#[test]
fn native_visible_table_name_supports_exact_transactions() -> TestResult {
    assert_exact_transactions(FIXTURES[0])
}

#[test]
fn source_built_table_name_supports_exact_transactions() -> TestResult {
    assert_exact_transactions(FIXTURES[1])
}

#[test]
fn native_saved_table_name_supports_exact_transactions() -> TestResult {
    assert_exact_transactions(FIXTURES[2])
}

#[test]
fn focused_rust_candidate_matches_regenerated_visible_rename() -> TestResult {
    let package = Package::from_bytes(NATIVE_VISIBLE)?;
    let changed = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name(FOCUSED_NAME)?
        .commit()?;
    assert_eq!(
        exact_bytes(changed.package())?,
        FOCUSED_RUST,
        "the retained focused candidate must be reproducible from the visible baseline"
    );
    Ok(())
}

#[test]
fn focused_rust_candidate_reads_by_index_and_name_with_exact_roundtrip() -> TestResult {
    assert_focused_name_readback(FOCUSED_RUST, "focused-rust")
}

#[test]
fn focused_native_saved_candidate_reads_by_index_and_name_with_exact_roundtrip() -> TestResult {
    assert_focused_name_readback(FOCUSED_NATIVE_SAVED, "focused-native-saved")
}

fn assert_unrelated_uuid_omission(fixture: Fixture) -> TestResult {
    let unrelated = first_unrelated_uuid(fixture.source, fixture.selected_model_identifier)?;
    let (mutated, removed) = remove_uuid_registration(fixture.source, unrelated)?;
    assert!(
        removed,
        "{}: unrelated registration was not found",
        fixture.label
    );
    let package = Package::from_bytes(&mutated)?;
    let name = package
        .body_table_name(BodyTableSelector::index(0))
        .map_err(|error| {
            format!(
                "{}: unrelated UUID omission blocked read: {error:?}",
                fixture.label
            )
        })?;
    assert_eq!(
        name.as_str(),
        fixture.expected_name,
        "{}: unrelated UUID omission blocked read",
        fixture.label
    );
    assert_eq!(exact_bytes(&package)?, mutated);
    Ok(())
}

#[test]
fn native_visible_unrelated_uuid_omission_remains_readable() -> TestResult {
    assert_unrelated_uuid_omission(FIXTURES[0])
}

#[test]
fn source_built_unrelated_uuid_omission_remains_readable() -> TestResult {
    assert_unrelated_uuid_omission(FIXTURES[1])
}

#[test]
fn native_saved_unrelated_uuid_omission_remains_readable() -> TestResult {
    assert_unrelated_uuid_omission(FIXTURES[2])
}

#[test]
fn selected_uuid_omission_is_rejected_atomically() -> TestResult {
    for fixture in FIXTURES {
        let (mutated, removed) =
            remove_uuid_registration(fixture.source, fixture.selected_model_identifier)?;
        assert!(
            removed,
            "{}: selected registration was not found",
            fixture.label
        );
        let package = Package::from_bytes(&mutated)?;
        let before = exact_bytes(&package)?;
        assert!(matches!(
            package.body_table_name(BodyTableSelector::index(0)),
            Err(BodyTableNameError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, before);
        assert_eq!(
            before, mutated,
            "{}: rejected source changed",
            fixture.label
        );
    }
    Ok(())
}

#[test]
fn source_name_rename_budget_failure_is_typed_and_atomic() -> TestResult {
    let fixture = FIXTURES[1];
    let limits = ArchiveLimits::new(
        u64::try_from(fixture.source.len())?,
        128,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let package = Package::from_bytes_with_limits(fixture.source, limits)?;
    let before = exact_bytes(&package)?;
    let replacement = "x".repeat(4_096);
    let error = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name(&replacement)?
        .commit()
        .expect_err("the exact input ceiling must reject output growth");
    assert!(matches!(
        error,
        BodyTableNameError::LimitExceeded {
            kind: BodyTableNameLimitKind::OutputBytes,
            ..
        }
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn selected_uuid_foreign_component_remains_rejected() -> TestResult {
    let fixture = FIXTURES[1];
    let mutated = rewrite_metadata(fixture.source, |metadata| {
        let component = metadata
            .components
            .iter_mut()
            .find(|component| {
                component
                    .object_uuid_map_entries
                    .iter()
                    .any(|entry| entry.identifier == fixture.selected_model_identifier)
            })
            .ok_or("missing selected model component")?;
        let entry = component
            .object_uuid_map_entries
            .iter_mut()
            .find(|entry| entry.identifier == fixture.selected_model_identifier)
            .ok_or("missing selected model registration")?;
        entry.uuid = tsp::Uuid {
            lower: 0x1111,
            upper: 0x2222,
        };
        component.preferred_locator = "Foreign".to_owned();
        component.locator = Some("Foreign".to_owned());
        Ok(())
    })?;
    let package = Package::from_bytes(&mutated)?;
    assert!(matches!(
        package.body_table_name(BodyTableSelector::index(0)),
        Err(BodyTableNameError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, mutated);
    Ok(())
}

#[test]
fn native_duplicate_current_locator_alias_without_selected_binding_is_rejected_atomically()
-> TestResult {
    let fixture = FIXTURES[0];
    let mutated = rewrite_metadata(fixture.source, |metadata| {
        let selected_index = metadata
            .components
            .iter()
            .position(|component| {
                component
                    .object_uuid_map_entries
                    .iter()
                    .any(|entry| entry.identifier == fixture.selected_model_identifier)
            })
            .ok_or("missing selected model component")?;
        let mut alias = metadata.components[selected_index].clone();
        alias.identifier = next_component_identifier(metadata)?;
        alias.object_uuid_map_entries.clear();
        metadata.components.push(alias);
        Ok(())
    })?;
    let package = Package::from_bytes(&mutated)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.body_table_name(BodyTableSelector::index(0)),
        Err(BodyTableNameError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(before, mutated);
    Ok(())
}

#[test]
fn native_selected_binding_moved_to_same_locator_alias_is_rejected_atomically() -> TestResult {
    let fixture = FIXTURES[0];
    let mutated = rewrite_metadata(fixture.source, |metadata| {
        let selected_index = metadata
            .components
            .iter()
            .position(|component| {
                component
                    .object_uuid_map_entries
                    .iter()
                    .any(|entry| entry.identifier == fixture.selected_model_identifier)
            })
            .ok_or("missing selected model component")?;
        let mut alias = metadata.components[selected_index].clone();
        alias.identifier = next_component_identifier(metadata)?;
        let binding_index = metadata.components[selected_index]
            .object_uuid_map_entries
            .iter()
            .position(|entry| entry.identifier == fixture.selected_model_identifier)
            .ok_or("missing selected model registration")?;
        let binding = metadata.components[selected_index]
            .object_uuid_map_entries
            .remove(binding_index);
        alias.object_uuid_map_entries.push(binding);
        metadata.components.push(alias);
        Ok(())
    })?;
    let package = Package::from_bytes(&mutated)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.body_table_name(BodyTableSelector::index(0)),
        Err(BodyTableNameError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(before, mutated);
    Ok(())
}
