//! Numbers builder options parity (record 0370i): the host builder's
//! `table_dimensions`, `language` and `locale` against the leaf
//! `litchi_numbers::Builder`.
//!
//! The host builder draws a random document identity, table UUID and
//! row/column UIDs. The leaf build takes the host output's identity, so
//! the document identifier, properties and package metadata compare
//! exactly; the table UUID (with the formula-owner UUID and owner-map entry
//! derived from it) and the row/column UID map are compared by form, a
//! documented policy.

use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::{Decoded, decoded};
use crate::leaf_bytes;

/// A hyphenated UUID string as its 16 bytes.
fn uuid_bytes(text: &str) -> TestResult<[u8; 16]> {
    let hex: String = text.chars().filter(|character| *character != '-').collect();
    let value = u128::from_str_radix(&hex, 16)?;
    Ok(value.to_be_bytes())
}

/// The identity of a built package: its document, version and private UUIDs.
fn identity(bytes: &[u8]) -> TestResult<litchi_numbers::Identity> {
    let properties = crate::archive_diff::members(bytes)?
        .into_iter()
        .find(|member| member.name == "Metadata/Properties.plist")
        .ok_or("no properties")?
        .data;
    let value = plist::Value::from_reader(std::io::Cursor::new(properties))?;
    let dictionary = value.as_dictionary().ok_or("properties")?;
    let field = |key: &str| -> TestResult<[u8; 16]> {
        uuid_bytes(
            dictionary
                .get(key)
                .and_then(plist::Value::as_string)
                .ok_or("missing UUID")?,
        )
    };
    Ok(litchi_numbers::Identity::from_bytes(
        field("documentUUID")?,
        field("versionUUID")?,
        field("privateUUID")?,
    )?)
}

/// `TST.TableModelArchive` without its table UUID.
pub(crate) fn model_without_table_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut model = litchi_iwa_protos::tst::TableModelArchive::decode(data)?;
    assert_eq!(model.table_id.len(), 36, "table UUID form");
    model.table_id.clear();
    Ok(decoded(model, data))
}

/// `TSCE.FormulaOwnerDependenciesArchive` without its owner UUID.
pub(crate) fn owner_without_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut owner = litchi_iwa_protos::tsce::FormulaOwnerDependenciesArchive::decode(data)?;
    owner.formula_owner_uid = Default::default();
    Ok(decoded(owner, data))
}

/// `TSCE.CalculationEngineArchive` without its owner map's UUIDs.
pub(crate) fn engine_without_owner_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut engine = litchi_iwa_protos::tsce::CalculationEngineArchive::decode(data)?;
    if let Some(map) = engine.dependency_tracker.owner_id_map.as_mut() {
        for entry in &mut map.map_entry {
            entry.owner_id = Default::default();
        }
    }
    Ok(decoded(engine, data))
}

/// `TST.ColumnRowUIDMapArchive` by form: its axis lengths, each index list
/// a permutation, and the sorted lists sorted.
pub(crate) fn uid_map_by_form(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let map = litchi_iwa_protos::tst::ColumnRowUidMapArchive::decode(data)?;
    let permutation = |values: &[u32], length: usize| {
        let mut sorted = values.to_vec();
        sorted.sort_unstable();
        sorted == (0..u32::try_from(length).unwrap_or(u32::MAX)).collect::<Vec<_>>()
    };
    let columns = map.sorted_column_uids.len();
    let rows = map.sorted_row_uids.len();
    assert!(permutation(&map.column_index_for_uid, columns));
    assert!(permutation(&map.column_uid_for_index, columns));
    assert!(permutation(&map.row_index_for_uid, rows));
    assert!(permutation(&map.row_uid_for_index, rows));
    let form = litchi_iwa_protos::tst::ColumnRowUidMapArchive {
        sorted_column_uids: vec![Default::default(); columns],
        sorted_row_uids: vec![Default::default(); rows],
        ..Default::default()
    };
    Ok(decoded(form, data))
}

/// `TSP.PackageMetadata` without its object UUIDs (the host draws one per
/// registered object at random; the leaf template carries fixed ones).
fn metadata_without_object_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut metadata = litchi_iwa_protos::tsp::PackageMetadata::decode(data)?;
    for component in metadata
        .components
        .iter_mut()
        .chain(metadata.versioned_components.iter_mut())
    {
        for entry in &mut component.object_uuid_map_entries {
            entry.uuid = Default::default();
        }
    }
    Ok(decoded(metadata, data))
}

/// The builder options one case sets.
#[derive(Clone, Copy, Default)]
struct Options<'a> {
    rows: Option<(usize, usize)>,
    language: Option<&'a str>,
    locale: Option<&'a str>,
    sheet: Option<&'a str>,
    table: Option<&'a str>,
}

fn run_builder(
    name: &'static str,
    rows: Option<(usize, usize)>,
    language: Option<&str>,
    locale: Option<&str>,
) -> TestResult {
    run_options(
        name,
        Options {
            rows,
            language,
            locale,
            ..Options::default()
        },
    )
}

fn run_options(name: &'static str, options: Options<'_>) -> TestResult {
    let Options {
        rows,
        language,
        locale,
        sheet,
        table,
    } = options;
    let mut host = litchi_iwa::numbers::Draft::builder();
    if let Some(sheet) = sheet {
        host = host.sheet_name(sheet);
    }
    if let Some(table) = table {
        host = host.table_name(table);
    }
    if let Some((rows, columns)) = rows {
        host = host.table_dimensions(rows, columns);
    }
    if let Some(language) = language {
        host = host.language(language);
    }
    if let Some(locale) = locale {
        host = host.locale(locale);
    }
    let host_bytes = host.build()?.to_bytes()?;
    let identity = identity(&host_bytes)?;
    let leaf = move |_source: &[u8]| -> TestResult<LeafRun> {
        let mut builder = litchi_numbers::Builder::new().identity(identity);
        if let Some(sheet) = sheet {
            builder = builder.sheet_name(sheet);
        }
        if let Some(table) = table {
            builder = builder.table_name(table);
        }
        if let Some((rows, columns)) = rows {
            builder = builder.table_dimensions(u32::try_from(rows)?, u32::try_from(columns)?);
        }
        if let Some(language) = language {
            builder = builder.language(language);
        }
        if let Some(locale) = locale {
            builder = builder.locale(locale);
        }
        Ok(LeafRun::output(leaf_bytes!(&builder.build()?)))
    };
    let input = Fixture {
        path: "builder",
        bytes: leaf_bytes!(&litchi_numbers::Package::new()),
    };
    let host_output = host_bytes.clone();
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, &input)
        .known_fields(|known| {
            known
                .with(
                    6_001,
                    "TST.TableModelArchive (table UUID by form)",
                    model_without_table_uuid,
                )
                .with(
                    4_008,
                    "TSCE.FormulaOwnerDependenciesArchive (owner UUID by form)",
                    owner_without_uuid,
                )
                .with(
                    4_000,
                    "TSCE.CalculationEngineArchive (owner-map UUIDs by form)",
                    engine_without_owner_uuids,
                )
                .with(
                    6_267,
                    "TST.ColumnRowUIDMapArchive (UIDs by form)",
                    uid_map_by_form,
                )
                .with(
                    11_006,
                    "TSP.PackageMetadata (object UUIDs by form)",
                    metadata_without_object_uuids,
                )
        })
        .host(move |_source| Ok(host_output.clone()))
        .leaf(leaf)
        .run()?;
    Ok(())
}

/// The builder's table dimensions.
#[test]
fn numbers_builder_table_dimensions_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for dimensions in [(10, 5), (20, 8), (3, 2), (2, 2), (40, 26)] {
        run_builder(
            "numbers_builder_table_dimensions_parity",
            Some(dimensions),
            None,
            None,
        )
        .map_err(|error| format!("{dimensions:?}: {error}"))?;
    }
    Ok(())
}

/// The builder's document language and locale.
#[test]
fn numbers_builder_language_locale_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (language, locale) in [("fr", "fr_FR"), ("ja", "ja_JP"), ("de-CH", "de_CH")] {
        run_builder(
            "numbers_builder_language_locale_parity",
            None,
            Some(language),
            Some(locale),
        )
        .map_err(|error| format!("{language}/{locale}: {error}"))?;
    }
    Ok(())
}

/// A table with no body row or column: the host builder writes one, the
/// leaf refuses it, as its resize does (record 0370i, a documented
/// difference: a native table keeps at least one body row and column).
#[test]
fn numbers_builder_headers_only_refusal() -> TestResult {
    assert!(
        litchi_iwa::numbers::Draft::builder()
            .table_dimensions(1, 1)
            .build()
            .is_ok()
    );
    let leaf = litchi_numbers::Builder::new()
        .identity(litchi_numbers::Identity::BLANK)
        .table_dimensions(1, 1)
        .build();
    assert!(matches!(
        leaf,
        Err(litchi_numbers::create::Error::Dimensions(_))
    ));
    Ok(())
}

/// The default build and the builder's sheet and table names.
#[test]
fn numbers_builder_names_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (sheet, table) in [
        (None, None),
        (Some("Budget"), None),
        (None, Some("Q1")),
        (Some("Résumé 2026"), Some("Données")),
    ] {
        run_options(
            "numbers_builder_names_parity",
            Options {
                sheet,
                table,
                ..Options::default()
            },
        )
        .map_err(|error| format!("{sheet:?}/{table:?}: {error}"))?;
    }
    Ok(())
}
