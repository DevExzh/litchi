//! Pages builder `language` and `locale` parity (wave W1-B, record 0371):
//! the host builder against `litchi_pages::Builder`.
//!
//! The host builder draws a random document identity and object UUIDs. The
//! leaf build takes the host output's document and private UUIDs, so the
//! identifier and properties compare exactly; the registry's object UUIDs
//! are compared by form, a documented policy.

use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::{Decoded, decoded};
use crate::leaf_bytes;

fn uuid_bytes(text: &str) -> TestResult<[u8; 16]> {
    let hex: String = text.chars().filter(|character| *character != '-').collect();
    Ok(u128::from_str_radix(&hex, 16)?.to_be_bytes())
}

/// The document and private UUIDs of a built package.
fn identity(bytes: &[u8]) -> TestResult<litchi_pages::builder::Identity> {
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
    Ok(litchi_pages::builder::Identity::from_parts(
        field("documentUUID")?,
        field("versionUUID")?,
        field("privateUUID")?,
    ))
}

/// `TSP.PackageMetadata` without its object UUIDs.
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

fn run_builder(name: &'static str, text: &str, language: &str, locale: &str) -> TestResult {
    let host_bytes = litchi_iwa::pages::Builder::new()
        .body_text(text)
        .language(language)
        .locale(locale)
        .build()?
        .to_bytes()?;
    let identity = identity(&host_bytes)?;
    let (text, language, locale) = (text.to_owned(), language.to_owned(), locale.to_owned());
    let leaf = move |_source: &[u8]| -> TestResult<LeafRun> {
        let package = litchi_pages::builder::Builder::new()
            .identity(identity)
            .body_text(text.clone())
            .language(language.clone())
            .locale(locale.clone())
            .build()?;
        Ok(LeafRun::output(leaf_bytes!(&package)))
    };
    let input = Fixture {
        path: "builder",
        bytes: leaf_bytes!(&litchi_pages::Package::new()?),
    };
    let host_output = host_bytes.clone();
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, &input)
        .known_fields(|known| {
            known.with(
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

/// The builder's document language and locale.
#[test]
fn pages_builder_language_locale_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (text, language, locale) in [
        ("", "fr", "fr_FR"),
        ("Bonjour\nMonde", "fr", "fr_FR"),
        ("日本語", "ja", "ja_JP"),
        ("Grüezi", "de-CH", "de_CH"),
    ] {
        run_builder(
            "pages_builder_language_locale_parity",
            text,
            language,
            locale,
        )
        .map_err(|error| format!("{language}/{locale}: {error}"))?;
    }
    Ok(())
}

/// Host `pages::Draft::{create, create_with_text}` (a default build, and
/// one with body text) against the leaf builder under the host output's
/// identity (record 0417).
#[test]
fn pages_builder_create_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for text in [None, Some("Bonjour"), Some("Line one\nLine two"), Some("")] {
        let host = match text {
            None => litchi_iwa::pages::Draft::create()?,
            Some(text) => litchi_iwa::pages::Draft::create_with_text(text)?,
        };
        let host_bytes = host.to_bytes()?;
        let identity = identity(&host_bytes)?;
        let leaf = move |_source: &[u8]| -> TestResult<LeafRun> {
            let mut builder = litchi_pages::builder::Builder::new().identity(identity);
            if let Some(text) = text {
                builder = builder.body_text(text);
            }
            Ok(LeafRun::output(leaf_bytes!(&builder.build()?)))
        };
        let input = Fixture {
            path: "builder",
            bytes: leaf_bytes!(&litchi_pages::Package::new()?),
        };
        Case::new(
            "pages_builder_create_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .known_fields(|known| {
            known.with(
                11_006,
                "TSP.PackageMetadata (object UUIDs by form)",
                metadata_without_object_uuids,
            )
        })
        .host(move |_source| Ok(host_bytes.clone()))
        .leaf(leaf)
        .run()
        .map_err(|error| format!("{text:?}: {error}"))?;
    }
    Ok(())
}

/// Host `pages::Builder::body_table` against the leaf builder's
/// from-scratch table graph (record 0420): one table after body text, and
/// two tables (the second cloned from the first by both sides). The host
/// draws the table, owner, UID-map and object UUIDs at random; they are
/// compared by form.
#[test]
fn pages_builder_body_table_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (text, tables) in [
        ("", vec![("Table 1", 3, 2)]),
        ("Budget", vec![("Q1", 5, 4)]),
        ("Two", vec![("First", 2, 2), ("Second", 4, 3)]),
    ] {
        let mut host = litchi_iwa::pages::Builder::new().body_text(text);
        for (name, rows, columns) in &tables {
            host = host.body_table(*name, *rows, *columns);
        }
        let host_bytes = host.build()?.to_bytes()?;
        let identity = identity(&host_bytes)?;
        let leaf_tables = tables.clone();
        let leaf = move |_source: &[u8]| -> TestResult<LeafRun> {
            let mut builder = litchi_pages::builder::Builder::new()
                .identity(identity)
                .body_text(text);
            for (name, rows, columns) in &leaf_tables {
                builder =
                    builder.body_table(*name, u32::try_from(*rows)?, u32::try_from(*columns)?);
            }
            Ok(LeafRun::output(leaf_bytes!(&builder.build()?)))
        };
        let input = Fixture {
            path: "builder",
            bytes: leaf_bytes!(&litchi_pages::Package::new()?),
        };
        Case::new(
            "pages_builder_body_table_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .known_fields(|known| {
            known
                .with(
                    11_006,
                    "TSP.PackageMetadata (object UUIDs by form)",
                    metadata_without_object_uuids,
                )
                .with(
                    6_001,
                    "TST.TableModelArchive (table UUID by form)",
                    crate::numbers_builder_cases::model_without_table_uuid,
                )
                .with(
                    4_008,
                    "TSCE.FormulaOwnerDependenciesArchive (owner UUID by form)",
                    crate::numbers_builder_cases::owner_without_uuid,
                )
                .with(
                    4_000,
                    "TSCE.CalculationEngineArchive (owner-map UUIDs by form)",
                    crate::numbers_builder_cases::engine_without_owner_uuids,
                )
                .with(
                    6_267,
                    "TST.ColumnRowUIDMapArchive (UIDs by form)",
                    crate::numbers_builder_cases::uid_map_by_form,
                )
        })
        .host(move |_source| Ok(host_bytes.clone()))
        .leaf(leaf)
        .run()
        .map_err(|error| format!("{text:?} {tables:?}: {error}"))?;
    }
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn builder_language_locale_cost_abba() -> TestResult {
    let host = |_source: &[u8]| -> TestResult<Vec<u8>> {
        Ok(litchi_iwa::pages::Builder::new()
            .body_text("Bonjour")
            .language("fr")
            .locale("fr_FR")
            .build()?
            .to_bytes()?)
    };
    let leaf = |_source: &[u8]| -> TestResult<Vec<u8>> {
        let package = litchi_pages::builder::Builder::new()
            .identity(litchi_pages::builder::Identity::from_bytes([7; 16]))
            .body_text("Bonjour")
            .language("fr")
            .locale("fr_FR")
            .build()?;
        Ok(leaf_bytes!(&package))
    };
    crate::pages_cost_probes::abba(
        &Fixture {
            path: "builder",
            bytes: Vec::new(),
        },
        host,
        leaf,
    )
}

/// The builder with body text and one 5×4 body table (record 0420).
#[test]
#[ignore = "callgrind cost probe (record 0420)"]
fn builder_body_table_cost_abba() -> TestResult {
    let host = |_source: &[u8]| -> TestResult<Vec<u8>> {
        Ok(litchi_iwa::pages::Builder::new()
            .body_text("Budget")
            .body_table("Q1", 5, 4)
            .build()?
            .to_bytes()?)
    };
    let leaf = |_source: &[u8]| -> TestResult<Vec<u8>> {
        let package = litchi_pages::builder::Builder::new()
            .identity(litchi_pages::builder::Identity::from_bytes([7; 16]))
            .body_text("Budget")
            .body_table("Q1", 5, 4)
            .build()?;
        Ok(leaf_bytes!(&package))
    };
    crate::pages_cost_probes::abba(
        &Fixture {
            path: "builder",
            bytes: Vec::new(),
        },
        host,
        leaf,
    )
}

/// Table insertion into a table-less built-in-template document (record
/// 0426): the host editor's `add_table` bootstraps the first table graph
/// (`bootstrap_first_table_graph`) where the leaf `table::creation` used to
/// refuse (`NoTemplate`); the leaf builds the same graph with the builder's
/// `builder_table`. Anchored at the start, inside and at the end of the
/// body text. The host draws the table, owner, UID-map and object UUIDs at
/// random; they are compared by form.
#[test]
fn pages_body_table_insert_without_template_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let text = "Hello table world";
    let input = Fixture {
        path: "built-in template + body text",
        bytes: leaf_bytes!(
            &litchi_pages::builder::Builder::new()
                .identity(litchi_pages::builder::Identity::from_bytes([7; 16]))
                .body_text(text)
                .build()?
        ),
    };
    for anchor in [0, 6, text.len()] {
        Case::new(
            "pages_body_table_insert_without_template_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .known_fields(|known| {
            known
                .with(
                    11_006,
                    "TSP.PackageMetadata (object UUIDs by form)",
                    metadata_without_object_uuids,
                )
                .with(
                    6_001,
                    "TST.TableModelArchive (table UUID by form)",
                    crate::numbers_builder_cases::model_without_table_uuid,
                )
                .with(
                    4_008,
                    "TSCE.FormulaOwnerDependenciesArchive (owner UUID by form)",
                    crate::numbers_builder_cases::owner_without_uuid,
                )
                .with(
                    4_000,
                    "TSCE.CalculationEngineArchive (owner-map UUIDs by form)",
                    crate::numbers_builder_cases::engine_without_owner_uuids,
                )
                .with(
                    6_267,
                    "TST.ColumnRowUIDMapArchive (UIDs by form)",
                    crate::numbers_builder_cases::uid_map_by_form,
                )
        })
        .host(|source| {
            let mut editor = litchi_iwa::pages::Draft::from_bytes(source)?;
            drop(editor.add_table(anchor, "Litchi table", 3, 4)?);
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_pages::Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_creation()
                .insert(
                    litchi_core::Position::new(anchor),
                    litchi_pages::table::creation::NewTable::new(3, 4).with_name("Litchi table"),
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
        .read("leaf tables", leaf_tables)
        .run()
        .map_err(|error| format!("anchor {anchor}: {error}"))?;
    }
    Ok(())
}

fn host_tables(bytes: &[u8]) -> TestResult<String> {
    let editor = litchi_iwa::pages::Draft::from_bytes(bytes)?;
    Ok(format!(
        "{:?} {:?}",
        editor.body_text()?,
        editor
            .tables()?
            .iter()
            .map(|table| (table.name.clone(), table.rows, table.columns))
            .collect::<Vec<_>>()
    ))
}

fn leaf_tables(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_pages::Package::from_bytes(bytes)?;
    let tables = package.body_tables()?;
    Ok(format!(
        "{:?} {:?}",
        package.text()?,
        tables
            .iter()
            .map(|table| (
                table.name().to_owned(),
                table.rows() as usize,
                table.columns() as usize
            ))
            .collect::<Vec<_>>()
    ))
}

#[inline(never)]
fn w4m_first_table_host(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = litchi_iwa::pages::Draft::from_bytes(source)?;
    drop(editor.add_table(6, "Litchi table", 3, 4)?);
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn w4m_first_table_leaf(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_pages::Package::from_bytes(source)?;
    let commit = package
        .edit_body_table_creation()
        .insert(
            litchi_core::Position::new(6),
            litchi_pages::table::creation::NewTable::new(3, 4).with_name("Litchi table"),
        )?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// Paired host-vs-leaf cost of the first-table insertion (record 0426):
/// open + insert + commit + serialize, ABBA; read the inclusive cost of
/// `w4m_first_table_{host,leaf}` under callgrind.
#[test]
#[ignore = "callgrind cost probe (record 0426)"]
fn w4m_first_table_cost_abba() -> TestResult {
    let input = leaf_bytes!(
        &litchi_pages::builder::Builder::new()
            .identity(litchi_pages::builder::Identity::from_bytes([7; 16]))
            .body_text("Hello table world")
            .build()?
    );
    let a1 = std::hint::black_box(w4m_first_table_host(&input)?);
    let b1 = std::hint::black_box(w4m_first_table_leaf(&input)?);
    let b2 = std::hint::black_box(w4m_first_table_leaf(&input)?);
    let a2 = std::hint::black_box(w4m_first_table_host(&input)?);
    assert_eq!(b1, b2, "the leaf route is deterministic");
    assert!(!a1.is_empty() && !a2.is_empty());
    Ok(())
}
