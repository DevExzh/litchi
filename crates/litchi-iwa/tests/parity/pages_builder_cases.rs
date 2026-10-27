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
