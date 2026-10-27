//! Keynote builder language/locale differential (wave W1-C, record 0372).
//!
//! The host builder draws random object UUIDs, so its output cannot be
//! compared byte- or object-wise with the leaf builder's; this test compares
//! every field the host's `language`/`locale` writes (document language and
//! locales, every text storage's language table, the calculation engine
//! locale) and the semantic reads instead. It is not a tiered parity case:
//! the ledger rows stay host-only until the harness can normalize builder
//! identities (record 0372 §gaps).

use litchi_iwa_protos::{kn, tsce, tswp};
use prost::Message as _;

use crate::archive_diff::{members, objects};
use crate::harness::TestResult;

/// (document language, locale, creation locale, storage languages sorted,
/// calculation-engine locale).
type LocaleFields = (
    Option<String>,
    Option<String>,
    Option<String>,
    Vec<String>,
    Option<String>,
);

#[allow(
    deprecated,
    reason = "the native document archive marks locale fields deprecated"
)]
fn locale_fields(bytes: &[u8]) -> TestResult<LocaleFields> {
    let mut fields: LocaleFields = (None, None, None, Vec::new(), None);
    for member in members(bytes)? {
        if !member.name.ends_with(".iwa") {
            continue;
        }
        for object in objects(&member.data)? {
            for message in &object.messages {
                match message.type_ {
                    1 if member.name == "Index/Document.iwa" => {
                        let document = kn::DocumentArchive::decode(message.data.as_slice())?;
                        fields.0 = document.super_.document_language.clone();
                        fields.1 = document.super_.super_.locale_identifier.clone();
                        fields.2 = document.super_.super_.creation_locale_identifier.clone();
                    },
                    2_001 => {
                        let storage = tswp::StorageArchive::decode(message.data.as_slice())?;
                        if let Some(table) = storage.table_language {
                            fields
                                .3
                                .extend(table.entries.into_iter().filter_map(|entry| entry.object));
                        }
                    },
                    4_000 => {
                        let engine =
                            tsce::CalculationEngineArchive::decode(message.data.as_slice())?;
                        fields.4 = engine.saved_locale_identifier;
                    },
                    _ => {},
                }
            }
        }
    }
    fields.3.sort();
    Ok(fields)
}

#[test]
fn keynote_builder_language_locale_differential() -> TestResult {
    for (language, locale) in [("fr", "fr_FR"), ("ja", "ja_JP"), ("de-CH", "de_CH")] {
        let host = litchi_iwa::keynote::Builder::new()
            .title("Titre")
            .subtitle("Sous-titre")
            .presenter_notes("Notes")
            .language(language)
            .locale(locale)
            .build()?
            .to_bytes()?;
        let leaf_package = litchi_keynote::Builder::new()
            .identity(litchi_keynote::Identity::BLANK)
            .title("Titre")
            .subtitle("Sous-titre")
            .notes("Notes")
            .language(language)
            .locale(locale)
            .build()?;
        let mut leaf = Vec::new();
        leaf_package
            .write_to(&mut leaf)
            .map_err(|error| format!("{error:?}"))?;
        assert_eq!(
            locale_fields(&leaf)?,
            locale_fields(&host)?,
            "{language}/{locale}"
        );
        let reopened = litchi_keynote::Package::from_bytes(&host)?;
        let slide = litchi_keynote::SlideSelector::index(0);
        assert_eq!(
            leaf_package.slide_title(slide)?,
            reopened.slide_title(slide)?
        );
        assert_eq!(
            leaf_package.slide_notes(slide)?,
            reopened.slide_notes(slide)?
        );
    }
    Ok(())
}
