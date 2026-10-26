//! Language-run parity (wave W1-D, record 0374): the host `text::Draft`
//! language family against `litchi_iwa_text_wire::language` published by
//! [`text_glue`](crate::text_glue). Language tables live in the storage
//! payload only, so every case is tier A.

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_text::language::{TextLanguage, TextLanguageRun};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::language::{self, Change};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::text_annotation_cases::text_storages;
use crate::text_glue::Workspace;
use crate::text_style_cases::FIXTURES;

const LIMITS: WireLimits = WireLimits::DEFAULT;

pub fn host_runs(bytes: &[u8], storage: u64) -> TestResult<Vec<TextLanguageRun>> {
    Ok(Draft::from_bytes(bytes)?.text_languages(TextStorageId::try_from(storage)?)?)
}

pub fn leaf_runs(bytes: &[u8], storage: u64) -> TestResult<Vec<TextLanguageRun>> {
    let workspace = Workspace::open(bytes)?;
    Ok(language::runs(
        &workspace.storage(storage)?.payload,
        LIMITS,
    )?)
}

pub fn host_route(source: &[u8], storage: u64, change: &Change) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match change.clone() {
        Change::Set { position, language } => draft.set_text_language(id, position, language)?,
        Change::RemoveBoundary(position) => {
            draft.remove_text_language_boundary(id, position)?;
        },
        Change::Reset => {
            draft.reset_text_languages(id)?;
        },
        other => return Err(format!("unmapped change {other:?}").into()),
    }
    Ok(draft.to_bytes()?)
}

pub fn leaf_route(source: &[u8], storage: u64, change: &Change) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    let located = workspace.storage(storage)?;
    if let Some(payload) = language::rewrite(&located.payload, change, LIMITS)? {
        workspace.replace_with_references(&located, payload, None, None)?;
    }
    workspace.into_bytes()
}

fn run(
    name: &'static str,
    format: Format,
    input: &Fixture,
    storage: u64,
    change: &Change,
) -> TestResult {
    Case::new(name, format, Tier::A, Evidence::Queued, input)
        .host(|source| host_route(source, storage, change))
        .leaf(|source| Ok(LeafRun::output(leaf_route(source, storage, change)?)))
        .read("host runs", move |bytes| host_runs(bytes, storage))
        .read("leaf runs", move |bytes| leaf_runs(bytes, storage))
        .run()?;
    Ok(())
}

fn tagged(input: &Fixture, storage: u64, position: u32) -> TestResult<Fixture> {
    let mut bytes = input.bytes.clone();
    for (at, tag) in [(0, "en"), (position, "fr")] {
        bytes = host_route(
            &bytes,
            storage,
            &Change::Set {
                position: TextPosition::from_utf16_code_units(at),
                language: TextLanguage::tag(tag)?,
            },
        )?;
    }
    Ok(Fixture {
        path: input.path,
        bytes,
    })
}

#[test]
fn text_language_set_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let end = range.end().utf16_index();
            let base = tagged(&input, storage, end)?;
            for (fixture, position, tag) in [
                (&input, 0, "de"),
                (&input, end, "ja"),
                (&base, end, "es"),
                (&base, 0, "fr"),
            ] {
                run(
                    "text_language_set_parity",
                    format,
                    fixture,
                    storage,
                    &Change::Set {
                        position: TextPosition::from_utf16_code_units(position),
                        language: TextLanguage::tag(tag)?,
                    },
                )
                .map_err(|e| format!("{path} storage {storage} {position} {tag}: {e}"))?;
                cases += 1;
            }
            let _ = Tier::A;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 16, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_language_remove_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let end = range.end().utf16_index();
            let base = tagged(&input, storage, end)?;
            run(
                "text_language_remove_parity",
                format,
                &base,
                storage,
                &Change::RemoveBoundary(TextPosition::from_utf16_code_units(end)),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
            let zero = Change::RemoveBoundary(TextPosition::ZERO);
            assert!(host_route(&base.bytes, storage, &zero).is_err());
            assert!(leaf_route(&base.bytes, storage, &zero).is_err());
            let _ = Tier::A;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_language_reset_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let base = tagged(&input, storage, range.end().utf16_index())?;
            run(
                "text_language_reset_parity",
                format,
                &base,
                storage,
                &Change::Reset,
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
            let _ = Tier::A;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_language_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 4)? {
            let end = range.end().utf16_index();
            let base = tagged(&input, storage, end)?;
            for bytes in [&input.bytes, &base.bytes] {
                assert_eq!(host_runs(bytes, storage)?, leaf_runs(bytes, storage)?);
                for position in [0, 1, end] {
                    let position = TextPosition::from_utf16_code_units(position);
                    let host = Draft::from_bytes(bytes)?
                        .text_language(TextStorageId::try_from(storage)?, position)?;
                    let workspace = Workspace::open(bytes)?;
                    let leaf =
                        language::at(&workspace.storage(storage)?.payload, position, LIMITS)?;
                    assert_eq!(host, leaf, "{path} storage {storage}");
                    reads += 1;
                }
            }
        }
    }
    assert!(reads >= 24, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
