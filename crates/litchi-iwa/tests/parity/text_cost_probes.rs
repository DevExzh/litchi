//! Callgrind ABBA cost probes of the W1-D text engines (record 0374): each
//! host `text::Draft` route against the engine route published through
//! [`text_glue`](crate::text_glue), on the same fixture bytes (open, edit,
//! serialize).
//!
//! ```sh
//! CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
//! taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*text_cost_probes::host_*' \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   text_cost_probes::text_engine_cost_abba --test-threads=1
//! # likewise with '--toggle-collect=*text_cost_probes::leaf_*'
//! ```
//!
//! `text_property_batch_wall_abba` times 1, 10 and 100 sequential property
//! sets on one storage (the 1k batch is too slow on the host for callgrind).

use std::hint::black_box;
use std::time::Instant;

use litchi_iwa_text::paragraph::format::Alignment;
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::language;
use litchi_iwa_text_wire::paragraph_style::{Change, Property, PropertyKind};

use crate::harness::{TestResult, fixture};
use crate::text_annotation_cases::{self, text_storages};
use crate::text_glue::Workspace;
use crate::text_style_cases::{self, apply_change, sample};

const FIXTURE: &str = "iwork/keynote/basic.key";

#[inline(never)]
fn host_property(source: &[u8], storage: u64, change: &Change) -> TestResult<Vec<u8>> {
    text_style_cases::host_change(source, storage, change)
}

#[inline(never)]
fn leaf_property(source: &[u8], storage: u64, change: &Change) -> TestResult<Vec<u8>> {
    text_style_cases::leaf_change(source, storage, change)
}

#[inline(never)]
fn host_annotation(
    source: &[u8],
    storage: u64,
    operation: &text_annotation_cases::Operation,
) -> TestResult<Vec<u8>> {
    text_annotation_cases::host_route(source, storage, operation)
}

#[inline(never)]
fn leaf_annotation(
    source: &[u8],
    storage: u64,
    operation: &text_annotation_cases::Operation,
) -> TestResult<Vec<u8>> {
    text_annotation_cases::leaf_route(source, storage, operation)
}

#[inline(never)]
fn host_language(source: &[u8], storage: u64, change: &language::Change) -> TestResult<Vec<u8>> {
    crate::text_language_cases::host_route(source, storage, change)
}

#[inline(never)]
fn leaf_language(source: &[u8], storage: u64, change: &language::Change) -> TestResult<Vec<u8>> {
    crate::text_language_cases::leaf_route(source, storage, change)
}

#[inline(never)]
fn host_batch(source: &[u8], storage: u64, count: usize) -> TestResult<Vec<u8>> {
    let mut draft = litchi_iwa::text::Draft::from_bytes(source)?;
    let id = litchi_iwa::text::TextStorageId::try_from(storage)?;
    for index in 0..count {
        draft.set_paragraph_alignment(id, alternate(index))?;
    }
    Ok(draft.to_bytes()?)
}

#[inline(never)]
fn leaf_batch(source: &[u8], storage: u64, count: usize) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    for index in 0..count {
        apply_change(
            &mut workspace,
            storage,
            &Change::Set(Property::Alignment(alternate(index))),
        )?;
    }
    workspace.into_bytes()
}

const fn alternate(index: usize) -> Alignment {
    if index.is_multiple_of(2) {
        Alignment::Right
    } else {
        Alignment::Center
    }
}

/// Single-operation routes: a derived set, an in-place rewrite, a reset,
/// a hyperlink add and a language set, each host/leaf/leaf/host.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn text_engine_cost_abba() -> TestResult {
    let input = fixture(FIXTURE)?;
    let (storage, range, _) = text_storages(&input.bytes, 1)?
        .into_iter()
        .next()
        .ok_or("no storage")?;
    let derived = text_style_cases::derived(&input, storage, PropertyKind::Alignment)?;
    let property = [
        (
            input.bytes.clone(),
            Change::Set(sample(PropertyKind::TextColor)?),
        ),
        (
            derived.bytes.clone(),
            Change::Set(Property::Alignment(Alignment::Center)),
        ),
        (
            derived.bytes.clone(),
            Change::Reset(PropertyKind::Alignment),
        ),
    ];
    for (bytes, change) in &property {
        black_box(host_property(bytes, storage, change)?);
        black_box(leaf_property(bytes, storage, change)?);
        black_box(leaf_property(bytes, storage, change)?);
        black_box(host_property(bytes, storage, change)?);
    }
    let link = text_annotation_cases::Operation::Add(
        range,
        text_annotation_cases::Field::Hyperlink(
            litchi_iwa_text::hyperlink::TextHyperlinkTarget::new("https://example.com")?,
        ),
    );
    black_box(host_annotation(&input.bytes, storage, &link)?);
    black_box(leaf_annotation(&input.bytes, storage, &link)?);
    black_box(leaf_annotation(&input.bytes, storage, &link)?);
    black_box(host_annotation(&input.bytes, storage, &link)?);
    let set = language::Change::Set {
        position: TextPosition::ZERO,
        language: litchi_iwa_text::language::TextLanguage::tag("fr")?,
    };
    black_box(host_language(&input.bytes, storage, &set)?);
    black_box(leaf_language(&input.bytes, storage, &set)?);
    black_box(leaf_language(&input.bytes, storage, &set)?);
    black_box(host_language(&input.bytes, storage, &set)?);
    Ok(())
}

/// Wall-clock ABBA of 1, 10 and 100 sequential sets on one storage.
#[test]
#[ignore = "wall-clock cost probe; run explicitly"]
fn text_property_batch_wall_abba() -> TestResult {
    let input = fixture(FIXTURE)?;
    let (storage, _, _) = text_storages(&input.bytes, 1)?
        .into_iter()
        .next()
        .ok_or("no storage")?;
    for count in [1usize, 10, 100] {
        let mut host = std::time::Duration::ZERO;
        let mut leaf = std::time::Duration::ZERO;
        for _ in 0..3 {
            let start = Instant::now();
            black_box(host_batch(&input.bytes, storage, count)?);
            host += start.elapsed();
            let start = Instant::now();
            black_box(leaf_batch(&input.bytes, storage, count)?);
            black_box(leaf_batch(&input.bytes, storage, count)?);
            leaf += start.elapsed() / 2;
            let start = Instant::now();
            black_box(host_batch(&input.bytes, storage, count)?);
            host += start.elapsed();
        }
        eprintln!(
            "BATCH {count}: host {:?} leaf {:?} per batch",
            host / 6,
            leaf / 3
        );
    }
    Ok(())
}
