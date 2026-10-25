#![allow(
    clippy::expect_used,
    clippy::too_many_lines,
    clippy::unwrap_used,
    reason = "canonical integration assertions use panic-driven test fixtures"
)]

//! Integration coverage for the canonical aggregate PPT shape-text API.
//!
//! The aggregate API accepts bounded replacement chunks and emits one native
//! append-only publication.  These tests deliberately inspect the persisted
//! `Current User`/`UserEditAtom` chain in addition to the public semantic
//! projection, so a test cannot pass merely because the final texts happen to
//! be right after several hidden publications.

use std::collections::BTreeMap;
use std::io::{self, Cursor};
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use litchi_cfb::{OleError, OleFile, OleWriter};
use litchi_core::{Position, ReadAt, SourceVersion};
use litchi_ppt::package::edit::{EditOptions, Error, ShapeTextReplacement};
use litchi_ppt::shapes::{Shape, ShapeEnum, ShapeType};
use litchi_ppt::text_edit::Target;
use litchi_ppt::{CurrentUser, OpenOptions, Package, Record, RecordType, SourceLimits};

const MAX_CHUNK: usize = 256;
const MAX_AGGREGATE: usize = 16384;

fn fixture(slide_count: usize, textboxes_per_slide: usize) -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    for slide_index in 0..slide_count {
        let slide = writer.add_slide().expect("add fixture slide");
        for shape_index in 0..textboxes_per_slide {
            let x = 10 + i32::try_from(shape_index % 16).expect("shape x") * 260;
            let y = 10 + i32::try_from(shape_index / 16).expect("shape y") * 60;
            writer
                .add_textbox(
                    slide,
                    x,
                    y,
                    240,
                    40,
                    &format!("source-{slide_index:02}-{shape_index:04}"),
                )
                .expect("add fixture textbox");
        }
        // The rectangle is outside every replacement plan.  Its identity and
        // shape kind make the locality assertion independent of text output.
        writer
            .add_rectangle(slide, 5000, 5000, 720, 480)
            .expect("add unrelated fixture rectangle");
    }

    writer
        .build()
        .expect("build fixture")
        .to_bytes()
        .expect("serialize fixture")
}

fn package(slide_count: usize, textboxes_per_slide: usize) -> Package {
    Package::from_bytes(&fixture(slide_count, textboxes_per_slide)).expect("open fixture package")
}

fn shape_texts(package: &Package) -> Vec<Vec<String>> {
    package
        .presentation()
        .expect("read fixture presentation")
        .slides()
        .expect("read fixture slides")
        .into_iter()
        .map(|slide| {
            slide
                .shapes()
                .expect("read fixture shapes")
                .iter()
                .map(|shape| shape.text().expect("read shape text"))
                .collect()
        })
        .collect()
}

fn unrelated_shapes(package: &Package, textboxes_per_slide: usize) -> Vec<(u32, ShapeType)> {
    package
        .presentation()
        .expect("read fixture presentation")
        .slides()
        .expect("read fixture slides")
        .into_iter()
        .map(|slide| {
            let shapes = slide.shapes().expect("read fixture shapes");
            let shape = shapes
                .get(textboxes_per_slide)
                .expect("unrelated rectangle exists");
            let ShapeEnum::AutoShape(shape) = shape else {
                panic!("fixture rectangle parsed as an unexpected shape")
            };
            (shape.id(), shape.shape_type())
        })
        .collect()
}

fn stream_map(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .expect("open CFB stream map");
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).expect("read CFB stream");
            (path, data)
        })
        .collect()
}

fn named_stream<'a>(streams: &'a BTreeMap<Vec<String>, Vec<u8>>, name: &str) -> &'a [u8] {
    streams
        .iter()
        .find(|(path, _)| {
            path.last()
                .is_some_and(|last| last.eq_ignore_ascii_case(name))
        })
        .map(|(_, data)| data.as_slice())
        .unwrap_or_else(|| panic!("missing {name} stream"))
}

fn current_edit_offset(bytes: &[u8]) -> usize {
    let streams = stream_map(bytes);
    let current_user = CurrentUser::parse(
        named_stream(&streams, "Current User"),
        &SourceLimits::default(),
    )
    .expect("parse Current User stream");
    usize::try_from(current_user.current_edit_offset()).expect("Current User offset fits usize")
}

fn with_document_stream<T>(bytes: &[u8], operation: impl FnOnce(&[u8]) -> T) -> T {
    let streams = stream_map(bytes);
    operation(named_stream(&streams, "PowerPoint Document"))
}

fn top_level_user_edit_offsets(bytes: &[u8]) -> Vec<usize> {
    with_document_stream(bytes, |document| {
        let mut offsets = Vec::new();
        let mut offset = 0usize;
        while offset < document.len() {
            let Ok((record, consumed)) = Record::parse(document, offset, &SourceLimits::default())
            else {
                break;
            };
            if consumed == 0 {
                break;
            }
            if record.record_type == RecordType::UserEditAtom {
                offsets.push(offset);
            }
            offset = offset
                .checked_add(consumed)
                .expect("document record offset does not overflow");
        }
        offsets
    })
}

fn user_edit_predecessor(bytes: &[u8]) -> usize {
    let offset = current_edit_offset(bytes);
    with_document_stream(bytes, |document| {
        let (record, consumed) = Record::parse(document, offset, &SourceLimits::default())
            .expect("parse current UserEdit");
        assert_eq!(record.record_type, RecordType::UserEditAtom);
        assert!(consumed >= 20, "UserEditAtom contains offsetLastEdit");
        usize::try_from(u32::from_le_bytes(
            record.data[8..12].try_into().expect("offsetLastEdit field"),
        ))
        .expect("UserEdit predecessor fits usize")
    })
}

fn assert_one_native_publication(source: &[u8], target: &[u8]) {
    let source_edits = top_level_user_edit_offsets(source);
    let target_edits = top_level_user_edit_offsets(target);
    assert_eq!(
        target_edits.len(),
        source_edits.len() + 1,
        "one aggregate call must append exactly one UserEditAtom"
    );
    let source_head = current_edit_offset(source);
    let target_head = current_edit_offset(target);
    assert!(
        target_head > source_head,
        "the target Current User must point at the appended UserEditAtom"
    );
    assert_eq!(
        user_edit_predecessor(target),
        source_head,
        "the aggregate publication must link directly to the source head"
    );
}

fn assert_exact_patch_round_trip(
    source: &Package,
    target: &Package,
    patch: &litchi_ppt::package::edit::Patch,
) {
    let source_bytes = source.to_bytes().expect("read source bytes");
    let target_bytes = target.to_bytes().expect("read target bytes");
    assert_eq!(patch.before(), source_bytes.as_slice());
    assert_eq!(patch.after(), target_bytes.as_slice());

    let applied = source
        .apply(patch)
        .expect("apply exact forward patch")
        .into_package();
    assert_eq!(
        applied.to_bytes().expect("read applied bytes"),
        target_bytes,
        "forward patch must publish the committed bytes"
    );

    let restored = target
        .apply(&patch.inverse())
        .expect("apply exact inverse patch")
        .into_package();
    assert_eq!(
        restored.to_bytes().expect("read restored bytes"),
        source_bytes,
        "inverse patch must restore the exact source artifact"
    );
}

fn assert_unrelated_streams_preserved(before: &[u8], after: &[u8]) {
    let before_streams = stream_map(before);
    let after_streams = stream_map(after);
    assert_eq!(
        before_streams.keys().collect::<Vec<_>>(),
        after_streams.keys().collect::<Vec<_>>(),
        "aggregate publication must preserve CFB stream topology"
    );
    for (path, data) in before_streams {
        if path.last().is_some_and(|name| {
            name.eq_ignore_ascii_case("PowerPoint Document")
                || name.eq_ignore_ascii_case("Current User")
        }) {
            continue;
        }
        assert_eq!(
            after_streams.get(&path),
            Some(&data),
            "unrelated stream changed: {}",
            path.join("/")
        );
    }
}

fn replacement_chunks(
    slide_count: usize,
    textboxes_per_slide: usize,
) -> Vec<Vec<ShapeTextReplacement>> {
    let all = (0..slide_count)
        .flat_map(|slide| {
            (0..textboxes_per_slide).map(move |shape| {
                ShapeTextReplacement::new(
                    Target::new(Position::new(slide), Position::new(shape)),
                    format!("target-{slide:02}-{shape:04}"),
                )
            })
        })
        .collect::<Vec<_>>();
    all.chunks(MAX_CHUNK).map(ToOwned::to_owned).collect()
}

fn assert_replaced_texts(package: &Package, slide_count: usize, textboxes_per_slide: usize) {
    let texts = shape_texts(package);
    assert_eq!(texts.len(), slide_count);
    for (slide, shapes) in texts.iter().enumerate() {
        assert_eq!(shapes.len(), textboxes_per_slide + 1);
        for (shape, text) in shapes.iter().enumerate().take(textboxes_per_slide) {
            assert_eq!(*text, format!("target-{slide:02}-{shape:04}"));
        }
    }
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn change(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x0220,
            self.revision.load(Ordering::Acquire),
        ))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset =
            usize::try_from(offset).map_err(|_| io::Error::from(io::ErrorKind::InvalidInput))?;
        let input = self.bytes.get(offset..).unwrap_or_default();
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }
}

#[test]
fn too_many_chunks_checks_source_freshness_before_limit_admission() {
    let source_owner = Arc::new(VersionedSource::new(fixture(1, 1)));
    let source = Package::from_read_at(source_owner.clone()).expect("open positional fixture");
    let mut edit = source.edit().expect("start positional chunk edit");
    let empty: &[ShapeTextReplacement] = &[];
    let too_many_chunks = vec![empty; 4097];

    // The caller's source changed after the edit was captured.  The stale
    // source refusal must win over the cheap chunk-count refusal, even though
    // this plan contains no target resolution work.
    source_owner.change();
    let error = match edit.set_shape_text_chunks(&too_many_chunks) {
        Ok(_) => panic!("stale source must be refused"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Package(litchi_ppt::Error::Ole(OleError::SourceChanged { .. }))
    ));
    assert_eq!(edit.operation_count(), 0);
}

#[test]
fn aggregate_replacement_byte_limit_precedes_target_and_signed_checks() {
    let unsigned_bytes = fixture(1, 4);
    let unsigned_limits = SourceLimits::builder()
        .max_input_bytes(unsigned_bytes.len())
        .expect("unsigned package byte limit")
        .build()
        .expect("build unsigned package limits");
    let unsigned = Package::from_bytes_with_options(
        &unsigned_bytes,
        OpenOptions::new().source_limits(unsigned_limits),
    )
    .expect("open byte-limited unsigned package");
    assert_eq!(
        unsigned.source_limits().max_input_bytes(),
        unsigned_bytes.len()
    );

    // The target is invalid on purpose.  The replacement is only a few
    // kilobytes, but exceeds the selected package ceiling.  The byte-bound
    // refusal must precede slide/shape resolution and leave the edit retryable.
    let oversized_text = "x".repeat(unsigned_bytes.len() + 1);
    let invalid_target = [ShapeTextReplacement::new(
        Target::new(Position::new(99), Position::new(99)),
        oversized_text,
    )];
    let mut unsigned_edit = unsigned.edit().expect("start byte-limited edit");
    let error = unsigned_edit
        .set_shape_text_chunks(&[invalid_target.as_slice()])
        .expect_err("aggregate replacement bytes must be bounded");
    assert!(matches!(
        error,
        Error::Native(litchi_ppt::slide_order::Error::Package(litchi_ppt::Error::ResourceLimit(message)))
            if message.contains("aggregate replacement bytes")
    ));
    assert_eq!(unsigned_edit.operation_count(), 0);
    assert_eq!(
        unsigned_edit.source().to_bytes().expect("source bytes"),
        unsigned_bytes
    );
    let noop = [ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(0)),
        "source-00-0000",
    )];
    assert_eq!(
        unsigned_edit
            .set_shape_text_chunks(&[noop.as_slice()])
            .expect("retry with bounded no-op"),
        0
    );

    // Repeat the same admission check on a signed source.  Without the
    // aggregate byte check this path would return Error::Signed instead.
    let signed_bytes = signed_cfb(&unsigned_bytes);
    let signed_limits = SourceLimits::builder()
        .max_input_bytes(signed_bytes.len())
        .expect("signed package byte limit")
        .build()
        .expect("build signed package limits");
    let signed = Package::from_bytes_with_options(
        &signed_bytes,
        OpenOptions::new().source_limits(signed_limits),
    )
    .expect("open byte-limited signed package");
    let signed_oversized = [ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(0)),
        "y".repeat(signed_bytes.len() + 1),
    )];
    let mut signed_edit = signed.edit().expect("start signed byte-limited edit");
    let error = signed_edit
        .set_shape_text_chunks(&[signed_oversized.as_slice()])
        .expect_err("signed aggregate byte bound must precede signature policy");
    assert!(matches!(
        error,
        Error::Native(litchi_ppt::slide_order::Error::Package(litchi_ppt::Error::ResourceLimit(message)))
            if message.contains("aggregate replacement bytes")
    ));
    assert_eq!(signed_edit.operation_count(), 0);
    assert_eq!(
        signed_edit
            .source()
            .to_bytes()
            .expect("signed source bytes"),
        signed_bytes
    );
}

#[test]
fn aggregate_exact_chunk_boundary_publishes_once_and_round_trips_exactly() {
    let source = package(2, 128);
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_unrelated = unrelated_shapes(&source, 128);
    let chunks = replacement_chunks(2, 128);
    assert_eq!(chunks.len(), 1);
    assert_eq!(chunks[0].len(), MAX_CHUNK);
    let chunk_refs = chunks.iter().map(Vec::as_slice).collect::<Vec<_>>();

    let mut edit = source.edit().expect("start aggregate edit");
    assert_eq!(
        edit.set_shape_text_chunks(&chunk_refs)
            .expect("publish exact 256-selection aggregate"),
        MAX_CHUNK
    );
    assert_eq!(edit.operation_count(), MAX_CHUNK);
    let commit = edit.commit().expect("commit exact chunk boundary");
    assert_eq!(commit.patch().operation_count(), MAX_CHUNK);
    assert!(!commit.patch().is_noop());
    assert_replaced_texts(commit.package(), 2, 128);
    assert_eq!(unrelated_shapes(commit.package(), 128), source_unrelated);

    let target_bytes = commit.package().to_bytes().expect("target bytes");
    assert_one_native_publication(&source_bytes, &target_bytes);
    let reopened = Package::from_bytes(&target_bytes).expect("reopen aggregate target");
    assert_eq!(shape_texts(&reopened), shape_texts(commit.package()));
    assert_exact_patch_round_trip(&source, commit.package(), commit.patch());
    assert_eq!(
        source.to_bytes().expect("source remains immutable"),
        source_bytes
    );
    assert_unrelated_streams_preserved(&source_bytes, &target_bytes);
}

#[test]
fn aggregate_exact_selection_limit_is_one_publication_and_matches_old_chunks() {
    let source = package(1, MAX_AGGREGATE);
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_unrelated = unrelated_shapes(&source, MAX_AGGREGATE);
    let chunks = replacement_chunks(1, MAX_AGGREGATE);
    assert_eq!(chunks.len(), MAX_AGGREGATE / MAX_CHUNK);
    assert!(chunks.iter().all(|chunk| chunk.len() == MAX_CHUNK));
    let chunk_refs = chunks.iter().map(Vec::as_slice).collect::<Vec<_>>();

    let mut aggregate = source.edit().expect("start exact-limit aggregate edit");
    assert_eq!(
        aggregate
            .set_shape_text_chunks(&chunk_refs)
            .expect("publish exact aggregate selection limit"),
        MAX_AGGREGATE
    );
    assert_eq!(aggregate.operation_count(), MAX_AGGREGATE);
    let aggregate_commit = aggregate.commit().expect("commit exact-limit aggregate");
    let aggregate_bytes = aggregate_commit
        .package()
        .to_bytes()
        .expect("aggregate target bytes");
    assert_one_native_publication(&source_bytes, &aggregate_bytes);
    assert_replaced_texts(aggregate_commit.package(), 1, MAX_AGGREGATE);
    assert_eq!(
        unrelated_shapes(aggregate_commit.package(), MAX_AGGREGATE),
        source_unrelated
    );
    assert_exact_patch_round_trip(
        &source,
        aggregate_commit.package(),
        aggregate_commit.patch(),
    );

    // The old public method remains valid for bounded chunks, but each call
    // appends its own UserEdit.  Semantic output must match even though exact
    // package bytes and history are allowed to differ.
    let mut sequential = source.edit().expect("start sequential old edit");
    for chunk in &chunks {
        assert_eq!(
            sequential
                .set_shape_texts(chunk)
                .expect("publish one legacy chunk"),
            chunk.len()
        );
    }
    assert_eq!(sequential.operation_count(), MAX_AGGREGATE);
    let sequential_commit = sequential.commit().expect("commit sequential chunks");
    let sequential_bytes = sequential_commit
        .package()
        .to_bytes()
        .expect("sequential target bytes");
    assert_eq!(
        shape_texts(sequential_commit.package()),
        shape_texts(aggregate_commit.package()),
        "aggregate and sequential chunks must have identical semantic output"
    );
    assert_eq!(
        top_level_user_edit_offsets(&sequential_bytes).len(),
        top_level_user_edit_offsets(&source_bytes).len() + chunks.len(),
        "legacy one-chunk calls append one native publication each"
    );
    assert_ne!(
        current_edit_offset(&aggregate_bytes),
        current_edit_offset(&sequential_bytes),
        "aggregate and sequential history heads should remain distinguishable"
    );
}

#[test]
fn chunk_count_chunk_size_and_aggregate_limits_are_atomic_and_retryable() {
    let source = package(1, 4);
    let empty: &[ShapeTextReplacement] = &[];

    let mut edit = source.edit().expect("start admission edit");
    let exact_empty_chunks = vec![empty; 4096];
    assert_eq!(
        edit.set_shape_text_chunks(&exact_empty_chunks)
            .expect("4096 empty chunks fit the chunk bound"),
        0
    );
    assert_eq!(edit.operation_count(), 0);

    let over_chunk_count = vec![empty; 4097];
    assert!(
        edit.set_shape_text_chunks(&over_chunk_count).is_err(),
        "4097 chunks must be refused before target resolution"
    );
    assert_eq!(edit.operation_count(), 0);

    let oversized = (0..257)
        .map(|shape| {
            ShapeTextReplacement::new(
                Target::new(Position::new(0), Position::new(shape)),
                "not resolved",
            )
        })
        .collect::<Vec<_>>();
    assert!(
        edit.set_shape_text_chunks(&[oversized.as_slice()]).is_err(),
        "one chunk over 256 selections must be refused"
    );
    assert_eq!(edit.operation_count(), 0);

    let over_aggregate = (0..MAX_AGGREGATE / MAX_CHUNK + 1)
        .map(|chunk| {
            (0..MAX_CHUNK)
                .map(|shape| {
                    ShapeTextReplacement::new(
                        Target::new(Position::new(99), Position::new(chunk * MAX_CHUNK + shape)),
                        "not resolved",
                    )
                })
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    let over_aggregate_refs = over_aggregate.iter().map(Vec::as_slice).collect::<Vec<_>>();
    assert!(
        edit.set_shape_text_chunks(&over_aggregate_refs).is_err(),
        "more than 16384 aggregate selections must be refused"
    );
    assert_eq!(edit.operation_count(), 0);

    let valid = [ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(0)),
        "retry-success",
    )];
    assert_eq!(
        edit.set_shape_text_chunks(&[valid.as_slice()])
            .expect("valid plan remains retryable after admission refusals"),
        1
    );
    assert_eq!(edit.operation_count(), 1);
    let commit = edit
        .commit()
        .expect("commit retry after admission refusals");
    assert_eq!(shape_texts(commit.package())[0][0], "retry-success");
}

#[test]
fn duplicate_targets_across_chunks_and_custom_budget_leave_edit_unchanged() {
    let source = package(1, 4);
    let first = ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(0)), "first");
    let duplicate =
        ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(0)), "duplicate");
    let second =
        ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(1)), "second");
    let first_chunk = [first.clone()];
    let duplicate_chunk = [second.clone(), duplicate];

    let mut edit = source.edit().expect("start duplicate-check edit");
    assert!(
        edit.set_shape_text_chunks(&[first_chunk.as_slice(), duplicate_chunk.as_slice()])
            .is_err(),
        "targets must be distinct across chunk boundaries"
    );
    assert_eq!(edit.operation_count(), 0);
    assert_eq!(shape_texts(edit.source()), shape_texts(&source));

    let retry_chunk = [second];
    assert_eq!(
        edit.set_shape_text_chunks(&[first_chunk.as_slice(), retry_chunk.as_slice()])
            .expect("duplicate refusal leaves the plan retryable"),
        2
    );
    assert_eq!(edit.operation_count(), 2);
    let commit = edit.commit().expect("commit retried distinct chunks");
    assert_eq!(shape_texts(commit.package())[0][0], "first");
    assert_eq!(shape_texts(commit.package())[0][1], "second");

    let mut bounded = source
        .edit_with_options(EditOptions::new().max_operations(1).expect("budget"))
        .expect("start custom budget edit");
    let two = [
        ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(0)), "a"),
        ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(1)), "b"),
    ];
    assert!(
        matches!(
            bounded.set_shape_text_chunks(&[two.as_slice()]),
            Err(Error::Limit {
                observed: 2,
                maximum: 1,
                ..
            })
        ),
        "the selected EditOptions budget must apply to aggregate selections"
    );
    assert_eq!(bounded.operation_count(), 0);
    let one = [two[0].clone()];
    assert_eq!(
        bounded
            .set_shape_text_chunks(&[one.as_slice()])
            .expect("custom budget retry"),
        1,
        "one replacement fits the selected custom budget"
    );
    assert_eq!(bounded.operation_count(), 1);
}

#[test]
fn later_invalid_encoding_or_missing_target_rolls_back_and_can_retry() {
    let source = package(1, 4);
    let source_bytes = source.to_bytes().expect("source bytes");
    let valid = ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(0)),
        "valid-but-unpublished",
    );
    let invalid = ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(1)),
        "invalid\0encoding",
    );
    let valid_chunk = [valid.clone()];
    let invalid_chunk = [invalid];

    let mut encoding_edit = source.edit().expect("start encoding rollback edit");
    assert!(
        encoding_edit
            .set_shape_text_chunks(&[valid_chunk.as_slice(), invalid_chunk.as_slice()])
            .is_err(),
        "a later incompatible encoding must reject the complete aggregate"
    );
    assert_eq!(encoding_edit.operation_count(), 0);
    assert_eq!(
        encoding_edit.source().to_bytes().expect("source bytes"),
        source_bytes
    );
    assert_eq!(
        encoding_edit
            .set_shape_text_chunks(&[valid_chunk.as_slice()])
            .expect("retry after encoding refusal"),
        1
    );

    let missing = ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(999)),
        "missing-target",
    );
    let missing_chunk = [missing];
    let mut missing_edit = source.edit().expect("start missing-target rollback edit");
    assert!(
        missing_edit
            .set_shape_text_chunks(&[valid_chunk.as_slice(), missing_chunk.as_slice()])
            .is_err(),
        "a later missing target must reject the complete aggregate"
    );
    assert_eq!(missing_edit.operation_count(), 0);
    assert_eq!(
        missing_edit.source().to_bytes().expect("source bytes"),
        source_bytes
    );
    assert_eq!(
        missing_edit
            .set_shape_text_chunks(&[valid_chunk.as_slice()])
            .expect("retry after missing-target refusal"),
        1
    );
}

#[test]
fn noops_count_as_operations_but_preserve_exact_source_and_signed_noops() {
    let source = package(1, 4);
    let source_bytes = source.to_bytes().expect("source bytes");
    let noops = (0..3)
        .map(|shape| {
            ShapeTextReplacement::new(
                Target::new(Position::new(0), Position::new(shape)),
                format!("source-00-{shape:04}"),
            )
        })
        .collect::<Vec<_>>();
    let noops_chunk = [noops[0].clone(), noops[1].clone()];
    let mixed_chunk = [
        noops[2].clone(),
        ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(3)), "changed"),
    ];

    let shared = source.shared_ole_file_arc();
    let mut edit = source.edit().expect("start mixed no-op aggregate");
    assert_eq!(
        edit.set_shape_text_chunks(&[noops_chunk.as_slice(), mixed_chunk.as_slice()])
            .expect("stage mixed aggregate"),
        1
    );
    assert_eq!(edit.operation_count(), 4);
    let commit = edit.commit().expect("commit mixed aggregate");
    assert!(!commit.patch().is_noop());
    assert_eq!(shape_texts(commit.package())[0][3], "changed");
    assert_eq!(commit.patch().operation_count(), 4);
    assert_eq!(source.to_bytes().expect("source immutable"), source_bytes);
    assert_unrelated_streams_preserved(
        &source_bytes,
        &commit.package().to_bytes().expect("mixed target bytes"),
    );

    let mut noop_edit = source.edit().expect("start exact no-op aggregate");
    let noop_refs = [&noops[..]];
    assert_eq!(
        noop_edit
            .set_shape_text_chunks(&noop_refs)
            .expect("stage semantic no-op aggregate"),
        0
    );
    assert_eq!(noop_edit.operation_count(), noops.len());
    let noop_commit = noop_edit.commit().expect("commit exact no-op aggregate");
    assert!(noop_commit.patch().is_noop());
    assert!(Arc::ptr_eq(
        &shared,
        &noop_commit.package().shared_ole_file_arc()
    ));
    assert_eq!(noop_commit.patch().before(), source_bytes.as_slice());
    assert_eq!(noop_commit.patch().after(), source_bytes.as_slice());

    let signed_bytes = signed_cfb(&source_bytes);
    let signed = Package::from_bytes(&signed_bytes).expect("open synthetic signed package");
    let mut signed_noop = signed.edit().expect("signed no-op planning");
    assert_eq!(
        signed_noop
            .set_shape_text_chunks(&noop_refs)
            .expect("signed semantic no-op aggregate"),
        0
    );
    let signed_commit = signed_noop.commit().expect("signed no-op commit");
    assert!(signed_commit.patch().is_noop());
    assert_eq!(signed_commit.patch().before(), signed_bytes.as_slice());
    assert_eq!(signed_commit.patch().after(), signed_bytes.as_slice());

    let changed = [ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(0)),
        "signed-change",
    )];
    let mut signed_changed = signed.edit().expect("signed changed planning");
    assert!(matches!(
        signed_changed.set_shape_text_chunks(&[changed.as_slice()]),
        Err(Error::Signed)
    ));
    assert_eq!(signed_changed.operation_count(), 0);
}

fn signed_cfb(unsigned: &[u8]) -> Vec<u8> {
    let mut source = OleFile::open(Cursor::new(unsigned), litchi_cfb::OleFileLimits::default())
        .expect("open unsigned CFB");
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let names = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = source.open_stream(&names).expect("read unsigned stream");
        writer
            .create_stream(&names, &data)
            .expect("copy unsigned stream");
    }
    writer
        .create_stream(&["_signatures"], b"synthetic signature")
        .expect("add signature stream");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write signed CFB");
    output.into_inner()
}
