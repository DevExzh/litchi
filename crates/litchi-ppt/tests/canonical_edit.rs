#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "canonical API contract tests use panic-driven assertions"
)]

//! Contract tests for the immutable, source-checked PPT edit root.
//!
//! These tests deliberately exercise the public `package::edit` family rather
//! than the native `slide_order` transaction directly.  A generated deck keeps
//! the semantic assertions small while still proving that structural, text,
//! and anchor operations compose into one exact reversible publication.

use litchi_cfb::OleError;
use litchi_core::{Position, ReadAt, SourceVersion};
use litchi_ppt::package::edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};
use litchi_ppt::shapes::{Shape, ShapeEnum};
use litchi_ppt::text_edit::Target;
use litchi_ppt::{Anchor, Leniency, OpenOptions, Package, SourceLimits};
use std::io::{self, Cursor};
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

fn generated_deck(suffix: &str) -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    for text in ["alpha", "beta", "gamma"] {
        let slide = writer.add_slide().expect("add generated slide");
        writer
            .add_textbox(slide, 10, 10, 200, 40, &format!("{text}{suffix}"))
            .expect("add generated textbox");
    }
    let mut output = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).expect("write generated deck");
    output.into_inner()
}

fn package() -> Package {
    Package::from_bytes(&generated_deck("")).expect("open generated deck")
}

fn slide_texts(package: &Package) -> Vec<String> {
    package
        .presentation()
        .expect("read presentation")
        .slides()
        .expect("read slides")
        .into_iter()
        .map(|slide| slide.text().expect("read slide text").to_owned())
        .collect()
}

fn first_shape_bounds(package: &Package, slide: usize) -> (i32, i32, i32, i32) {
    let presentation = package.presentation().expect("read presentation");
    let slide = presentation
        .slide_at(slide)
        .expect("resolve slide")
        .expect("slide exists");
    let shape = &slide.shapes().expect("read shapes")[0];
    match shape {
        ShapeEnum::TextBox(textbox) => textbox.bounds(),
        other => panic!("expected generated textbox, found {other:?}"),
    }
}

fn type_checks<T>() {}

fn is_source_changed(error: &Error) -> bool {
    matches!(
        error,
        Error::Package(litchi_ppt::Error::Ole(OleError::SourceChanged { .. }))
    )
}

#[test]
fn canonical_edit_family_is_public_and_sendable() {
    type_checks::<Edit>();
    type_checks::<Commit>();
    type_checks::<Patch>();
    type_checks::<Diagnostics>();
    type_checks::<Error>();
    type_checks::<LimitKind>();
    type_checks::<EditOptions>();

    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Package>();
    assert_send_sync::<Edit>();
    assert_send_sync::<Patch>();
}

#[test]
fn structural_text_and_anchor_operations_publish_one_semantic_commit() {
    let source = package();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("start canonical edit");

    // The target positions below are projected positions: after the move the
    // old beta slide is at zero and gamma is at one.
    edit.move_slide(Position::new(0), Position::new(2))
        .expect("move alpha after gamma");
    edit.set_shape_text(
        Target::new(Position::new(0), Position::new(0)),
        "beta-updated".to_owned(),
    )
    .expect("replace beta text");
    edit.set_shape_anchor(
        Target::new(Position::new(1), Position::new(0)),
        Anchor::small(20, 20, 220, 80).expect("valid anchor"),
    )
    .expect("replace gamma anchor");
    edit.remove_slide(Position::new(2))
        .expect("remove moved alpha");

    let commit = edit.commit().expect("commit canonical edit");
    let target = commit.package();
    assert_eq!(
        slide_texts(target),
        vec!["beta-updated", "gamma"],
        "structural and text operations retain projected semantic order"
    );
    assert_eq!(
        first_shape_bounds(target, 1),
        (20, 20, 200, 60),
        "anchor edit is visible through the normal presentation projection"
    );

    let patch = commit.patch();
    assert!(!patch.is_noop());
    assert_eq!(patch.before(), source_bytes.as_slice());
    assert_eq!(patch.after(), target.to_bytes().expect("target bytes"));
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().operation_count() >= 1);
}

#[test]
fn patch_replay_inverse_and_wrong_source_refusal_are_exact() {
    let source = package();
    let mut edit = source.edit().expect("start canonical edit");
    edit.move_slide(Position::new(0), Position::new(1))
        .expect("stage move");
    edit.set_shape_text(
        Target::new(Position::new(0), Position::new(0)),
        "beta".to_owned(),
    )
    .expect("stage exact-length text replacement");
    let commit = edit.commit().expect("commit edit");

    let replay = source
        .apply(commit.patch())
        .expect("replay patch on exact source");
    assert!(replay.diagnostics().changed());
    assert_eq!(replay.patch().after(), commit.patch().after());
    let replayed = replay.into_package();
    assert_eq!(
        replayed.to_bytes().expect("replayed bytes"),
        commit.package().to_bytes().expect("target bytes")
    );

    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .expect("apply exact inverse")
        .into_package();
    assert_eq!(
        restored.to_bytes().expect("restored bytes"),
        source.to_bytes().expect("source bytes")
    );

    let wrong_source = Package::from_bytes(&generated_deck("-other")).expect("wrong source");
    assert!(matches!(
        wrong_source.apply(commit.patch()),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn exact_noop_preserves_the_source_owner_and_reports_no_change() {
    let source = package();
    let shared = source.shared_ole_file_arc();
    let mut edit = source.edit().expect("start canonical edit");
    edit.move_slide(Position::new(0), Position::new(0))
        .expect("same-position move is a no-op");
    let commit = edit.commit().expect("commit no-op");

    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert!(Arc::ptr_eq(
        &shared,
        &commit.package().shared_ole_file_arc()
    ));
    assert_eq!(
        commit.package().to_bytes().expect("no-op bytes"),
        source.to_bytes().expect("source bytes")
    );
}

#[test]
fn failed_stage_keeps_prior_valid_stage_and_drop_keeps_source_immutable() {
    let source = package();
    let source_bytes = source.to_bytes().expect("source bytes");
    {
        let mut edit = source.edit().expect("start canonical edit");
        edit.move_slide(Position::new(0), Position::new(1))
            .expect("valid first stage");
        assert!(
            edit.remove_slide(Position::new(99)).is_err(),
            "failed stage is reported"
        );
        edit.set_shape_text(
            Target::new(Position::new(0), Position::new(0)),
            "beta".to_owned(),
        )
        .expect("valid retry after failed stage");
        let commit = edit.commit().expect("commit surviving stages");
        assert_eq!(
            slide_texts(commit.package()),
            vec!["beta", "alpha", "gamma"]
        );
    }
    assert_eq!(
        source.to_bytes().expect("source remains readable"),
        source_bytes
    );

    // Dropping an uncommitted edit must not publish its candidate.
    {
        let mut edit = source.edit().expect("start second edit");
        edit.remove_slide(Position::new(0)).expect("stage removal");
    }
    assert_eq!(slide_texts(&source), vec!["alpha", "beta", "gamma"]);
}

#[test]
fn edit_output_retains_source_limits_and_leniency() {
    let bytes = generated_deck("");
    let limits = SourceLimits::builder()
        // Append-only PPT publication can grow the source.  The selected
        // policy therefore leaves a finite, explicit headroom for the target.
        .max_input_bytes(bytes.len().saturating_mul(2).saturating_add(64 * 1024))
        .expect("input limit")
        .max_events(10_000)
        .expect("event limit")
        .build()
        .expect("source limits");
    let options = OpenOptions::new()
        .source_limits(limits)
        .leniency(Leniency::TolerateUserEditChainDefects);
    let source = Package::from_bytes_with_options(&bytes, options).expect("open with policies");
    let mut edit = source.edit().expect("start policy-preserving edit");
    edit.set_shape_text(
        Target::new(Position::new(0), Position::new(0)),
        "ALPHA".to_owned(),
    )
    .expect("stage policy-preserving text change");
    let target = edit
        .commit()
        .expect("commit policy-preserving edit")
        .into_package();

    assert_eq!(target.source_limits(), limits);
    assert_eq!(target.leniency(), Leniency::TolerateUserEditChainDefects);
    assert_eq!(
        target
            .presentation()
            .expect("target presentation")
            .source_limits(),
        limits
    );
    assert!(slide_texts(&target)[0].contains("ALPHA"));
}

#[test]
fn operation_limits_are_finite_and_failed_operations_can_be_retried() {
    let invalid = EditOptions::new().max_operations(0).unwrap_err();
    assert!(matches!(
        invalid,
        Error::InvalidLimit {
            kind: LimitKind::Operations,
            value: 0,
            maximum
        } if maximum > 0
    ));

    let options = EditOptions::new()
        .max_operations(2)
        .expect("finite operation limit");
    assert_eq!(options.max_operations_value(), 2);
    let source = package();
    let mut edit = source
        .edit_with_options(options)
        .expect("start bounded edit");
    edit.move_slide(Position::new(0), Position::new(1))
        .expect("first operation fits bound");
    assert!(edit.remove_slide(Position::new(99)).is_err());
    edit.set_shape_text(
        Target::new(Position::new(0), Position::new(0)),
        "beta".to_owned(),
    )
    .expect("valid operation can retry after failed operation");
    assert!(matches!(
        edit.move_slide(Position::new(1), Position::new(2)),
        Err(Error::Limit {
            kind: LimitKind::Operations,
            observed: 3,
            maximum: 2
        })
    ));
    let commit = edit.commit().expect("commit bounded edit");
    assert_eq!(
        slide_texts(commit.package()),
        vec!["beta", "alpha", "gamma"]
    );
}

#[test]
fn debug_output_does_not_expose_package_payload() {
    let package = package();
    let options = EditOptions::new();
    let package_debug = format!("{package:?}");
    let options_debug = format!("{options:?}");
    let mut edit = package.edit().expect("start debug edit");
    edit.set_shape_text(
        Target::new(Position::new(0), Position::new(0)),
        "debug-secret".to_owned(),
    )
    .expect("stage debug edit");
    let edit_debug = format!("{edit:?}");
    let commit = edit.commit().expect("commit debug edit");
    let commit_debug = format!("{commit:?}");
    let patch_debug = format!("{:?}", commit.patch());
    assert!(!package_debug.contains("alpha"));
    assert!(!package_debug.contains("beta"));
    assert!(!options_debug.is_empty());
    assert!(!edit_debug.contains("debug-secret"));
    assert!(!commit_debug.contains("debug-secret"));
    assert!(!patch_debug.contains("debug-secret"));
}

struct FakeReadAt {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl FakeReadAt {
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

impl ReadAt for FakeReadAt {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x0212,
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
fn source_changes_are_rejected_before_edit_commit_and_apply() {
    let bytes = generated_deck("");

    let source = Arc::new(FakeReadAt::new(bytes.clone()));
    let package = Package::from_read_at(source.clone()).expect("open positional package");
    source.change();
    assert!(
        matches!(package.edit(), Err(error) if is_source_changed(&error)),
        "stale source rejected before edit"
    );

    let source = Arc::new(FakeReadAt::new(bytes.clone()));
    let package = Package::from_read_at(source.clone()).expect("open positional package");
    let mut edit = package.edit().expect("start positional edit");
    edit.move_slide(Position::new(0), Position::new(1))
        .expect("stage positional edit");
    source.change();
    assert!(
        matches!(edit.commit(), Err(error) if is_source_changed(&error)),
        "stale source rejected before commit"
    );

    let source = Arc::new(FakeReadAt::new(bytes));
    let package = Package::from_read_at(source.clone()).expect("open positional package");
    let mut edit = package.edit().expect("start source for patch");
    edit.move_slide(Position::new(0), Position::new(1))
        .expect("stage source patch");
    let commit = edit.commit().expect("commit source patch");
    source.change();
    assert!(
        matches!(package.apply(commit.patch()), Err(error) if is_source_changed(&error)),
        "stale source rejected on apply"
    );
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_edit_refuses_before_capture_and_preserves_ciphertext() {
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 0, 0, 100, 100, "secret").unwrap();
    writer
        .set_password(
            "password",
            litchi_ppt::EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.set_encryption_entropy(litchi_ppt::EncryptionEntropy::generate().unwrap());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    let bytes = output.into_inner();
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().password("password")).unwrap();
    assert!(matches!(
        package.edit(),
        Err(Error::Package(litchi_ppt::Error::EncryptionPolicy))
    ));
    assert_eq!(package.to_encrypted_bytes().unwrap(), bytes);
}

#[test]
fn changed_publication_cannot_widen_input_limits() {
    let bytes = generated_deck("");
    let limits = SourceLimits::builder()
        .max_input_bytes(bytes.len())
        .unwrap()
        .build()
        .unwrap();
    let source =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(limits)).unwrap();
    let mut edit = source.edit().unwrap();
    edit.move_slide(Position::new(0), Position::new(2)).unwrap();
    assert!(
        edit.commit().is_err(),
        "append-only publication exceeds the retained input ceiling"
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);

    let wide = Package::from_bytes(&bytes).unwrap();
    let mut edit = wide.edit().unwrap();
    edit.move_slide(Position::new(0), Position::new(2)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().after().len() > bytes.len());
    assert!(source.apply(commit.patch()).is_err());
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn canonical_transfer_and_text_batch_reuse_native_closure() {
    use litchi_ppt::package::edit::ShapeTextReplacement;
    let donor = package();
    let mut remove = donor.edit().unwrap();
    remove.remove_slide(Position::new(0)).unwrap();
    let receiver = remove.commit().unwrap().into_package();
    let plan = receiver
        .plan_transfer_from(&donor, Position::new(0))
        .unwrap();
    let mut edit = receiver.edit().unwrap();
    edit.insert_transfer(Position::new(0), &plan).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(slide_texts(commit.package()), slide_texts(&donor));
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        receiver.to_bytes().unwrap()
    );

    let source = package();
    let mut edit = source
        .edit_with_options(EditOptions::new().max_operations(2).unwrap())
        .unwrap();
    let replacements = [
        ShapeTextReplacement::new(Target::new(Position::new(0), Position::new(0)), "ALPHA"),
        ShapeTextReplacement::new(Target::new(Position::new(1), Position::new(0)), "BETA"),
    ];
    assert_eq!(edit.set_shape_texts(&replacements).unwrap(), 2);
    assert_eq!(edit.operation_count(), 2);
    assert!(matches!(
        edit.remove_slide(Position::new(2)),
        Err(Error::Limit { .. })
    ));
    let commit = edit.commit().unwrap();
    assert_eq!(
        slide_texts(commit.package()),
        vec!["ALPHA", "BETA", "gamma"]
    );
}
