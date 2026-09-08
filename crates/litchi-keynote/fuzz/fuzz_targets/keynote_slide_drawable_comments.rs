#![no_main]

//! Bounded selector-first fuzzing for Keynote direct-drawable comments.
//!
//! The target keeps comment graphs behind the semantic drawable and reply
//! selectors.  It reads every supported drawable, exercises root and reply
//! edits, and requires exact source preservation on rejected selectors,
//! hostile graphs, and tight semantic budgets.  Successful edits must reopen,
//! replay, invert, and double-invert byte-for-byte.  A private archive oracle
//! also mutates the checked-in native comment component in bounded ways so
//! cycle, missing, duplicate, unknown-payload, and unknown-header paths stay
//! reachable.

mod support;

#[path = "../../tests/support/drawable_comment_fixtures.rs"]
mod drawable_comment_fixtures;

use std::{collections::BTreeSet, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_archive::{Limits as ArchiveLimits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{tsd, tsp};
use litchi_keynote::{
    DrawableSelector, DrawableSummary, Limits, Package, ReadOptions, ReplySelector, SemanticLimits,
    SlideSelector,
};
use prost::Message as _;

use self::support::{
    MAX_ENTRY_BYTES, MAX_EXPANDED_BYTES, MAX_INPUT_BYTES, MAX_IWA_STREAM_BYTES, MAX_OBJECTS,
    MAX_PACKAGE_BYTES, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, control,
    observe_error, package_bytes, read_options,
};

use self::drawable_comment_fixtures as fixtures;

const TARGET_SET: &[u8] = b"target-drawable-comment-set";
const TARGET_CLEAR: &[u8] = b"target-drawable-comment-clear";
const TARGET_REPLY: &[u8] = b"target-drawable-comment-reply";
const TARGET_SET_REPLY: &[u8] = b"target-drawable-comment-set-reply";
const TARGET_REMOVE_REPLY: &[u8] = b"target-drawable-comment-remove-reply";
const NATIVE_ROOT_TEXT: &str = "native drawable contract root — 北区";
const NATIVE_REPLY_TEXT: &str = "native drawable contract reply — 北区";
const NATIVE_MAX_ENTRIES: usize = 16 * 1024;
const NATIVE_MAX_REFERENCES: usize = 256 * 1024;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const NATIVE_KEYNOTE: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/drawable-comments-source-native.key");

fuzz_target!(|data: &[u8]| {
    exercise_arbitrary_input(data);
    exercise_native_seed(data);
    exercise_cross_component_seed(data);
    exercise_hostile_native_once();
    exercise_limit_budget_once();
});

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, native_read_options())
            .unwrap_or_else(|error| panic!("native drawable-comments seed must open: {error}"));
        verify_native_contract(&package);
        package
    })
}

fn native_read_options() -> ReadOptions {
    // The native comment routes perform a package-wide ownership census.  The
    // checked fixture contains 990 IWA objects, so the shared 256-entry fuzz
    // profile would reject this deterministic seed before semantic coverage.
    // Keep the widened profile local to native and hostile-native inputs.
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_PACKAGE_BYTES,
            NATIVE_MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid native archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            NATIVE_MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid native semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn exercise_arbitrary_input(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }
    match Package::from_bytes_with_options(data, read_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }
}

fn exercise_native_seed(data: &[u8]) {
    exercise_package(native_package(), data);
    exercise_package(native_reply_package(), data);
}

fn verify_native_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    let inventory = read_inventory(package, &source_bytes)
        .unwrap_or_else(|| panic!("native drawable-comments inventory must read"));
    assert!(!inventory.is_empty());
    let target = inventory
        .iter()
        .find(|summary| summary.has_comment())
        .unwrap_or_else(|| panic!("native drawable-comments seed has no root comment"))
        .selector();
    let commit = package
        .edit_slide_drawable_comment(SlideSelector::index(0), target)
        .unwrap_or_else(|error| panic!("native root contract edit must stage: {error}"))
        .set(NATIVE_ROOT_TEXT)
        .unwrap_or_else(|error| panic!("native root contract set must stage: {error}"))
        .commit()
        .unwrap_or_else(|error| panic!("native root contract edit must commit: {error}"));
    assert!(
        !commit.patch().is_noop(),
        "native root contract edit must change the source"
    );
    assert_eq!(
        commit
            .package()
            .slide_drawable_comment(SlideSelector::index(0), target)
            .unwrap_or_else(|error| panic!("native root contract read must succeed: {error}"))
            .as_ref()
            .map(|comment| comment.text()),
        Some(NATIVE_ROOT_TEXT)
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn native_reply_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = native_package();
        let source_bytes = package_bytes(package);
        let inventory = read_inventory(package, &source_bytes)
            .unwrap_or_else(|| panic!("native reply seed inventory must read"));
        let target = inventory
            .iter()
            .find(|summary| summary.has_comment())
            .unwrap_or_else(|| panic!("native reply seed has no root comment"))
            .selector();
        let commit = package
            .edit_slide_drawable_comment(SlideSelector::index(0), target)
            .unwrap_or_else(|error| panic!("native reply seed edit must stage: {error}"))
            .add_reply(NATIVE_REPLY_TEXT)
            .unwrap_or_else(|error| panic!("native reply seed add must stage: {error}"))
            .commit()
            .unwrap_or_else(|error| panic!("native reply seed add must commit: {error}"));
        assert!(
            !commit.patch().is_noop(),
            "native reply seed add must change the source"
        );
        let candidate = commit.package().clone();
        let replies = candidate
            .slide_drawable_comment_replies(SlideSelector::index(0), target)
            .unwrap_or_else(|error| panic!("native reply seed read must succeed: {error}"));
        assert_eq!(
            replies.last().map(|reply| reply.text()),
            Some(NATIVE_REPLY_TEXT)
        );
        assert_eq!(package_bytes(package), source_bytes);
        candidate
    })
}

#[derive(Debug, Clone, Copy)]
enum CrossComponentKind {
    ForeignRoot,
    ForeignReply,
    SharedForeignRoot,
    SharedForeignReply,
    ForeignDrawable,
}

struct CrossComponentSeed {
    package: Package,
    kind: CrossComponentKind,
    target_position: usize,
    sibling_position: Option<usize>,
}

fn exercise_cross_component_seed(data: &[u8]) {
    let seed = match usize::from(control(data, 5)) % 5 {
        0 => foreign_root_seed(),
        1 => foreign_reply_seed(),
        2 => shared_foreign_root_seed(),
        3 => shared_foreign_reply_seed(),
        _ => foreign_drawable_seed(),
    };
    // The fixture is selected per input so one fuzz iteration pays for one
    // package-wide semantic census. The seed initializer performs the fixed
    // success contract once, while this path keeps arbitrary operations and
    // source-preserving replay under fuzz control.
    exercise_package(&seed.package, data);
}

fn foreign_root_seed() -> &'static CrossComponentSeed {
    static SEED: OnceLock<CrossComponentSeed> = OnceLock::new();
    SEED.get_or_init(|| {
        build_cross_component_seed(
            fixtures::root_foreign_fixture,
            CrossComponentKind::ForeignRoot,
            "foreign-root",
        )
    })
}

fn foreign_reply_seed() -> &'static CrossComponentSeed {
    static SEED: OnceLock<CrossComponentSeed> = OnceLock::new();
    SEED.get_or_init(|| {
        build_cross_component_seed(
            fixtures::reply_foreign_fixture,
            CrossComponentKind::ForeignReply,
            "foreign-reply",
        )
    })
}

fn shared_foreign_root_seed() -> &'static CrossComponentSeed {
    static SEED: OnceLock<CrossComponentSeed> = OnceLock::new();
    SEED.get_or_init(|| {
        build_cross_component_seed(
            fixtures::shared_foreign_root_fixture,
            CrossComponentKind::SharedForeignRoot,
            "shared-foreign-root",
        )
    })
}

fn shared_foreign_reply_seed() -> &'static CrossComponentSeed {
    static SEED: OnceLock<CrossComponentSeed> = OnceLock::new();
    SEED.get_or_init(|| {
        build_cross_component_seed(
            fixtures::shared_foreign_reply_fixture,
            CrossComponentKind::SharedForeignReply,
            "shared-foreign-reply",
        )
    })
}

fn foreign_drawable_seed() -> &'static CrossComponentSeed {
    static SEED: OnceLock<CrossComponentSeed> = OnceLock::new();
    SEED.get_or_init(|| {
        build_cross_component_seed(
            fixtures::foreign_drawable_fixture,
            CrossComponentKind::ForeignDrawable,
            "foreign-drawable",
        )
    })
}

fn build_cross_component_seed(
    build: fn() -> fixtures::TestResult<fixtures::RelocatedFixture>,
    kind: CrossComponentKind,
    label: &str,
) -> CrossComponentSeed {
    let fixture = build().unwrap_or_else(|error| panic!("{label} fixture must build: {error}"));
    let moved = match kind {
        CrossComponentKind::ForeignDrawable => fixture.target.identifier,
        CrossComponentKind::ForeignReply | CrossComponentKind::SharedForeignReply => fixture
            .reply_identifier
            .unwrap_or_else(|| panic!("{label} fixture must expose its relocated reply")),
        CrossComponentKind::ForeignRoot | CrossComponentKind::SharedForeignRoot => fixture
            .root_identifier
            .unwrap_or_else(|| panic!("{label} fixture must expose its relocated root")),
    };
    fixtures::assert_metadata_relocated(&fixture.metadata_source, &fixture.bytes, moved)
        .unwrap_or_else(|error| panic!("{label} fixture metadata must relocate: {error}"));
    let target_position = fixture.target.position;
    let sibling_position = fixture.sibling.as_ref().map(|target| target.position);
    let package = Package::from_bytes_with_options(&fixture.bytes, native_read_options())
        .unwrap_or_else(|error| panic!("{label} fixture must open: {error}"));
    let seed = CrossComponentSeed {
        package,
        kind,
        target_position,
        sibling_position,
    };
    verify_cross_component_contract(&seed, label);
    seed
}

fn verify_cross_component_contract(seed: &CrossComponentSeed, label: &str) {
    let source_bytes = package_bytes(&seed.package);
    let inventory = read_inventory(&seed.package, &source_bytes)
        .unwrap_or_else(|| panic!("{label} fixture inventory must read"));
    let target = inventory
        .get(seed.target_position)
        .unwrap_or_else(|| panic!("{label} fixture target is outside source order"));
    let slide = SlideSelector::index(0);
    let selector = target.selector();
    match seed.kind {
        CrossComponentKind::ForeignRoot => {
            assert!(target.has_comment(), "{label} fixture must retain its root");
            let before = seed
                .package
                .slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} root read must succeed: {error}"));
            let commit = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} edit must stage: {error}"))
                .set("fuzz foreign root contract")
                .unwrap_or_else(|error| panic!("{label} set must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} set must commit: {error}"));
            assert_eq!(
                commit
                    .package()
                    .slide_drawable_comment(slide, selector)
                    .unwrap_or_else(|error| panic!("{label} candidate read failed: {error}"))
                    .as_ref()
                    .map(|comment| comment.text()),
                Some("fuzz foreign root contract")
            );
            assert_ne!(
                before,
                commit
                    .package()
                    .slide_drawable_comment(slide, selector)
                    .unwrap()
            );
            let restored = commit
                .package()
                .apply_slide_drawable_comment(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("{label} inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);
        },
        CrossComponentKind::ForeignReply => {
            assert!(target.has_comment(), "{label} fixture must retain its root");
            assert!(
                target.reply_count() > 0,
                "{label} fixture must retain its reply"
            );
            let before = seed
                .package
                .slide_drawable_comment_replies(slide, selector)
                .unwrap_or_else(|error| panic!("{label} replies read must succeed: {error}"));
            let registered_set = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} registered edit must stage: {error}"))
                .set_reply(ReplySelector::index(0), "fuzz registered foreign reply")
                .unwrap_or_else(|error| panic!("{label} registered set must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} registered set must commit: {error}"));
            let after_registered_set = registered_set
                .package()
                .slide_drawable_comment_replies(slide, selector)
                .unwrap_or_else(|error| {
                    panic!("{label} registered set candidate read failed: {error}")
                });
            assert_eq!(after_registered_set.len(), before.len());
            assert_eq!(
                after_registered_set.first().map(|reply| reply.text()),
                Some("fuzz registered foreign reply")
            );
            let restored_registered_set = registered_set
                .package()
                .apply_slide_drawable_comment(&registered_set.patch().inverse())
                .unwrap_or_else(|error| panic!("{label} registered set inverse failed: {error}"));
            assert_eq!(
                package_bytes(restored_registered_set.package()),
                source_bytes
            );

            let registered_remove = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} registered remove must stage: {error}"))
                .remove_reply(ReplySelector::index(0))
                .unwrap_or_else(|error| panic!("{label} registered remove must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} registered remove must commit: {error}"));
            let after_registered_remove = registered_remove
                .package()
                .slide_drawable_comment_replies(slide, selector)
                .unwrap_or_else(|error| {
                    panic!("{label} registered remove candidate read failed: {error}")
                });
            assert_eq!(after_registered_remove.len(), before.len() - 1);
            let restored_registered_remove = registered_remove
                .package()
                .apply_slide_drawable_comment(&registered_remove.patch().inverse())
                .unwrap_or_else(|error| {
                    panic!("{label} registered remove inverse failed: {error}")
                });
            assert_eq!(
                package_bytes(restored_registered_remove.package()),
                source_bytes
            );

            let commit = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} edit must stage: {error}"))
                .add_reply("fuzz foreign reply contract")
                .unwrap_or_else(|error| panic!("{label} add must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} add must commit: {error}"));
            let after = commit
                .package()
                .slide_drawable_comment_replies(slide, selector)
                .unwrap_or_else(|error| panic!("{label} candidate replies failed: {error}"));
            assert_eq!(after.len(), before.len() + 1);
            assert_eq!(
                after.last().map(|reply| reply.text()),
                Some("fuzz foreign reply contract")
            );
            let restored = commit
                .package()
                .apply_slide_drawable_comment(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("{label} inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);
        },
        CrossComponentKind::SharedForeignRoot => {
            assert!(target.has_comment(), "{label} fixture must retain its root");
            let sibling_position = seed
                .sibling_position
                .unwrap_or_else(|| panic!("{label} fixture must have a sibling"));
            let sibling = DrawableSelector::index(sibling_position);
            let before_sibling = seed
                .package
                .slide_drawable_comment(slide, sibling)
                .unwrap_or_else(|error| panic!("{label} sibling read failed: {error}"));
            let commit = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} edit must stage: {error}"))
                .set("fuzz shared foreign root contract")
                .unwrap_or_else(|error| panic!("{label} set must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} set must commit: {error}"));
            assert_eq!(
                commit
                    .package()
                    .slide_drawable_comment(slide, selector)
                    .unwrap_or_else(|error| panic!("{label} selected read failed: {error}"))
                    .as_ref()
                    .map(|comment| comment.text()),
                Some("fuzz shared foreign root contract")
            );
            assert_eq!(
                commit
                    .package()
                    .slide_drawable_comment(slide, sibling)
                    .unwrap_or_else(|error| panic!(
                        "{label} sibling candidate read failed: {error}"
                    )),
                before_sibling
            );
            let restored = commit
                .package()
                .apply_slide_drawable_comment(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("{label} inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);
        },
        CrossComponentKind::SharedForeignReply => {
            assert!(target.has_comment(), "{label} fixture must retain its root");
            let sibling_position = seed
                .sibling_position
                .unwrap_or_else(|| panic!("{label} fixture must have a sibling"));
            let sibling = DrawableSelector::index(sibling_position);
            let before_sibling = seed
                .package
                .slide_drawable_comment_replies(slide, sibling)
                .unwrap_or_else(|error| panic!("{label} sibling replies failed: {error}"));
            assert!(
                !before_sibling.is_empty(),
                "{label} sibling must retain a reply"
            );
            let commit = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} edit must stage: {error}"))
                .clear()
                .unwrap_or_else(|error| panic!("{label} clear must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} clear must commit: {error}"));
            assert!(
                commit
                    .package()
                    .slide_drawable_comment(slide, selector)
                    .unwrap_or_else(|error| panic!("{label} selected read failed: {error}"))
                    .is_none()
            );
            assert_eq!(
                commit
                    .package()
                    .slide_drawable_comment_replies(slide, sibling)
                    .unwrap_or_else(|error| panic!(
                        "{label} sibling candidate replies failed: {error}"
                    )),
                before_sibling
            );
            let restored = commit
                .package()
                .apply_slide_drawable_comment(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("{label} inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);
        },
        CrossComponentKind::ForeignDrawable => {
            assert!(
                !target.has_comment(),
                "{label} fixture starts without a root"
            );
            let commit = seed
                .package
                .edit_slide_drawable_comment(slide, selector)
                .unwrap_or_else(|error| panic!("{label} edit must stage: {error}"))
                .set("fuzz foreign drawable contract")
                .unwrap_or_else(|error| panic!("{label} set must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| panic!("{label} set must commit: {error}"));
            assert_eq!(
                commit
                    .package()
                    .slide_drawable_comment(slide, selector)
                    .unwrap_or_else(|error| panic!("{label} candidate read failed: {error}"))
                    .as_ref()
                    .map(|comment| comment.text()),
                Some("fuzz foreign drawable contract")
            );
            let restored = commit
                .package()
                .apply_slide_drawable_comment(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("{label} inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);
        },
    }
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let Some(inventory) = read_inventory(package, &source_bytes) else {
        // Accepted hostile comment references, cycles, and malformed payloads
        // are rejected by the semantic reader without publishing anything.
        return;
    };
    let Some(existing) = inventory.iter().find(|summary| summary.has_comment()) else {
        exercise_invalid_selectors(package, &source_bytes);
        return;
    };
    let existing_selector = existing.selector();
    let existing_replies = existing.reply_count();
    exercise_edit(
        package,
        existing_selector,
        existing_replies,
        false,
        data,
        &source_bytes,
    );

    if let Some(empty) = inventory.iter().find(|summary| !summary.has_comment()) {
        exercise_edit(package, empty.selector(), 0, true, data, &source_bytes);
    }
    exercise_invalid_selectors(package, &source_bytes);
    assert_eq!(package_bytes(package), source_bytes);
}

fn read_inventory(package: &Package, source_bytes: &[u8]) -> Option<Box<[DrawableSummary]>> {
    let summaries = match package.slide_drawables(SlideSelector::index(0)) {
        Ok(summaries) => summaries,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return None;
        },
    };
    for summary in &summaries {
        let selector = summary.selector();
        let comment = match package.slide_drawable_comment(SlideSelector::index(0), selector) {
            Ok(comment) => comment,
            Err(error) => {
                observe_error(error);
                assert_eq!(package_bytes(package), source_bytes);
                return None;
            },
        };
        let replies =
            match package.slide_drawable_comment_replies(SlideSelector::index(0), selector) {
                Ok(replies) => replies,
                Err(error) => {
                    observe_error(error);
                    assert_eq!(package_bytes(package), source_bytes);
                    return None;
                },
            };
        assert_eq!(comment.is_some(), summary.has_comment());
        assert_eq!(replies.len(), summary.reply_count());
        if comment.is_none() {
            assert!(replies.is_empty());
        }
        for reply in &replies {
            assert!(reply.text().len() <= 64 * 1024 * 1024);
        }
    }
    Some(summaries)
}

#[derive(Clone, Debug)]
enum Operation {
    Set(String),
    Clear,
    AddReply(String),
    SetReply(usize, String),
    RemoveReply(usize),
}

fn exercise_edit(
    package: &Package,
    selector: DrawableSelector,
    reply_count: usize,
    empty_target: bool,
    data: &[u8],
    source_bytes: &[u8],
) {
    let slide = SlideSelector::index(0);
    let before_comment = match package.slide_drawable_comment(slide, selector) {
        Ok(comment) => comment,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let before_replies = match package.slide_drawable_comment_replies(slide, selector) {
        Ok(replies) => replies,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    assert_eq!(before_replies.len(), reply_count);
    let operation = requested_operation(
        data,
        before_comment.is_some(),
        before_replies.len(),
        empty_target,
    );
    let edit = match package.edit_slide_drawable_comment(slide, selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let staged = match &operation {
        Operation::Set(text) => edit.set(text),
        Operation::Clear => edit.clear(),
        Operation::AddReply(text) => edit.add_reply(text),
        Operation::SetReply(index, text) => edit.set_reply(ReplySelector::index(*index), text),
        Operation::RemoveReply(index) => edit.remove_reply(ReplySelector::index(*index)),
    };
    let edit = match staged {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let patch = commit.patch().clone();
    let candidate_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), before_comment.as_ref());
    assert_eq!(patch.is_noop(), candidate_bytes == source_bytes);
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    assert_operation_state(
        commit.package(),
        selector,
        &operation,
        &before_comment,
        &before_replies,
    );
    let candidate_comment = commit
        .package()
        .slide_drawable_comment(slide, selector)
        .unwrap_or_else(|error| panic!("candidate comment read must succeed: {error}"));
    assert_eq!(patch.after(), candidate_comment.as_ref());
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("drawable-comments candidate must validate: {error}"));
    assert_eq!(package_bytes(package), source_bytes);

    let replay = package
        .apply_slide_drawable_comment(&patch)
        .unwrap_or_else(|error| panic!("drawable-comments forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = commit
        .package()
        .apply_slide_drawable_comment(&inverse)
        .unwrap_or_else(|error| panic!("drawable-comments inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_state_matches_source(
        restored.package(),
        slide,
        selector,
        &before_comment,
        &before_replies,
    );

    let twice = package
        .apply_slide_drawable_comment(&inverse.inverse())
        .unwrap_or_else(|error| panic!("drawable-comments double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    if !patch.is_noop() {
        assert!(
            commit
                .package()
                .apply_slide_drawable_comment(&patch)
                .is_err(),
            "changed drawable-comments patch must reject its post-state source"
        );
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn requested_operation(
    data: &[u8],
    has_comment: bool,
    reply_count: usize,
    empty_target: bool,
) -> Operation {
    if data.starts_with(TARGET_SET_REPLY) {
        return if reply_count == 0 {
            Operation::AddReply("fuzz drawable reply — 北区".to_owned())
        } else {
            Operation::SetReply(reply_count - 1, "fuzz reply updated — 北区".to_owned())
        };
    }
    if data.starts_with(TARGET_SET) {
        return Operation::Set("fuzz drawable root set — 北区".to_owned());
    }
    if data.starts_with(TARGET_CLEAR) {
        return Operation::Clear;
    }
    if data.starts_with(TARGET_REPLY) {
        return Operation::AddReply("fuzz drawable reply — 北区".to_owned());
    }
    if data.starts_with(TARGET_REMOVE_REPLY) {
        return if reply_count == 0 {
            Operation::Clear
        } else {
            Operation::RemoveReply(reply_count - 1)
        };
    }
    if empty_target || !has_comment {
        return if control(data, 0) & 1 == 0 {
            Operation::Set(text_from_data(data, 2, "fuzz drawable root"))
        } else {
            Operation::Clear
        };
    }
    match control(data, 0) % 5 {
        0 => Operation::Set(text_from_data(data, 2, "fuzz drawable root")),
        1 => Operation::Clear,
        2 => Operation::AddReply(text_from_data(data, 3, "fuzz drawable reply")),
        3 if reply_count > 0 => Operation::SetReply(
            usize::from(control(data, 1)) % reply_count,
            text_from_data(data, 4, "fuzz reply updated"),
        ),
        4 if reply_count > 0 => Operation::RemoveReply(usize::from(control(data, 1)) % reply_count),
        _ => Operation::Set(text_from_data(data, 2, "fuzz drawable root")),
    }
}

fn text_from_data(data: &[u8], offset: usize, fallback: &str) -> String {
    let start = offset.min(data.len());
    let length = usize::from(control(data, offset)) % 96;
    let end = start.saturating_add(length).min(data.len());
    let mut text = String::from(fallback);
    if start < end {
        text.push(' ');
        text.push_str(&String::from_utf8_lossy(&data[start..end]));
    }
    while text.len() > 256 {
        // `String::truncate` requires a character boundary; lossy UTF-8 input
        // can otherwise leave the 256-byte cap in the middle of a code point.
        text.pop();
    }
    text
}

fn assert_operation_state(
    package: &Package,
    selector: DrawableSelector,
    operation: &Operation,
    before_comment: &Option<litchi_keynote::Comment>,
    before_replies: &[litchi_keynote::Reply],
) {
    let slide = SlideSelector::index(0);
    let comment = package
        .slide_drawable_comment(slide, selector)
        .unwrap_or_else(|error| panic!("candidate root comment read must succeed: {error}"));
    let replies = package
        .slide_drawable_comment_replies(slide, selector)
        .unwrap_or_else(|error| panic!("candidate replies read must succeed: {error}"));
    match operation {
        Operation::Set(text) => {
            assert_eq!(
                comment.as_ref().map(|comment| comment.text()),
                Some(text.as_str())
            );
            if let (Some(actual), Some(before)) = (comment.as_ref(), before_comment.as_ref()) {
                assert_eq!(actual.timestamp(), before.timestamp());
                assert_eq!(actual.author(), before.author());
            }
            assert_eq!(replies.len(), before_replies.len());
            assert_reply_prefix(&replies, before_replies);
        },
        Operation::Clear => {
            assert!(comment.is_none());
            assert!(replies.is_empty());
        },
        Operation::AddReply(text) => {
            assert_eq!(comment.as_ref(), before_comment.as_ref());
            assert_eq!(replies.len(), before_replies.len() + 1);
            assert_eq!(
                replies.last().map(|reply| reply.text()),
                Some(text.as_str())
            );
            assert_reply_prefix(&replies, before_replies);
        },
        Operation::SetReply(index, text) => {
            assert_eq!(comment.as_ref(), before_comment.as_ref());
            assert_eq!(replies.len(), before_replies.len());
            assert_eq!(replies[*index].text(), text);
            for (position, reply) in replies.iter().enumerate() {
                if position == *index {
                    assert_eq!(reply.timestamp(), before_replies[position].timestamp());
                    assert_eq!(reply.author(), before_replies[position].author());
                } else {
                    assert_eq!(reply, &before_replies[position]);
                }
            }
        },
        Operation::RemoveReply(index) => {
            assert_eq!(comment.as_ref(), before_comment.as_ref());
            assert_eq!(replies.len() + 1, before_replies.len());
            for (position, reply) in replies.iter().enumerate() {
                let source_position = if position < *index {
                    position
                } else {
                    position + 1
                };
                assert_eq!(reply, &before_replies[source_position]);
            }
        },
    }
}

fn assert_reply_prefix(actual: &[litchi_keynote::Reply], before: &[litchi_keynote::Reply]) {
    assert!(actual.len() >= before.len());
    for (actual, before) in actual.iter().zip(before.iter()) {
        assert_eq!(actual, before);
    }
}

fn assert_state_matches_source(
    package: &Package,
    slide: SlideSelector,
    selector: DrawableSelector,
    comment: &Option<litchi_keynote::Comment>,
    replies: &[litchi_keynote::Reply],
) {
    let actual = package
        .slide_drawable_comment(slide, selector)
        .unwrap_or_else(|error| panic!("restored root comment read must succeed: {error}"));
    assert_eq!(actual.as_ref(), comment.as_ref());
    let actual_replies = package
        .slide_drawable_comment_replies(slide, selector)
        .unwrap_or_else(|error| panic!("restored replies read must succeed: {error}"));
    assert_eq!(actual_replies.len(), replies.len());
    assert_eq!(actual_replies.as_ref(), replies);
}

fn exercise_invalid_selectors(package: &Package, source_bytes: &[u8]) {
    let slide = SlideSelector::index(0);
    let invalid_drawable = DrawableSelector::index(usize::MAX);
    assert!(
        package
            .slide_drawable_comment(slide, invalid_drawable)
            .is_err()
    );
    assert!(
        package
            .slide_drawable_comment_replies(slide, invalid_drawable)
            .is_err()
    );
    assert!(
        package
            .edit_slide_drawable_comment(slide, invalid_drawable)
            .is_err()
    );
    assert!(
        package
            .slide_drawables(SlideSelector::index(usize::MAX))
            .is_err()
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_limit_budget_once() {
    static RUN: OnceLock<()> = OnceLock::new();
    RUN.get_or_init(|| {
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            1,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| panic!("finite comment semantic limit must build: {error}"));
        match Package::from_bytes_with_options(
            NATIVE_KEYNOTE,
            ReadOptions::new(Limits::default(), semantic),
        ) {
            Err(error) => observe_error(error),
            Ok(package) => {
                let source_bytes = package_bytes(&package);
                let result = package
                    .edit_slide_drawable_comment(
                        SlideSelector::index(0),
                        DrawableSelector::index(0),
                    )
                    .and_then(|edit| edit.set("budget probe"))
                    .and_then(|edit| edit.commit());
                assert!(result.is_err(), "tight comment budget must reject mutation");
                assert_eq!(package_bytes(&package), source_bytes);
            },
        }
    });
}

#[derive(Debug, Clone, Copy)]
enum HostileMutation {
    Cycle,
    MissingReply,
    DuplicateReply,
    UnknownPayload,
    UnknownHeader,
}

fn exercise_hostile_native_once() {
    static RUN: OnceLock<()> = OnceLock::new();
    RUN.get_or_init(|| {
        for mutation in [
            HostileMutation::Cycle,
            HostileMutation::MissingReply,
            HostileMutation::DuplicateReply,
            HostileMutation::UnknownPayload,
            HostileMutation::UnknownHeader,
        ] {
            let hostile = match mutation {
                HostileMutation::UnknownHeader => mutate_native_comment_header(NATIVE_KEYNOTE),
                _ => mutate_native_comment(NATIVE_KEYNOTE, mutation),
            }
            .unwrap_or_else(|| panic!("native fixture must support {mutation:?}"));
            let package = match Package::from_bytes_with_options(&hostile, native_read_options()) {
                Ok(package) => package,
                Err(error) => {
                    observe_error(error);
                    assert!(
                        !matches!(mutation, HostileMutation::UnknownPayload),
                        "unknown comment payload must remain readable"
                    );
                    continue;
                },
            };
            let source_bytes = package_bytes(&package);
            match mutation {
                HostileMutation::UnknownPayload => {
                    assert_unknown_payload_is_readable(&package, &source_bytes);
                },
                HostileMutation::UnknownHeader => {
                    assert_hostile_graph_is_rejected(&package, &source_bytes);
                },
                HostileMutation::Cycle
                | HostileMutation::MissingReply
                | HostileMutation::DuplicateReply => {
                    assert_hostile_graph_is_rejected(&package, &source_bytes);
                },
            }
        }
    });
}

fn assert_hostile_graph_is_rejected(package: &Package, source_bytes: &[u8]) {
    let summaries = match package.slide_drawables(SlideSelector::index(0)) {
        Ok(summaries) => summaries,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let mut rejected = false;
    for summary in &summaries {
        let selector = summary.selector();
        if package
            .slide_drawable_comment_replies(SlideSelector::index(0), selector)
            .is_err()
        {
            rejected = true;
            break;
        }
        let comment = match package.slide_drawable_comment(SlideSelector::index(0), selector) {
            Ok(comment) => comment,
            Err(_) => {
                rejected = true;
                break;
            },
        };
        let Some(comment) = comment else {
            continue;
        };
        let result = package
            .edit_slide_drawable_comment(SlideSelector::index(0), selector)
            .and_then(|edit| edit.set(comment.text()))
            .and_then(|edit| edit.commit());
        if result.is_err() {
            rejected = true;
            break;
        }
    }
    assert!(
        rejected,
        "hostile comment graph was accepted by every public route"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn assert_unknown_payload_is_readable(package: &Package, source_bytes: &[u8]) {
    let summaries = package
        .slide_drawables(SlideSelector::index(0))
        .unwrap_or_else(|error| panic!("unknown comment payload inventory must read: {error}"));
    for summary in &summaries {
        let selector = summary.selector();
        let comment = package
            .slide_drawable_comment(SlideSelector::index(0), selector)
            .unwrap_or_else(|error| panic!("unknown comment payload root must read: {error}"));
        package
            .slide_drawable_comment_replies(SlideSelector::index(0), selector)
            .unwrap_or_else(|error| panic!("unknown comment payload replies must read: {error}"));
        if let Some(comment) = comment {
            let noop = package
                .edit_slide_drawable_comment(SlideSelector::index(0), selector)
                .and_then(|edit| edit.set(comment.text()))
                .unwrap_or_else(|error| panic!("unknown comment payload no-op must stage: {error}"))
                .commit()
                .unwrap_or_else(|error| {
                    panic!("unknown comment payload no-op must commit: {error}")
                });
            assert!(noop.patch().is_noop());
            assert_eq!(package_bytes(noop.package()), source_bytes);
        }
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn mutate_native_comment(source: &[u8], mutation: HostileMutation) -> Option<Vec<u8>> {
    if matches!(mutation, HostileMutation::UnknownHeader) {
        return None;
    }
    let catalog = Catalog::from_bytes(source).ok()?;
    let mut comments = Vec::new();
    let mut referenced = BTreeSet::new();
    let mut duplicate_reply = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let Ok(stream) = SnappyStream::decompress(entry.data()) else {
            continue;
        };
        let Ok(archive) = Archive::parse(&stream.into_bytes()) else {
            continue;
        };
        for object in &archive.objects {
            let Some((message_index, message)) = object
                .messages
                .iter()
                .enumerate()
                .find(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            else {
                continue;
            };
            let Ok(decoded) = tsd::CommentStorageArchive::decode(message.data.as_slice()) else {
                continue;
            };
            for reply in &decoded.replies {
                referenced.insert(reply.identifier);
                duplicate_reply.get_or_insert(reply.identifier);
            }
            comments.push((
                entry.name().to_owned(),
                object.archive_info.identifier?,
                message_index,
                decoded.replies.first().map(|reply| reply.identifier),
            ));
        }
    }
    let target = comments
        .iter()
        .find(|(_, identifier, _, _)| !referenced.contains(identifier))
        .cloned()
        .or_else(|| comments.first().cloned())?;
    let (component, identifier, message_index, target_reply) = target;
    let reference_identifier = match mutation {
        HostileMutation::Cycle => identifier,
        HostileMutation::MissingReply => u64::MAX - 19,
        HostileMutation::DuplicateReply => target_reply
            .or_else(|| {
                comments
                    .iter()
                    .map(|(_, candidate, _, _)| *candidate)
                    .find(|candidate| *candidate != identifier)
            })
            .or(duplicate_reply)
            .unwrap_or(identifier),
        HostileMutation::UnknownPayload => 0,
        HostileMutation::UnknownHeader => return None,
    };

    let mut replacement = None;
    for entry in catalog.iter() {
        if entry.name() != component {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data()).ok()?;
        let mut archive = Archive::parse(&stream.into_bytes()).ok()?;
        let object = archive.object_mut(identifier)?;
        let message = object.messages.get_mut(message_index)?;
        let mut payload = message.data.clone();
        if matches!(mutation, HostileMutation::UnknownPayload) {
            append_length_delimited_field(&mut payload, 9_001, b"fuzz unknown comment metadata")
                .ok()?;
        } else {
            let reference = tsp::Reference {
                identifier: reference_identifier,
                ..Default::default()
            }
            .encode_to_vec();
            let copies = usize::from(matches!(mutation, HostileMutation::DuplicateReply)) + 1;
            for _ in 0..copies {
                append_length_delimited_field(&mut payload, 4, &reference).ok()?;
                object
                    .archive_info
                    .message_infos
                    .get_mut(message_index)?
                    .object_references
                    .push(reference_identifier);
            }
        }
        message.data = payload;
        replacement = Some(SnappyStream::compress(&archive.to_bytes().ok()?).ok()?);
        break;
    }
    let replacement = replacement?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component {
                (entry.name(), replacement.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default()).ok()
}

/// Add one unknown field to the raw ArchiveInfo header of the same native
/// comment object selected by [`mutate_native_comment`].  Re-serializing an
/// [`Archive`] would canonicalize that header, so this narrow framing helper
/// edits the decompressed IWA bytes directly and lets the archive parser prove
/// that the resulting source remains well-framed.
fn mutate_native_comment_header(source: &[u8]) -> Option<Vec<u8>> {
    let catalog = Catalog::from_bytes(source).ok()?;
    let mut comments = Vec::new();
    let mut referenced = BTreeSet::new();
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let Ok(stream) = SnappyStream::decompress(entry.data()) else {
            continue;
        };
        let Ok(archive) = Archive::parse(&stream.into_bytes()) else {
            continue;
        };
        for object in &archive.objects {
            let Some(message) = object
                .messages
                .iter()
                .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            else {
                continue;
            };
            let Ok(decoded) = tsd::CommentStorageArchive::decode(message.data.as_slice()) else {
                continue;
            };
            for reply in &decoded.replies {
                referenced.insert(reply.identifier);
            }
            comments.push((entry.name().to_owned(), object.archive_info.identifier?));
        }
    }
    let (component, identifier) = comments
        .iter()
        .find(|(_, identifier)| !referenced.contains(identifier))
        .cloned()
        .or_else(|| comments.first().cloned())?;

    let mut replacement = None;
    for entry in catalog.iter() {
        if entry.name() != component {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data()).ok()?.into_bytes();
        let archive = Archive::parse(&stream).ok()?;
        let object = archive.object(identifier)?;
        let header_offset = usize::try_from(object.header_offset).ok()?;
        let data_offset = usize::try_from(object.data_offset).ok()?;
        let (header_length, prefix_length) = decode_varint(stream.get(header_offset..)?)?;
        let header_length = usize::try_from(header_length).ok()?;
        let header_start = header_offset.checked_add(prefix_length)?;
        let header_end = header_start.checked_add(header_length)?;
        if header_end != data_offset {
            return None;
        }
        let mut header = stream.get(header_start..header_end)?.to_vec();
        append_varint_field(&mut header, 9_002, 0xD00D).ok()?;
        let mut rewritten = Vec::with_capacity(
            stream
                .len()
                .saturating_add(header.len().saturating_sub(header_length)),
        );
        rewritten.extend_from_slice(stream.get(..header_offset)?);
        append_varint(&mut rewritten, u64::try_from(header.len()).ok()?);
        rewritten.extend_from_slice(&header);
        rewritten.extend_from_slice(stream.get(data_offset..)?);
        Archive::parse(&rewritten).ok()?;
        replacement = Some(SnappyStream::compress(&rewritten).ok()?);
        break;
    }
    let replacement = replacement?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component {
                (entry.name(), replacement.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default()).ok()
}

fn decode_varint(bytes: &[u8]) -> Option<(u64, usize)> {
    let mut value = 0_u64;
    for (index, byte) in bytes.iter().copied().enumerate().take(10) {
        if index == 9 && byte > 1 {
            return None;
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Some((value, index + 1));
        }
    }
    None
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}
