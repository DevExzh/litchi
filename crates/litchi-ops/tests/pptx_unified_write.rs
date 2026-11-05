//! Public immutable PPTX text writes through both native owner routes.
#![cfg(feature = "pptx")]
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};
const FIRST: &str = "/slide[1]/shape[1]";
fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    for texts in [["first", "other"], ["second", "untouched"]] {
        let slide = builder.presentation().unwrap().add_slide().unwrap();
        for text in texts {
            slide.add_text_box(text, 0, 0, 914_400, 914_400);
        }
    }
    builder.build().unwrap().to_bytes().unwrap()
}
fn open(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}
fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}

#[test]
fn shape_text_round_trips_and_preserves_input_for_both_owners() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let updated = source
            .set(FIRST, props! { "text" => "changed & <escaped>" })
            .unwrap();
        assert_eq!(text(&updated, FIRST), "changed & <escaped>");
        assert_eq!(text(&source, FIRST), "first");
        assert_eq!(source.to_bytes().unwrap(), before);
        assert_eq!(
            text(&open(&updated.to_bytes().unwrap(), eager), FIRST),
            "changed & <escaped>"
        );
        assert_eq!(text(&updated, "/slide[2]/shape[2]"), "untouched");
        let noop = source.set(FIRST, props! { "text" => "first" }).unwrap();
        assert_eq!(noop.to_bytes().unwrap(), before);
        let cleared = updated
            .set(FIRST, props! { "text" => litchi_ops::PropValue::Null })
            .unwrap();
        assert_eq!(text(&cleared, FIRST), "");
        assert_eq!(
            text(
                &cleared.set(FIRST, props! { "text" => "refilled" }).unwrap(),
                FIRST
            ),
            "refilled"
        );
    }
}

#[test]
fn scoped_selector_batches_and_failed_properties_are_atomic() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let updated = source
            .set("/slide[1]/*", props! { "text" => "both" })
            .unwrap();
        assert_eq!(text(&updated, FIRST), "both");
        assert_eq!(text(&updated, "/slide[1]/shape[2]"), "both");
        assert_eq!(text(&updated, "/slide[2]/shape[1]"), "second");
        assert!(
            source
                .set("shape", props! { "text" => "unscoped" })
                .is_err()
        );
        assert_eq!(
            source
                .set(
                    FIRST,
                    props! { "text" => "partial", "highlight" => "FFFF00" }
                )
                .unwrap_err()
                .code(),
            Code::UnsupportedProperty
        );
        assert_eq!(text(&source, FIRST), "first");
        assert_eq!(
            source
                .set("/slide[1]/*[@name='missing']", props! { "text" => "none" })
                .unwrap_err()
                .code(),
            Code::ZeroMatches
        );
        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            source
                .edit_with_options(
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                    |edit| edit.set("/slide[1]/*", props! { "text" => "too many" })
                )
                .unwrap_err()
                .code(),
            Code::LimitExceeded
        );
    }
}

#[test]
fn transactions_replay_inverse_poison_and_detect_net_noop() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let commit = source
            .edit(|edit| {
                edit.set(FIRST, props! { "text" => "staged" })?;
                edit.set("/slide[2]/shape[1]", props! { "text" => "second staged" })
            })
            .unwrap();
        assert!(commit.diagnostics().changed());
        assert_eq!(text(commit.file(), FIRST), "staged");
        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            commit.file().to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            before
        );
        assert!(commit.patch().apply(&open(&bytes, eager)).is_ok());
        let noop = source
            .edit(|edit| {
                edit.set(FIRST, props! { "text" => "temporary" })?;
                edit.set(FIRST, props! { "text" => "first" })
            })
            .unwrap();
        assert!(noop.patch().is_noop());
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.file().to_bytes().unwrap(), before);
        let poisoned = source.edit(|edit| {
            edit.set(FIRST, props! { "text" => "discarded" })?;
            let _ = edit.set(FIRST, props! { "highlight" => "FFFF00" });
            Ok(())
        });
        assert!(poisoned.is_err());
        assert_eq!(source.to_bytes().unwrap(), before);
        let limits = SemanticLimits::builder()
            .max_snapshot_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            source
                .edit_with_options(
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                    |edit| edit.set(FIRST, props! { "text" => "no budget" })
                )
                .unwrap_err()
                .code(),
            Code::LimitExceeded
        );
    }
}

#[test]
fn atomic_save_preserves_origin_and_no_clobber_policy() {
    let bytes = fixture();
    for eager in [false, true] {
        let dir = tempfile::tempdir().unwrap();
        let origin = dir.path().join("source.pptx");
        std::fs::write(&origin, &bytes).unwrap();
        let source = litchi_ops::open_with(
            &origin,
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let updated = source
            .set(FIRST, props! { "text" => "saved revision" })
            .unwrap();
        updated.save().unwrap();
        assert_eq!(
            text(&litchi_ops::open(&origin).unwrap(), FIRST),
            "saved revision"
        );
        assert_eq!(text(&source, FIRST), "first");
        let target = dir.path().join("target.pptx");
        std::fs::write(&target, b"keep").unwrap();
        let result = updated.save_with(
            litchi_ops::SaveTo::Path(&target),
            &litchi_ops::SaveOptions::new().overwrite(litchi_ops::Overwrite::Refuse),
        );
        assert!(result.is_err());
        assert_eq!(std::fs::read(&target).unwrap(), b"keep");
        assert_eq!(
            updated
                .save_as(dir.path().join("wrong.docx"))
                .unwrap_err()
                .code(),
            Code::InvalidArgument
        );
    }
}

#[test]
fn signed_and_released_source_policies_match_capabilities() {
    use litchi_ops::capabilities::Verb;
    let bytes = fixture();
    let released =
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().retain_source(false)).unwrap();
    assert!(!released.capabilities().verb(Verb::Set).is_supported());
    assert_eq!(
        released
            .set(FIRST, props! { "text" => "refused" })
            .unwrap_err()
            .code(),
        Code::UnsupportedCapability
    );
    let bytes =
        std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx").unwrap();
    let signed = open(&bytes, false);
    assert!(!signed.capabilities().verb(Verb::Set).is_supported());
    assert_eq!(signed.to_bytes().unwrap(), bytes);
    assert_eq!(
        signed
            .set(FIRST, props! { "text" => "refused" })
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );
    let File::Signed(wrapper) = signed else {
        panic!("signed fixture");
    };
    let unsigned = wrapper.clone().unsign();
    let updated = unsigned
        .set(FIRST, props! { "text" => "explicit unsigned" })
        .unwrap();
    assert_eq!(text(&updated, FIRST), "explicit unsigned");
    assert!(
        !litchi_ops::open_bytes(updated.to_bytes().unwrap())
            .unwrap()
            .is_signed()
    );
    assert!(
        wrapper
            .file()
            .set(FIRST, props! { "text" => "still refused" })
            .is_err()
    );
    assert_eq!(wrapper.file().to_bytes().unwrap(), bytes);
}

#[test]
fn invalid_text_and_staging_limits_leave_input_unchanged() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        assert_eq!(
            source
                .set(FIRST, props! { "text" => "bad\u{0001}" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        let limits = SemanticLimits::builder()
            .max_staged_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            source
                .edit_with_options(
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                    |edit| edit.set(FIRST, props! { "text" => "too large" })
                )
                .unwrap_err()
                .code(),
            Code::LimitExceeded
        );
        assert_eq!(text(&source, FIRST), "first");
    }
}

#[cfg(any(unix, windows))]
#[test]
fn unsigned_revision_keeps_source_freshness_for_reads_edits_and_save() {
    use std::io::Write;
    let dir = tempfile::tempdir().unwrap();
    let origin = dir.path().join("signed.pptx");
    std::fs::copy(
        "../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx",
        &origin,
    )
    .unwrap();
    let File::Signed(signed) = litchi_ops::open(&origin).unwrap() else {
        panic!("signed fixture")
    };
    let updated = signed
        .unsign()
        .set(FIRST, props! { "text" => "unsigned" })
        .unwrap();
    assert_eq!(text(&updated, FIRST), "unsigned");
    std::fs::OpenOptions::new()
        .append(true)
        .open(&origin)
        .unwrap()
        .write_all(b"changed")
        .unwrap();
    assert_eq!(updated.get(FIRST).unwrap_err().code(), Code::IoError);
    assert_eq!(
        updated
            .set(FIRST, props! { "text" => "again" })
            .unwrap_err()
            .code(),
        Code::IoError
    );
    assert_eq!(updated.to_bytes().unwrap_err().code(), Code::IoError);
    let target = dir.path().join("target.pptx");
    std::fs::write(&target, b"unchanged").unwrap();
    assert_eq!(updated.save_as(&target).unwrap_err().code(), Code::IoError);
    assert_eq!(std::fs::read(&target).unwrap(), b"unchanged");
}

#[test]
fn batches_keep_atomic_and_best_effort_semantics() {
    use litchi_ops::{BatchItem, BatchMode, ItemResult, Props};
    fn item(path: &'static str, props: Props<'static>) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(props);
        item
    }
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let commands = || {
            [
                item(FIRST, Props::new().with("text", "changed")),
                item(FIRST, Props::new().with("missingProperty", true)),
                item("/slide[2]/shape[1]", Props::new().with("text", "later")),
            ]
        };
        assert_eq!(
            source
                .batch(commands(), BatchMode::Atomic)
                .unwrap_err()
                .item(),
            Some(1)
        );
        assert_eq!(text(&source, FIRST), "first");
        let result = source.batch(commands(), BatchMode::BestEffort).unwrap();
        assert!(matches!(&result.value()[1], ItemResult::Failed(error) if error.item() == Some(1)));
        assert_eq!(text(result.file(), FIRST), "changed");
        assert_eq!(text(result.file(), "/slide[2]/shape[1]"), "later");
        assert_eq!(text(&source, FIRST), "first");
    }
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_publication_retains_profile_origin_and_plaintext_refusal() {
    use litchi_core::Durability;
    use litchi_ops::{Overwrite, Password, SaveOptions, SaveTo};
    use litchi_pptx::encryption::{self, Mode};
    let clear = fixture();
    let encrypted = encryption::encrypt(
        clear.clone(),
        "old",
        Mode::Standard,
        &encryption::Limits::default(),
    )
    .unwrap();
    for eager in [false, true] {
        let dir = tempfile::tempdir().unwrap();
        let origin = dir.path().join("encrypted.pptx");
        std::fs::write(&origin, &encrypted).unwrap();
        let locked = litchi_ops::open_with(
            &origin,
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let File::Locked(locked) = locked else {
            panic!("encrypted fixture")
        };
        let source = locked.unlock(Password::new("old".to_owned())).unwrap();
        assert_eq!(
            source.to_bytes().unwrap_err().code(),
            Code::DocumentProtected
        );
        assert_eq!(source.save().unwrap_err().code(), Code::DocumentProtected);
        assert_eq!(
            source
                .set(FIRST, props! { "text" => "refused" })
                .unwrap_err()
                .code(),
            Code::DocumentProtected
        );
        for durability in [Durability::Full, Durability::FileOnly, Durability::NoSync] {
            let target = dir.path().join("output.pptx");
            source
                .save_with(
                    SaveTo::Path(&target),
                    &SaveOptions::new()
                        .durability(durability)
                        .reencrypt(Password::new("new".to_owned())),
                )
                .unwrap();
            let result = encryption::open(
                std::fs::read(&target).unwrap(),
                "new",
                &encryption::Limits::default(),
            )
            .unwrap();
            assert_eq!(result.mode(), Some(Mode::Standard));
            assert_eq!(result.into_bytes(), clear);
            let bytes = std::fs::read(&target).unwrap();
            assert!(
                source
                    .save_with(
                        SaveTo::Path(&target),
                        &SaveOptions::new()
                            .overwrite(Overwrite::Refuse)
                            .reencrypt(Password::new("other".to_owned()))
                    )
                    .is_err()
            );
            assert_eq!(std::fs::read(&target).unwrap(), bytes);
            assert_eq!(
                source
                    .save_with(
                        SaveTo::Path(&target),
                        &SaveOptions::new().reencrypt(Password::new(String::new()))
                    )
                    .unwrap_err()
                    .code(),
                Code::InvalidArgument
            );
            assert_eq!(std::fs::read(&target).unwrap(), bytes);
        }
        assert_eq!(std::fs::read(&origin).unwrap(), encrypted);
        source
            .save_with(
                SaveTo::Origin,
                &SaveOptions::new().reencrypt(Password::new("new".to_owned())),
            )
            .unwrap();
        assert_eq!(
            encryption::open(
                std::fs::read(&origin).unwrap(),
                "new",
                &encryption::Limits::default(),
            )
            .unwrap()
            .into_bytes(),
            clear
        );
        assert_eq!(text(&source, FIRST), "first");
        assert_eq!(
            source.to_bytes().unwrap_err().code(),
            Code::DocumentProtected
        );
    }
}

#[cfg(feature = "encryption")]
#[test]
fn reencryption_rejects_plain_input_and_preserves_signed_encrypted_content() {
    use litchi_ops::{Password, SaveOptions, SaveTo};
    use litchi_pptx::encryption::{self, Mode};
    let dir = tempfile::tempdir().unwrap();
    let target = dir.path().join("output.pptx");
    std::fs::write(&target, b"keep").unwrap();
    let options = || SaveOptions::new().reencrypt(Password::new("new".to_owned()));
    for eager in [false, true] {
        assert_eq!(
            open(&fixture(), eager)
                .save_with(SaveTo::Path(&target), &options())
                .unwrap_err()
                .code(),
            Code::InvalidArgument
        );
        assert_eq!(std::fs::read(&target).unwrap(), b"keep");
    }
    let signed =
        std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx").unwrap();
    let encrypted = encryption::encrypt(
        signed.clone(),
        "old",
        Mode::Standard,
        &encryption::Limits::default(),
    )
    .unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(encrypted).unwrap() else {
        panic!("encrypted")
    };
    let source = locked.unlock(Password::new("old".to_owned())).unwrap();
    assert!(source.is_signed());
    source.save_with(SaveTo::Path(&target), &options()).unwrap();
    assert_eq!(
        encryption::open(
            std::fs::read(&target).unwrap(),
            "new",
            &encryption::Limits::default(),
        )
        .unwrap()
        .into_bytes(),
        signed
    );
    let before = std::fs::read(&target).unwrap();
    let File::Signed(wrapper) = source else {
        panic!("signed")
    };
    let unsigned = wrapper.unsign();
    assert_eq!(
        unsigned
            .save_with(SaveTo::Path(&target), &options())
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );
    assert_eq!(std::fs::read(&target).unwrap(), before);
}

#[test]
fn created_presentation_is_an_immutable_native_owner_and_publishes() {
    use litchi_core::FileFormat;
    use litchi_ops::capabilities::{Capabilities, Verb};
    assert!(
        Capabilities::of(FileFormat::Pptx)
            .verb(Verb::Create)
            .is_supported()
    );
    let file = litchi_ops::create(FileFormat::Pptx).unwrap();
    let File::Pptx(handle) = &file else {
        panic!("PPTX creator")
    };
    assert!(!handle.is_source_backed());
    assert_eq!(handle.slide_count().unwrap(), 0);
    assert_eq!(handle.text().unwrap(), "");
    let bytes = file.to_bytes().unwrap();
    assert_eq!(
        bytes,
        litchi_ops::create(FileFormat::Pptx)
            .unwrap()
            .to_bytes()
            .unwrap()
    );
    let commit = file.edit(|_| Ok(())).unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.file().to_bytes().unwrap(), bytes);
    for eager in [false, true] {
        let File::Pptx(reopened) = open(&bytes, eager) else {
            panic!("PPTX")
        };
        assert_eq!(reopened.slide_count().unwrap(), 0);
    }
    assert_eq!(file.save().unwrap_err().code(), Code::FileRequired);
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("new.pptx");
    let saved = file.save_as(&path).unwrap();
    assert_eq!(saved.bytes, bytes.len() as u64);
    assert_eq!(std::fs::read(&path).unwrap(), bytes);
}

#[test]
fn pptx_creation_refuses_unimplemented_options_explicitly() {
    use litchi_core::FileFormat;
    use litchi_ops::{Conformance, CreateOptions, Flavor};
    for options in [
        CreateOptions::new().flavor(Flavor::Template),
        CreateOptions::new().conformance(Conformance::Strict),
    ] {
        let error = litchi_ops::create_with(FileFormat::Pptx, &options).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(FileFormat::Pptx));
    }
}

#[test]
fn slide_addition_composes_with_text_edits_and_reversible_transactions() {
    use litchi_ops::{NewNode, Placement};
    let source = open(&fixture(), true);
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            edit.set(FIRST, props! { "text" => "edited before insertion" })?;
            let path = edit.add(
                "/",
                NewNode::new("slide")
                    .prop("layout", "Blank")
                    .at(Placement::Start),
            )?;
            assert_eq!(path.to_string(), "/slide[1]");
            edit.set(
                "/slide[2]/shape[1]",
                props! { "text" => "edited after insertion" },
            )?;
            let path = edit.add(
                "/",
                NewNode::new("slide")
                    .prop("layout", "Blank")
                    .at(Placement::after("/slide[2]")?),
            )?;
            assert_eq!(path.to_string(), "/slide[3]");
            Ok(())
        })
        .unwrap();
    assert_eq!(
        text(commit.file(), "/slide[2]/shape[1]"),
        "edited after insertion"
    );
    assert_eq!(text(commit.file(), "/slide[4]/shape[1]"), "second");
    let result = commit.file().to_bytes().unwrap();
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        result
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    for eager in [false, true] {
        let reopened = open(&result, eager);
        assert_eq!(
            text(&reopened, "/slide[2]/shape[1]"),
            "edited after insertion"
        );
        assert!(reopened.get("/slide[4]").is_ok());
    }
}

#[test]
fn new_presentations_accept_all_slide_placements() {
    use litchi_core::{FileFormat, Position};
    use litchi_ops::{NewNode, Placement};
    let source = litchi_ops::create(FileFormat::Pptx).unwrap();
    let node = || NewNode::new("slide").prop("layout", "Blank");
    let first = source.add("/", node()).unwrap();
    assert_eq!(first.path.to_string(), "/slide[1]");
    for (placement, expected) in [
        (Placement::Start, "/slide[1]"),
        (Placement::End, "/slide[2]"),
        (Placement::Position(Position::new(1)), "/slide[2]"),
        (Placement::before("/slide[1]").unwrap(), "/slide[1]"),
        (Placement::after("/slide[1]").unwrap(), "/slide[2]"),
        (Placement::inside("/").unwrap(), "/slide[2]"),
    ] {
        let added = first.file.add("/", node().at(placement)).unwrap();
        assert_eq!(added.path.to_string(), expected);
        assert!(added.file.get("/slide[2]").is_ok());
        assert!(first.file.get("/slide[2]").is_err());
    }
    assert!(source.get("/slide[1]").is_err());
}

#[test]
fn slide_addition_refuses_properties_placements_and_limits_atomically() {
    use litchi_core::Position;
    use litchi_ops::{NewNode, Placement};
    let source = open(&fixture(), true);
    let before = source.to_bytes().unwrap();
    for (node, code) in [
        (NewNode::new("slide"), Code::MissingProperty),
        (NewNode::new("picture"), Code::UnsupportedType),
        (NewNode::new("chart"), Code::UnsupportedType),
        (NewNode::new("ole"), Code::UnsupportedElement),
        (
            NewNode::new("slide").prop("layout", "missing"),
            Code::NotFound,
        ),
        (
            NewNode::new("slide").prop("bogus", "x"),
            Code::UnsupportedProperty,
        ),
        (
            NewNode::new("slide")
                .prop("layout", "Blank")
                .prop("layout", "Blank"),
            Code::InvalidArgument,
        ),
        (
            NewNode::new("slide")
                .prop("layout", "Blank")
                .at(Placement::Position(Position::new(99))),
            Code::OutOfRange,
        ),
    ] {
        assert_eq!(source.add("/", node).unwrap_err().code(), code);
        assert_eq!(source.to_bytes().unwrap(), before);
    }
    let limits = SemanticLimits::builder()
        .max_staged_bytes(4)
        .unwrap()
        .build()
        .unwrap();
    let result = source.edit_with_options(
        litchi_ops::EditOptions::new().semantic_limits(limits),
        |edit| {
            edit.add("/", NewNode::new("slide").prop("layout", "Blank"))?;
            Ok(())
        },
    );
    assert_eq!(result.unwrap_err().code(), Code::LimitExceeded);
    let result = source.edit(|edit| {
        edit.add("/", NewNode::new("slide").prop("layout", "Blank"))?;
        let _ = edit.add("/", NewNode::new("slide").prop("layout", "missing"));
        Ok(())
    });
    assert!(result.is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
    let deferred = open(&fixture(), false);
    let before = deferred.to_bytes().unwrap();
    let added = deferred
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap();
    assert!(added.file.get(&added.path.to_string()).is_ok());
    assert_eq!(deferred.to_bytes().unwrap(), before);
}

#[test]
fn slide_addition_batches_keep_prior_success_only_in_best_effort_mode() {
    use litchi_core::FileFormat;
    use litchi_ops::{BatchItem, BatchMode, ItemResult, Props};
    let source = litchi_ops::create(FileFormat::Pptx).unwrap();
    let before = source.to_bytes().unwrap();
    let commands = || {
        ["Blank", "missing", "Blank"].map(|layout| {
            let mut item = BatchItem::new("add");
            item.parent = Some("/".into());
            item.kind = Some("slide".into());
            item.props = Some(Props::new().with("layout", layout));
            item
        })
    };
    assert_eq!(
        source
            .batch(commands(), BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let result = source.batch(commands(), BatchMode::BestEffort).unwrap();
    assert!(matches!(&result.value()[1], ItemResult::Failed(error) if error.item() == Some(1)));
    assert!(result.file().get("/slide[2]").is_ok());
    assert!(result.file().get("/slide[3]").is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}
