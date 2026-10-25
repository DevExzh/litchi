//! Atomic closure execution and opaque process-local patch contracts.
#![cfg(feature = "docx")]
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, Props, SourceMode, props};
use std::io::{Cursor, Write as _};

const XML: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>FIRST_PRIVATE</w:t></w:r><w:r><w:t>tail</w:t></w:r></w:p><!--between--><w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;

fn bytes() -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, text) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        ("word/document.xml", XML),
        ("opaque.xml", "<opaque>UNTOUCHED_PRIVATE</opaque>"),
    ] {
        writer.start_file(name, options).unwrap();
        writer.write_all(text.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn open(eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}
fn text(file: &File) -> String {
    file.get("/body/p[1]/r[1]")
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}
fn conflict(error: litchi_ops::Error) {
    assert_eq!(error.code(), Code::InvalidInput);
    assert!(matches!(
        error.leaf(),
        Source::Transaction(Error::PatchConflict)
    ));
}

#[test]
fn sequential_addressing_and_all_five_verbs_match_single_operation_oracle() {
    for eager in [false, true] {
        let file = open(eager);
        let original = file.to_bytes().unwrap();
        let commit = file
            .edit(|tx| {
                let path = tx.add("/body", NewNode::new("paragraph").prop("text", "third"))?;
                assert_eq!(path.to_string(), "/body/p[3]");
                assert_eq!(tx.get(&path)?.view().text()?.as_deref(), Some("third"));
                tx.set("/body/p[3]/r[1]", props! {"text" => "last"})?;
                tx.move_to(&path, Placement::Start)?;
                tx.swap("/body/p[1]", "/body/p[2]")?;
                tx.remove("/body/p[3]")?;
                assert!(tx.find("/body/p[3]")?.is_none());
                Ok("PRIVATE_CLOSURE_VALUE")
            })
            .unwrap();
        let oracle = file
            .add("/body", NewNode::new("paragraph").prop("text", "third"))
            .unwrap()
            .file
            .set("/body/p[3]/r[1]", props! {"text" => "last"})
            .unwrap()
            .move_to("/body/p[3]", Placement::Start)
            .unwrap()
            .swap("/body/p[1]", "/body/p[2]")
            .unwrap()
            .remove("/body/p[3]")
            .unwrap();
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(file.to_bytes().unwrap(), original);
        assert_eq!(*commit.value(), "PRIVATE_CLOSURE_VALUE");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 5);
        assert!(commit.diagnostics().leaf_operations() >= 5);
        assert!(commit.diagnostics().snapshot_bytes() > 0);
        assert!(!format!("{commit:?} {:?}", commit.patch()).contains("PRIVATE"));
        assert_eq!(
            commit
                .patch()
                .apply(&file.clone())
                .unwrap()
                .to_bytes()
                .unwrap(),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(&commit.file().clone())
                .unwrap()
                .to_bytes()
                .unwrap(),
            original
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .inverse()
                .apply(&file)
                .unwrap()
                .to_bytes()
                .unwrap(),
            oracle.to_bytes().unwrap()
        );
    }
}

#[test]
fn replay_accepts_byte_identical_reopens_and_rejects_other_content() {
    for eager in [false, true] {
        let file = open(eager);
        let commit = file
            .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => "new"}))
            .unwrap();
        assert_eq!(text(&commit.patch().apply(&open(eager)).unwrap()), "new");
        conflict(commit.patch().apply(commit.file()).unwrap_err());
        conflict(commit.patch().inverse().apply(&file).unwrap_err());
        let reopened = litchi_ops::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
        assert_eq!(
            text(&commit.patch().inverse().apply(&reopened).unwrap()),
            "FIRST_PRIVATE"
        );
        let first = commit.patch().apply(&file).unwrap();
        assert_eq!(
            text(&commit.patch().inverse().apply(&first).unwrap()),
            "FIRST_PRIVATE"
        );
        assert_eq!(text(&commit.patch().apply(&file).unwrap()), "new");
        let noop = file.edit(|_| Ok(())).unwrap();
        assert!(noop.patch().is_noop());
        assert_eq!(
            noop.patch()
                .apply(&open(eager))
                .unwrap()
                .to_bytes()
                .unwrap(),
            file.to_bytes().unwrap()
        );
    }
}

#[test]
fn empty_noop_and_cancelled_edits_preserve_exact_archive_and_identity() {
    for eager in [false, true] {
        let file = open(eager);
        let empty = file.edit(|_| Ok(7)).unwrap();
        assert_eq!(*empty.value(), 7);
        assert_eq!(empty.diagnostics().operations(), 0);
        assert_eq!(empty.diagnostics().snapshot_bytes(), 0);
        assert!(empty.patch().is_noop());
        let noop = file
            .edit(|tx| {
                tx.set("/body/p[1]/r[1]", Props::new())?;
                tx.swap("/body/p[1]", "/body/p[1]")?;
                tx.set("/body/p[1]/r[1]", props! {"text" => "FIRST_PRIVATE"})
            })
            .unwrap();
        assert!(noop.patch().is_noop());
        assert_eq!(noop.diagnostics().operations(), 3);
        let cancel = file
            .edit(|tx| {
                tx.set("/body/p[1]/r[1]", props! {"text" => "temporary"})?;
                tx.set("/body/p[1]/r[1]", props! {"text" => "FIRST_PRIVATE"})
            })
            .unwrap();
        assert!(cancel.patch().is_noop());
        assert!(!cancel.diagnostics().changed());
        assert_eq!(cancel.diagnostics().operations(), 2);
        assert!(cancel.diagnostics().leaf_operations() > 0);
        for result in [empty.file(), noop.file(), cancel.file()] {
            assert_eq!(result.to_bytes().unwrap(), bytes());
            assert_eq!(
                cancel.patch().apply(result).unwrap().to_bytes().unwrap(),
                bytes()
            );
            assert_eq!(
                cancel
                    .patch()
                    .inverse()
                    .apply(result)
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                bytes()
            );
        }
    }
}

#[test]
fn caught_mutation_failure_poisoning_closure_error_and_panic_roll_back() {
    for eager in [false, true] {
        let file = open(eager);
        let error = file
            .edit(|tx| {
                tx.set("/body/p[1]/r[1]", props! {"text" => "staged"})?;
                let failure = tx
                    .set("/body/p[1]/r[1]", props! {"not-a-property" => true})
                    .unwrap_err();
                assert_eq!(failure.item(), Some(1));
                let next = tx.remove("/body/p[1]").unwrap_err();
                assert_eq!(next.item(), Some(1));
                assert_eq!(next.code(), failure.code());
                assert_eq!(tx.get("/body").unwrap_err().item(), Some(1));
                Ok(())
            })
            .unwrap_err();
        assert_eq!(error.item(), Some(1));
        let error = file
            .edit::<()>(|tx| {
                tx.remove("/body/p[1]")?;
                Err(litchi_ops::Error::new(
                    Code::InvalidArgument,
                    "closure refusal",
                ))
            })
            .unwrap_err();
        assert_eq!(error.item(), None);
        let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let _ = file.edit::<()>(|tx| {
                tx.remove("/body/p[1]")?;
                panic!("user closure panic")
            });
        }));
        assert!(panic.is_err());
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}

#[test]
fn read_errors_can_be_handled_without_poisoning_mutations() {
    let file = open(false);
    let result = file
        .edit(|tx| {
            assert!(tx.get("/body/p[99]").is_err());
            tx.set("/body/p[1]/r[1]", props! {"text" => "new"})
        })
        .unwrap();
    assert_eq!(text(result.file()), "new");
}

#[test]
fn swapped_nested_proxies_cannot_transfer_limits_or_failure_state() {
    let file = open(false);
    let error = file
        .edit(|outer| {
            let inner_error = file
                .edit(|inner| {
                    std::mem::swap(outer, inner);
                    Ok(())
                })
                .unwrap_err();
            conflict(inner_error);
            Ok(())
        })
        .unwrap_err();
    conflict(error);
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn operation_and_retention_bounds_are_typed_and_discard_candidates() {
    let file = open(false);
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let error = file
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |tx| {
                tx.set("/body/p[1]/r[1]", props! {"text" => "new"})?;
                tx.remove("/body/p[2]")
            },
        )
        .unwrap_err();
    assert_eq!(error.item(), Some(1));
    assert_eq!(error.code(), Code::LimitExceeded);
    assert!(matches!(
        error.leaf(),
        Source::Transaction(Error::Limit {
            kind: LimitKind::Operations,
            observed: 2,
            maximum: 1
        })
    ));
    let commit = file
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |tx| tx.remove("/body/p[2]"),
        )
        .unwrap();
    let charge = commit.diagnostics().snapshot_bytes();
    for (bound, success) in [(charge, true), (charge - 1, false)] {
        let limits = SemanticLimits::builder()
            .max_snapshot_bytes(bound)
            .unwrap()
            .build()
            .unwrap();
        let result = file.edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |tx| tx.remove("/body/p[2]"),
        );
        assert_eq!(result.is_ok(), success);
        if let Err(error) = result {
            assert_eq!(error.item(), Some(0));
            assert!(
                matches!(error.leaf(), Source::Transaction(Error::Limit {kind: LimitKind::SnapshotBytes, observed, maximum}) if *observed == charge && *maximum == bound)
            );
        }
    }
    assert_eq!(file.to_bytes().unwrap(), bytes());
}

#[test]
fn stale_physical_sources_refuse_transaction_and_both_replay_directions() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    std::fs::write(&path, bytes()).unwrap();
    let file = litchi_ops::open(&path).unwrap();
    let commit = file.edit(|tx| tx.remove("/body/p[2]")).unwrap();
    std::fs::write(&path, b"changed physical source").unwrap();
    for error in [
        commit.patch().apply(&file).unwrap_err(),
        commit.patch().inverse().apply(commit.file()).unwrap_err(),
        file.edit(|_| Ok(())).unwrap_err(),
    ] {
        assert_eq!(error.code(), Code::IoError);
    }
}

#[test]
fn final_source_fence_includes_closure_failures_and_mutations() {
    for fail in [false, true] {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("source.docx");
        std::fs::write(&path, bytes()).unwrap();
        let file = litchi_ops::open(&path).unwrap();
        let error = file
            .edit(|tx| {
                tx.remove("/body/p[2]")?;
                std::fs::write(&path, b"changed physical source").unwrap();
                if fail {
                    Err(litchi_ops::Error::new(
                        Code::InvalidArgument,
                        "closure refusal",
                    ))
                } else {
                    Ok(())
                }
            })
            .unwrap_err();
        assert_eq!(error.code(), Code::IoError);
    }
}

#[test]
fn read_only_handles_refuse_before_running_closure() {
    let file =
        litchi_ops::open_bytes_with(bytes(), &OpenOptions::new().retain_source(false)).unwrap();
    let error = file
        .edit::<()>(|_| panic!("read-only closure ran"))
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

const SIGNED: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
));
const SIGNED_OTHER: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/poi/test-data/xmldsign/hello-world-signed-twice.docx"
));
#[test]
fn signature_authorization_is_bound_and_inverse_never_restores_signatures() {
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            SIGNED.to_vec(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let File::Signed(signed) = file else {
            panic!("signed fixture")
        };
        let signed_outer = File::Signed(signed.clone());
        let signed_inner = signed.file().clone();
        let signed_noop = signed_outer.edit(|_| Ok(())).unwrap();
        assert!(signed_noop.patch().is_noop());
        let reopened_signed = litchi_ops::open_bytes(SIGNED.to_vec()).unwrap();
        let replayed = signed_noop.patch().apply(&reopened_signed).unwrap();
        assert!(replayed.is_signed());
        assert_eq!(replayed.to_bytes().unwrap(), SIGNED);
        let signed_leaf_noop = signed_inner.edit(|_| Ok(())).unwrap();
        let reopened_signed_leaf = match litchi_ops::open_bytes(SIGNED.to_vec()).unwrap() {
            File::Signed(signed) => signed.file().clone(),
            _ => panic!("signed fixture"),
        };
        let replayed_leaf = signed_leaf_noop
            .patch()
            .apply(&reopened_signed_leaf)
            .unwrap();
        assert!(replayed_leaf.is_signed());
        assert_eq!(replayed_leaf.to_bytes().unwrap(), SIGNED);
        let original_text = text(&signed_inner);
        let setter_noop = signed_outer
            .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => original_text.clone()}))
            .unwrap_or_else(|error| panic!("setter no-op: {error}; leaf={:?}", error.leaf()));
        assert!(setter_noop.patch().is_noop());
        assert!(setter_noop.file().is_signed());
        assert_eq!(setter_noop.file().to_bytes().unwrap(), SIGNED);
        let leaf_setter_noop = signed_inner
            .set("/body/p[1]/r[1]", props! {"text" => original_text.clone()})
            .unwrap_or_else(|error| panic!("leaf setter no-op: {error}; leaf={:?}", error.leaf()));
        assert!(leaf_setter_noop.is_signed());
        assert_eq!(leaf_setter_noop.to_bytes().unwrap(), SIGNED);
        let facade_setter_noop = signed_outer
            .set("/body/p[1]/r[1]", props! {"text" => original_text})
            .unwrap_or_else(|error| {
                panic!("facade setter no-op: {error}; leaf={:?}", error.leaf())
            });
        assert!(facade_setter_noop.is_signed());
        assert_eq!(facade_setter_noop.to_bytes().unwrap(), SIGNED);
        assert_eq!(
            signed_inner
                .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => "blocked"}))
                .unwrap_err()
                .code(),
            Code::DocumentProtected
        );
        assert_eq!(
            signed_outer
                .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => "blocked"}))
                .unwrap_err()
                .code(),
            Code::DocumentProtected
        );
        assert_eq!(
            signed_inner
                .set("/body/p[1]/r[1]", props! {"text" => "blocked"})
                .unwrap_err()
                .code(),
            Code::DocumentProtected
        );
        assert_eq!(
            signed_outer
                .set("/body/p[1]/r[1]", props! {"text" => "blocked"})
                .unwrap_err()
                .code(),
            Code::DocumentProtected
        );
        let unsigned = signed.unsign();
        let baseline = unsigned.to_bytes().unwrap();
        let commit = unsigned
            .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => "unsigned edit"}))
            .unwrap();
        assert_eq!(
            commit.patch().apply(&signed_inner).unwrap_err().code(),
            Code::DocumentProtected
        );
        assert_eq!(
            commit.patch().apply(&signed_outer).unwrap_err().code(),
            Code::DocumentProtected
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(&signed_outer)
                .unwrap_err()
                .code(),
            Code::DocumentProtected,
            "an inverse of an authorized changed edit cannot restore a signed state"
        );
        let restored = commit.patch().inverse().apply(commit.file()).unwrap();
        assert!(!restored.is_signed());
        assert_eq!(restored.to_bytes().unwrap(), baseline);
        assert!(
            !litchi_ops::open_bytes(restored.to_bytes().unwrap())
                .unwrap()
                .is_signed()
        );
        let noop = unsigned.edit(|_| Ok(())).unwrap();
        assert!(noop.patch().is_noop());
        assert_eq!(noop.file().to_bytes().unwrap(), baseline);
        let genuinely_unsigned = litchi_ops::open_bytes(baseline.clone()).unwrap();
        assert!(!genuinely_unsigned.is_signed());
        let unsigned_noop = genuinely_unsigned.edit(|_| Ok(())).unwrap();
        for recipient in [&genuinely_unsigned, &unsigned] {
            let replayed = noop.patch().apply(recipient).unwrap();
            assert!(!replayed.is_signed());
            assert_eq!(replayed.to_bytes().unwrap(), baseline);
        }
        assert_eq!(
            unsigned_noop
                .patch()
                .apply(&signed_outer)
                .unwrap_err()
                .code(),
            Code::InvalidInput,
            "an ordinary unsigned base cannot match a signed recipient"
        );
        assert_eq!(
            noop.patch().apply(&signed_outer).unwrap_err().code(),
            Code::InvalidInput,
            "an authorized unsigned base cannot match a signed recipient"
        );
        let other = litchi_ops::open_bytes(SIGNED_OTHER.to_vec()).unwrap();
        assert!(other.is_signed());
        assert_eq!(
            signed_noop.patch().apply(&other).unwrap_err().code(),
            Code::InvalidInput,
            "a signed recipient with different recorded bytes is a patch conflict"
        );
    }
}

#[test]
fn source_guards_survive_signature_removal_promotion_to_eager_owner() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("signed.docx");
    std::fs::write(&path, SIGNED).unwrap();
    let File::Signed(signed) = litchi_ops::open(&path).unwrap() else {
        panic!("signed fixture")
    };
    let signed_target = File::Signed(signed.clone());
    let signed_noop = signed_target.edit(|_| Ok(())).unwrap();
    let unsigned = signed.unsign();
    let noop = unsigned.edit(|_| Ok(())).unwrap();
    let commit = unsigned
        .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => "unsigned edit"}))
        .unwrap();
    std::fs::write(&path, b"changed physical source").unwrap();
    assert_eq!(
        signed_noop
            .patch()
            .apply(&signed_target)
            .unwrap_err()
            .code(),
        Code::IoError,
        "a signed no-op checks its retained source before replay"
    );
    assert_eq!(
        noop.patch().apply(&signed_target).unwrap_err().code(),
        Code::IoError,
        "a signed no-op recipient still checks its retained source"
    );
    assert_eq!(
        commit.patch().apply(&unsigned).unwrap_err().code(),
        Code::IoError
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap_err()
            .code(),
        Code::IoError
    );
}

#[test]
fn unsigned_noop_and_cancelled_transactions_keep_the_logical_unsigned_baseline() {
    for eager in [false, true] {
        let file = litchi_ops::open_bytes_with(
            SIGNED.to_vec(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .unwrap();
        let File::Signed(signed) = file else {
            panic!("signed fixture")
        };
        let file = signed.unsign();
        let baseline = file.to_bytes().unwrap();
        let original_text = text(&file);
        for cancel in [false, true] {
            let commit = file
                .edit(|tx| {
                    if cancel {
                        tx.set("/body/p[1]/r[1]", props! {"text" => "temporary"})?;
                    }
                    tx.set("/body/p[1]/r[1]", props! {"text" => original_text.as_str()})
                })
                .unwrap();
            assert!(!commit.diagnostics().changed());
            assert!(commit.patch().is_noop());
            for endpoint in [
                commit.file().clone(),
                commit.patch().apply(&file).unwrap(),
                commit.patch().inverse().apply(commit.file()).unwrap(),
            ] {
                assert!(!endpoint.is_signed());
                assert_eq!(endpoint.to_bytes().unwrap(), baseline);
                assert!(
                    !litchi_ops::open_bytes(endpoint.to_bytes().unwrap())
                        .unwrap()
                        .is_signed()
                );
            }
        }
    }
}

#[test]
fn replay_preserves_origin_for_both_endpoints() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("origin.docx");
    std::fs::write(&path, bytes()).unwrap();
    let file =
        litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Eager)).unwrap();
    let commit = file
        .edit(|tx| tx.set("/body/p[1]/r[1]", props! {"text" => "new"}))
        .unwrap();
    let forward = commit.patch().apply(&file).unwrap();
    let inverse = commit.patch().inverse().apply(&forward).unwrap();
    assert_eq!(forward.save().unwrap().path, path);
    assert_eq!(text(&litchi_ops::open(&path).unwrap()), "new");
    assert_eq!(inverse.save().unwrap().path, path);
    assert_eq!(std::fs::read(path).unwrap(), bytes());
}

#[cfg(feature = "encryption")]
#[test]
fn locked_input_refuses_before_closure_execution() {
    let encrypted = litchi_crypto::ooxml::encrypt(
        bytes(),
        "transaction",
        litchi_crypto::ooxml::Mode::Agile,
        &litchi_crypto::ooxml::Limits::default(),
    )
    .unwrap();
    let locked = litchi_ops::open_bytes(encrypted).unwrap();
    let error = locked
        .edit::<()>(|_| panic!("locked closure ran"))
        .unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    let plain = open(false);
    let noop = plain.edit(|_| Ok(())).unwrap();
    assert_eq!(
        noop.patch().apply(&locked).unwrap_err().code(),
        Code::DocumentProtected,
        "a no-op patch must not bypass a locked recipient"
    );
}

#[cfg(feature = "encryption")]
#[test]
fn unlocked_encrypted_input_refuses_without_dropping_publication_protection() {
    let encrypted = litchi_crypto::ooxml::encrypt(
        bytes(),
        "transaction",
        litchi_crypto::ooxml::Mode::Agile,
        &litchi_crypto::ooxml::Limits::default(),
    )
    .unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(encrypted).unwrap() else {
        panic!("locked fixture")
    };
    let file = locked
        .unlock(litchi_ops::Password::new("transaction".to_owned()))
        .unwrap();
    let error = file
        .edit::<()>(|_| panic!("unsupported encrypted closure ran"))
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(
        file.to_bytes().unwrap_err().code(),
        Code::UnsupportedCapability
    );
    let plain = open(false);
    let noop = plain.edit(|_| Ok(())).unwrap();
    assert_eq!(
        noop.patch().apply(&file).unwrap_err().code(),
        Code::UnsupportedCapability,
        "a no-op patch must not bypass encrypted publication policy"
    );
    let report = serde_json::to_value(file.capabilities()).unwrap();
    assert!(
        report["gaps"]
            .as_array()
            .unwrap()
            .iter()
            .any(|gap| gap["verb"] == "edit" && gap["context"] == "unlocked-encrypted-owner")
    );
}

#[test]
fn closure_formatting_coalesces_typed_targets_and_preserves_exact_inverse() {
    for eager in [false, true] {
        let file = open(eager);
        let path = litchi_ops::path::Path::parse("/body/p[1]/r[1]").unwrap();
        let commit = file
            .edit(|tx| {
                tx.set(&path, props! {"bold" => true})?;
                tx.set(&path, props! {"bold" => true})?;
                tx.set("/body/p[2]", props! {"align" => "center"})?;
                tx.set(&path, props! {"italic" => true})
            })
            .unwrap();
        let oracle = file
            .set(&path, props! {"bold"=>true})
            .unwrap()
            .set("/body/p[2]", props! {"align"=>"center"})
            .unwrap()
            .set(&path, props! {"italic"=>true})
            .unwrap();
        assert_eq!(commit.diagnostics().operations(), 4);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
    }
}

#[test]
fn closure_reads_publish_pending_formatting_and_structural_barriers_readdress() {
    for eager in [false, true] {
        let file = open(eager);
        let commit = file
            .edit(|tx| {
                tx.set("/body/p[1]/r[1]", props! {"bold"=>true})?;
                assert!(
                    tx.get("/body/p[1]/r[1]")?
                        .view()
                        .format()?
                        .any(|p| p.key() == "bold" && p.to_string() == "true")
                );
                assert!(tx.find("/body/p[9]")?.is_none());
                tx.set("/body/p[2]", props! {"align"=>"center"})?;
                tx.remove("/body/p[1]")?;
                tx.set("/body/p[1]/r[1]", props! {"italic"=>true})
            })
            .unwrap();
        let oracle = file
            .set("/body/p[1]/r[1]", props! {"bold"=>true})
            .unwrap()
            .set("/body/p[2]", props! {"align"=>"center"})
            .unwrap()
            .remove("/body/p[1]")
            .unwrap()
            .set("/body/p[1]/r[1]", props! {"italic"=>true})
            .unwrap();
        assert_eq!(commit.diagnostics().commits(), 4);
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
    }
}

#[test]
fn pending_formatting_is_poisoned_by_caught_input_errors_and_operation_limits() {
    let file = open(false);
    let error = file
        .edit(|tx| {
            tx.set("/body/p[1]/r[1]", props! {"bold"=>true})?;
            let first = tx.set("/body/p[1]/r[1]", &["malformed"][..]).unwrap_err();
            assert_eq!(first.item(), Some(1));
            assert_eq!(tx.get("/body/p[1]").unwrap_err().code(), first.code());
            assert_eq!(
                tx.set("/body/p[2]", props! {"align"=>"center"})
                    .unwrap_err()
                    .item(),
                Some(1)
            );
            Ok(())
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    assert_eq!(file.to_bytes().unwrap(), bytes());
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let error = file
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |tx| {
                tx.set("/body/p[1]/r[1]", props! {"bold"=>true})?;
                assert_eq!(
                    tx.set("bad[", props! {"unknown"=>true}).unwrap_err().code(),
                    Code::LimitExceeded
                );
                Ok(())
            },
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(error.item(), Some(1));
}

#[test]
fn closure_group_retention_is_bounded_and_cancellation_keeps_exact_baseline() {
    let file = open(true)
        .set("/body/p[1]/r[1]", props! {"bold"=>false})
        .unwrap();
    let original = file.to_bytes().unwrap();
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(2048)
        .unwrap()
        .build()
        .unwrap();
    let commit = file
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |tx| {
                for index in 0..100 {
                    tx.set("/body/p[1]/r[1]", props! {"bold"=>index % 2 == 0})?;
                }
                Ok(())
            },
        )
        .unwrap();
    assert_eq!(commit.diagnostics().operations(), 100);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.file().to_bytes().unwrap(), original);
}

#[test]
fn closure_selectors_observe_pending_formatting_before_matching() {
    for eager in [false, true] {
        let file = open(eager);
        let commit = file
            .edit(|tx| {
                tx.set("/body/p[1]", props! {"align"=>"center"})?;
                tx.set("/body/p[align=center]", props! {"style"=>"Heading1"})
            })
            .unwrap();
        let oracle = file
            .set("/body/p[1]", props! {"align"=>"center"})
            .unwrap()
            .set("/body/p[align=center]", props! {"style"=>"Heading1"})
            .unwrap();
        assert_eq!(commit.diagnostics().commits(), 2);
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            oracle.to_bytes().unwrap()
        );
    }
}

#[test]
fn native_paragraph_text_controls_preserve_slots_and_inverse() {
    for eager in [false, true] {
        let file = open(eager);
        let commit = file
            .edit(|edit| {
                edit.set("/body/p[1]", props! { "text" => "first\tline\nsecond" })?;
                Ok(())
            })
            .unwrap();
        let bytes_after = commit.file().to_bytes().unwrap();
        let package = litchi_docx::Package::from_bytes(&bytes_after).unwrap();
        let paragraphs = package.document().unwrap().paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].text().unwrap(), "first\tline\nsecond");
        assert_eq!(paragraphs[0].runs().unwrap().len(), 2);
        let repeat = commit
            .file()
            .edit(|edit| edit.set("/body/p[1]", props! { "text" => "first\tline\nsecond" }))
            .unwrap();
        assert!(!repeat.diagnostics().changed());
        assert_eq!(repeat.file().to_bytes().unwrap(), bytes_after);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes()
        );
        for invalid in ["bad\rtext", "bad\u{b}text", "bad\u{1}text"] {
            assert!(
                commit
                    .file()
                    .set("/body/p[1]", props! { "text" => invalid })
                    .is_err()
            );
        }
        assert_eq!(file.to_bytes().unwrap(), bytes());
    }
}
