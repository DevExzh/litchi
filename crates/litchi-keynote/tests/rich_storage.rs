use std::io;
use std::path::PathBuf;
use std::sync::Arc;

use litchi_iwa_archive::{ComponentCatalog, Limits};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp, tswp};
use litchi_keynote::{
    MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES,
    Package, ReadError, ReadOptions, SemanticLimitKind, SemanticLimits, SemanticPath,
    TextStorageFailure,
};
use prost::Message as _;

const TITLE: &[&str] = &["Rich ", "title"];
const BODY: &[&str] = &["Body ", "fragment"];
const SHAPE: &[&str] = &["Shape", " text"];
const NOTES: &[&str] = &["Note ", "run"];

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, Copy, Default)]
enum Malformation {
    #[default]
    None,
    DuplicateBodyStorage,
    WrongBodyType,
    WrongBodyWire,
    Legacy2022Sibling,
    OversizedBuildIdentifier,
    OversizedTransitionIdentifier,
    MissingShowTheme,
    MissingDocumentShow,
    MissingSlideNodeFlag,
    MissingSlideStyle,
    MissingBuildAttributes,
    MissingPlaceholderSuper,
    MissingShapeSuper,
    MissingNoteStorage,
    AmbiguousTitleOwner,
    DuplicateObjectIdentity,
    DuplicateSlideReference,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, value: &impl prost::Message) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_,
            data: value.encode_to_vec(),
        }],
    )?)
}

fn storage(identifier: u64, fragments: &[&str]) -> TestResult<ArchiveObject> {
    object(
        identifier,
        2_001,
        &tswp::StorageArchive {
            text: fragments
                .iter()
                .map(|fragment| (*fragment).to_owned())
                .collect(),
            ..Default::default()
        },
    )
}

fn synthetic_package(malformed: Malformation) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let mut slide_references = vec![reference(3)];
    if matches!(malformed, Malformation::DuplicateSlideReference) {
        slide_references.push(reference(3));
    }
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: slide_references,
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..Default::default()
    };
    #[allow(deprecated, reason = "native schema requires legacy cache fields")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(4)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    };
    #[allow(
        deprecated,
        reason = "the adversarial fixture exercises the legacy transition fallback read path"
    )]
    let transition = if matches!(malformed, Malformation::OversizedTransitionIdentifier) {
        kn::TransitionArchive {
            attributes: kn::TransitionAttributesArchive {
                database_effect: Some("future-transition".to_owned()),
                ..Default::default()
            },
        }
    } else {
        kn::TransitionArchive::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        builds: matches!(
            malformed,
            Malformation::OversizedBuildIdentifier | Malformation::MissingBuildAttributes
        )
        .then(|| reference(14))
        .into_iter()
        .collect(),
        transition,
        title_placeholder: Some(reference(5)),
        body_placeholder: Some(reference(6)),
        owned_drawables: vec![reference(5), reference(6), reference(7)],
        drawables_z_order: vec![reference(5), reference(6), reference(7)],
        note: Some(reference(8)),
        in_document: true,
        ..Default::default()
    };
    let title = kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            owned_storage: Some(reference(10)),
            ..Default::default()
        },
        ..Default::default()
    };
    let body = kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            owned_storage: Some(reference(11)),
            ..Default::default()
        },
        ..Default::default()
    };
    let shape = tswp::ShapeInfoArchive {
        owned_storage: Some(reference(12)),
        ..Default::default()
    };
    let speaker_note = kn::NoteArchive {
        contained_storage: reference(13),
    };

    let false_text = tswp::StorageArchive {
        text: vec!["UNREACHABLE FALSE POSITIVE".to_owned()],
        ..Default::default()
    }
    .encode_to_vec();
    let show_payload = if matches!(malformed, Malformation::MissingShowTheme) {
        litchi_iwa_common::wire::patch_length_delimited_field(&show.encode_to_vec(), 2, true, None)?
    } else {
        show.encode_to_vec()
    };
    let show_object = ArchiveObject::new(
        2,
        vec![
            RawMessage {
                type_: 2,
                data: show_payload,
            },
            RawMessage {
                type_: 999,
                data: false_text.clone(),
            },
        ],
    )?;
    let mut title_messages = vec![RawMessage {
        type_: 7,
        data: title.encode_to_vec(),
    }];
    if matches!(malformed, Malformation::AmbiguousTitleOwner) {
        title_messages.push(RawMessage {
            type_: 2_011,
            data: tswp::ShapeInfoArchive {
                owned_storage: Some(reference(10)),
                ..Default::default()
            }
            .encode_to_vec(),
        });
    }
    let title_object = ArchiveObject::new(5, title_messages)?;

    let body_payload = tswp::StorageArchive {
        text: BODY.iter().map(|fragment| (*fragment).to_owned()).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    let mut body_messages = vec![RawMessage {
        type_: if matches!(malformed, Malformation::WrongBodyType) {
            999
        } else {
            2_001
        },
        data: if matches!(malformed, Malformation::WrongBodyWire) {
            vec![0x18, 0x01]
        } else {
            body_payload.clone()
        },
    }];
    if matches!(malformed, Malformation::DuplicateBodyStorage) {
        body_messages.push(RawMessage {
            type_: 2_001,
            data: body_payload,
        });
    }
    if matches!(malformed, Malformation::Legacy2022Sibling) {
        body_messages.push(RawMessage {
            type_: 2_022,
            data: vec![0x18, 0x01],
        });
    }
    let body_storage = ArchiveObject::new(11, body_messages)?;

    let node_payload = if matches!(malformed, Malformation::MissingSlideNodeFlag) {
        litchi_iwa_common::wire::patch_varint_field(&node.encode_to_vec(), 6, true, None)?
    } else {
        node.encode_to_vec()
    };
    let slide_payload = if matches!(malformed, Malformation::MissingSlideStyle) {
        litchi_iwa_common::wire::patch_length_delimited_field(
            &slide.encode_to_vec(),
            1,
            true,
            None,
        )?
    } else {
        slide.encode_to_vec()
    };
    let document_payload = if matches!(malformed, Malformation::MissingDocumentShow) {
        litchi_iwa_common::wire::patch_length_delimited_field(
            &document.encode_to_vec(),
            2,
            true,
            None,
        )?
    } else {
        document.encode_to_vec()
    };
    let body_owner_payload = if matches!(malformed, Malformation::MissingPlaceholderSuper) {
        litchi_iwa_common::wire::patch_length_delimited_field(&body.encode_to_vec(), 1, true, None)?
    } else {
        body.encode_to_vec()
    };
    let shape_payload = if matches!(malformed, Malformation::MissingShapeSuper) {
        litchi_iwa_common::wire::patch_length_delimited_field(
            &shape.encode_to_vec(),
            1,
            true,
            None,
        )?
    } else {
        shape.encode_to_vec()
    };
    let speaker_note_payload = if matches!(malformed, Malformation::MissingNoteStorage) {
        litchi_iwa_common::wire::patch_length_delimited_field(
            &speaker_note.encode_to_vec(),
            1,
            true,
            None,
        )?
    } else {
        speaker_note.encode_to_vec()
    };

    let mut objects = vec![
        ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 1,
                data: document_payload,
            }],
        )?,
        show_object,
        ArchiveObject::new(
            3,
            vec![RawMessage {
                type_: 4,
                data: node_payload,
            }],
        )?,
        ArchiveObject::new(
            4,
            vec![RawMessage {
                type_: 5,
                data: slide_payload,
            }],
        )?,
        title_object,
        ArchiveObject::new(
            6,
            vec![RawMessage {
                type_: 7,
                data: body_owner_payload,
            }],
        )?,
        ArchiveObject::new(
            7,
            vec![RawMessage {
                type_: 2_011,
                data: shape_payload,
            }],
        )?,
        ArchiveObject::new(
            8,
            vec![RawMessage {
                type_: 15,
                data: speaker_note_payload,
            }],
        )?,
        storage(10, TITLE)?,
        body_storage,
        storage(12, SHAPE)?,
        storage(13, NOTES)?,
        ArchiveObject::new(
            99,
            vec![RawMessage {
                type_: 999,
                data: false_text,
            }],
        )?,
    ];
    if matches!(
        malformed,
        Malformation::OversizedBuildIdentifier | Malformation::MissingBuildAttributes
    ) {
        let delivery = b"future-effect";
        let mut payload = vec![0x12, u8::try_from(delivery.len())?];
        payload.extend_from_slice(delivery);
        if !matches!(malformed, Malformation::MissingBuildAttributes) {
            payload.extend_from_slice(&[0x22, 0x00]);
        }
        objects.push(ArchiveObject::new(
            14,
            vec![RawMessage {
                type_: 8,
                data: payload,
            }],
        )?);
    }
    let archive = Archive { objects };
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    if matches!(malformed, Malformation::DuplicateObjectIdentity) {
        let duplicate = Archive {
            objects: vec![ArchiveObject::new(
                13,
                vec![RawMessage {
                    type_: 999,
                    data: Vec::new(),
                }],
            )?],
        };
        let duplicate_compressed = SnappyStream::compress(&duplicate.to_bytes()?)?;
        return Ok(litchi_iwa_archive::package::to_bytes(
            [
                ("Index/Document.iwa", compressed.as_slice()),
                ("Index/Duplicate.iwa", duplicate_compressed.as_slice()),
            ],
            Limits::default(),
        )?);
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", compressed.as_slice())],
        Limits::default(),
    )?)
}

fn semantic_text_bytes() -> usize {
    [TITLE, BODY, SHAPE, NOTES]
        .into_iter()
        .flatten()
        .map(|fragment| fragment.len())
        .sum()
}

fn semantic_reference_count() -> usize {
    13
}

fn semantic_fragment_count() -> usize {
    8
}

#[test]
fn reachable_buffa_projection_preserves_runs_and_rejects_false_text() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::default())?;
    let package = Package::from_bytes(&bytes)?;
    let show = package.show()?;
    let slide = &show.slides()[0];

    assert_eq!(slide.title(), Some("Rich title"));
    assert_eq!(slide.notes(), Some("Note run"));
    assert!(slide.text_content().is_empty());
    assert_eq!(slide.text_storages().len(), 2);
    assert_eq!(slide.text_storages()[0].text(), "Body fragment");
    assert_eq!(slide.text_storages()[1].text(), "Shape text");
    assert_eq!(slide.text_storages()[0].runs().len(), 2);
    assert_eq!(slide.text_storages()[1].runs().len(), 2);
    assert_eq!(slide.text_storages()[0].runs()[0].len(), BODY[0].len());
    assert_eq!(slide.text_storages()[0].runs()[1].len(), BODY[1].len());

    let text = package.text()?;
    assert_eq!(text, "Rich title\nBody fragment\nShape text\nNote run");
    assert!(!text.contains("FALSE POSITIVE"));
    Ok(())
}

#[test]
fn semantic_limits_are_inclusive_typed_and_semantically_located() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::default())?;
    let text_bytes = semantic_text_bytes();
    let references = semantic_reference_count();
    let fragments = semantic_fragment_count();
    let exact = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        references,
        4,
        fragments,
        text_bytes,
    )?;
    let exact_package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), exact))?;
    assert_eq!(exact_package.semantic_limits(), exact);
    assert_eq!(exact_package.show()?.slides().len(), 1);

    let reference_limited = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        references - 1,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )?;
    let reference_package = Package::from_bytes_with_options(
        &bytes,
        ReadOptions::new(Limits::default(), reference_limited),
    )?;
    assert!(matches!(
        reference_package.show(),
        Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::References,
            observed,
            maximum,
            path: SemanticPath::SlideNotes { index: 0 },
        }) if observed == references && maximum == references - 1
    ));

    let storage_limited = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        MAX_REFERENCES,
        3,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )?;
    let storage_package = Package::from_bytes_with_options(
        &bytes,
        ReadOptions::new(Limits::default(), storage_limited),
    )?;
    let storage_error = storage_package.show().err();
    assert!(
        matches!(
            storage_error,
            Some(ReadError::SemanticLimit {
                kind: SemanticLimitKind::TextStorages,
                observed: 4,
                maximum: 3,
                path: SemanticPath::SlideNotes { index: 0 },
            })
        ),
        "unexpected storage-limit result: {storage_error:?}"
    );

    let fragment_limited = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        fragments - 1,
        MAX_TEXT_BYTES,
    )?;
    let fragment_package = Package::from_bytes_with_options(
        &bytes,
        ReadOptions::new(Limits::default(), fragment_limited),
    )?;
    assert!(matches!(
        fragment_package.show(),
        Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::TextFragments,
            observed,
            maximum,
            path: SemanticPath::SlideNotes { index: 0 },
        }) if observed == fragments && maximum == fragments - 1
    ));

    let text_limited = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        text_bytes - 1,
    )?;
    let text_package = Package::from_bytes_with_options(
        &bytes,
        ReadOptions::new(Limits::default(), text_limited),
    )?;
    assert!(matches!(
        text_package.show(),
        Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::TextBytes,
            maximum,
            path: SemanticPath::SlideNotes { index: 0 },
            ..
        }) if maximum == text_bytes - 1
    ));
    Ok(())
}

#[test]
fn package_wide_object_limit_is_checked_before_index_allocation() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::default())?;
    let limits = SemanticLimits::new(
        12,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )?;
    let error =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), limits))
            .err()
            .ok_or_else(|| io::Error::other("object limit should reject the package"))?;
    assert!(matches!(
        error,
        ReadError::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed: 13,
            maximum: 12,
            path: SemanticPath::Package,
        }
    ));
    Ok(())
}

#[test]
fn slide_limit_is_checked_before_record_allocation() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::DuplicateSlideReference)?;
    let limits = SemanticLimits::new(
        MAX_OBJECTS,
        1,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )?;
    let package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), limits))?;
    assert!(matches!(
        package.show(),
        Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::Slides,
            observed: 2,
            maximum: 1,
            path: SemanticPath::Show,
        })
    ));
    Ok(())
}

#[test]
fn duplicate_native_object_identity_is_rejected_at_ingress() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::DuplicateObjectIdentity)?;
    let error = Package::from_bytes(&bytes)
        .err()
        .ok_or_else(|| io::Error::other("duplicate object identity should fail closed"))?;
    assert!(matches!(error, ReadError::InvalidFormat(_)));
    Ok(())
}

#[test]
fn duplicate_wrong_type_wrong_wire_and_ambiguous_owners_fail_closed() -> TestResult<()> {
    let cases = [
        (Malformation::DuplicateBodyStorage, "duplicate storage"),
        (Malformation::WrongBodyType, "wrong storage type"),
        (Malformation::WrongBodyWire, "wrong storage wire type"),
        (
            Malformation::AmbiguousTitleOwner,
            "ambiguous drawable owner",
        ),
    ];

    for (malformed, label) in cases {
        let bytes = synthetic_package(malformed)?;
        let package = Package::from_bytes(&bytes)?;
        let error = package
            .show()
            .err()
            .ok_or_else(|| io::Error::other(format!("{label} should fail closed")))?;
        match label {
            "wrong storage wire type" => assert!(matches!(
                error,
                ReadError::TextStorage {
                    reason: TextStorageFailure::WrongWireType,
                    path: SemanticPath::SlideBody { index: 0 },
                }
            )),
            _ => assert!(matches!(
                error,
                ReadError::InvalidFormat(_) | ReadError::Decode(_)
            )),
        }
    }
    Ok(())
}

#[test]
fn missing_required_proto2_envelope_fields_fail_closed() -> TestResult<()> {
    let cases = [
        Malformation::MissingShowTheme,
        Malformation::MissingSlideNodeFlag,
        Malformation::MissingSlideStyle,
        Malformation::MissingBuildAttributes,
        Malformation::MissingPlaceholderSuper,
        Malformation::MissingShapeSuper,
        Malformation::MissingNoteStorage,
    ];
    for malformed in cases {
        let bytes = synthetic_package(malformed)?;
        let package = Package::from_bytes(&bytes)?;
        assert!(matches!(package.show(), Err(ReadError::InvalidFormat(_))));
    }
    Ok(())
}

#[test]
fn missing_required_document_envelope_fails_at_ingress() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::MissingDocumentShow)?;
    let error = Package::from_bytes(&bytes).err();
    assert!(
        matches!(
            error,
            Some(ReadError::InvalidFormat(_) | ReadError::Detection(_))
        ),
        "unexpected missing-document-envelope result: {error:?}"
    );
    Ok(())
}

#[test]
fn unproven_legacy_type_2022_sibling_is_not_misclassified_as_storage() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::Legacy2022Sibling)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.show()?.slides()[0].text_storages()[0].text(),
        "Body fragment"
    );
    Ok(())
}

#[test]
fn build_identifier_budget_is_checked_before_prost_materialization() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::OversizedBuildIdentifier)?;
    let limits = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        "future-effect".len() - 1,
    )?;
    let package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), limits))?;
    assert!(matches!(
        package.show(),
        Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::TextBytes,
            observed,
            maximum,
            path: SemanticPath::SlideBuild { slide: 0, index: 0 },
        }) if observed == "future-effect".len() && maximum == "future-effect".len() - 1
    ));
    Ok(())
}

#[test]
fn transition_identifier_budget_is_checked_before_slide_materialization() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::OversizedTransitionIdentifier)?;
    let limits = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        "future-transition".len() - 1,
    )?;
    let package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), limits))?;
    assert!(matches!(
        package.show(),
        Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::TextBytes,
            observed,
            maximum,
            path: SemanticPath::SlideTransition { index: 0 },
        }) if observed == "future-transition".len()
            && maximum == "future-transition".len() - 1
    ));
    Ok(())
}

#[test]
fn concurrent_first_access_is_deterministic() -> TestResult<()> {
    let bytes = synthetic_package(Malformation::default())?;
    let package = Arc::new(Package::from_bytes(&bytes)?);
    let mut handles = Vec::new();
    handles.try_reserve_exact(8)?;
    for _ in 0..8 {
        let task_package = Arc::clone(&package);
        handles.push(std::thread::spawn(move || {
            let slides = task_package.show()?.slides().len();
            let text = task_package.text()?;
            Ok::<_, ReadError>((slides, text))
        }));
    }
    for handle in handles {
        let result = handle
            .join()
            .map_err(|_panic| io::Error::other("Keynote reader thread panicked"))??;
        assert_eq!(result.0, 1);
        assert_eq!(result.1, "Rich title\nBody fragment\nShape text\nNote run");
    }
    Ok(())
}

#[test]
fn native_storage_messages_match_the_prost_oracle() -> TestResult<()> {
    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key");
    let bytes = std::fs::read(path)?;
    let catalog = ComponentCatalog::from_bytes(&bytes)?;
    let mut checked = 0usize;
    for component in catalog.iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                // Type 2001 is the schema-owned TSWP.StorageArchive identity.
                // Legacy dispatch also guessed at 2022, but native packages
                // contain type-2022 payloads that the generated schema rejects.
                if message.type_ != 2_001 {
                    continue;
                }
                let oracle = tswp::StorageArchive::decode(message.data.as_slice())?;
                let projected = litchi_iwa_text_wire::from_bytes(&message.data)?;
                assert_eq!(projected.text(), oracle.text.concat());
                assert_eq!(projected.runs().len(), oracle.text.len());
                checked = checked
                    .checked_add(1)
                    .ok_or_else(|| io::Error::other("native storage count overflowed usize"))?;
            }
        }
    }
    assert!(checked > 0);
    Ok(())
}
