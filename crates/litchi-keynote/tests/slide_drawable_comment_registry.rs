//! Focused author-registry compatibility coverage for direct drawable comments.
//!
//! The native source owns its annotation-author storage through the root
//! Document.  Older hosts also emitted a valid type-213 registry without that
//! edge, so the focused owner has a deliberately narrow compatibility route:
//! a single valid unrooted registry may be used, while a malformed, duplicate,
//! or dangling rooted witness remains a hard error.  The wire helpers below
//! manufacture those package shapes without exposing archive identifiers to
//! the public assertions.

use std::{
    collections::{BTreeMap, BTreeSet},
    env, fs, io,
    path::Path,
};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, append_varint_field, patch_length_delimited_field,
    transform_length_delimited_field,
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::tsp;
use litchi_keynote::{DrawableSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;
type TestError = Box<dyn std::error::Error>;

const SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/drawable-comments-source-native.key");
const OUTPUT_ENV: &str = "LITCHI_KEYNOTE_DRAWABLE_COMMENT_REGISTRY_OUTPUT_DIR";
const NATIVE_DIR_ENV: &str = "LITCHI_KEYNOTE_DRAWABLE_COMMENT_REGISTRY_NATIVE_DIR";
const NATIVE_FIXTURE_PATH: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/drawable-comments-cross-component-native/",
    "unrooted-unique-native-create-reply.key"
);
const NATIVE_CANDIDATE: &str = "unrooted-unique-native-create-reply.key";
const NATIVE_ROOT_TEXT: &str = "native unrooted registry";
const NATIVE_REPLY_TEXT: &str = "native registry reply";

const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const DOCUMENT_OBJECT_IDENTIFIER: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const DOCUMENT_SUPER_FIELD: u32 = 3;
const TSA_DOCUMENT_SUPER_FIELD: u32 = 1;
const TSK_ANNOTATION_AUTHOR_STORAGE_FIELD: u32 = 7;

#[derive(Debug, Clone, PartialEq, Eq)]
struct RegistrySnapshot {
    author_objects: BTreeSet<u64>,
    storage_references: BTreeMap<u64, Vec<u64>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawableSemantic {
    kind: litchi_keynote::DrawableKind,
    has_comment: bool,
    reply_count: usize,
    comment_text: Option<String>,
    comment_has_author: bool,
    reply_texts: Vec<String>,
    reply_authors: Vec<bool>,
}

#[derive(Debug, Clone, Copy)]
enum RootMutation {
    Duplicate,
    WrongWire,
    Dangling,
    Malformed,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn target(package: &Package) -> TestResult<DrawableSelector> {
    package
        .slide_drawables(SlideSelector::index(0))?
        .iter()
        .enumerate()
        .find_map(|(index, summary)| {
            (!summary.has_comment()).then_some(DrawableSelector::index(index))
        })
        .ok_or_else(|| io::Error::other("native source has no empty drawable").into())
}

fn export_if_requested(bytes: &[u8], name: &str) -> TestResult<()> {
    let Some(directory) = env::var_os(OUTPUT_ENV) else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    let path = directory.join(name);
    fs::write(&path, bytes)?;
    eprintln!("exported Keynote registry candidate to {}", path.display());
    Ok(())
}

fn drawable_semantics(package: &Package) -> TestResult<Vec<DrawableSemantic>> {
    let slide = SlideSelector::index(0);
    package
        .slide_drawables(slide)?
        .iter()
        .enumerate()
        .map(|(index, summary)| {
            let selector = DrawableSelector::index(index);
            let comment = package.slide_drawable_comment(slide, selector)?;
            let replies = package.slide_drawable_comment_replies(slide, selector)?;
            Ok(DrawableSemantic {
                kind: summary.kind(),
                has_comment: summary.has_comment(),
                reply_count: summary.reply_count(),
                comment_text: comment.as_ref().map(|comment| comment.text().to_owned()),
                comment_has_author: comment
                    .as_ref()
                    .and_then(|comment| comment.author())
                    .is_some(),
                reply_texts: replies
                    .iter()
                    .map(|reply| reply.text().to_owned())
                    .collect(),
                reply_authors: replies
                    .iter()
                    .map(|reply| reply.author().is_some())
                    .collect(),
            })
        })
        .collect()
}

fn component_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        archives.push((entry.name().to_owned(), archive));
    }
    Ok(archives)
}

fn replace_component(source: &[u8], component_name: &str, archive: Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component_name {
                (entry.name(), compressed.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn max_identifier(source: &[u8]) -> TestResult<u64> {
    component_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| archive.objects.into_iter())
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .ok_or_else(|| io::Error::other("native source has no archive objects").into())
}

fn storage_locations(source: &[u8]) -> TestResult<Vec<(String, u64, usize)>> {
    let mut locations = Vec::new();
    for (component_name, archive) in component_archives(source)? {
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE {
                    locations.push((component_name.clone(), identifier, message_index));
                }
            }
        }
    }
    Ok(locations)
}

fn annotation_author_registry(source: &[u8]) -> TestResult<RegistrySnapshot> {
    let mut author_objects = BTreeSet::new();
    let mut storage_references = BTreeMap::new();
    for (_, archive) in component_archives(source)? {
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("registry object has no identifier"))?;
            for message in &object.messages {
                match message.type_ {
                    ANNOTATION_AUTHOR_MESSAGE_TYPE => {
                        author_objects.insert(identifier);
                    },
                    ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE => {
                        let references = WireView::parse(&message.data)?
                            .fields()
                            .filter(|field| field.number() == 1)
                            .map(|field| {
                                if field.wire_type() != 2 {
                                    return Err(io::Error::other(
                                        "author registry reference has the wrong wire type",
                                    )
                                    .into());
                                }
                                let reference = tsp::Reference::decode(field.payload())?;
                                if reference.identifier == 0 {
                                    return Err(io::Error::other(
                                        "author registry contains a zero identifier",
                                    )
                                    .into());
                                }
                                Ok(reference.identifier)
                            })
                            .collect::<TestResult<Vec<_>>>()?;
                        storage_references.insert(identifier, references);
                    },
                    _ => {},
                }
            }
        }
    }
    Ok(RegistrySnapshot {
        author_objects,
        storage_references,
    })
}

fn document_root_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let (_, archive) = component_archives(source)?
        .into_iter()
        .find(|(name, _)| name == DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("native Document component is missing"))?;
    let object = archive
        .object(DOCUMENT_OBJECT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("native Document root is missing"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("native Document message is missing").into())
}

fn reference_at_path(payload: &[u8], path: &[u32]) -> TestResult<u64> {
    let mut current = payload;
    for field_number in path {
        let fields = WireView::parse(current)?
            .fields()
            .filter(|field| field.number() == *field_number)
            .collect::<Vec<_>>();
        if fields.len() != 1 || fields[0].wire_type() != 2 {
            return Err(io::Error::other("Document registry path is malformed").into());
        }
        current = fields[0].payload();
    }
    let reference = tsp::Reference::decode(current)?;
    if reference.identifier == 0 {
        return Err(io::Error::other("Document registry reference is zero").into());
    }
    Ok(reference.identifier)
}

fn rewrite_document_root(
    source: &[u8],
    mutate: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
    update_references: impl FnOnce(&mut Vec<u64>, u64) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let storage_identifier = reference_at_path(
        &document_root_payload(source)?,
        &[
            DOCUMENT_SUPER_FIELD,
            TSA_DOCUMENT_SUPER_FIELD,
            TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
        ],
    )?;
    let (_, mut archive) = component_archives(source)?
        .into_iter()
        .find(|(name, _)| name == DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("native Document component is missing"))?;
    let root = archive
        .object_mut(DOCUMENT_OBJECT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("native Document root is missing"))?;
    let message_index = root
        .messages
        .iter()
        .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native Document message is missing"))?;
    let data = mutate(&root.messages[message_index].data)?;
    update_references(
        &mut root.archive_info.message_infos[message_index].object_references,
        storage_identifier,
    )?;
    root.replace_message(
        message_index,
        RawMessage {
            type_: DOCUMENT_MESSAGE_TYPE,
            data,
        },
    )?;
    replace_component(source, DOCUMENT_COMPONENT, archive)
}

fn remove_root_storage_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_root(
        source,
        |document| -> TestResult<Vec<u8>> {
            transform_length_delimited_field(
                document,
                DOCUMENT_SUPER_FIELD,
                |tsa| -> TestResult<Vec<u8>> {
                    transform_length_delimited_field(
                        tsa,
                        TSA_DOCUMENT_SUPER_FIELD,
                        |tsk| -> TestResult<Vec<u8>> {
                            Ok(patch_length_delimited_field(
                                tsk,
                                TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
                                true,
                                None,
                            )?)
                        },
                    )
                },
            )
        },
        |references, storage_identifier| {
            references.retain(|identifier| *identifier != storage_identifier);
            Ok(())
        },
    )
}

fn mutate_root_storage_reference(source: &[u8], mutation: RootMutation) -> TestResult<Vec<u8>> {
    let root = document_root_payload(source)?;
    let storage_identifier = reference_at_path(
        &root,
        &[
            DOCUMENT_SUPER_FIELD,
            TSA_DOCUMENT_SUPER_FIELD,
            TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
        ],
    )?;
    let replacement = tsp::Reference {
        identifier: storage_identifier,
        ..Default::default()
    }
    .encode_to_vec();
    rewrite_document_root(
        source,
        |document| -> TestResult<Vec<u8>> {
            transform_length_delimited_field(
                document,
                DOCUMENT_SUPER_FIELD,
                |tsa| -> TestResult<Vec<u8>> {
                    transform_length_delimited_field(
                        tsa,
                        TSA_DOCUMENT_SUPER_FIELD,
                        |tsk| -> TestResult<Vec<u8>> {
                            match mutation {
                                RootMutation::Duplicate => {
                                    let mut output = tsk.to_vec();
                                    append_length_delimited_field(
                                        &mut output,
                                        TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
                                        &replacement,
                                    )?;
                                    Ok(output)
                                },
                                RootMutation::WrongWire => {
                                    let mut output = patch_length_delimited_field(
                                        tsk,
                                        TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
                                        true,
                                        None,
                                    )?;
                                    append_varint_field(
                                        &mut output,
                                        TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
                                        storage_identifier,
                                    )?;
                                    Ok(output)
                                },
                                RootMutation::Dangling => patch_length_delimited_field(
                                    tsk,
                                    TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
                                    true,
                                    Some(
                                        &tsp::Reference {
                                            identifier: u64::MAX - 37,
                                            ..Default::default()
                                        }
                                        .encode_to_vec(),
                                    ),
                                ),
                                RootMutation::Malformed => patch_length_delimited_field(
                                    tsk,
                                    TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
                                    true,
                                    Some(&[0xff]),
                                ),
                            }
                            .map_err(|error| -> TestError { Box::new(error) })
                        },
                    )
                },
            )
        },
        |references, storage_identifier| {
            if matches!(mutation, RootMutation::Dangling) {
                references.retain(|identifier| *identifier != storage_identifier);
                references.push(u64::MAX - 37);
            }
            Ok(())
        },
    )
}

fn without_storage_object(source: &[u8]) -> TestResult<Vec<u8>> {
    let locations = storage_locations(source)?;
    if locations.len() != 1 {
        return Err(io::Error::other("source does not have one author registry").into());
    }
    let (component_name, identifier, _) = &locations[0];
    let (_, mut archive) = component_archives(source)?
        .into_iter()
        .find(|(name, _)| name == component_name)
        .ok_or_else(|| io::Error::other("author registry component is missing"))?;
    archive
        .remove_object(*identifier)
        .ok_or_else(|| io::Error::other("author registry object disappeared"))?;
    replace_component(source, component_name, archive)
}

fn with_detached_registry(source: &[u8]) -> TestResult<(Vec<u8>, u64)> {
    let locations = storage_locations(source)?;
    if locations.len() != 1 {
        return Err(io::Error::other("source does not have one author registry").into());
    }
    let (component_name, identifier, _) = &locations[0];
    let (_, mut archive) = component_archives(source)?
        .into_iter()
        .find(|(name, _)| name == component_name)
        .ok_or_else(|| io::Error::other("author registry component is missing"))?;
    let source_object = archive
        .object(*identifier)
        .ok_or_else(|| io::Error::other("author registry object disappeared"))?;
    let new_identifier = max_identifier(source)?
        .checked_add(17)
        .ok_or_else(|| io::Error::other("archive identifier overflow"))?;
    let clone = source_object.clone_with_identity_remap(
        new_identifier,
        &[(*identifier, new_identifier)],
        &source_object.messages,
    )?;
    archive.insert_object(clone)?;
    Ok((
        replace_component(source, component_name, archive)?,
        new_identifier,
    ))
}

fn assert_rejected_atomically(source: &[u8], label: &str) -> TestResult<()> {
    let package = Package::from_bytes(source)?;
    let before = exact_bytes(&package)?;
    let target = target(&package)?;
    let result = package
        .edit_slide_drawable_comment(SlideSelector::index(0), target)
        .and_then(|edit| edit.set(format!("rejected {label}")))
        .and_then(|edit| edit.commit());
    assert!(result.is_err(), "malformed registry was accepted: {label}");
    assert_eq!(
        exact_bytes(&package)?,
        before,
        "mutation was not atomic: {label}"
    );
    Ok(())
}

#[test]
fn unrooted_unique_registry_creates_and_cleans_generated_author() -> TestResult {
    let source = remove_root_storage_reference(SOURCE)?;
    assert_eq!(storage_locations(&source)?.len(), 1);
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let before_registry = annotation_author_registry(&before)?;
    export_if_requested(&before, "unrooted-unique-source.key")?;
    let selected = target(&package)?;

    let created = package
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set("unrooted unique registry")?
        .commit()?;
    let created_comment = created
        .package()
        .slide_drawable_comment(SlideSelector::index(0), selected)?
        .ok_or_else(|| io::Error::other("unrooted registry comment was not created"))?;
    assert!(created_comment.author().is_some());
    let created_registry = annotation_author_registry(&exact_bytes(created.package())?)?;
    assert_eq!(
        created_registry.author_objects.len(),
        before_registry.author_objects.len() + 1
    );
    assert_eq!(
        created_registry.storage_references.len(),
        before_registry.storage_references.len()
    );
    let noop = created
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set("unrooted unique registry")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(
        exact_bytes(noop.package())?,
        exact_bytes(created.package())?
    );
    export_if_requested(
        &exact_bytes(created.package())?,
        "unrooted-unique-create.key",
    )?;

    let native_root = package
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set(NATIVE_ROOT_TEXT)?
        .commit()?;
    let native_candidate = native_root
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .add_reply(NATIVE_REPLY_TEXT)?
        .commit()?;
    export_if_requested(&exact_bytes(native_candidate.package())?, NATIVE_CANDIDATE)?;

    let cleared = created
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(SlideSelector::index(0), selected)?
            .is_none()
    );
    assert_eq!(
        annotation_author_registry(&exact_bytes(cleared.package())?)?,
        before_registry
    );
    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(created.package())?
    );
    Ok(())
}

#[test]
fn rooted_registry_is_authoritative_and_detached_registry_survives() -> TestResult {
    let (source, detached_identifier) = with_detached_registry(SOURCE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let before_registry = annotation_author_registry(&before)?;
    assert!(
        before_registry
            .storage_references
            .contains_key(&detached_identifier)
    );
    let selected = target(&package)?;
    let noop = package
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .clear()?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, before);

    let created = package
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set("rooted registry authority")?
        .commit()?;
    let created_bytes = exact_bytes(created.package())?;
    let created_registry = annotation_author_registry(&created_bytes)?;
    assert_eq!(
        created_registry
            .storage_references
            .get(&detached_identifier),
        before_registry.storage_references.get(&detached_identifier)
    );
    assert_eq!(
        created_registry.storage_references.len(),
        before_registry.storage_references.len()
    );
    let noop = created
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set("rooted registry authority")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, created_bytes);
    export_if_requested(&created_bytes, "root-create-registry.key")?;

    let cleared = created
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .clear()?
        .commit()?;
    assert_eq!(
        annotation_author_registry(&exact_bytes(cleared.package())?)?,
        before_registry
    );
    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, created_bytes);
    Ok(())
}

#[test]
fn no_registry_fallback_creates_an_authorless_comment() -> TestResult {
    let source = without_storage_object(&remove_root_storage_reference(SOURCE)?)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let selected = target(&package)?;
    let created = package
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set("authorless registry fallback")?
        .commit()?;
    let comment = created
        .package()
        .slide_drawable_comment(SlideSelector::index(0), selected)?
        .ok_or_else(|| io::Error::other("authorless comment was not created"))?;
    assert!(comment.author().is_none());
    let noop = created
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .set("authorless registry fallback")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(
        exact_bytes(noop.package())?,
        exact_bytes(created.package())?
    );
    let cleared = created
        .package()
        .edit_slide_drawable_comment(SlideSelector::index(0), selected)?
        .clear()?
        .commit()?;
    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(created.package())?
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn unrooted_duplicate_or_malformed_registry_is_rejected_atomically() -> TestResult {
    let unrooted = remove_root_storage_reference(SOURCE)?;
    let (duplicate, _) = with_detached_registry(&unrooted)?;
    assert_rejected_atomically(&duplicate, "duplicate unrooted registry")?;

    let malformed = rewrite_storage_payload(&unrooted, |_| Ok(vec![0xff]))?;
    assert_rejected_atomically(&malformed, "malformed unrooted registry")?;
    Ok(())
}

#[test]
fn present_root_registry_witness_never_falls_back() -> TestResult {
    for mutation in [
        RootMutation::Duplicate,
        RootMutation::WrongWire,
        RootMutation::Dangling,
        RootMutation::Malformed,
    ] {
        let source = mutate_root_storage_reference(SOURCE, mutation)?;
        assert_rejected_atomically(&source, "present malformed registry root")?;
    }
    Ok(())
}

#[test]
fn native_resaved_unique_unrooted_registry_readback() -> TestResult {
    let candidate_bytes = if let Some(directory) = env::var_os(NATIVE_DIR_ENV) {
        let path = Path::new(&directory).join(NATIVE_CANDIDATE);
        assert!(
            path.is_file(),
            "missing native registry candidate {}; export the source with {OUTPUT_ENV}, save it in Keynote, and rerun with {NATIVE_DIR_ENV}",
            path.display()
        );
        fs::read(path)?
    } else {
        fs::read(NATIVE_FIXTURE_PATH).map_err(|error| {
            io::Error::other(format!(
                "read permanent native registry fixture {NATIVE_FIXTURE_PATH}: {error}"
            ))
        })?
    };

    let source_package = Package::from_bytes(SOURCE)?;
    let source_semantics = drawable_semantics(&source_package)?;
    let target_index = source_semantics
        .iter()
        .position(|drawable| !drawable.has_comment)
        .ok_or_else(|| io::Error::other("native source has no empty drawable"))?;
    let candidate = Package::from_bytes(&candidate_bytes)?;
    let candidate_semantics = drawable_semantics(&candidate)?;
    assert_eq!(candidate_semantics.len(), source_semantics.len());

    for (index, (source, candidate)) in source_semantics
        .iter()
        .zip(&candidate_semantics)
        .enumerate()
    {
        if index == target_index {
            assert_eq!(candidate.kind, source.kind);
            assert!(candidate.has_comment);
            assert_eq!(candidate.comment_text.as_deref(), Some(NATIVE_ROOT_TEXT));
            assert!(candidate.comment_has_author);
            assert_eq!(candidate.reply_count, 1);
            assert_eq!(candidate.reply_texts, vec![NATIVE_REPLY_TEXT.to_owned()]);
            assert_eq!(candidate.reply_authors, vec![true]);
        } else {
            assert_eq!(
                candidate, source,
                "native sibling drawable changed at {index}"
            );
        }
    }
    Ok(())
}

fn rewrite_storage_payload(
    source: &[u8],
    mutate: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let locations = storage_locations(source)?;
    if locations.len() != 1 {
        return Err(io::Error::other("source does not have one author registry").into());
    }
    let (component_name, identifier, message_index) = &locations[0];
    let (_, mut archive) = component_archives(source)?
        .into_iter()
        .find(|(name, _)| name == component_name)
        .ok_or_else(|| io::Error::other("author registry component is missing"))?;
    let object = archive
        .object_mut(*identifier)
        .ok_or_else(|| io::Error::other("author registry object disappeared"))?;
    let payload = mutate(
        &object
            .messages
            .get(*message_index)
            .ok_or_else(|| io::Error::other("author registry message disappeared"))?
            .data,
    )?;
    object.replace_message(
        *message_index,
        RawMessage {
            type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
            data: payload,
        },
    )?;
    replace_component(source, component_name, archive)
}
