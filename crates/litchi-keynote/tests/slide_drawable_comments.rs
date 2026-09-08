//! Selector-first integration coverage for Keynote drawable comments.
//!
//! The public assertions in this file use only slide, drawable, and reply
//! selectors.  Native identifiers are confined to the small wire oracle at
//! the bottom of the file; that oracle checks the metadata and copy-on-write
//! invariants which are deliberately not part of the focused value API.

use std::{
    collections::{BTreeMap, BTreeSet},
    env, fs, io,
    path::Path,
};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    decode_varint_from_bytes, encode_varint_into,
    wire::{WireView, append_length_delimited_field},
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsd, tsp};
use litchi_keynote::{
    DrawableSelector, MovieInfo, Package, ReadOptions, ReplySelector, SemanticLimits,
    SlideDrawableCommentError, SlideDrawableCommentLimitKind, SlideSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/drawable-comments-source-native.key");
const OUTPUT_ENV: &str = "LITCHI_KEYNOTE_DRAWABLE_COMMENTS_OUTPUT_DIR";
const NATIVE_DIR_ENV: &str = "LITCHI_KEYNOTE_DRAWABLE_COMMENTS_NATIVE_DIR";
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const UNKNOWN_COMMENT_EXTENSION_FIELD: u32 = 9_001;
const UNKNOWN_COMMENT_EXTENSION_PAYLOAD: &[u8] =
    b"litchi-keynote-drawable-comment-unknown-extension";
const UNKNOWN_COMMENT_HEADER_FIELD: u32 = 9_002;
const NATIVE_CANDIDATES: [&str; 6] = [
    "root-create.key",
    "root-update.key",
    "root-clear.key",
    "reply-add.key",
    "reply-update.key",
    "reply-remove.key",
];

fn source_bytes() -> TestResult<Vec<u8>> {
    Ok(SOURCE.to_vec())
}

fn with_unknown_comment_extensions(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut locations = BTreeMap::<u64, (String, usize)>::new();
    let mut roots = Vec::<(u64, Vec<u64>)>::new();
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
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("comment object has no archive identifier"))?;
            let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
            let replies = comment
                .replies
                .into_iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            locations.insert(identifier, (entry.name().to_owned(), message_index));
            if !replies.is_empty() {
                roots.push((identifier, replies));
            }
        }
    }

    let mut targets = BTreeMap::<String, BTreeMap<u64, usize>>::new();
    for (root, replies) in roots {
        let Some(reply) = replies
            .into_iter()
            .find(|identifier| locations.contains_key(identifier))
        else {
            continue;
        };
        for identifier in [root, reply] {
            let (component, message_index) = locations
                .get(&identifier)
                .ok_or_else(|| io::Error::other("comment reply location disappeared"))?;
            targets
                .entry(component.clone())
                .or_default()
                .insert(identifier, *message_index);
        }
    }
    if targets.is_empty() {
        return Err(io::Error::other("fixture has no rooted comment reply").into());
    }

    let mut replacements = BTreeMap::<String, Vec<u8>>::new();
    for entry in catalog.iter() {
        let Some(component_targets) = targets.get(entry.name()) else {
            continue;
        };
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        let mut archive = Archive::parse(&stream)?;
        for (identifier, message_index) in component_targets {
            let object = archive
                .object_mut(*identifier)
                .ok_or_else(|| io::Error::other("comment target disappeared"))?;
            let (type_, mut payload) = object
                .messages
                .get(*message_index)
                .map(|message| (message.type_, message.data.clone()))
                .ok_or_else(|| io::Error::other("comment payload disappeared"))?;
            append_length_delimited_field(
                &mut payload,
                UNKNOWN_COMMENT_EXTENSION_FIELD,
                UNKNOWN_COMMENT_EXTENSION_PAYLOAD,
            )?;
            object.replace_message(
                *message_index,
                RawMessage {
                    type_,
                    data: payload,
                },
            )?;
        }
        replacements.insert(
            entry.name().to_owned(),
            SnappyStream::compress(&archive.to_bytes()?)?,
        );
    }

    let entries = catalog
        .iter()
        .map(|entry| {
            replacements.get(entry.name()).map_or_else(
                || (entry.name(), entry.data()),
                |replacement| (entry.name(), replacement.as_slice()),
            )
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn unknown_comment_extension_spans(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let mut spans = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
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
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                spans.extend(
                    WireView::parse(&message.data)?
                        .fields()
                        .filter(|field| field.number() == UNKNOWN_COMMENT_EXTENSION_FIELD)
                        .map(|field| field.raw().to_vec()),
                );
            }
        }
    }
    spans.sort();
    Ok(spans)
}

fn with_unknown_comment_header(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected = None;
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
        if let Some(object) = archive.objects.iter().find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        }) {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("comment object has no archive identifier"))?;
            selected = Some((entry.name().to_owned(), identifier));
            break;
        }
    }
    let (component, identifier) =
        selected.ok_or_else(|| io::Error::other("fixture has no comment object"))?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == component)
        .ok_or_else(|| io::Error::other("comment component disappeared"))?;
    let bytes = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&bytes)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("comment object disappeared"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(
        bytes
            .get(header_offset..)
            .ok_or_else(|| io::Error::other("comment header offset is invalid"))?,
    )?;
    let header_length = usize::try_from(header_length)?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("comment header start overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("comment header end overflow"))?;
    if header_end != data_offset {
        return Err(io::Error::other("comment archive offsets disagree").into());
    }
    let mut header = bytes
        .get(header_start..header_end)
        .ok_or_else(|| io::Error::other("comment header range is invalid"))?
        .to_vec();
    append_length_delimited_field(
        &mut header,
        UNKNOWN_COMMENT_HEADER_FIELD,
        UNKNOWN_COMMENT_EXTENSION_PAYLOAD,
    )?;
    let mut modified = Vec::with_capacity(bytes.len().saturating_add(header.len()));
    modified.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut modified, u64::try_from(header.len())?);
    modified.extend_from_slice(&header);
    modified.extend_from_slice(&bytes[data_offset..]);
    assert_eq!(Archive::parse(&modified)?.to_bytes()?, modified);

    let replacement = SnappyStream::compress(&modified)?;
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
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn export_if_requested(package: &Package, name: &str) -> TestResult<()> {
    let Some(directory) = env::var_os(OUTPUT_ENV) else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    let path = directory.join(name);
    fs::write(&path, exact_bytes(package)?)?;
    eprintln!(
        "exported Keynote drawable-comment candidate to {}",
        path.display()
    );
    Ok(())
}

fn summaries(package: &Package) -> TestResult<Vec<(DrawableSelector, bool)>> {
    Ok(package
        .slide_drawables(SlideSelector::index(0))?
        .iter()
        .enumerate()
        .map(|(index, summary)| (DrawableSelector::index(index), summary.has_comment()))
        .collect())
}

fn slide_semantic_baseline(
    package: &Package,
) -> TestResult<Vec<(Option<String>, Option<String>, Vec<MovieInfo>)>> {
    Ok(package
        .slides()?
        .iter()
        .map(|slide| {
            (
                slide.name().map(str::to_owned),
                slide.title().map(str::to_owned),
                slide.movies().to_vec(),
            )
        })
        .collect())
}

fn existing_target(package: &Package) -> TestResult<DrawableSelector> {
    summaries(package)?
        .into_iter()
        .find_map(|(selector, has_comment)| has_comment.then_some(selector))
        .ok_or_else(|| io::Error::other("fixture has no commented drawable").into())
}

fn empty_target(package: &Package) -> TestResult<DrawableSelector> {
    summaries(package)?
        .into_iter()
        .find_map(|(selector, has_comment)| (!has_comment).then_some(selector))
        .ok_or_else(|| io::Error::other("fixture has no empty drawable").into())
}

#[test]
fn permanent_native_fixture_is_pinned() {
    assert_eq!(SOURCE.len(), 753_541);
}

#[test]
fn entry_limit_is_typed_and_source_stays_unchanged() -> TestResult {
    let source = source_bytes()?;
    let limits = Limits::new(
        8 * 1024 * 1024,
        256,
        2 * 1024 * 1024,
        8 * 1024 * 1024,
        2 * 1024 * 1024,
    )?;
    let semantic = SemanticLimits::new(
        16 * 1024,
        512,
        32 * 1024,
        8 * 1024,
        32 * 1024,
        2 * 1024 * 1024,
    )?;
    let package = Package::from_bytes_with_options(&source, ReadOptions::new(limits, semantic))?;
    let slide = SlideSelector::index(0);
    let inventory = package.slide_drawables(slide)?;
    let target = inventory
        .iter()
        .find(|summary| summary.has_comment())
        .map(|summary| summary.selector())
        .ok_or_else(|| io::Error::other("fixture has no commented drawable"))?;
    let comment = package.slide_drawable_comment(slide, target);
    assert!(
        matches!(
            comment.as_ref(),
            Err(SlideDrawableCommentError::LimitExceeded {
                kind: SlideDrawableCommentLimitKind::Entries,
                observed: 257,
                maximum: 256,
            })
        ),
        "unexpected comment read error: {:?}",
        comment.as_ref().err()
    );
    assert_eq!(exact_bytes(&package)?, source);

    let package = Package::from_bytes_with_options(&source, ReadOptions::new(limits, semantic))?;
    let mutation = package.edit_slide_drawable_comment(slide, target);
    assert!(
        matches!(
            mutation.as_ref(),
            Err(SlideDrawableCommentError::LimitExceeded {
                kind: SlideDrawableCommentLimitKind::Entries,
                observed: 257,
                maximum: 256,
            })
        ),
        "unexpected mutation error: {:?}",
        mutation.as_ref().err()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn native_resaved_drawable_comment_candidates_read_back() -> TestResult {
    let Some(directory) = env::var_os(NATIVE_DIR_ENV) else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    let source = source_bytes()?;
    let source_package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let root_target = empty_target(&source_package)?;
    let thread_target = existing_target(&source_package)?;
    let source_replies = source_package.slide_drawable_comment_replies(slide, thread_target)?;
    let source_thread_text = source_package
        .slide_drawable_comment(slide, thread_target)?
        .ok_or_else(|| io::Error::other("fixture comment target disappeared"))?
        .text()
        .to_owned();
    let source_baseline = slide_semantic_baseline(&source_package)?;

    for name in NATIVE_CANDIDATES {
        let path = directory.join(name);
        assert!(
            path.is_file(),
            "missing native candidate {}",
            path.display()
        );
        let bytes = fs::read(&path)?;
        let package = Package::from_bytes(&bytes)?;
        assert_eq!(
            slide_semantic_baseline(&package)?,
            source_baseline,
            "native comment candidate changed title or media semantics: {}",
            path.display()
        );

        match name {
            "root-create.key" => assert_eq!(
                package
                    .slide_drawable_comment(slide, root_target)?
                    .as_ref()
                    .map(|comment| comment.text()),
                Some("drawable root created")
            ),
            "root-update.key" => assert_eq!(
                package
                    .slide_drawable_comment(slide, root_target)?
                    .as_ref()
                    .map(|comment| comment.text()),
                Some("drawable root updated")
            ),
            "root-clear.key" => assert!(
                package
                    .slide_drawable_comment(slide, root_target)?
                    .is_none()
            ),
            "reply-add.key" => {
                let replies = package.slide_drawable_comment_replies(slide, thread_target)?;
                assert_eq!(replies.len(), source_replies.len() + 1);
                assert_eq!(
                    replies.last().map(|reply| reply.text()),
                    Some("duplicate reply")
                );
                assert_eq!(
                    package
                        .slide_drawable_comment(slide, thread_target)?
                        .as_ref()
                        .map(|comment| comment.text()),
                    Some(source_thread_text.as_str())
                );
            },
            "reply-update.key" => {
                let replies = package.slide_drawable_comment_replies(slide, thread_target)?;
                assert_eq!(replies.len(), source_replies.len() + 2);
                assert_eq!(
                    replies.last().map(|reply| reply.text()),
                    Some("duplicate reply updated")
                );
            },
            "reply-remove.key" => {
                let replies = package.slide_drawable_comment_replies(slide, thread_target)?;
                assert_eq!(replies.len(), source_replies.len() + 1);
                assert_eq!(
                    replies.last().map(|reply| reply.text()),
                    Some("duplicate reply")
                );
            },
            _ => unreachable!("candidate list is exhaustive"),
        }
    }
    Ok(())
}

#[test]
fn every_fixture_drawable_route_reads_through_selector_order() -> TestResult {
    let source = source_bytes()?;
    let package = Package::from_bytes(&source)?;
    let summaries = package.slide_drawables(SlideSelector::index(0))?;
    assert!(!summaries.is_empty(), "fixture has no owned drawables");

    for (index, summary) in summaries.iter().enumerate() {
        let selector = DrawableSelector::index(index);
        let comment = package.slide_drawable_comment(SlideSelector::index(0), selector)?;
        assert_eq!(comment.is_some(), summary.has_comment());
        let replies = package.slide_drawable_comment_replies(SlideSelector::index(0), selector)?;
        if let Some(comment) = comment {
            assert!(comment.text().len() <= 64 * 1024);
            for reply in &replies {
                assert!(reply.text().len() <= 64 * 1024);
            }
        } else {
            assert!(replies.is_empty());
        }
    }
    Ok(())
}

#[test]
fn root_comment_create_update_clear_and_inverse_are_exact() -> TestResult {
    let source = source_bytes()?;
    let package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let target = empty_target(&package)?;
    let before = exact_bytes(&package)?;

    let noop_clear = package
        .edit_slide_drawable_comment(slide, target)?
        .clear()?
        .commit()?;
    assert!(noop_clear.patch().is_noop());
    assert_eq!(exact_bytes(noop_clear.package())?, before);

    let created = package
        .edit_slide_drawable_comment(slide, target)?
        .set("drawable root created")?
        .commit()?;
    assert!(!created.patch().is_noop());
    assert_eq!(
        created
            .package()
            .slide_drawable_comment(slide, target)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("drawable root created")
    );
    let root_noop = created
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set("drawable root created")?
        .commit()?;
    assert!(root_noop.patch().is_noop());
    assert_eq!(
        exact_bytes(root_noop.package())?,
        exact_bytes(created.package())?
    );
    export_if_requested(created.package(), "root-create.key")?;

    let updated = created
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set("drawable root updated")?
        .commit()?;
    assert_eq!(
        updated
            .package()
            .slide_drawable_comment(slide, target)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("drawable root updated")
    );
    assert_metadata_preserved(
        &comment_signatures(&exact_bytes(created.package())?)?,
        &comment_signatures(&exact_bytes(updated.package())?)?,
        "drawable root created",
        "drawable root updated",
    );
    assert_created_metadata(
        &comment_signatures(&exact_bytes(created.package())?)?,
        "drawable root created",
    );
    export_if_requested(updated.package(), "root-update.key")?;

    let cleared = updated
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(slide, target)?
            .is_none()
    );
    export_if_requested(cleared.package(), "root-clear.key")?;

    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    assert_eq!(
        restored
            .package()
            .slide_drawable_comment(slide, target)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("drawable root updated")
    );
    let restored = restored
        .package()
        .apply_slide_drawable_comment(&updated.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(created.package())?
    );

    let original = package
        .edit_slide_drawable_comment(slide, target)?
        .set("drawable root created")?
        .commit()?;
    let conflict_target = original
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set("conflicting root")?
        .commit()?;
    let conflict_before = exact_bytes(conflict_target.package())?;
    assert!(
        conflict_target
            .package()
            .apply_slide_drawable_comment(created.patch())
            .is_err()
    );
    assert_eq!(exact_bytes(conflict_target.package())?, conflict_before);
    Ok(())
}

#[test]
fn root_set_preserves_replies_and_inverse_restores_exactly() -> TestResult {
    let source = source_bytes()?;
    let package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let target = existing_target(&package)?;
    let added = package
        .edit_slide_drawable_comment(slide, target)?
        .add_reply("reply retained across root set")?
        .commit()?;
    let before_set_bytes = exact_bytes(added.package())?;
    let before_replies = added
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    let changed = added
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set("root text changed with replies")?
        .commit()?;
    let after_replies = changed
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_replies.len(), before_replies.len());
    for (before, after) in before_replies.iter().zip(after_replies.iter()) {
        assert_eq!(after.text(), before.text());
        assert_eq!(after.timestamp(), before.timestamp());
        assert_eq!(after.author(), before.author());
    }

    let restored = changed
        .package()
        .apply_slide_drawable_comment(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, before_set_bytes);
    Ok(())
}

#[test]
fn unknown_comment_payload_extensions_survive_root_and_reply_rewrites() -> TestResult {
    let source_package = Package::from_bytes(SOURCE)?;
    let slide = SlideSelector::index(0);
    let target = existing_target(&source_package)?;
    let with_reply = source_package
        .edit_slide_drawable_comment(slide, target)?
        .add_reply("unknown extension setup reply")?
        .commit()?;
    let source = exact_bytes(with_reply.package())?;
    let mutated = with_unknown_comment_extensions(&source)?;
    let expected = {
        let mut field = Vec::new();
        append_length_delimited_field(
            &mut field,
            UNKNOWN_COMMENT_EXTENSION_FIELD,
            UNKNOWN_COMMENT_EXTENSION_PAYLOAD,
        )?;
        field
    };
    let before = unknown_comment_extension_spans(&mutated)?;
    assert!(
        before.len() >= 2,
        "fixture has no root and reply extensions"
    );
    assert!(before.iter().all(|span| span == &expected));

    let package = Package::from_bytes(&mutated)?;
    let replies = package.slide_drawable_comment_replies(slide, target)?;
    assert!(!replies.is_empty(), "extension target has no reply");
    let root_set = package
        .edit_slide_drawable_comment(slide, target)?
        .set("root extension preservation")?
        .commit()?;
    assert_eq!(
        unknown_comment_extension_spans(&exact_bytes(root_set.package())?)?,
        before
    );
    let reply_set = root_set
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set_reply(
            ReplySelector::index(replies.len() - 1),
            "reply extension preservation",
        )?
        .commit()?;
    assert_eq!(
        unknown_comment_extension_spans(&exact_bytes(reply_set.package())?)?,
        before
    );
    assert_eq!(
        reply_set
            .package()
            .slide_drawable_comment_replies(slide, target)?
            .last()
            .map(|reply| reply.text()),
        Some("reply extension preservation")
    );
    Ok(())
}

#[test]
fn unknown_comment_archive_headers_remain_rejected_atomically() -> TestResult {
    let source = with_unknown_comment_header(SOURCE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let slide = SlideSelector::index(0);
    let rejected = match package.slide_drawables(slide) {
        Err(SlideDrawableCommentError::InvalidSource) => true,
        Err(error) => panic!("unexpected inventory error: {error:?}"),
        Ok(summaries) => summaries
            .iter()
            .filter(|summary| summary.has_comment())
            .any(|summary| {
                matches!(
                    package.slide_drawable_comment(slide, summary.selector()),
                    Err(SlideDrawableCommentError::InvalidSource)
                )
            }),
    };
    assert!(rejected, "unknown comment header was accepted");
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn every_empty_drawable_route_supports_root_create_and_clear() -> TestResult {
    let source = source_bytes()?;
    let package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let mut exercised = 0usize;
    for (index, summary) in package.slide_drawables(slide)?.iter().enumerate() {
        if summary.has_comment() {
            continue;
        }
        let target = DrawableSelector::index(index);
        let text = format!("route root {index}");
        let created = package
            .edit_slide_drawable_comment(slide, target)?
            .set(&text)?
            .commit()?;
        assert_eq!(
            created
                .package()
                .slide_drawable_comment(slide, target)?
                .as_ref()
                .map(|comment| comment.text()),
            Some(text.as_str())
        );
        let cleared = created
            .package()
            .edit_slide_drawable_comment(slide, target)?
            .clear()?
            .commit()?;
        assert!(
            cleared
                .package()
                .slide_drawable_comment(slide, target)?
                .is_none()
        );
        exercised += 1;
    }
    assert!(exercised > 0, "fixture has no empty drawable route");
    Ok(())
}

#[test]
fn replies_are_ordered_by_ordinal_and_copy_on_write() -> TestResult {
    let source = source_bytes()?;
    let package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let target = existing_target(&package)?;
    let before = package.slide_drawable_comment_replies(slide, target)?;
    let root = package
        .slide_drawable_comment(slide, target)?
        .ok_or_else(|| io::Error::other("comment target disappeared"))?;

    let added = package
        .edit_slide_drawable_comment(slide, target)?
        .add_reply("duplicate reply")?
        .commit()?;
    let after_add = added
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_add.len(), before.len() + 1);
    assert_eq!(
        after_add.last().map(|reply| reply.text()),
        Some("duplicate reply")
    );
    assert_created_metadata(
        &comment_signatures(&exact_bytes(added.package())?)?,
        "duplicate reply",
    );
    export_if_requested(added.package(), "reply-add.key")?;

    let added_twice = added
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .add_reply("duplicate reply")?
        .commit()?;
    let duplicate_replies = added_twice
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(
        duplicate_replies.last().map(|reply| reply.text()),
        Some("duplicate reply")
    );
    assert_eq!(
        duplicate_replies
            .iter()
            .filter(|reply| reply.text() == "duplicate reply")
            .count(),
        2
    );

    let reply_index = ReplySelector::index(duplicate_replies.len() - 1);
    let updated = added_twice
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set_reply(reply_index, "duplicate reply updated")?
        .commit()?;
    let after_set = updated
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_set.len(), duplicate_replies.len());
    assert_eq!(
        after_set.last().map(|reply| reply.text()),
        Some("duplicate reply updated")
    );
    let reply_noop = added_twice
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set_reply(reply_index, "duplicate reply")?
        .commit()?;
    assert!(reply_noop.patch().is_noop());
    assert_eq!(
        exact_bytes(reply_noop.package())?,
        exact_bytes(added_twice.package())?
    );
    assert_reply_metadata_preserved(
        &comment_signatures(&exact_bytes(added_twice.package())?)?,
        &comment_signatures(&exact_bytes(updated.package())?)?,
        "duplicate reply",
        "duplicate reply updated",
    );
    export_if_requested(updated.package(), "reply-update.key")?;

    let removed = updated
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .remove_reply(ReplySelector::index(after_set.len() - 1))?
        .commit()?;
    let after_remove = removed
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_remove.len(), after_set.len() - 1);
    assert_eq!(
        removed
            .package()
            .slide_drawable_comment(slide, target)?
            .as_ref()
            .map(|comment| comment.text()),
        Some(root.text())
    );
    export_if_requested(removed.package(), "reply-remove.key")?;

    let restored = removed
        .package()
        .apply_slide_drawable_comment(&removed.patch().inverse())?;
    assert_eq!(
        restored
            .package()
            .slide_drawable_comment_replies(slide, target)?
            .len(),
        after_set.len()
    );
    Ok(())
}

#[test]
fn independent_comments_isolate_the_selected_drawable() -> TestResult {
    let source = source_bytes()?;
    let package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let targets = summaries(&package)?
        .into_iter()
        .filter_map(|(selector, has_comment)| has_comment.then_some(selector))
        .collect::<Vec<_>>();
    assert!(
        targets.len() >= 2,
        "fixture must contain two independently commented drawables"
    );
    let selected = targets[0];
    let sibling = targets[1];
    let sibling_before = package
        .slide_drawable_comment(slide, sibling)?
        .ok_or_else(|| io::Error::other("shared sibling comment disappeared"))?;
    let changed = package
        .edit_slide_drawable_comment(slide, selected)?
        .set("selected drawable only")?
        .commit()?;
    assert_eq!(
        changed
            .package()
            .slide_drawable_comment(slide, selected)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("selected drawable only")
    );
    assert_eq!(
        changed
            .package()
            .slide_drawable_comment(slide, sibling)?
            .as_ref()
            .map(|comment| comment.text()),
        Some(sibling_before.text())
    );
    Ok(())
}

#[test]
fn malformed_comment_reply_graphs_are_rejected_atomically() -> TestResult {
    let source = source_bytes()?;
    for mutation in [
        HostileMutation::Cycle,
        HostileMutation::MissingReply,
        HostileMutation::MalformedWire,
    ] {
        let hostile = with_hostile_comment(&source, mutation)?;
        let Ok(package) = Package::from_bytes(&hostile) else {
            continue;
        };
        let before = exact_bytes(&package)?;
        let routes = match summaries(&package) {
            Ok(routes) => routes,
            Err(_) => {
                // Strict graph validation may reject the archive while
                // inventorying drawables. The parsed package remains
                // immutable and must still round-trip byte-for-byte.
                assert_eq!(exact_bytes(&package)?, before);
                continue;
            },
        };
        let mut rejected = false;
        for (selector, has_comment) in routes {
            if !has_comment {
                continue;
            }
            if package
                .slide_drawable_comment_replies(SlideSelector::index(0), selector)
                .is_err()
            {
                rejected = true;
                break;
            }
        }
        assert!(rejected, "hostile fixture was accepted: {mutation:?}");
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq)]
struct CommentSignature {
    identifier: u64,
    text: Option<String>,
    creation_date: Option<f64>,
    author: Option<u64>,
    uuid: Option<(u64, u64)>,
    replies: Vec<u64>,
}

fn comment_signatures(source: &[u8]) -> TestResult<Vec<CommentSignature>> {
    let mut signatures = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
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
            let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("missing archive object identifier"))?;
            signatures.push(CommentSignature {
                identifier,
                text: comment.text,
                creation_date: comment.creation_date.map(|date| date.seconds),
                author: comment.author.map(|reference| reference.identifier),
                uuid: comment.storage_uuid.map(|uuid| (uuid.lower, uuid.upper)),
                replies: comment
                    .replies
                    .into_iter()
                    .map(|reference| reference.identifier)
                    .collect(),
            });
        }
    }
    Ok(signatures)
}

fn assert_metadata_preserved(
    before: &[CommentSignature],
    after: &[CommentSignature],
    old: &str,
    new: &str,
) {
    let source = before
        .iter()
        .find(|signature| signature.text.as_deref() == Some(old))
        .expect("source comment text");
    let candidate = after
        .iter()
        .find(|signature| signature.text.as_deref() == Some(new))
        .expect("candidate comment text");
    assert_eq!(candidate.creation_date, source.creation_date);
    assert_eq!(candidate.author, source.author);
    assert_eq!(candidate.uuid, source.uuid);
    assert_ne!(candidate.identifier, 0);
    assert!(!candidate.replies.contains(&candidate.identifier));
}

fn assert_reply_metadata_preserved(
    before: &[CommentSignature],
    after: &[CommentSignature],
    old: &str,
    new: &str,
) {
    let root = before
        .iter()
        .find(|signature| {
            signature
                .replies
                .iter()
                .filter(|identifier| {
                    before.iter().any(|reply| {
                        reply.identifier == **identifier && reply.text.as_deref() == Some(old)
                    })
                })
                .count()
                >= 2
        })
        .expect("root with duplicate reply text");
    let selected_identifier = root
        .replies
        .iter()
        .rev()
        .find(|identifier| {
            before
                .iter()
                .any(|reply| reply.identifier == **identifier && reply.text.as_deref() == Some(old))
        })
        .copied()
        .expect("selected duplicate reply");
    let source = before
        .iter()
        .find(|signature| signature.identifier == selected_identifier)
        .expect("selected reply signature");
    let uuid = source.uuid.expect("selected reply UUID");
    let candidate = after
        .iter()
        .find(|signature| signature.uuid == Some(uuid) && signature.text.as_deref() == Some(new))
        .expect("rewritten selected reply signature");
    assert_eq!(candidate.creation_date, source.creation_date);
    assert_eq!(candidate.author, source.author);
    assert_eq!(candidate.uuid, source.uuid);
    assert_ne!(candidate.identifier, 0);
    assert!(!candidate.replies.contains(&candidate.identifier));
}

fn assert_created_metadata(signatures: &[CommentSignature], text: &str) {
    let signature = signatures
        .iter()
        .find(|signature| signature.text.as_deref() == Some(text))
        .expect("created comment text");
    assert!(signature.creation_date.is_some());
    assert!(signature.author.is_some());
    assert!(signature.uuid.is_some());
    assert_ne!(signature.identifier, 0);
}

#[derive(Debug, Clone, Copy)]
enum HostileMutation {
    Cycle,
    MissingReply,
    MalformedWire,
}

fn with_hostile_comment(source: &[u8], mutation: HostileMutation) -> TestResult<Vec<u8>> {
    let mut comments = Vec::new();
    let mut referenced = BTreeSet::new();
    for entry in Catalog::from_bytes(source)?.iter() {
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
            let Some((index, message)) = object
                .messages
                .iter()
                .enumerate()
                .find(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            else {
                continue;
            };
            let decoded = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
            for reply in &decoded.replies {
                referenced.insert(reply.identifier);
            }
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("missing archive object identifier"))?;
            comments.push((entry.name().to_owned(), identifier, index));
        }
    }
    let (component, identifier, message_index) = comments
        .iter()
        .find(|(_, identifier, _)| !referenced.contains(identifier))
        .cloned()
        .or_else(|| comments.first().cloned())
        .ok_or_else(|| io::Error::other("fixture has no comment storage"))?;

    let mut replacement = None;
    for entry in Catalog::from_bytes(source)?.iter() {
        if entry.name() != component {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        let mut parsed = Archive::parse(&stream)?;
        let object = parsed
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("comment object disappeared"))?;
        let message = object
            .messages
            .get_mut(message_index)
            .ok_or_else(|| io::Error::other("comment payload disappeared"))?;
        let mut payload = message.data.clone();
        let referenced_identifier = match mutation {
            HostileMutation::Cycle => identifier,
            HostileMutation::MissingReply => u64::MAX - 19,
            HostileMutation::MalformedWire => 0,
        };
        if !matches!(mutation, HostileMutation::MalformedWire) {
            let reference = tsp::Reference {
                identifier: referenced_identifier,
                ..Default::default()
            }
            .encode_to_vec();
            append_length_delimited_field(&mut payload, 4, &reference)?;
        } else {
            payload = vec![0xff];
        }
        message.data = payload;
        if !matches!(mutation, HostileMutation::MalformedWire) {
            object.archive_info.message_infos[message_index]
                .object_references
                .push(referenced_identifier);
        }
        replacement = Some(SnappyStream::compress(&parsed.to_bytes()?)?);
        break;
    }
    let replacement = replacement.ok_or_else(|| io::Error::other("comment component missing"))?;
    let catalog = Catalog::from_bytes(source)?;
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
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}
