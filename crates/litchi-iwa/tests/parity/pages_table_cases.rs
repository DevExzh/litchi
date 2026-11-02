//! Pages body-table parity cases of wave W1-B (record 0371): cell values
//! through the leaf `table::values` family and the cell reads.

use std::collections::{BTreeMap, BTreeSet};
use std::hint::black_box;

use litchi_iwa::pages::{self, CellValue, TableCellUpdate};
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::comment::AuthorId;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_protos::{tsd, tsk, tsp, tst};
use litchi_pages::Package;
use litchi_pages::table::values::Value;
use prost::Message as _;

use crate::archive_diff::{self, MetadataPolicy};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, KnownFields, top_level_fields};
use crate::leaf_bytes;
use crate::text_glue::Workspace;

const CELLS: &str = "iwork/pages/body-table-cells-native.pages";
/// Native two-table source without formulas (the values family refuses
/// cells that take part in a formula graph, as `body-table-cells-native`
/// has).
const CATALOG: &str = "iwork/pages/body-table-catalog-native.pages";
pub(crate) const TABLE: usize = 0;
pub(crate) const COMMENT_ROW: usize = 1;
pub(crate) const COMMENT_COLUMN: usize = 1;
const AUTHOR_MESSAGE_TYPE: u32 = 212;

pub(crate) fn comment_position() -> TestResult<At> {
    Ok(At::new(
        u32::try_from(COMMENT_ROW)?,
        u32::try_from(COMMENT_COLUMN)?,
    ))
}

pub(crate) fn host_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = editor.table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
    let replies = if comment.is_some() {
        editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?
    } else {
        Vec::new()
    };
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.comment.text.as_str()),
        replies
            .iter()
            .map(|reply| reply.comment.text.as_str())
            .collect::<Vec<_>>()
    ))
}

pub(crate) fn leaf_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let comment = package.body_table_cell_comment(TABLE, position)?;
    let replies = comment
        .as_ref()
        .and_then(|comment| comment.replies())
        .unwrap_or(&[]);
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.text()),
        replies.iter().map(|reply| reply.text()).collect::<Vec<_>>()
    ))
}

fn host_populated_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = editor
        .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
        .ok_or("host read did not find the populated root comment")?;
    if comment.comment.text.is_empty()
        || comment.comment.creation_date_seconds.is_none()
        || comment.comment.author_id.is_none()
        || comment.comment.storage_uuid.is_none()
    {
        return Err("host root comment is missing text/date/author/storage metadata".into());
    }
    let replies = editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?;
    if replies.is_empty() {
        return Err("host read did not find a populated reply".into());
    }
    let reply_metadata = replies
        .iter()
        .map(|reply| {
            if reply.comment.text.is_empty()
                || reply.comment.creation_date_seconds.is_none()
                || reply.comment.author_id.is_none()
                || reply.comment.storage_uuid.is_none()
            {
                return Err::<String, Box<dyn std::error::Error>>(
                    "host reply is missing text/date/author/storage metadata".into(),
                );
            }
            Ok(format!(
                "{}|{:?}|{}|{}",
                reply.comment.text,
                reply.comment.creation_date_seconds,
                reply.comment.author_id.is_some(),
                reply.comment.storage_uuid.is_some()
            ))
        })
        .collect::<TestResult<Vec<_>>>()?;
    Ok(format!(
        "root={}|{:?}|{}|{};replies={reply_metadata:?}",
        comment.comment.text,
        comment.comment.creation_date_seconds,
        comment.comment.author_id.is_some(),
        comment.comment.storage_uuid.is_some()
    ))
}

fn leaf_populated_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let comment = package
        .body_table_cell_comment(TABLE, position)?
        .ok_or("leaf read did not find the populated root comment")?;
    if comment.text().is_empty() || comment.timestamp().is_none() || comment.author().is_none() {
        return Err("leaf root comment is missing text/date/author metadata".into());
    }
    let replies = comment.replies().unwrap_or(&[]);
    if replies.is_empty() {
        return Err("leaf read did not find a populated reply".into());
    }
    let reply_metadata = replies
        .iter()
        .map(|reply| {
            if reply.text().is_empty() || reply.timestamp().is_none() || reply.author().is_none() {
                return Err::<String, Box<dyn std::error::Error>>(
                    "leaf reply is missing text/date/author metadata".into(),
                );
            }
            Ok(format!(
                "{}|{:?}|{}",
                reply.text(),
                reply.timestamp().map(|value| value.as_f64()),
                reply.author().is_some()
            ))
        })
        .collect::<TestResult<Vec<_>>>()?;
    Ok(format!(
        "root={}|{:?}|{};replies={reply_metadata:?}",
        comment.text(),
        comment.timestamp().map(|value| value.as_f64()),
        comment.author().is_some()
    ))
}

fn author_metadata(
    bytes: &[u8],
    author_id: Option<AuthorId>,
) -> TestResult<(Option<String>, Option<String>)> {
    let Some(author_id) = author_id else {
        return Ok((None, None));
    };
    let located = Workspace::open(bytes)?.message(author_id.get(), AUTHOR_MESSAGE_TYPE)?;
    let author = tsk::AnnotationAuthorArchive::decode(located.payload.as_slice())?;
    Ok((author.name, author.public_id))
}

/// Compare both readers on the same host-authored source.  `Case::read`
/// intentionally invokes each reader on its own host/leaf output, so that
/// harness result is not enough to establish semantic reader parity.
fn assert_populated_comment_semantics(bytes: &[u8]) -> TestResult {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let host = editor
        .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
        .ok_or("host read did not find the populated root comment")?;
    let host_replies = editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?;
    if host_replies.is_empty() {
        return Err("host source must contain a non-empty reply list".into());
    }

    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let leaf = package
        .body_table_cell_comment(TABLE, position)?
        .ok_or("leaf read did not find the populated root comment")?;
    let leaf_replies = leaf.replies().unwrap_or(&[]);
    if leaf_replies.is_empty() {
        return Err("leaf source must contain a non-empty reply list".into());
    }

    assert_eq!(host.comment.text, leaf.text(), "root text");
    assert_eq!(
        host.comment.creation_date_seconds,
        leaf.timestamp().map(|value| value.as_f64()),
        "root timestamp"
    );
    assert_eq!(
        author_metadata(bytes, host.comment.author_id)?,
        {
            let author = leaf.author().ok_or("leaf root author is missing")?;
            (
                author.display_name().map(str::to_owned),
                author.public_id().map(str::to_owned),
            )
        },
        "root author metadata"
    );
    assert_eq!(host_replies.len(), leaf_replies.len(), "reply count");
    for (index, (host_reply, leaf_reply)) in
        host_replies.iter().zip(leaf_replies.iter()).enumerate()
    {
        assert_eq!(
            host_reply.comment.text,
            leaf_reply.text(),
            "reply {index} text"
        );
        assert_eq!(
            host_reply.comment.creation_date_seconds,
            leaf_reply.timestamp().map(|value| value.as_f64()),
            "reply {index} timestamp"
        );
        assert_eq!(
            author_metadata(bytes, host_reply.comment.author_id)?,
            {
                let author = leaf_reply
                    .author()
                    .ok_or_else(|| format!("leaf reply {index} author is missing"))?;
                (
                    author.display_name().map(str::to_owned),
                    author.public_id().map(str::to_owned),
                )
            },
            "reply {index} author metadata"
        );
    }
    Ok(())
}

pub(crate) fn comment_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &crate::harness::Fixture,
    input_path: &'static str,
    host: impl FnOnce(&mut pages::Draft, u64) -> TestResult + 'static,
    leaf: impl FnOnce(Package, At) -> TestResult<LeafRun> + 'static,
) -> TestResult {
    debug_assert_eq!(input.path, input_path);
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(move |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            leaf(package, comment_position()?)
        })
        .leaf_structural_metadata()
        .read("host table comment thread", host_comment_thread)
        .read("leaf table comment thread", leaf_comment_thread)
        .run()?;
    Ok(())
}

const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

fn is_pages_author_component(name: &str) -> bool {
    name.strip_prefix("Index/AnnotationAuthorStorage-")
        .and_then(|name| name.strip_suffix(".iwa"))
        .is_some_and(|identifier| {
            !identifier.is_empty() && identifier.bytes().all(|byte| byte.is_ascii_digit())
        })
}

// The fresh Pages-create comparison is intentionally narrow.  It binds the
// allowance to the selected cell's actual root and list owner, and never excludes that new root from
// payload comparison.
fn narrow_object_ids(bytes: &[u8]) -> TestResult<BTreeMap<String, BTreeSet<u64>>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut result = BTreeMap::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        result.insert(
            entry.name().to_owned(),
            archive
                .objects
                .into_iter()
                .filter_map(|object| object.archive_info.identifier)
                .collect(),
        );
    }
    Ok(result)
}

fn narrow_metadata_last(bytes: &[u8]) -> TestResult<u64> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut values = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for object in archive.objects {
            for message in object.messages {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    values.push(
                        tsp::PackageMetadata::decode(message.data.as_slice())?
                            .last_object_identifier,
                    );
                }
            }
        }
    }
    match values.as_slice() {
        [value] => Ok(*value),
        _ => Err(format!("expected one package watermark, found {values:?}").into()),
    }
}

fn narrow_metadata_payload(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut payload = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for object in archive.objects {
            for message in object.messages {
                if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                    continue;
                }
                if payload.replace(message.data).is_some() {
                    return Err("Pages package has multiple metadata payloads".into());
                }
            }
        }
    }
    payload.ok_or_else(|| "Pages package has no metadata payload".into())
}

fn narrow_selected_root(bytes: &[u8], leaf: bool, row: usize, column: usize) -> TestResult<u64> {
    if leaf {
        let comment = Package::from_bytes(bytes)?
            .body_table_cell_comment(TABLE, At::new(u32::try_from(row)?, u32::try_from(column)?))?
            .ok_or("Pages leaf create did not persist the selected root")?;
        if comment.text() != "W1-G create" {
            return Err("Pages leaf selected root text differs".into());
        }
    }
    // The legacy reader is the selected-cell oracle for both outputs.  It
    // returns the native storage identity attached to this exact cell, so a
    // similarly named comment elsewhere cannot satisfy the proof.
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = editor
        .table_cell_comment(model, row, column)?
        .ok_or("Pages selected-cell reader did not persist the root")?;
    if comment.comment.text != "W1-G create" {
        return Err("Pages selected-cell reader returned the wrong root text".into());
    }
    Ok(comment.storage_id.get())
}

fn narrow_list_owner(bytes: &[u8], root: u64) -> TestResult<(String, u64)> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut matches = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or("Pages table object has no identifier")?;
            for (index, info) in object.archive_info.message_infos.iter().enumerate() {
                if info.type_ != TABLE_DATA_LIST_MESSAGE_TYPE
                    || !info.object_references.contains(&root)
                {
                    continue;
                }
                let message = object
                    .messages
                    .get(index)
                    .ok_or("Pages table list index is out of range")?;
                let list = tst::TableDataList::decode(message.data.as_slice())?;
                let count = list
                    .entries
                    .iter()
                    .filter(|entry| {
                        entry
                            .comment_storage
                            .is_some_and(|reference| reference.identifier == root)
                    })
                    .count();
                if count != 1 {
                    return Err(format!("selected list has {count} entries for root {root}").into());
                }
                matches.push((entry.name().to_owned(), identifier));
            }
        }
    }
    match matches.as_slice() {
        [(member, object)] => Ok((member.clone(), *object)),
        _ => Err(format!("selected root has {matches:?} list owners").into()),
    }
}

fn narrow_validate_list_refs(
    host: &[u8],
    leaf: &[u8],
    host_root: u64,
    leaf_root: u64,
    selected_member: &str,
    selected_object: u64,
) -> TestResult<()> {
    let host_catalog = Catalog::from_bytes(host, Default::default())?;
    let leaf_catalog = Catalog::from_bytes(leaf, Default::default())?;
    let mut selected = 0;
    for host_entry in host_catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let leaf_entry = leaf_catalog
            .iter()
            .find(|entry| entry.name() == host_entry.name())
            .ok_or("Pages create leaf dropped a host IWA member")?;
        let host_stream = SnappyStream::decompress(host_entry.data(), Default::default())?;
        let leaf_stream = SnappyStream::decompress(leaf_entry.data(), Default::default())?;
        let host_archive = Archive::parse(host_stream.as_bytes(), Default::default())?;
        let leaf_archive = Archive::parse(leaf_stream.as_bytes(), Default::default())?;
        for host_object in &host_archive.objects {
            let identifier = host_object
                .archive_info
                .identifier
                .ok_or("Pages host object has no identifier")?;
            // The selected root is intentionally a fresh object, so its
            // source-relative identity differs between the two writers.  The
            // selected table-list owner remains stable; all other objects
            // must retain their identities exactly.
            let leaf_identifier = if host_entry.name() == selected_member && identifier == host_root
            {
                leaf_root
            } else {
                identifier
            };
            let Some(leaf_object) = leaf_archive
                .objects
                .iter()
                .find(|object| object.archive_info.identifier == Some(leaf_identifier))
            else {
                // The host's fresh generated-author archive is intentionally
                // absent from the authorless leaf output.  Admit only that
                // exact type-212 allocation here; all other missing objects
                // remain a hard locality failure.
                let known_generated_author = is_pages_author_component(host_entry.name())
                    && host_object.messages.len() == 1
                    && host_object.messages[0].type_ == 212
                    && {
                        let fields = litchi_iwa_common::wire::WireView::parse(
                            host_object.messages[0].data.as_slice(),
                            litchi_iwa_common::WireLimits::DEFAULT,
                        )?;
                        fields
                            .fields()
                            .all(|field| (1..=5).contains(&field.number()))
                    };
                if known_generated_author {
                    continue;
                }
                return Err(format!(
                    "Pages create changed an unrelated object identity: member {}, host id {}, mapped leaf id {}, selected member {}, selected object {}, host root {}, leaf root {}",
                    host_entry.name(),
                    identifier,
                    leaf_identifier,
                    selected_member,
                    selected_object,
                    host_root,
                    leaf_root,
                )
                .into());
            };
            if host_object.archive_info.message_infos.len()
                != leaf_object.archive_info.message_infos.len()
                || host_object.messages.len() != leaf_object.messages.len()
            {
                return Err("Pages create changed an unrelated object's message shape".into());
            }
            for (host_info, leaf_info) in host_object
                .archive_info
                .message_infos
                .iter()
                .zip(&leaf_object.archive_info.message_infos)
            {
                if host_info.type_ != leaf_info.type_ {
                    return Err("Pages create changed an unrelated message type".into());
                }
                if host_info.type_ != TABLE_DATA_LIST_MESSAGE_TYPE {
                    continue;
                }
                if host_info.data_references != leaf_info.data_references {
                    return Err("Pages create changed an unrelated 6005 data reference".into());
                }
                let host_selected = host_info.object_references.contains(&host_root);
                let leaf_selected = leaf_info.object_references.contains(&leaf_root);
                let normalise = |references: &[u64], selected_root| {
                    references
                        .iter()
                        .map(|reference| {
                            if *reference == selected_root {
                                0
                            } else {
                                *reference
                            }
                        })
                        .collect::<Vec<_>>()
                };
                if host_selected || leaf_selected {
                    let host_selected_count = host_info
                        .object_references
                        .iter()
                        .filter(|reference| **reference == host_root)
                        .count();
                    let leaf_selected_count = leaf_info
                        .object_references
                        .iter()
                        .filter(|reference| **reference == leaf_root)
                        .count();
                    if !host_selected
                        || !leaf_selected
                        || host_selected_count != 1
                        || leaf_selected_count != 1
                        || identifier != selected_object
                        || host_entry.name() != selected_member
                        || normalise(&host_info.object_references, host_root)
                            != normalise(&leaf_info.object_references, leaf_root)
                    {
                        return Err(
                            "Pages selected 6005 root reference is not the only changed reference"
                                .into(),
                        );
                    }
                    selected += 1;
                } else if host_info.object_references != leaf_info.object_references {
                    return Err("Pages create changed an unrelated 6005 reference".into());
                }
            }
        }
    }
    if selected != 1 {
        return Err(
            format!("Pages create expected one selected 6005 reference, found {selected}").into(),
        );
    }
    Ok(())
}

fn narrow_archive_object(
    bytes: &[u8],
    member_name: &str,
    identifier: u64,
) -> TestResult<litchi_iwa_archive::iwa::ArchiveObject> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| format!("missing IWA member {member_name}"))?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let archive = Archive::parse(stream.as_bytes(), Default::default())?;
    archive
        .objects
        .into_iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| format!("missing object {identifier} in {member_name}"))
        .map_err(Into::into)
}

fn narrow_validate_generated_author(data: &[u8]) -> TestResult<()> {
    let author = tsk::AnnotationAuthorArchive::decode(data)?;
    let expected = tsk::AnnotationAuthorArchive {
        name: Some("litchi-iwa".to_owned()),
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(0.368_627_46),
            g: Some(0.568_627_5),
            b: Some(0.937_254_9),
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
            a: Some(1.0),
            ..Default::default()
        }),
        public_id: None,
        is_public_author: Some(false),
        public_ids: Vec::new(),
    };
    if author != expected {
        return Err(format!("Pages generated author profile changed: {author:?}").into());
    }
    if data != expected.encode_to_vec().as_slice() {
        return Err("Pages generated author contains noncanonical or unknown bytes".into());
    }
    let fields =
        litchi_iwa_common::wire::WireView::parse(data, litchi_iwa_common::WireLimits::DEFAULT)?;
    if fields
        .fields()
        .any(|field| !(1..=5).contains(&field.number()))
    {
        return Err("Pages generated author contains unknown wire fields".into());
    }
    Ok(())
}

fn narrow_validate_author_storage(
    source: &[u8],
    host: &[u8],
    leaf: &[u8],
    member_name: &str,
    storage_id: u64,
    author_id: u64,
) -> TestResult<u64> {
    let source = narrow_archive_object(source, member_name, storage_id)?;
    let host = narrow_archive_object(host, member_name, storage_id)?;
    let leaf = narrow_archive_object(leaf, member_name, storage_id)?;
    let message = |object: &litchi_iwa_archive::iwa::ArchiveObject| -> TestResult<Vec<u8>> {
        if object.messages.len() != 1 || object.messages[0].type_ != 213 {
            return Err("Pages author-storage object is not exactly one type-213 message".into());
        }
        if object.archive_info.message_infos.len() != 1 {
            return Err("Pages author-storage object has an unexpected message-info count".into());
        }
        Ok(object.messages[0].data.clone())
    };
    let source_data = message(&source)?;
    let host_data = message(&host)?;
    let leaf_data = message(&leaf)?;
    if source_data != leaf_data || !source_data.is_empty() {
        return Err("Pages author-storage source/leaf changed the existing empty payload".into());
    }
    let reference = tsp::Reference {
        identifier: author_id,
        ..Default::default()
    }
    .encode_to_vec();
    let expected_host = litchi_iwa_common::wire::append_repeated_length_delimited_field(
        &source_data,
        1,
        &reference,
    )?;
    if host_data != expected_host {
        return Err(
            "Pages author-storage changed bytes outside the one new author reference".into(),
        );
    }
    let source_info = &source.archive_info.message_infos[0].object_references;
    let host_info = &host.archive_info.message_infos[0].object_references;
    let leaf_info = &leaf.archive_info.message_infos[0].object_references;
    if !source_info.is_empty() || !host_info.is_empty() || !leaf_info.is_empty() {
        return Err(format!(
            "Pages author-storage references are not source-relative: source {source_info:?}, host {host_info:?}, leaf {leaf_info:?}"
        )
        .into());
    }
    let storage = tsk::AnnotationAuthorStorageArchive::decode(host_data.as_slice())?;
    if storage.annotation_author.len() != 1 || storage.annotation_author[0].identifier != author_id
    {
        return Err("Pages author-storage did not add exactly the generated author".into());
    }
    Ok(source
        .archive_info
        .identifier
        .ok_or("author storage has no id")?)
}

fn narrow_single_message_object(
    bytes: &[u8],
    member_name: &str,
    message_type: u32,
) -> TestResult<u64> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| format!("missing IWA member {member_name}"))?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let matches = archive
        .objects
        .iter()
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == message_type)
        })
        .filter_map(|object| object.archive_info.identifier)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [identifier] => Ok(*identifier),
        _ => Err(format!(
            "expected one type-{message_type} object in {member_name}, found {matches:?}"
        )
        .into()),
    }
}

fn narrow_validate_selected_comment_author(
    bytes: &[u8],
    member_name: &str,
    root: u64,
    expected_author: Option<u64>,
) -> TestResult<Vec<u64>> {
    let object = narrow_archive_object(bytes, member_name, root)?;
    if object.messages.len() != 1 || object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE {
        return Err("Pages selected comment is not exactly one 3056 message".into());
    }
    if object.archive_info.message_infos.len() != 1 {
        return Err("Pages selected comment has an unexpected message-info count".into());
    }
    let comment = tsd::CommentStorageArchive::decode(object.messages[0].data.as_slice())?;
    if comment.text.as_deref() != Some("W1-G create") {
        return Err("Pages selected comment object does not carry the selected-cell text".into());
    }
    let known_fields = BTreeSet::from([1, 2, 3, 4, 5]);
    let unknown_fields = top_level_fields(&object.messages[0].data)
        .difference(&known_fields)
        .copied()
        .collect::<Vec<_>>();
    if !unknown_fields.is_empty() {
        return Err(format!(
            "Pages selected comment has unknown top-level fields {unknown_fields:?}"
        )
        .into());
    }
    let actual_author = comment
        .author
        .as_ref()
        .map(|reference| reference.identifier);
    if actual_author != expected_author {
        return Err(format!(
            "Pages selected comment author differs: expected {expected_author:?}, got {actual_author:?}"
        )
        .into());
    }
    Ok(object.archive_info.message_infos[0]
        .object_references
        .clone())
}

/// Existing comment-storage payloads, including neighboring roots and
/// replies, are source-owned bytes. The creation-only decoder below must
/// never mask a change to one of these objects, so compare every source
/// type-3056 payload byte-for-byte in both outputs before using the broader
/// archive diff.
fn narrow_validate_surviving_comment_storage(
    source: &[u8],
    host: &[u8],
    leaf: &[u8],
) -> TestResult<usize> {
    let source_catalog = Catalog::from_bytes(source, Default::default())?;
    let host_catalog = Catalog::from_bytes(host, Default::default())?;
    let leaf_catalog = Catalog::from_bytes(leaf, Default::default())?;
    let mut checked = 0;
    for source_entry in source_catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let host_entry = host_catalog
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or("Pages create host dropped a source IWA member")?;
        let leaf_entry = leaf_catalog
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or("Pages create leaf dropped a source IWA member")?;
        let source_stream = SnappyStream::decompress(source_entry.data(), Default::default())?;
        let host_stream = SnappyStream::decompress(host_entry.data(), Default::default())?;
        let leaf_stream = SnappyStream::decompress(leaf_entry.data(), Default::default())?;
        let source_archive = Archive::parse(source_stream.as_bytes(), Default::default())?;
        let host_archive = Archive::parse(host_stream.as_bytes(), Default::default())?;
        let leaf_archive = Archive::parse(leaf_stream.as_bytes(), Default::default())?;
        for source_object in &source_archive.objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or("Pages source comment object has no identifier")?;
            for (index, source_info) in source_object.archive_info.message_infos.iter().enumerate()
            {
                if source_info.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
                    continue;
                }
                let source_message = source_object
                    .messages
                    .get(index)
                    .ok_or("Pages source comment message-info index is out of range")?;
                let host_object = host_archive
                    .objects
                    .iter()
                    .find(|object| object.archive_info.identifier == Some(identifier))
                    .ok_or("Pages create host dropped a source comment object")?;
                let leaf_object = leaf_archive
                    .objects
                    .iter()
                    .find(|object| object.archive_info.identifier == Some(identifier))
                    .ok_or("Pages create leaf dropped a source comment object")?;
                for (label, output_object) in [("host", host_object), ("leaf", leaf_object)] {
                    let output_info = output_object
                        .archive_info
                        .message_infos
                        .get(index)
                        .ok_or_else(|| format!("Pages {label} dropped a source comment info"))?;
                    let output_message = output_object
                        .messages
                        .get(index)
                        .ok_or_else(|| format!("Pages {label} dropped a source comment message"))?;
                    if output_info.type_ != source_info.type_
                        || output_info.object_references != source_info.object_references
                        || output_info.data_references != source_info.data_references
                        || output_message.type_ != source_message.type_
                        || output_message.data != source_message.data
                    {
                        return Err(format!(
                            "Pages {label} changed source-owned type-3056 payload {identifier} in {}",
                            source_entry.name()
                        )
                        .into());
                    }
                }
                checked += 1;
            }
        }
    }
    if checked == 0 {
        return Err("Pages create source has no surviving type-3056 comment payload".into());
    }
    Ok(checked)
}

fn narrow_first_comment_storage(source: &[u8]) -> TestResult<(String, u64)> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for object in archive.objects {
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                return Ok((
                    entry.name().to_owned(),
                    object
                        .archive_info
                        .identifier
                        .ok_or("Pages source comment object has no identifier")?,
                ));
            }
        }
    }
    Err("Pages source has no type-3056 comment object".into())
}

fn narrow_corrupt_existing_comment(
    bytes: &[u8],
    member_name: &str,
    identifier: u64,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or("Pages surviving comment member disappeared for negative control")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let mut changed = 0;
    for object in &mut archive.objects {
        if object.archive_info.identifier != Some(identifier) {
            continue;
        }
        for message in &mut object.messages {
            if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE {
                message.data = litchi_iwa_common::wire::patch_length_delimited_field(
                    &message.data,
                    1,
                    true,
                    Some(b"W1-G surviving-comment negative-control"),
                )?;
                changed += 1;
            }
        }
    }
    if changed != 1 {
        return Err(
            format!("Pages surviving-comment negative control patched {changed} payloads").into(),
        );
    }
    let replacement =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member_name, &replacement)],
        Default::default(),
    )?)
}

fn narrow_corrupt_metadata_watermark(bytes: &[u8], replacement: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut edits = Vec::new();
    let mut changed = 0;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
        let mut archive_changed = false;
        for object in &mut archive.objects {
            for message in &mut object.messages {
                if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                    continue;
                }
                message.data = litchi_iwa_common::wire::patch_varint_field(
                    &message.data,
                    1,
                    true,
                    Some(replacement),
                )?;
                changed += 1;
                archive_changed = true;
            }
        }
        if archive_changed {
            edits.push((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?,
            ));
        }
    }
    if changed != 1 {
        return Err(format!(
            "Pages metadata watermark negative control patched {changed} messages"
        )
        .into());
    }
    let entry_edits = edits
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&entry_edits, Default::default())?)
}

fn narrow_add_extra_author_object(
    bytes: &[u8],
    member_name: &str,
    source_author_id: u64,
    extra_author_id: u64,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or("Pages author component disappeared for extra-allocation control")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let source = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(source_author_id))
        .ok_or("Pages generated author disappeared for extra-allocation control")?;
    let extra = source.clone_with_identity_remap(
        extra_author_id,
        &[],
        &source.messages,
        Default::default(),
    )?;
    archive.objects.push(extra);
    let replacement =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member_name, &replacement)],
        Default::default(),
    )?)
}

fn narrow_corrupt_author_info_reference(
    bytes: &[u8],
    member_name: &str,
    author_id: u64,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or("Pages author component disappeared for reference-shape control")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let mut changed = 0;
    for object in &mut archive.objects {
        if object.archive_info.identifier != Some(author_id) {
            continue;
        }
        for info in &mut object.archive_info.message_infos {
            if info.type_ == AUTHOR_MESSAGE_TYPE {
                info.object_references.push(u64::MAX);
                changed += 1;
            }
        }
    }
    if changed != 1 {
        return Err(format!(
            "Pages author reference-shape control changed {changed} message infos"
        )
        .into());
    }
    let replacement =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member_name, &replacement)],
        Default::default(),
    )?)
}

fn narrow_pages_create_policy(
    source: &[u8],
    host: &[u8],
    leaf: &[u8],
    host_root: u64,
    leaf_root: u64,
) -> TestResult<(String, u64, String, u64, BTreeSet<u64>, u64)> {
    let source_ids = narrow_object_ids(source)?;
    let host_ids = narrow_object_ids(host)?;
    let leaf_ids = narrow_object_ids(leaf)?;
    let source_all = source_ids
        .values()
        .flat_map(|ids| ids.iter())
        .copied()
        .collect::<BTreeSet<_>>();
    if source_all.contains(&host_root) || source_all.contains(&leaf_root) {
        return Err("Pages selected roots are not fresh relative to the whole source".into());
    }
    if host_root == leaf_root {
        return Err("Pages selected roots are not two fresh identities".into());
    }
    let (selected_member, selected_object) = narrow_list_owner(host, host_root)?;
    if narrow_list_owner(leaf, leaf_root)? != (selected_member.clone(), selected_object) {
        return Err("Pages selected list owner changed identity".into());
    }
    let source_selected_ids = source_ids
        .get(&selected_member)
        .cloned()
        .unwrap_or_default();
    if source_selected_ids.contains(&host_root) || source_selected_ids.contains(&leaf_root) {
        return Err("Pages selected root was already present in the source cell component".into());
    }
    narrow_validate_list_refs(
        host,
        leaf,
        host_root,
        leaf_root,
        &selected_member,
        selected_object,
    )?;

    let mut host_only = Vec::new();
    let mut leaf_fresh = Vec::new();
    for (member, ids) in &host_ids {
        let source_member = source_ids.get(member).cloned().unwrap_or_default();
        for identifier in ids.difference(&source_member) {
            if !(member == &selected_member && *identifier == host_root) {
                host_only.push((member.clone(), *identifier));
            }
        }
    }
    for (member, ids) in &leaf_ids {
        let source_member = source_ids.get(member).cloned().unwrap_or_default();
        for identifier in ids.difference(&source_member) {
            if !(member == &selected_member && *identifier == leaf_root) {
                leaf_fresh.push((member.clone(), *identifier));
            }
        }
    }
    if !leaf_fresh.is_empty() {
        return Err(
            format!("Pages create has unexpected leaf fresh objects {leaf_fresh:?}").into(),
        );
    }
    if host_only.len() != 1 {
        return Err(format!(
            "Pages create expected one host-only author object, found {host_only:?}"
        )
        .into());
    }
    let (author_member, author_id) = &host_only[0];
    let source_author_members = source_ids
        .keys()
        .filter(|member| is_pages_author_component(member))
        .cloned()
        .collect::<Vec<_>>();
    let [source_author_member] = source_author_members.as_slice() else {
        return Err(format!(
            "Pages source has an ambiguous author component: {source_author_members:?}"
        )
        .into());
    };
    if author_member != source_author_member
        || author_member == &selected_member
        || !host_ids.contains_key(author_member)
        || !leaf_ids.contains_key(author_member)
    {
        return Err("Pages host-only object is not in the author-storage component".into());
    }
    if source_ids
        .iter()
        .any(|(member, ids)| member != author_member && ids.contains(author_id))
        || host_ids
            .iter()
            .any(|(member, ids)| member != author_member && ids.contains(author_id))
    {
        return Err(
            "Pages generated author identifier collides with an unrelated component object".into(),
        );
    }
    let host_catalog = Catalog::from_bytes(host, Default::default())?;
    let author_entry = host_catalog
        .iter()
        .find(|entry| entry.name() == author_member)
        .ok_or("Pages author component disappeared")?;
    let stream = SnappyStream::decompress(author_entry.data(), Default::default())?;
    let archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let author_object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(*author_id))
        .ok_or("Pages host-only author object disappeared")?;
    if author_object.messages.len() != 1 || author_object.messages[0].type_ != 212 {
        return Err("Pages host-only allocation is not exactly one author object".into());
    }
    if author_object.archive_info.message_infos.len() != 1 {
        return Err("Pages host-only author object has an unexpected message-info count".into());
    }
    let author_info = &author_object.archive_info.message_infos[0];
    if author_info.type_ != 212
        || !author_info.object_references.is_empty()
        || !author_info.data_references.is_empty()
    {
        return Err(
            "Pages host-only author object has an unexpected archive reference shape".into(),
        );
    }
    narrow_validate_generated_author(author_object.messages[0].data.as_slice())?;
    let author_storage_object = narrow_single_message_object(source, author_member, 213)?;
    narrow_validate_author_storage(
        source,
        host,
        leaf,
        author_member,
        author_storage_object,
        *author_id,
    )?;
    let host_author_refs = narrow_validate_selected_comment_author(
        host,
        &selected_member,
        host_root,
        Some(*author_id),
    )?;
    let leaf_author_refs =
        narrow_validate_selected_comment_author(leaf, &selected_member, leaf_root, None)?;
    if host_author_refs != [*author_id] || !leaf_author_refs.is_empty() {
        return Err(format!(
            "Pages selected comment archive references differ: host {host_author_refs:?}, leaf {leaf_author_refs:?}"
        )
        .into());
    }
    let source_last = narrow_metadata_last(source)?;
    let host_last = narrow_metadata_last(host)?;
    let leaf_last = narrow_metadata_last(leaf)?;
    let expected_host_last = source_last
        .checked_add(u64::try_from(host_only.len() + 1)?)
        .ok_or("Pages host metadata watermark overflow")?;
    let expected_leaf_last = source_last
        .checked_add(1)
        .ok_or("Pages leaf metadata watermark overflow")?;
    if host_last != expected_host_last || leaf_last != expected_leaf_last {
        return Err(format!(
            "Pages metadata watermarks are not allocation-derived: source {source_last}, expected host {expected_host_last}, host {host_last}, expected leaf {expected_leaf_last}, leaf {leaf_last}"
        )
        .into());
    }
    Ok((
        selected_member,
        selected_object,
        author_member.clone(),
        author_storage_object,
        BTreeSet::from([*author_id]),
        host_last,
    ))
}

fn narrow_patch_reference(
    data: &[u8],
    field: u32,
    from: u64,
    to: u64,
) -> Result<Vec<u8>, litchi_iwa_common::Error> {
    let values = litchi_iwa_common::wire::repeated_varint_values(data, field)?;
    if values.len() == 1 && values[0] == from {
        litchi_iwa_common::wire::patch_varint_field(data, field, true, Some(to))
    } else {
        Ok(data.to_vec())
    }
}

/// Patch exactly one selected `comment_storage` reference while retaining all
/// other wire fields and their original encodings.
fn narrow_remap_selected_list(data: &[u8], from: u64, to: u64) -> TestResult<Vec<u8>> {
    let mut matches = 0;
    let output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        data,
        &[3, 10],
        |reference| -> Result<Vec<u8>, litchi_iwa_common::Error> {
            let values = litchi_iwa_common::wire::repeated_varint_values(reference, 1)?;
            if values.len() == 1 && values[0] == from {
                matches += 1;
                litchi_iwa_common::wire::patch_varint_field(reference, 1, true, Some(to))
            } else {
                Ok(reference.to_vec())
            }
        },
    )?;
    if matches != 1 {
        return Err(format!("selected list reference matched {matches} entries").into());
    }
    Ok(output)
}

fn narrow_corrupt_unrelated_list_reference(
    data: &[u8],
    selected_root: u64,
) -> TestResult<Option<Vec<u8>>> {
    let mut matches = 0;
    let output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        data,
        &[3, 10],
        |reference| -> Result<Vec<u8>, litchi_iwa_common::Error> {
            let values = litchi_iwa_common::wire::repeated_varint_values(reference, 1)?;
            if values.len() == 1 && values[0] != selected_root {
                matches += 1;
                litchi_iwa_common::wire::patch_varint_field(reference, 1, true, Some(u64::MAX))
            } else {
                Ok(reference.to_vec())
            }
        },
    )?;
    if matches == 0 {
        return Ok(None);
    }
    if matches != 1 {
        return Err(format!("unrelated 6005 negative control changed {matches} references").into());
    }
    Ok(Some(output))
}

fn narrow_corrupt_selected_list(
    bytes: &[u8],
    member_name: &str,
    selected_object: u64,
    selected_root: u64,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or("Pages selected list member disappeared for negative control")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let mut changed = 0;
    for object in &mut archive.objects {
        if object.archive_info.identifier != Some(selected_object) {
            continue;
        }
        let mut changed_payload = false;
        for message in &mut object.messages {
            if message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE
                && let Some(data) =
                    narrow_corrupt_unrelated_list_reference(&message.data, selected_root)?
            {
                message.data = data;
                changed_payload = true;
            }
        }
        if !changed_payload {
            let mut changed_info = 0;
            for info in &mut object.archive_info.message_infos {
                if info.type_ == TABLE_DATA_LIST_MESSAGE_TYPE {
                    info.object_references.push(u64::MAX);
                    changed_info += 1;
                }
            }
            if changed_info != 1 {
                return Err(format!(
                    "Pages unrelated 6005 negative control found {changed_info} list infos"
                )
                .into());
            }
        }
        changed += 1;
    }
    if changed != 1 {
        return Err(format!(
            "Pages unrelated 6005 negative control patched {changed} list payloads"
        )
        .into());
    }
    let replacement =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member_name, &replacement)],
        Default::default(),
    )?)
}

fn narrow_selected_list_payload(
    bytes: &[u8],
    selected_member: &str,
    selected_object: u64,
) -> TestResult<Vec<u8>> {
    let object = narrow_archive_object(bytes, selected_member, selected_object)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    match messages.as_slice() {
        [message] => Ok(message.data.clone()),
        _ => Err(format!(
            "Pages selected list owner has {} table-list payloads",
            messages.len()
        )
        .into()),
    }
}

fn narrow_remap_metadata(data: &[u8], from: u64, to: u64, host_last: u64) -> TestResult<Vec<u8>> {
    let mut output = litchi_iwa_common::wire::patch_varint_field(data, 1, true, Some(host_last))?;
    for path in [[3, 6], [3, 18], [11, 6], [11, 18]] {
        output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
            &output,
            &path,
            |reference| narrow_patch_reference(reference, 2, from, to),
        )?;
    }
    for path in [[3, 7, 2], [11, 7, 2]] {
        output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
            &output,
            &path,
            |reference| narrow_patch_reference(reference, 1, from, to),
        )?;
    }
    for path in [[3, 11], [11, 11]] {
        output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
            &output,
            &path,
            |reference| narrow_patch_reference(reference, 1, from, to),
        )?;
    }
    output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        &output,
        &[10],
        |reference| narrow_patch_reference(reference, 1, from, to),
    )?;
    Ok(output)
}

fn narrow_remap_leaf(
    leaf: &[u8],
    selected_member: &str,
    selected_object: u64,
    author_member: &str,
    author_storage_object: u64,
    host_author_id: u64,
    leaf_root: u64,
    host_root: u64,
    host_last: u64,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(leaf, Default::default())?;
    let remap = [(leaf_root, host_root)];
    let mut edits = Vec::new();
    let mut selected_payloads = 0;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let selected_component = entry.name() == selected_member;
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
        let mut archive_changed = false;
        for object in &mut archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or("Pages leaf object has no identifier")?;
            let new_identifier = if selected_component && identifier == leaf_root {
                host_root
            } else {
                identifier
            };
            let mut messages = object.messages.clone();
            let mut object_changed = false;
            for (index, message) in messages.iter_mut().enumerate() {
                if selected_component
                    && identifier == selected_object
                    && message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE
                {
                    message.data = narrow_remap_selected_list(&message.data, leaf_root, host_root)?;
                    selected_payloads += 1;
                    object_changed = true;
                } else if selected_component
                    && identifier == leaf_root
                    && message.type_ == COMMENT_STORAGE_MESSAGE_TYPE
                {
                    if !litchi_iwa_common::wire::repeated_length_delimited_payloads(
                        &message.data,
                        3,
                    )?
                    .is_empty()
                    {
                        return Err(
                            "Pages leaf selected comment already contains an author payload".into(),
                        );
                    }
                    let reference = tsp::Reference {
                        identifier: host_author_id,
                        ..Default::default()
                    }
                    .encode_to_vec();
                    litchi_iwa_common::wire::append_length_delimited_field(
                        &mut message.data,
                        3,
                        &reference,
                        litchi_iwa_common::WireLimits::DEFAULT,
                    )?;
                    object_changed = true;
                } else if entry.name() == author_member
                    && identifier == author_storage_object
                    && message.type_ == 213
                {
                    let reference = tsp::Reference {
                        identifier: host_author_id,
                        ..Default::default()
                    }
                    .encode_to_vec();
                    message.data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
                        &message.data,
                        1,
                        &reference,
                    )?;
                    object_changed = true;
                } else if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    message.data =
                        narrow_remap_metadata(&message.data, leaf_root, host_root, host_last)?;
                    object_changed |= message.data != object.messages[index].data;
                }
            }
            let has_reference_remap = selected_component
                && object
                    .archive_info
                    .message_infos
                    .iter()
                    .flat_map(|message| &message.object_references)
                    .any(|reference| *reference == leaf_root);
            if new_identifier != identifier || object_changed || has_reference_remap {
                let remap_entries: &[(u64, u64)] = if selected_component { &remap } else { &[] };
                let mut replacement = object.clone_with_identity_remap(
                    new_identifier,
                    remap_entries,
                    &messages,
                    Default::default(),
                )?;
                if selected_component && identifier == leaf_root {
                    let mut selected_comment_infos = 0;
                    for info in &mut replacement.archive_info.message_infos {
                        if info.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
                            continue;
                        }
                        if !info.object_references.is_empty() {
                            return Err(
                                "Pages normalized selected comment already has an author reference"
                                    .into(),
                            );
                        }
                        info.object_references.push(host_author_id);
                        selected_comment_infos += 1;
                    }
                    if selected_comment_infos != 1 {
                        return Err(format!(
                            "Pages normalized selected comment has {selected_comment_infos} message-info entries"
                        )
                        .into());
                    }
                }
                *object = replacement;
                archive_changed = true;
            }
        }
        if archive_changed {
            edits.push((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?,
            ));
        }
    }
    if selected_payloads != 1 {
        return Err(format!("Pages remapped {selected_payloads} selected list payloads").into());
    }
    let entry_edits = edits
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&entry_edits, Default::default())?)
}

fn narrow_corrupt_selected_comment(
    bytes: &[u8],
    selected_member: &str,
    root: u64,
    append_unknown: bool,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == selected_member)
        .ok_or("Pages selected comment member disappeared")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let mut changed = 0;
    for object in &mut archive.objects {
        if object.archive_info.identifier != Some(root) {
            continue;
        }
        for message in &mut object.messages {
            if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
                continue;
            }
            if append_unknown {
                litchi_iwa_common::wire::append_length_delimited_field(
                    &mut message.data,
                    99,
                    b"W1-G selected-root unknown-field negative-control",
                    litchi_iwa_common::WireLimits::DEFAULT,
                )?;
            } else {
                message.data = litchi_iwa_common::wire::patch_length_delimited_field(
                    &message.data,
                    1,
                    true,
                    Some(b"W1-G selected-root negative-control"),
                )?;
            }
            changed += 1;
        }
    }
    if changed != 1 {
        return Err(format!(
            "Pages selected-root negative control patched {changed} comment payloads"
        )
        .into());
    }
    let replacement =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(selected_member, &replacement)],
        Default::default(),
    )?)
}

fn narrow_corrupt_author_storage(
    bytes: &[u8],
    author_member: &str,
    storage_object: u64,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == author_member)
        .ok_or("Pages author component disappeared for negative control")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    let mut changed = 0;
    for object in &mut archive.objects {
        if object.archive_info.identifier != Some(storage_object) {
            continue;
        }
        for message in &mut object.messages {
            if message.type_ != 213 {
                continue;
            }
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut message.data,
                99,
                b"W1-G author-storage unknown negative-control",
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            changed += 1;
        }
    }
    if changed != 1 {
        return Err(
            format!("Pages author-storage negative control patched {changed} payloads").into(),
        );
    }
    let replacement =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(author_member, &replacement)],
        Default::default(),
    )?)
}

fn raw_comment_fields<'a>(
    data: &'a [u8],
) -> Result<Vec<litchi_iwa_common::wire::RawWireField<'a>>, String> {
    let mut parser = litchi_iwa_common::wire::RawWireFields::new(
        data,
        litchi_iwa_common::wire::RawWireLimits::default(),
    );
    let mut fields = Vec::new();
    while let Some(field) = parser.next().map_err(|error| error.to_string())? {
        fields.push(field);
    }
    Ok(fields)
}

fn validate_comment_reference_wire(data: &[u8]) -> Result<(), String> {
    let fields = raw_comment_fields(data)?;
    let mut seen = BTreeSet::new();
    for field in fields {
        if !field.key_is_canonical() || !field.value_is_canonical() {
            return Err("comment reference is not canonically encoded".to_owned());
        }
        if !matches!(field.number(), 1..=3) || field.wire_type() != 0 {
            return Err("comment reference has an unknown or wrong-type field".to_owned());
        }
        if !seen.insert(field.number()) {
            return Err("comment reference has a duplicate field".to_owned());
        }
    }
    if !seen.contains(&1) {
        return Err("comment reference has no identifier".to_owned());
    }
    Ok(())
}

fn validate_comment_date_wire(data: &[u8]) -> Result<(), String> {
    let fields = raw_comment_fields(data)?;
    if fields.len() != 1 {
        return Err("comment date has unknown or duplicate fields".to_owned());
    }
    let field = fields[0];
    if field.number() != 1
        || field.wire_type() != 1
        || !field.key_is_canonical()
        || !field.value_is_canonical()
    {
        return Err("comment date has an unknown or noncanonical field".to_owned());
    }
    Ok(())
}

fn validate_comment_uuid_wire(data: &[u8]) -> Result<(), String> {
    let fields = raw_comment_fields(data)?;
    let mut seen = BTreeSet::new();
    for field in fields {
        if !field.key_is_canonical() || !field.value_is_canonical() {
            return Err("comment UUID is not canonically encoded".to_owned());
        }
        if !matches!(field.number(), 1..=2) || field.wire_type() != 0 {
            return Err("comment UUID has an unknown or wrong-type field".to_owned());
        }
        if !seen.insert(field.number()) {
            return Err("comment UUID has a duplicate field".to_owned());
        }
    }
    if seen != BTreeSet::from([1, 2]) {
        return Err("comment UUID is missing a required field".to_owned());
    }
    Ok(())
}

fn validate_new_comment_storage_wire(data: &[u8]) -> Result<(), String> {
    let fields = raw_comment_fields(data)?;
    let mut seen = BTreeSet::new();
    for field in fields {
        if !field.key_is_canonical() || !field.length_is_canonical() {
            return Err("new comment has noncanonical wire framing".to_owned());
        }
        match field.number() {
            1 => {
                if field.wire_type() != 2 || !seen.insert(1) {
                    return Err("new comment has an invalid or duplicate text field".to_owned());
                }
            },
            2 => {
                if field.wire_type() != 2 || !seen.insert(2) {
                    return Err("new comment has an invalid or duplicate date field".to_owned());
                }
                validate_comment_date_wire(field.payload())?;
            },
            3 => {
                if field.wire_type() != 2 || !seen.insert(3) {
                    return Err("new comment has an invalid or duplicate author field".to_owned());
                }
                validate_comment_reference_wire(field.payload())?;
            },
            4 => {
                if field.wire_type() != 2 {
                    return Err("new comment has an invalid replies field".to_owned());
                }
                validate_comment_reference_wire(field.payload())?;
            },
            5 => {
                if field.wire_type() != 2 || !seen.insert(5) {
                    return Err("new comment has an invalid or duplicate UUID field".to_owned());
                }
                validate_comment_uuid_wire(field.payload())?;
            },
            number => return Err(format!("new comment has unknown field {number}")),
        }
    }
    if !seen.contains(&2) || !seen.contains(&5) {
        return Err("new comment is missing creation date or storage UUID".to_owned());
    }
    Ok(())
}

#[allow(deprecated)]
pub(crate) fn masked_new_pages_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    validate_new_comment_storage_wire(data).map_err(prost::DecodeError::new)?;
    // Patch only the two creation-time fields in place. This keeps the
    // selected payload's field order and every untouched byte intact while
    // rejecting unknown or duplicate fields (including nested date, UUID and
    // reference fields) before normalization.
    let date = tsp::Date { seconds: 0.0 }.encode_to_vec();
    let uuid = tsp::Uuid { lower: 1, upper: 1 }.encode_to_vec();
    let normalized =
        litchi_iwa_common::wire::patch_length_delimited_field(data, 2, true, Some(&date))
            .and_then(|data| {
                litchi_iwa_common::wire::patch_length_delimited_field(&data, 5, true, Some(&uuid))
            })
            .map_err(|error| prost::DecodeError::new(error.to_string()))?;
    let mut decoded = crate::known_fields::decode_as::<tsd::CommentStorageArchive>(&normalized)?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn narrow_append_raw_at_path(
    data: &[u8],
    path: &[u32],
    bytes: &[u8],
) -> TestResult<(Vec<u8>, usize)> {
    let mut matches = 0;
    let output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        data,
        path,
        |nested| -> Result<Vec<u8>, litchi_iwa_common::Error> {
            matches += 1;
            let mut output = nested.to_vec();
            output.extend_from_slice(bytes);
            Ok(output)
        },
    )?;
    Ok((output, matches))
}

fn narrow_append_nested_comment_bytes(
    data: &[u8],
    field: u32,
    bytes: &[u8],
) -> TestResult<Vec<u8>> {
    let (output, matches) = narrow_append_raw_at_path(data, &[field], bytes)?;
    if matches == 0 {
        return Err(format!("comment nested field {field} was not present").into());
    }
    Ok(output)
}

pub(crate) fn host_model(editor: &pages::Draft) -> TestResult<u64> {
    Ok(editor
        .tables()?
        .get(TABLE)
        .ok_or("body table missing")?
        .model_object_id)
}

fn host_cells(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let table = editor.table(host_model(&editor)?)?;
    let mut cells = table
        .iter_cells()
        .map(|((row, column), value)| format!("{row},{column}: {value:?}"))
        .collect::<Vec<_>>();
    cells.push(format!("count {}", table.cell_count()));
    Ok(cells)
}

fn leaf_cells(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?.body_table_cells(TABLE)?
    ))
}

fn host_value(value: &Value) -> TestResult<CellValue> {
    Ok(match value {
        Value::Text(text) => CellValue::Text(text.clone()),
        Value::Number(number) => CellValue::number(*number)?,
        Value::Boolean(flag) => CellValue::Boolean(*flag),
        Value::Empty => CellValue::Empty,
        _ => return Err("unmapped value".into()),
    })
}

fn values_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    updates: Vec<(u32, u32, Value)>,
) -> TestResult {
    let input = fixture(CATALOG)?;
    let updates = &updates;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            if let [(row, column, value)] = updates.as_slice() {
                editor.set_table_cell(
                    model,
                    *row as usize,
                    *column as usize,
                    host_value(value)?,
                )?;
            } else {
                let batch = updates
                    .iter()
                    .map(|(row, column, value)| {
                        Ok(TableCellUpdate::new(
                            *row as usize,
                            *column as usize,
                            host_value(value)?,
                        ))
                    })
                    .collect::<TestResult<Vec<_>>>()?;
                editor.set_table_cells(model, batch)?;
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            for (row, column, value) in updates {
                edit.set(*row, *column, value.clone())?;
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_table_values(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table cells", host_cells)
        .read("leaf body_table_cells", leaf_cells)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_cell_parity() -> TestResult {
    values_case(
        "pages_body_table_cell_parity",
        Tier::A,
        Evidence::Queued,
        vec![(1, 3, Value::Text("Zürich".to_owned()))],
    )
}

#[test]
fn pages_body_table_cells_batch_parity() -> TestResult {
    values_case(
        "pages_body_table_cells_batch_parity",
        Tier::A,
        Evidence::Queued,
        vec![
            (2, 3, Value::Number(42.25)),
            (1, 2, Value::Boolean(false)),
            (2, 0, Value::Empty),
        ],
    )
}

#[test]
fn pages_body_table_cells_read_parity() -> TestResult {
    let input = fixture(CELLS)?;
    Case::new(
        "pages_body_table_cells_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-cells-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host table cells", host_cells)
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    // The two reads agree value for value.
    let host = pages::Draft::from_bytes(&input.bytes)?;
    let host_table = host.table(host_model(&host)?)?;
    let leaf = Package::from_bytes(&input.bytes)?
        .body_table_cells(TABLE)?
        .into_table();
    for ((row, column), value) in host_table.iter_cells() {
        let at = litchi_iwa_common::table::coordinate::At::new(
            u32::try_from(row)?,
            u32::try_from(column)?,
        );
        assert_eq!(leaf.get(at), Some(value), "{row},{column}");
    }
    Ok(())
}

pub(crate) const COMMENTS: &str = "iwork/pages/body-table-comment-hidden-native-saved.pages";

/// The table list, merges and cell comments read the same through the host
/// (`tables`, `Table::{merges, get_comment, iter_comments, comment_count}`,
/// `table_cell_comment[_replies]`) and the leaf (`body_tables`,
/// `body_table_merges`, `body_table_cells`) on native fixtures.
#[test]
fn pages_body_table_reads_parity() -> TestResult {
    let input = fixture(CELLS)?;
    Case::new(
        "pages_body_table_reads_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-cells-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    for path in [
        COMMENTS,
        CELLS,
        CATALOG,
        "iwork/pages/body-table-merges-native.pages",
    ] {
        let input = fixture(path)?;
        let host = pages::Draft::from_bytes(&input.bytes)?;
        let package = Package::from_bytes(&input.bytes)?;
        let catalog = package.body_tables()?;
        let tables = host.tables()?;
        assert_eq!(tables.len(), catalog.len(), "{path}");
        for (index, info) in tables.iter().enumerate() {
            let entry = catalog.get(index).ok_or("catalog entry")?;
            assert_eq!(entry.name(), info.name, "{path}");
            let host_table = host.table(info.model_object_id)?;
            let leaf_read = package.body_table_cells(index)?;
            assert_eq!(
                host_table.comment_count(),
                leaf_read.comment_count(),
                "{path}"
            );
            for ((row, column), comment) in host_table.iter_comments() {
                let at = litchi_iwa_common::table::coordinate::At::new(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                );
                assert_eq!(leaf_read.get_comment(at), Some(comment), "{path}");
                assert_eq!(host_table.get_comment(row, column), Some(comment));
                let native = host
                    .table_cell_comment(info.model_object_id, row, column)?
                    .ok_or("host cell comment")?;
                assert_eq!(native.comment.text, comment.text(), "{path}");
                assert_eq!(
                    host.table_cell_comment_replies(info.model_object_id, row, column)?
                        .len(),
                    comment.reply_count(),
                    "{path}"
                );
            }
            assert_eq!(
                host_table.merges(),
                package.body_table_merges(index)?.as_slice(),
                "{path}"
            );
        }
    }
    Ok(())
}

fn pages_body_table_comment_read_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    let input = fixture(COMMENTS)?;
    assert_populated_comment_semantics(&input.bytes)?;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .read_only()
        .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read(
            "host populated table comment thread",
            host_populated_comment_thread,
        )
        .read(
            "leaf populated table comment thread",
            leaf_populated_comment_thread,
        )
        .run()?;
    Ok(())
}

#[test]
#[rustfmt::skip]
fn pages_body_table_comment_read_parity() -> TestResult {
    let evidence = Evidence::Native("test-data/iwork/pages/body-table-comment-hidden-native-saved-receipt.json");
    pages_body_table_comment_read_case("pages_body_table_comment_read_parity", Tier::A, evidence)
}

#[test]
#[rustfmt::skip]
fn pages_body_table_comment_replies_read_parity() -> TestResult {
    let evidence = Evidence::Native("test-data/iwork/pages/body-table-comment-hidden-native-saved-receipt.json");
    pages_body_table_comment_read_case(
        "pages_body_table_comment_replies_read_parity",
        Tier::A,
        evidence,
    )
}

#[test]
fn pages_body_table_comment_set_parity() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G original")?;
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-for-set",
        bytes: authored.to_bytes()?,
    };
    comment_case(
        "pages_body_table_comment_set_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        "generated:pages-table-comment-root-for-set",
        |editor, model| {
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G root")?;
            Ok(())
        },
        |package, position| {
            let commit = package
                .edit_body_table_cell_comment(TABLE, position)?
                .set("W1-G root")?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_table_cell_comment(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

#[test]
fn pages_body_table_comment_create_parity() -> TestResult {
    let input = fixture(CELLS)?;
    const CREATE_ROW: usize = COMMENT_ROW;
    const CREATE_COLUMN: usize = COMMENT_COLUMN;
    let editor = pages::Draft::from_bytes(&input.bytes)?;
    let model = host_model(&editor)?;
    if editor
        .table(model)?
        .get_cell(CREATE_ROW, CREATE_COLUMN)
        .is_none()
    {
        return Err("Pages create target is not a stored table cell".into());
    }
    if editor
        .table_cell_comment(model, CREATE_ROW, CREATE_COLUMN)?
        .is_some()
    {
        return Err("Pages create target already has a root comment".into());
    }
    let position = At::new(u32::try_from(CREATE_ROW)?, u32::try_from(CREATE_COLUMN)?);
    let host = {
        let mut editor = pages::Draft::from_bytes(&input.bytes)?;
        let model = host_model(&editor)?;
        editor.set_table_cell_comment(model, CREATE_ROW, CREATE_COLUMN, "W1-G create")?;
        editor.to_bytes()?
    };
    let leaf_run = {
        let package = Package::from_bytes(&input.bytes)?;
        let commit = package
            .edit_body_table_cell_comment(TABLE, position)?
            .set("W1-G create")?
            .commit()?;
        let restored = commit
            .package()
            .apply_body_table_cell_comment(&commit.patch().inverse())?;
        LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        )
    };
    if host == input.bytes {
        return Err("Pages host create produced a no-op".into());
    }
    if leaf_run
        .restored
        .as_deref()
        .is_none_or(|restored| restored != input.bytes.as_slice())
    {
        return Err("Pages leaf create inverse did not restore the exact source".into());
    }
    Format::Pages.reopen(&host)?;
    Format::Pages.reopen(&leaf_run.output)?;
    for (route, bytes) in [
        ("host", host.as_slice()),
        ("leaf", leaf_run.output.as_slice()),
    ] {
        if route == "host" {
            let editor = pages::Draft::from_bytes(bytes)?;
            if editor
                .table_cell_comment(host_model(&editor)?, CREATE_ROW, CREATE_COLUMN)?
                .is_none()
            {
                return Err("Pages host create did not persist the stored target cell".into());
            }
        } else if Package::from_bytes(bytes)?
            .body_table_cell_comment(TABLE, position)?
            .is_none()
        {
            return Err("Pages leaf create did not persist the stored target cell".into());
        }
    }
    let host_root = narrow_selected_root(&host, false, CREATE_ROW, CREATE_COLUMN)?;
    let leaf_root = narrow_selected_root(&leaf_run.output, true, CREATE_ROW, CREATE_COLUMN)?;
    let (
        selected_member,
        selected_object,
        author_member,
        author_storage_object,
        excluded,
        host_last,
    ) = narrow_pages_create_policy(&input.bytes, &host, &leaf_run.output, host_root, leaf_root)?;
    let selected_list_payload =
        narrow_selected_list_payload(&leaf_run.output, &selected_member, selected_object)?;
    assert!(
        narrow_remap_selected_list(&selected_list_payload, u64::MAX, host_root).is_err(),
        "Pages list remapper accepted a payload without the selected root"
    );
    let normalized_leaf = narrow_remap_leaf(
        &leaf_run.output,
        &selected_member,
        selected_object,
        &author_member,
        author_storage_object,
        *excluded
            .iter()
            .next()
            .ok_or("missing generated author id")?,
        leaf_root,
        host_root,
        host_last,
    )?;
    let generated_author_id = *excluded
        .iter()
        .next()
        .ok_or("missing generated author id")?;
    let nested_unknown = [0x9a, 0x06, 0x01, b'x'];
    let (unknown_list, list_matches) =
        narrow_append_raw_at_path(&selected_list_payload, &[3, 10], &nested_unknown)?;
    assert_eq!(
        list_matches, 1,
        "Pages selected 6005 payload did not expose one comment reference"
    );
    let remapped_unknown_list = narrow_remap_selected_list(&unknown_list, leaf_root, host_root)?;
    assert!(
        remapped_unknown_list
            .windows(nested_unknown.len())
            .any(|window| window == nested_unknown),
        "Pages selected 6005 normalization dropped an unknown nested field"
    );
    let leaf_metadata = narrow_metadata_payload(&leaf_run.output)?;
    let mut metadata_unknown_output = None;
    let metadata_paths: &[&[u32]] = &[
        &[3, 6],
        &[3, 18],
        &[11, 6],
        &[11, 18],
        &[3, 7, 2],
        &[11, 7, 2],
        &[3, 11],
        &[11, 11],
        &[10],
    ];
    for path in metadata_paths {
        let (candidate, matches) =
            narrow_append_raw_at_path(&leaf_metadata, path, &nested_unknown)?;
        if matches != 0 {
            metadata_unknown_output = Some(narrow_remap_metadata(
                &candidate, leaf_root, host_root, host_last,
            )?);
            break;
        }
    }
    let metadata_unknown_output = metadata_unknown_output
        .ok_or("Pages metadata had no selected nested reference path for preservation control")?;
    assert!(
        metadata_unknown_output
            .windows(nested_unknown.len())
            .any(|window| window == nested_unknown),
        "Pages metadata normalization dropped an unknown nested field"
    );
    let corrupted_list = narrow_corrupt_selected_list(
        &leaf_run.output,
        &selected_member,
        selected_object,
        leaf_root,
    )?;
    assert!(
        narrow_pages_create_policy(&input.bytes, &host, &corrupted_list, host_root, leaf_root)
            .is_err(),
        "Pages unrelated 6005 mutation bypassed the full graph policy"
    );
    let corrupted_list_leaf = narrow_remap_leaf(
        &corrupted_list,
        &selected_member,
        selected_object,
        &author_member,
        author_storage_object,
        *excluded
            .iter()
            .next()
            .ok_or("missing generated author id")?,
        leaf_root,
        host_root,
        host_last,
    )?;
    let known = KnownFields::for_format(Format::Pages).with(
        COMMENT_STORAGE_MESSAGE_TYPE,
        "TSD.CommentStorageArchive (new-object date and storage UUID masked)",
        masked_new_pages_comment_storage,
    );
    let diff = archive_diff::compare_with_policies_except_objects_and_members(
        &input.bytes,
        &host,
        &normalized_leaf,
        &known,
        &[],
        MetadataPolicy {
            save_token: true,
            fresh_uuids: true,
        },
        &excluded,
        &[],
    )?;
    let unrelated_list_diff = archive_diff::compare_with_policies_except_objects_and_members(
        &input.bytes,
        &host,
        &corrupted_list_leaf,
        &known,
        &[],
        MetadataPolicy {
            save_token: true,
            fresh_uuids: true,
        },
        &excluded,
        &[],
    )?;
    assert!(
        !unrelated_list_diff.violations.is_empty(),
        "Pages unrelated 6005 mutation was hidden by the archive comparison policy"
    );
    let corrupted_watermark = narrow_corrupt_metadata_watermark(
        &host,
        host_last
            .checked_add(1)
            .ok_or("watermark negative control overflow")?,
    )?;
    assert!(
        narrow_pages_create_policy(
            &input.bytes,
            &corrupted_watermark,
            &leaf_run.output,
            host_root,
            leaf_root,
        )
        .is_err(),
        "Pages metadata watermark mutation bypassed allocation validation"
    );
    let host_ids = narrow_object_ids(&host)?;
    let mut extra_author_id = host_last
        .checked_add(1)
        .ok_or("extra author negative control overflow")?;
    while host_ids
        .values()
        .any(|identifiers| identifiers.contains(&extra_author_id))
    {
        extra_author_id = extra_author_id
            .checked_add(1)
            .ok_or("extra author negative control overflow")?;
    }
    let extra_author = narrow_add_extra_author_object(
        &host,
        &author_member,
        generated_author_id,
        extra_author_id,
    )?;
    assert!(
        narrow_pages_create_policy(
            &input.bytes,
            &extra_author,
            &leaf_run.output,
            host_root,
            leaf_root
        )
        .is_err(),
        "Pages extra type-212 allocation bypassed graph scope validation"
    );
    let extra_author_reference =
        narrow_corrupt_author_info_reference(&host, &author_member, generated_author_id)?;
    assert!(
        narrow_pages_create_policy(
            &input.bytes,
            &extra_author_reference,
            &leaf_run.output,
            host_root,
            leaf_root,
        )
        .is_err(),
        "Pages generated type-212 archive reference mutation bypassed locality validation"
    );
    assert!(
        diff.violations.is_empty(),
        "Pages fresh create had an unapproved difference after the pinned author-storage identity policy:\n{}",
        diff.render()
    );
    for append_unknown in [false, true] {
        let corrupted_leaf = narrow_corrupt_selected_comment(
            &normalized_leaf,
            &selected_member,
            host_root,
            append_unknown,
        )?;
        assert!(
            narrow_validate_selected_comment_author(
                &corrupted_leaf,
                &selected_member,
                host_root,
                Some(generated_author_id),
            )
            .is_err(),
            "Pages selected-root corruption (unknown={append_unknown}) bypassed the selected-cell validator"
        );
        if append_unknown {
            continue;
        }
        let negative = archive_diff::compare_with_policies_except_objects_and_members(
            &input.bytes,
            &host,
            &corrupted_leaf,
            &known,
            &[],
            MetadataPolicy {
                save_token: true,
                fresh_uuids: true,
            },
            &excluded,
            &[],
        )?;
        assert!(
            !negative.violations.is_empty(),
            "Pages selected-root corruption (unknown={append_unknown}) was hidden by the allocation policy"
        );
    }
    let corrupted_host =
        narrow_corrupt_author_storage(&host, &author_member, author_storage_object)?;
    assert!(
        narrow_validate_author_storage(
            &input.bytes,
            &corrupted_host,
            &leaf_run.output,
            &author_member,
            author_storage_object,
            generated_author_id,
        )
        .is_err(),
        "Pages unknown author-storage bytes were admitted by the graph policy"
    );
    eprintln!(
        "parity case pages_body_table_comment_create_parity [Pages, tier B, evidence Queued] on {}: passed\n{}",
        input.path,
        diff.render()
    );
    Ok(())
}

#[test]
fn pages_body_table_comment_surviving_storage_exact_parity() -> TestResult {
    let input = fixture(COMMENTS)?;
    let host = pages::Draft::from_bytes(&input.bytes)?.to_bytes()?;
    let leaf = Package::from_bytes(&input.bytes)?.to_bytes()?;
    let checked = narrow_validate_surviving_comment_storage(&input.bytes, &host, &leaf)?;
    let (member, object) = narrow_first_comment_storage(&input.bytes)?;
    let comment = narrow_archive_object(&input.bytes, &member, object)?
        .messages
        .into_iter()
        .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or("Pages source comment payload disappeared")?
        .data;
    let corrupted = narrow_corrupt_existing_comment(&leaf, &member, object)?;
    assert!(
        narrow_validate_surviving_comment_storage(&input.bytes, &host, &corrupted).is_err(),
        "Pages source-owned neighboring/reply comment corruption bypassed exact storage validation"
    );
    let nested_unknown = [0x9a, 0x06, 0x01, b'x'];
    let nested_noncanonical = [0x9a, 0x86, 0x00, 0x01, b'x'];
    for field in [2, 3, 4, 5] {
        let unknown = narrow_append_nested_comment_bytes(&comment, field, &nested_unknown)?;
        assert!(
            masked_new_pages_comment_storage(&unknown).is_err(),
            "Pages nested unknown field {field} bypassed creation-only comment validation"
        );
        let noncanonical =
            narrow_append_nested_comment_bytes(&comment, field, &nested_noncanonical)?;
        assert!(
            masked_new_pages_comment_storage(&noncanonical).is_err(),
            "Pages nested noncanonical field {field} bypassed creation-only comment validation"
        );
    }
    eprintln!(
        "parity case pages_body_table_comment_surviving_storage_exact_parity [Pages, proof-only] on {}: passed ({} source type-3056 payloads checked)",
        input.path, checked
    );
    Ok(())
}

#[test]
fn pages_body_table_comment_clear_parity() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G clear root")?;
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-for-clear",
        bytes: authored.to_bytes()?,
    };
    comment_case(
        "pages_body_table_comment_clear_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        "generated:pages-table-comment-root-for-clear",
        |editor, model| {
            editor.clear_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |package, position| {
            let clear = package
                .edit_body_table_cell_comment(TABLE, position)?
                .clear()?
                .commit()?;
            let restored = clear
                .package()
                .apply_body_table_cell_comment(&clear.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(clear.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

/// Full Pages facade cost probe for the root-comment rewrite.  The source is
/// host-authored before ABBA, so the host and leaf both measure replacement of
/// the same existing root rather than construction of different inputs.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_cost_abba() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G original")?;
    let input = authored.to_bytes()?;
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &crate::harness::Fixture {
            path: "generated:pages-table-comment-root-cost",
            bytes: input,
        },
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G measured")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_cell_comment(TABLE, position)?
                .set("W1-G measured")?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Pages facade cost probe for reading a populated root/reply thread.
/// The source is host-authored before ABBA and the semantic values are fed to
/// `black_box`, so this measures the read route rather than ZIP serialization
/// alone.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_read_cost_abba() -> TestResult {
    let input = fixture(COMMENTS)?;
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let comment = editor
                .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
                .ok_or("host read did not find the populated root comment")?;
            let replies = editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?;
            black_box((comment.comment.text.len(), replies.len()));
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let comment = package
                .body_table_cell_comment(TABLE, position)?
                .ok_or("leaf read did not find the populated root comment")?;
            let replies = comment.replies().unwrap_or(&[]);
            black_box((comment.text().len(), replies.len()));
            Ok(leaf_bytes!(package))
        },
    )
}

/// Full Pages facade cost probe for creating a root comment on an existing
/// stored cell without a prior root comment.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_create_cost_abba() -> TestResult {
    let base = fixture(CELLS)?;
    let authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    if authored
        .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
        .is_some()
    {
        return Err("Pages create probe target already has a root comment".into());
    }
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-create-cost",
        bytes: base.bytes,
    };
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G create")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_cell_comment(TABLE, position)?
                .set("W1-G create")?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Pages facade cost probe for clearing a host-authored populated root.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_clear_cost_abba() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G clear")?;
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-clear-cost",
        bytes: authored.to_bytes()?,
    };
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.clear_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let clear = package
                .edit_body_table_cell_comment(TABLE, position)?
                .clear()?
                .commit()?;
            Ok(leaf_bytes!(clear.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_cell_value_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(CATALOG)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_table_cell(model, 1, 3, CellValue::Text("Zürich".to_owned()))?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            edit.set(1, 3, Value::Text("Zürich".to_owned()))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_cell_value_batch_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(CATALOG)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let updates = (0..1_000)
                .map(|index| {
                    TableCellUpdate::new(
                        1 + index % 4,
                        index / 4 % 4,
                        CellValue::Text(format!("v{index}")),
                    )
                })
                .collect::<Vec<_>>();
            // Later updates of one coordinate replace earlier ones; keep one per cell.
            let mut seen = std::collections::BTreeMap::new();
            for update in updates {
                seen.insert((update.row, update.column), update);
            }
            editor.set_table_cells(model, seen.into_values())?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            let mut seen = std::collections::BTreeMap::new();
            for index in 0..1_000u32 {
                seen.insert((1 + index % 4, index / 4 % 4), format!("v{index}"));
            }
            for ((row, column), text) in seen {
                edit.set(row, column, Value::Text(text))?;
            }
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// `clear_table_cell` against the values family's `Value::Empty` on a cell
/// the host filled first (both drop the emptied row and its bare header).
#[test]
fn pages_body_table_cell_clear_parity() -> TestResult {
    clear_case(Tier::A, Evidence::Queued)
}

fn clear_case(tier: Tier, evidence: Evidence) -> TestResult {
    let source = fixture(CATALOG)?;
    let mut editor = pages::Draft::from_bytes(&source.bytes)?;
    let model = host_model(&editor)?;
    editor.set_table_cell(model, 1, 3, CellValue::Text("Zürich".to_owned()))?;
    let input = crate::harness::Fixture {
        path: "iwork/pages/body-table-catalog-native.pages + host D2 \"Zürich\"",
        bytes: editor.to_bytes()?,
    };
    Case::new(
        "pages_body_table_cell_clear_parity",
        Format::Pages,
        tier,
        evidence,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = host_model(&editor)?;
        editor.clear_table_cell(model, 1, 3)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_table_values(TABLE)?;
        edit.set(1, 3, Value::Empty)?;
        let commit = edit.commit()?;
        let restored = commit
            .package()
            .apply_body_table_values(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host table cells", host_cells)
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    Ok(())
}
