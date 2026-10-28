//! Numbers table-cell comment mutation parity (wave W1-G, record 0379).
//!
//! The first shared slice covers only root replacement.  Direct reply
//! mutations remain an explicit host-only follow-up until their graph
//! transitions are lifted into the shared native engine.

#[cfg(feature = "parity-host-oracle")]
use std::collections::{BTreeMap, BTreeSet};
#[cfg(feature = "parity-host-oracle")]
use std::hint::black_box;

use litchi_iwa::numbers;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_common::WireLimits;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_common::comment::AuthorId;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_common::wire::{WireView, append_varint_field, patch_varint_field};
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_protos::tsd;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_protos::tsk;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_protos::tst;
use litchi_numbers::table::At;
use litchi_numbers::{Package, SheetSelector, TableSelector};
#[cfg(feature = "parity-host-oracle")]
use prost::Message as _;

#[cfg(feature = "parity-host-oracle")]
use crate::harness::fixture;
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
#[cfg(feature = "parity-host-oracle")]
use crate::known_fields::{Decoded, decode_as, top_level_fields};
use crate::leaf_bytes;
#[cfg(feature = "parity-host-oracle")]
use crate::text_glue::Workspace;

#[cfg(feature = "parity-host-oracle")]
const FIXTURE: &str = "iwork/numbers/comment-reader-native-created.numbers";
const ROW: usize = 1;
const COLUMN: usize = 1;
#[cfg(feature = "parity-host-oracle")]
const AUTHOR_MESSAGE_TYPE: u32 = 212;
#[cfg(feature = "parity-host-oracle")]
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

#[cfg(feature = "parity-host-oracle")]
#[derive(Clone, Copy, Debug)]
enum Operation {
    Set,
    Create,
    Clear,
}

#[cfg(feature = "parity-host-oracle")]
#[derive(Clone, Debug)]
struct CommentListSnapshot {
    entry: String,
    object: u64,
    index: usize,
    list: tst::TableDataList,
    storage_id: u64,
    entry_key: u32,
}

#[cfg(feature = "parity-host-oracle")]
#[derive(Clone, Debug)]
struct DataListSnapshot {
    list: tst::TableDataList,
    payload: Vec<u8>,
}

fn position() -> At {
    At::new(ROW as u32, COLUMN as u32)
}

fn host_model(editor: &numbers::Draft) -> TestResult<u64> {
    crate::numbers_cases::host_first_table_model_id(editor)
}

fn populated_comment_fixture() -> TestResult<Fixture> {
    // Keep the source-built marker alive while the host writes both records:
    // `Draft::set_cell_comment` therefore takes its legacy fallback, rather
    // than delegating this fixture's authoring to the focused leaf owner.
    let mut editor = numbers::Builder::new().table_dimensions(4, 4).build()?;
    let model = host_model(&editor)?;
    editor.set_cell_comment(model, ROW, COLUMN, "W1-G read root")?;
    editor.add_cell_comment_reply(model, ROW, COLUMN, "W1-G read reply")?;
    Ok(Fixture {
        path: "generated:numbers-table-comment-root-and-reply",
        bytes: editor.to_bytes()?,
    })
}

#[cfg(feature = "parity-host-oracle")]
fn native_fixture() -> TestResult<Fixture> {
    fixture(FIXTURE)
}

#[cfg(feature = "parity-host-oracle")]
fn rootless_fixture() -> TestResult<Fixture> {
    let native = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    editor.__parity_clear_table_cell_comment_legacy(TableSelector::index(0), ROW, COLUMN)?;
    let input = Fixture {
        path: "generated:numbers-table-comment-root-create-parity",
        bytes: editor.to_bytes()?,
    };
    if numbers::Draft::from_bytes(&input.bytes)?
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .is_some()
    {
        return Err("Numbers create target still has a root comment".into());
    }
    Ok(input)
}

#[cfg(feature = "parity-host-oracle")]
fn populated_neighbor_clear_fixture() -> TestResult<Fixture> {
    let mut editor = numbers::Builder::new().table_dimensions(4, 4).build()?;
    let model = host_model(&editor)?;
    editor.set_cell_comment(model, ROW, COLUMN, "W1-G clear target")?;
    editor.set_cell_comment(model, ROW, COLUMN + 1, "W1-G populated neighbor")?;
    editor.add_cell_comment_reply(model, ROW, COLUMN + 1, "W1-G neighbor reply")?;
    let input = Fixture {
        path: "generated:numbers-table-comment-clear-with-populated-neighbor",
        bytes: editor.to_bytes()?,
    };
    let selected = comment_list_snapshot(&input.bytes, position())?;
    if selected.entry_key.checked_add(1) == Some(selected.list.next_list_id) {
        return Err(
            "Numbers clear locality fixture did not make the target a non-last comment key".into(),
        );
    }
    assert_populated_comment_semantics_at(
        &input.bytes,
        At::new(ROW as u32, (COLUMN + 1) as u32),
        ROW,
        COLUMN + 1,
    )?;
    Ok(input)
}

#[cfg(feature = "parity-host-oracle")]
fn data_list_messages(bytes: &[u8]) -> TestResult<BTreeMap<(String, u64, usize), Vec<u8>>> {
    let mut messages = BTreeMap::new();
    for located in Workspace::open(bytes)?.messages(6_005)? {
        messages.insert(
            (located.entry, located.object, located.index),
            located.payload,
        );
    }
    Ok(messages)
}

#[cfg(feature = "parity-host-oracle")]
fn comment_list_snapshot(bytes: &[u8], position: At) -> TestResult<CommentListSnapshot> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let comment = editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position)?
        .ok_or("selected source has no root comment")?;
    let storage_id = comment.storage_id.get();
    let mut matches = Vec::new();
    for located in Workspace::open(bytes)?.messages(6_005)? {
        let list = tst::TableDataList::decode(located.payload.as_slice())?;
        if list.list_type != tst::table_data_list::ListType::CommentStorage as i32 {
            continue;
        }
        for entry in &list.entries {
            if entry
                .comment_storage
                .as_ref()
                .is_some_and(|reference| reference.identifier == storage_id)
            {
                matches.push(CommentListSnapshot {
                    entry: located.entry.clone(),
                    object: located.object,
                    index: located.index,
                    list: list.clone(),
                    storage_id,
                    entry_key: entry.key,
                });
            }
        }
    }
    match matches.len() {
        1 => Ok(matches.remove(0)),
        count => Err(format!(
            "expected one direct Numbers comment list entry for storage {storage_id}, found {count}"
        )
        .into()),
    }
}

#[cfg(feature = "parity-host-oracle")]
fn data_list_by_key(bytes: &[u8], key: &(String, u64, usize)) -> TestResult<DataListSnapshot> {
    let located = Workspace::open(bytes)?
        .messages(6_005)?
        .into_iter()
        .find(|located| located.entry == key.0 && located.object == key.1 && located.index == key.2)
        .ok_or_else(|| format!("Numbers comment list message {key:?} disappeared"))?;
    let list = tst::TableDataList::decode(located.payload.as_slice())?;
    if list.list_type != tst::table_data_list::ListType::CommentStorage as i32 {
        return Err(format!("selected Numbers list {key:?} changed list type").into());
    }
    Ok(DataListSnapshot {
        payload: located.payload,
        list,
    })
}

#[cfg(feature = "parity-host-oracle")]
fn list_wire_without_target(data: &[u8], target_key: u32) -> TestResult<Vec<Vec<u8>>> {
    let view = WireView::parse(data, WireLimits::DEFAULT)?;
    let mut fields = Vec::new();
    for field in view.fields() {
        if field.number() == 2 {
            continue;
        }
        if field.number() == 3 {
            let entry = tst::table_data_list::ListEntry::decode(field.payload())?;
            if entry.key == target_key {
                continue;
            }
        }
        fields.push(field.raw().to_vec());
    }
    Ok(fields)
}

#[cfg(feature = "parity-host-oracle")]
fn expected_host_clear_next(source: &CommentListSnapshot) -> u32 {
    if source.entry_key.checked_add(1) == Some(source.list.next_list_id) {
        source.entry_key
    } else {
        source.list.next_list_id
    }
}

#[cfg(feature = "parity-host-oracle")]
fn validate_clear_candidate(
    source: &[u8],
    candidate: &[u8],
    source_selected: &CommentListSnapshot,
) -> TestResult<DataListSnapshot> {
    let key = (
        source_selected.entry.clone(),
        source_selected.object,
        source_selected.index,
    );
    let candidate_selected = data_list_by_key(candidate, &key)?;
    if candidate_selected
        .list
        .entries
        .iter()
        .any(|entry| entry.key == source_selected.entry_key)
        || candidate_selected.list.entries.iter().any(|entry| {
            entry
                .comment_storage
                .as_ref()
                .is_some_and(|reference| reference.identifier == source_selected.storage_id)
        })
    {
        return Err("Numbers clear candidate retained the selected comment entry".into());
    }
    let expected_host_next = expected_host_clear_next(source_selected);
    if candidate_selected.list.next_list_id != source_selected.list.next_list_id
        && candidate_selected.list.next_list_id != expected_host_next
    {
        return Err(format!(
            "Numbers clear candidate changed selected next_list_id from source {} to unexpected {} (expected source {} or host reclaim {})",
            source_selected.list.next_list_id,
            candidate_selected.list.next_list_id,
            source_selected.list.next_list_id,
            expected_host_next,
        )
        .into());
    }
    let source_lists = data_list_messages(source)?;
    let candidate_lists = data_list_messages(candidate)?;
    if source_lists.keys().collect::<Vec<_>>() != candidate_lists.keys().collect::<Vec<_>>() {
        return Err("Numbers clear candidate changed the 6005 message inventory".into());
    }
    for (message_key, source_payload) in source_lists {
        let candidate_payload = candidate_lists
            .get(&message_key)
            .ok_or_else(|| format!("Numbers clear candidate lost list {message_key:?}"))?;
        if message_key == key {
            if list_wire_without_target(source_payload.as_slice(), source_selected.entry_key)?
                != list_wire_without_target(candidate_payload, source_selected.entry_key)?
            {
                return Err(format!(
                    "Numbers clear candidate changed selected list fields outside entry {} and its source-derived next_list_id",
                    source_selected.entry_key
                )
                .into());
            }
        } else if source_payload.as_slice() != candidate_payload.as_slice() {
            return Err(format!(
                "Numbers clear candidate changed unrelated 6005 list {message_key:?}"
            )
            .into());
        }
    }
    Ok(candidate_selected)
}

#[cfg(feature = "parity-host-oracle")]
fn assert_clear_list_policy(source: &[u8], host: &[u8], leaf: &[u8], position: At) -> TestResult {
    let source_selected = comment_list_snapshot(source, position)?;
    let host_selected = validate_clear_candidate(source, host, &source_selected)?;
    let leaf_selected = validate_clear_candidate(source, leaf, &source_selected)?;
    let expected_host = expected_host_clear_next(&source_selected);
    if host_selected.list.next_list_id != expected_host {
        return Err(format!(
            "legacy host clear changed the source-derived next_list_id to {} (expected {})",
            host_selected.list.next_list_id, expected_host
        )
        .into());
    }
    if leaf_selected.list.next_list_id != source_selected.list.next_list_id {
        return Err(format!(
            "focused clear changed the source watermark to {} (expected {})",
            leaf_selected.list.next_list_id, source_selected.list.next_list_id
        )
        .into());
    }
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
fn normalize_clear_candidate(source: &[u8], candidate: &[u8]) -> TestResult<Vec<u8>> {
    // The legacy host reclaims a last comment key, while the focused
    // non-axis clear deliberately preserves that source watermark.  This is
    // a comparison-only projection: validate_clear_candidate has already
    // checked the raw route for the selected list, every retained entry, all
    // other 6005 payloads, and their unknown bytes before this one-field
    // source-derived alignment for Case::run.
    let source_selected = comment_list_snapshot(source, position())?;
    let candidate_selected = validate_clear_candidate(source, candidate, &source_selected)?;
    let expected = expected_host_clear_next(&source_selected);
    if candidate_selected.list.next_list_id == expected {
        return Ok(candidate.to_vec());
    }
    let key = (
        source_selected.entry,
        source_selected.object,
        source_selected.index,
    );
    let patched = patch_varint_field(
        &candidate_selected.payload,
        2,
        true,
        Some(u64::from(expected)),
    )?;
    let mut workspace = Workspace::open(candidate)?;
    let located = workspace
        .messages(6_005)?
        .into_iter()
        .find(|located| located.entry == key.0 && located.object == key.1 && located.index == key.2)
        .ok_or_else(|| format!("Numbers clear candidate message {key:?} disappeared"))?;
    workspace.replace(&located, patched, None)?;
    workspace.into_bytes()
}

#[cfg(feature = "parity-host-oracle")]
fn patch_data_list_next(
    bytes: &[u8],
    key: &(String, u64, usize),
    replacement: u32,
) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(bytes)?;
    let located = workspace
        .messages(6_005)?
        .into_iter()
        .find(|located| located.entry == key.0 && located.object == key.1 && located.index == key.2)
        .ok_or_else(|| format!("Numbers list {key:?} disappeared"))?;
    let payload = patch_varint_field(&located.payload, 2, true, Some(u64::from(replacement)))?;
    workspace.replace(&located, payload, None)?;
    workspace.into_bytes()
}

#[cfg(feature = "parity-host-oracle")]
fn append_data_list_unknown(bytes: &[u8], key: &(String, u64, usize)) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(bytes)?;
    let located = workspace
        .messages(6_005)?
        .into_iter()
        .find(|located| located.entry == key.0 && located.object == key.1 && located.index == key.2)
        .ok_or_else(|| format!("Numbers list {key:?} disappeared"))?;
    let mut payload = located.payload.clone();
    append_varint_field(&mut payload, 90, 1)?;
    workspace.replace(&located, payload, None)?;
    workspace.into_bytes()
}

fn host_populated_replies(bytes: &[u8]) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let replies = editor.cell_comment_replies(model, ROW, COLUMN)?;
    if replies.is_empty() {
        return Err("host read did not find a populated comment reply".into());
    }
    let metadata = replies
        .iter()
        .map(|reply| {
            if reply.comment.text.is_empty()
                || reply.comment.creation_date_seconds.is_none()
                || reply.comment.author_id.is_none()
            {
                return Err::<String, Box<dyn std::error::Error>>(
                    "host reply is missing text/date/author metadata".into(),
                );
            }
            Ok(format!(
                "{}|{:?}|{}",
                reply.comment.text,
                reply.comment.creation_date_seconds,
                reply.comment.author_id.is_some()
            ))
        })
        .collect::<TestResult<Vec<_>>>()?;
    Ok(format!("replies={metadata:?}"))
}

fn leaf_populated_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let root = package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("leaf read did not find a populated root comment")?;
    if root.text().is_empty() || root.timestamp().is_none() || root.author().is_none() {
        return Err("leaf root comment is missing text/date/author metadata".into());
    }
    let replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        position(),
    )?;
    if replies.is_empty() {
        return Err("leaf read did not find a populated comment reply".into());
    }
    let metadata = replies
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
    // The host's focused Numbers facade exposes reply metadata here; the
    // root assertion above still proves that the same host-authored source
    // contains a populated root, but it is not projected into this legacy
    // host read string.
    Ok(format!("replies={metadata:?}"))
}

#[cfg(feature = "parity-host-oracle")]
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

#[cfg(feature = "parity-host-oracle")]
fn decode_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let decoded = decode_as::<tsd::CommentStorageArchive>(data)?;
    Ok(decoded)
}

#[cfg(feature = "parity-host-oracle")]
fn masked_new_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tsd::CommentStorageArchive::decode(data)?;
    if message.creation_date.is_some() {
        message.creation_date = Some(litchi_iwa_protos::tsp::Date { seconds: 0.0 });
    }
    if message.storage_uuid.is_some() {
        message.storage_uuid = Some(litchi_iwa_protos::tsp::Uuid { lower: 1, upper: 1 });
    }
    let mut decoded = decode_as::<tsd::CommentStorageArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

#[cfg(feature = "parity-host-oracle")]
fn existing_comment_metadata_at(
    bytes: &[u8],
    position: At,
) -> TestResult<(Vec<u8>, Option<f64>, Option<(u64, u64)>)> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let comment = editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position)?
        .ok_or("source has no existing root comment at selected position")?;
    let located = Workspace::open(bytes)?.message(comment.storage_id.get(), 3056)?;
    let decoded = tsd::CommentStorageArchive::decode(located.payload.as_slice())?;
    let view = litchi_iwa_common::wire::WireView::parse(
        located.payload.as_slice(),
        litchi_iwa_common::WireLimits::DEFAULT,
    )?;
    let mut non_text_fields = Vec::new();
    for field in view.fields().filter(|field| field.number() != 1) {
        non_text_fields.extend_from_slice(field.raw());
    }
    Ok((
        non_text_fields,
        decoded.creation_date.map(|date| date.seconds),
        decoded.storage_uuid.map(|uuid| (uuid.lower, uuid.upper)),
    ))
}

#[cfg(feature = "parity-host-oracle")]
fn existing_comment_metadata(
    bytes: &[u8],
) -> TestResult<(Vec<u8>, Option<f64>, Option<(u64, u64)>)> {
    existing_comment_metadata_at(bytes, position())
}

/// Compare host and leaf semantic values on the same host-authored bytes.
/// The `Case::read` callbacks otherwise consume separate route outputs and
/// cannot establish that the readers project an untouched source identically.
#[cfg(feature = "parity-host-oracle")]
fn assert_populated_comment_semantics_at(
    bytes: &[u8],
    position: At,
    row: usize,
    column: usize,
) -> TestResult {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let host = editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position)?
        .ok_or("host read did not find the populated root comment")?;
    let host_replies = editor.cell_comment_replies(model, row, column)?;
    if host_replies.is_empty() {
        return Err("host source must contain a non-empty reply list".into());
    }

    let package = Package::from_bytes(bytes)?;
    let leaf = package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position)?
        .ok_or("leaf read did not find the populated root comment")?;
    let leaf_replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    )?;
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

#[cfg(feature = "parity-host-oracle")]
fn assert_populated_comment_semantics(bytes: &[u8]) -> TestResult {
    assert_populated_comment_semantics_at(bytes, position(), ROW, COLUMN)
}

#[cfg(feature = "parity-host-oracle")]
fn preview_names(bytes: &[u8]) -> TestResult<BTreeSet<String>> {
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(bytes, Default::default())?;
    Ok(catalog
        .iter()
        .filter(|entry| ROOT_PREVIEWS.contains(&entry.name()))
        .map(|entry| entry.name().to_owned())
        .collect())
}

#[cfg(feature = "parity-host-oracle")]
fn host_route(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match operation {
        Operation::Set | Operation::Create => {
            // This is the deliberately independent legacy oracle.  The
            // helper resolves the selector and calls the old package writer
            // directly; it never constructs a focused `litchi_numbers`
            // package.
            editor.__parity_set_table_cell_comment_legacy(
                TableSelector::index(0),
                ROW,
                COLUMN,
                "W1-G revised root",
            )?;
        },
        Operation::Clear => {
            editor.__parity_clear_table_cell_comment_legacy(
                TableSelector::index(0),
                ROW,
                COLUMN,
            )?;
        },
    }
    Ok(editor.to_bytes()?)
}

#[cfg(feature = "parity-host-oracle")]
fn leaf_forward(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let commit = match operation {
        Operation::Set | Operation::Create => package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W1-G revised root",
        )?,
        Operation::Clear => package.clear_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
        )?,
    };
    Ok(leaf_bytes!(commit.package()))
}

#[cfg(feature = "parity-host-oracle")]
fn leaf_route(source: &[u8], operation: Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = match operation {
        Operation::Set | Operation::Create => package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W1-G revised root",
        )?,
        Operation::Clear => package.clear_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
        )?,
    };
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    let output = leaf_bytes!(commit.package());
    let restored = leaf_bytes!(restored.package());
    if restored != source {
        return Err("Numbers comment inverse did not restore the exact source".into());
    }
    Ok(LeafRun::with_inverse(output, restored))
}

#[cfg(feature = "parity-host-oracle")]
fn run(
    name: &'static str,
    operation: Operation,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
) -> TestResult {
    if matches!(operation, Operation::Clear) {
        let case = Case::new(name, Format::Numbers, tier, evidence, input)
            .host(move |source| {
                let output = host_route(source, operation)?;
                normalize_clear_candidate(source, &output)
            })
            .leaf(move |source| {
                let actual = leaf_route(source, operation)?;
                let restored = actual
                    .restored
                    .ok_or("Numbers clear route did not publish an inverse")?;
                let output = normalize_clear_candidate(source, &actual.output)?;
                Ok(LeafRun::with_inverse(output, restored))
            })
            .known_fields(|known| {
                known
                    .with(3_056, "TSD.CommentStorageArchive", decode_comment_storage)
                    .with(
                        212,
                        "TSK.AnnotationAuthorArchive",
                        decode_as::<litchi_iwa_protos::tsk::AnnotationAuthorArchive>,
                    )
                    .with(
                        213,
                        "TSK.AnnotationAuthorStorageArchive",
                        decode_as::<litchi_iwa_protos::tsk::AnnotationAuthorStorageArchive>,
                    )
            })
            .leaf_invalidates_root_previews()
            .leaf_structural_metadata();
        case.run()
            .map_err(|error| format!("{operation:?}: {error}"))?;
        return Ok(());
    }
    let mut case = Case::new(name, Format::Numbers, tier, evidence, input)
        .host(move |source| host_route(source, operation))
        .leaf(move |source| leaf_route(source, operation))
        .known_fields(|known| {
            known
                .with(3_056, "TSD.CommentStorageArchive", decode_comment_storage)
                .with(
                    212,
                    "TSK.AnnotationAuthorArchive",
                    decode_as::<litchi_iwa_protos::tsk::AnnotationAuthorArchive>,
                )
                .with(
                    213,
                    "TSK.AnnotationAuthorStorageArchive",
                    decode_as::<litchi_iwa_protos::tsk::AnnotationAuthorStorageArchive>,
                )
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata();
    if matches!(operation, Operation::Create) {
        case = case.known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (new-object date and storage UUID masked)",
                masked_new_comment_storage,
            )
        });
    }
    case.run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
fn comment_storage_objects(bytes: &[u8]) -> TestResult<BTreeSet<u64>> {
    Ok(Workspace::open(bytes)?
        .messages(3_056)?
        .into_iter()
        .map(|located| located.object)
        .collect())
}

#[cfg(feature = "parity-host-oracle")]
fn assert_created_root_semantics(source: &[u8], host: &[u8], leaf: &[u8]) -> TestResult {
    let host_root = numbers::Draft::from_bytes(host)?
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .ok_or("legacy host create did not persist the root comment")?;
    let leaf_package = Package::from_bytes(leaf)?;
    let leaf_root = leaf_package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("leaf create did not persist the root comment")?;
    assert_eq!(host_root.comment.text, "W1-G revised root");
    assert_eq!(leaf_root.text(), "W1-G revised root");
    assert!(
        host_root.comment.creation_date_seconds.is_some(),
        "legacy host create must stamp a creation date"
    );
    assert!(
        leaf_root.timestamp().is_some(),
        "focused create must retain a creation date"
    );
    let host_author = host_root
        .comment
        .author_id
        .ok_or("legacy host create did not publish an author")?;
    let leaf_author = leaf_root
        .author()
        .ok_or("focused create did not publish an author")?;
    assert_eq!(
        author_metadata(host, Some(host_author))?,
        (
            leaf_author.display_name().map(str::to_owned),
            leaf_author.public_id().map(str::to_owned),
        ),
        "created root author metadata"
    );
    let source_storage_objects = comment_storage_objects(source)?;
    let host_storage = comment_list_snapshot(host, position())?.storage_id;
    let leaf_storage = comment_list_snapshot(leaf, position())?.storage_id;
    assert!(
        !source_storage_objects.contains(&host_storage),
        "host create reused a source comment-storage object"
    );
    assert!(
        !source_storage_objects.contains(&leaf_storage),
        "focused create reused a source comment-storage object"
    );
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_replies_read_parity() -> TestResult {
    let input = populated_comment_fixture()?;
    assert_populated_comment_semantics(&input.bytes)?;
    Case::new(
        "numbers_table_cell_comment_replies_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host table-cell replies", host_populated_replies)
    .read("leaf table-cell comment thread", leaf_populated_thread)
    .run()?;
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_set_parity() -> TestResult {
    let input = native_fixture()?;
    let before = existing_comment_metadata(&input.bytes)?;
    if before.1.is_none() || before.2.is_none() {
        return Err("native root comment is missing source-owned date/UUID metadata".into());
    }
    run(
        "numbers_table_cell_comment_set_parity",
        Operation::Set,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Set)?;
    let leaf = leaf_route(&input.bytes, Operation::Set)?.output;
    for (route, output) in [("host", host), ("leaf", leaf)] {
        let after = existing_comment_metadata(&output)?;
        assert_eq!(
            after.0, before.0,
            "{route} rewrote existing comment metadata"
        );
        assert_eq!(after.1, before.1, "{route} rewrote existing comment date");
        assert_eq!(after.2, before.2, "{route} rewrote existing comment UUID");
    }
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_create_parity() -> TestResult {
    let input = rootless_fixture()?;
    run(
        "numbers_table_cell_comment_create_parity",
        Operation::Create,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Create)?;
    let leaf = leaf_route(&input.bytes, Operation::Create)?.output;
    assert_created_root_semantics(&input.bytes, &host, &leaf)?;
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_create_identity_mask_negative_control() -> TestResult {
    let input = native_fixture()?;
    let editor = numbers::Draft::from_bytes(&input.bytes)?;
    let source_comment = editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .ok_or("native source has no comment-storage object")?;
    let located = Workspace::open(&input.bytes)?.message(source_comment.storage_id.get(), 3_056)?;
    let baseline = masked_new_comment_storage(&located.payload)?.canonical;
    let mut changed = tsd::CommentStorageArchive::decode(located.payload.as_slice())?;
    changed.text = Some("W1-G identity-mask negative control".to_owned());
    let changed = masked_new_comment_storage(&changed.encode_to_vec())?.canonical;
    assert_ne!(
        baseline, changed,
        "new-comment identity mask unexpectedly hid a changed text field"
    );
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_clear_parity() -> TestResult {
    let input = native_fixture()?;
    let editor = numbers::Draft::from_bytes(&input.bytes)?;
    let model = host_model(&editor)?;
    if editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .is_none()
    {
        return Err("Numbers clear target has no root comment".into());
    }
    if !editor.cell_comment_replies(model, ROW, COLUMN)?.is_empty() {
        return Err("Numbers clear target unexpectedly has replies".into());
    }
    run(
        "numbers_table_cell_comment_clear_parity",
        Operation::Clear,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Clear)?;
    let leaf = leaf_route(&input.bytes, Operation::Clear)?.output;
    assert_clear_list_policy(&input.bytes, &host, &leaf, position())?;
    assert!(
        numbers::Draft::from_bytes(&host)?
            .__parity_read_table_cell_comment(TableSelector::index(0), position())?
            .is_none(),
        "legacy host clear left a root comment"
    );
    assert!(
        Package::from_bytes(&leaf)?
            .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
            .is_none(),
        "leaf clear left a root comment"
    );
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_clear_preserves_populated_neighbor() -> TestResult {
    let input = populated_neighbor_clear_fixture()?;
    let neighbor_position = At::new(ROW as u32, (COLUMN + 1) as u32);
    let neighbor_before = existing_comment_metadata_at(&input.bytes, neighbor_position)?;
    run(
        "numbers_table_cell_comment_clear_preserves_populated_neighbor",
        Operation::Clear,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Clear)?;
    let leaf = leaf_route(&input.bytes, Operation::Clear)?.output;
    assert_clear_list_policy(&input.bytes, &host, &leaf, position())?;
    for (route, output) in [("host", host), ("leaf", leaf)] {
        assert!(
            numbers::Draft::from_bytes(&output)?
                .__parity_read_table_cell_comment(TableSelector::index(0), position())?
                .is_none(),
            "{route} clear left the selected root comment"
        );
        assert_eq!(
            existing_comment_metadata_at(&output, neighbor_position)?,
            neighbor_before,
            "{route} clear rewrote unrelated populated comment metadata"
        );
        assert_populated_comment_semantics_at(&output, neighbor_position, ROW, COLUMN + 1)?;
    }
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_clear_policy_negative_controls() -> TestResult {
    let input = native_fixture()?;
    let host = host_route(&input.bytes, Operation::Clear)?;
    let leaf = leaf_route(&input.bytes, Operation::Clear)?.output;
    let source_selected = comment_list_snapshot(&input.bytes, position())?;
    let selected_key = (
        source_selected.entry.clone(),
        source_selected.object,
        source_selected.index,
    );
    let host_selected = data_list_by_key(&host, &selected_key)?;
    let wrong_next = host_selected
        .list
        .next_list_id
        .checked_add(1)
        .ok_or("clear negative-control next_list_id overflow")?;
    let wrong_selected_counter = patch_data_list_next(&host, &selected_key, wrong_next)?;
    assert!(
        assert_clear_list_policy(&input.bytes, &wrong_selected_counter, &leaf, position()).is_err(),
        "clear policy accepted a wrong selected next_list_id"
    );

    let selected_unknown_tail = append_data_list_unknown(&leaf, &selected_key)?;
    assert!(
        assert_clear_list_policy(&input.bytes, &host, &selected_unknown_tail, position()).is_err(),
        "clear policy accepted an unknown tail on the selected list"
    );

    let unrelated_key = data_list_messages(&input.bytes)?
        .into_keys()
        .find(|key| key != &selected_key)
        .ok_or("Numbers fixture has no unrelated 6005 list for the negative control")?;
    let unrelated_unknown_tail = append_data_list_unknown(&leaf, &unrelated_key)?;
    assert!(
        assert_clear_list_policy(&input.bytes, &host, &unrelated_unknown_tail, position()).is_err(),
        "clear policy accepted an unrelated 6005 unknown tail mutation"
    );
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_host_oracle_control() -> TestResult {
    let input = native_fixture()?;
    let output = host_route(&input.bytes, Operation::Set)?;
    assert_ne!(output, input.bytes, "legacy oracle must publish a change");
    let comment = Package::from_bytes(&output)?
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("legacy host oracle did not persist the comment")?;
    assert_eq!(comment.text(), "W1-G revised root");

    // The legacy helper above calls `set_cell_comment_in_package` directly;
    // it does not construct or invoke the focused package owner.  Preserve
    // the three preview members as the discriminating legacy behavior: the
    // focused leaf route intentionally invalidates them after its write.
    let input_previews = preview_names(&input.bytes)?;
    assert_eq!(input_previews.len(), ROOT_PREVIEWS.len());
    assert_eq!(preview_names(&output)?, input_previews);
    let leaf_output = leaf_route(&input.bytes, Operation::Set)?.output;
    assert!(
        preview_names(&leaf_output)?.len() < input_previews.len(),
        "focused route should invalidate at least one root preview"
    );
    Ok(())
}

/// Full Numbers facade cost probe for the focused root replacement.  This
/// uses the same native source as the parity case and the independent legacy
/// host oracle; run it with `--features parity-host-oracle` under callgrind.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_cost_abba() -> TestResult {
    let input = native_fixture()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, Operation::Set),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package.set_table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                "W1-G revised root",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Numbers facade cost probe for reading a populated root/reply thread.
/// This uses the feature-gated legacy host reader as the differential oracle;
/// it is not a supported host API and is not native application evidence.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_read_cost_abba() -> TestResult {
    let input = populated_comment_fixture()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let editor = numbers::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let root = editor
                .__parity_read_table_cell_comment(TableSelector::index(0), position())?
                .ok_or("host read did not find the populated root comment")?;
            let replies = editor.cell_comment_replies(model, ROW, COLUMN)?;
            black_box((root.comment.text.len(), replies.len()));
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let root = package
                .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
                .ok_or("leaf read did not find the populated root comment")?;
            let replies = package.table_cell_comment_replies(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
            )?;
            black_box((root.text().len(), replies.len()));
            Ok(leaf_bytes!(package))
        },
    )
}

/// Full Numbers facade cost probe for creating a root comment.  The input is
/// host-authored and cleared before ABBA, and both mutation routes use the
/// feature-gated legacy writer/reader oracle boundary described above.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_create_cost_abba() -> TestResult {
    let native = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    editor.__parity_clear_table_cell_comment_legacy(TableSelector::index(0), ROW, COLUMN)?;
    let input = Fixture {
        path: "generated:numbers-table-comment-root-create-cost",
        bytes: editor.to_bytes()?,
    };
    if numbers::Draft::from_bytes(&input.bytes)?
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .is_some()
    {
        return Err("Numbers create probe target still has a root comment".into());
    }
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, Operation::Create),
        |source| leaf_forward(source, Operation::Create),
    )
}

/// Full Numbers facade cost probe for clearing a host-authored root comment.
/// The root has no replies, matching the supported root-only clear profile.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_clear_cost_abba() -> TestResult {
    let native = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    editor.__parity_set_table_cell_comment_legacy(
        TableSelector::index(0),
        ROW,
        COLUMN,
        "W1-G clear",
    )?;
    let input = Fixture {
        path: "generated:numbers-table-comment-root-clear-cost",
        bytes: editor.to_bytes()?,
    };
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, Operation::Clear),
        |source| leaf_forward(source, Operation::Clear),
    )
}
