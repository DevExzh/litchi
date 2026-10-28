//! Numbers table-cell comment mutation parity (wave W1-G, record 0379).
//!
//! The focused host routes still take a native table-model identity and reply
//! storage identity; the leaf routes use only sheet/table positions, a checked
//! cell coordinate, and a checked reply ordinal.  The fixture is native-authored
//! and starts with one root at B2 and no direct replies.  Reply replacement and
//! removal use a host-authored setup containing one reply so the compared
//! operation itself remains a single bounded route.

use litchi_iwa::numbers;
use litchi_numbers::cell::comment::CommentReplyIndex;
use litchi_numbers::table::At;
use litchi_numbers::{Package, SheetSelector, TableSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;

const FIXTURE: &str = "iwork/numbers/comment-reader-native-created.numbers";
const ROW: usize = 1;
const COLUMN: usize = 1;

#[derive(Clone, Copy, Debug)]
enum Operation {
    Set,
    AddReply,
    SetReply,
    RemoveReply,
}

fn position() -> At {
    At::new(ROW as u32, COLUMN as u32)
}

fn host_model(editor: &numbers::Draft) -> TestResult<u64> {
    crate::numbers_cases::host_first_table_model_id(editor)
}

fn host_replies(bytes: &[u8]) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let replies = editor.cell_comment_replies(model, ROW, COLUMN)?;
    Ok(format!(
        "{:?}",
        replies
            .iter()
            .map(|reply| reply.comment.text.as_str())
            .collect::<Vec<_>>()
    ))
}

fn leaf_replies(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        position(),
    )?;
    Ok(format!(
        "{:?}",
        replies.iter().map(|reply| reply.text()).collect::<Vec<_>>()
    ))
}

fn host_route(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    match operation {
        Operation::Set => {
            editor.set_cell_comment(model, ROW, COLUMN, "W1-G revised root")?;
        },
        Operation::AddReply => {
            editor.add_cell_comment_reply(model, ROW, COLUMN, "W1-G added reply")?;
        },
        Operation::SetReply => {
            let reply = editor
                .cell_comment_replies(model, ROW, COLUMN)?
                .into_iter()
                .next()
                .ok_or("host setup has no reply")?;
            editor.set_cell_comment_reply(
                model,
                ROW,
                COLUMN,
                reply.storage_id.get(),
                "W1-G revised reply",
            )?;
        },
        Operation::RemoveReply => {
            let reply = editor
                .cell_comment_replies(model, ROW, COLUMN)?
                .into_iter()
                .next()
                .ok_or("host setup has no reply")?;
            editor.remove_cell_comment_reply(model, ROW, COLUMN, reply.storage_id.get())?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], operation: Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = match operation {
        Operation::Set => package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W1-G revised root",
        )?,
        Operation::AddReply => package.add_table_cell_comment_reply(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W1-G added reply",
        )?,
        Operation::SetReply => package.set_table_cell_comment_reply(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            CommentReplyIndex::new(0),
            "W1-G revised reply",
        )?,
        Operation::RemoveReply => package.remove_table_cell_comment_reply(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            CommentReplyIndex::new(0),
        )?,
    };
    let restored = commit
        .package()
        .apply_table_cell_comment(commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn reply_setup() -> TestResult<Fixture> {
    let input = fixture(FIXTURE)?;
    let bytes = host_route(&input.bytes, Operation::AddReply)?;
    Ok(Fixture {
        path: "iwork/numbers/comment-reader-native-created.numbers + host reply setup",
        bytes,
    })
}

fn run(name: &'static str, operation: Operation, input: &Fixture) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .host(move |source| host_route(source, operation))
        .leaf(move |source| leaf_route(source, operation))
        .read("host table-cell replies", host_replies)
        .read("leaf table-cell replies", leaf_replies)
        .known_fields(|known| {
            crate::text_comment_cases::masked(known)
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
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

#[test]
fn numbers_table_cell_comment_set_parity() -> TestResult {
    let input = fixture(FIXTURE)?;
    run(
        "numbers_table_cell_comment_set_parity",
        Operation::Set,
        &input,
    )
}

#[test]
fn numbers_table_cell_comment_reply_add_parity() -> TestResult {
    let input = fixture(FIXTURE)?;
    run(
        "numbers_table_cell_comment_reply_add_parity",
        Operation::AddReply,
        &input,
    )
}

#[test]
fn numbers_table_cell_comment_reply_set_parity() -> TestResult {
    let input = reply_setup()?;
    run(
        "numbers_table_cell_comment_reply_set_parity",
        Operation::SetReply,
        &input,
    )
}

#[test]
fn numbers_table_cell_comment_reply_remove_parity() -> TestResult {
    let input = reply_setup()?;
    run(
        "numbers_table_cell_comment_reply_remove_parity",
        Operation::RemoveReply,
        &input,
    )
}
