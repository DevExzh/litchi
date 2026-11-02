//! Pages body-table comment threads (record 0412c): set and clear of a
//! reply-bearing root, and the direct-reply lifecycle (append, replace,
//! remove) through `litchi_iwa_table::comments_native::rewrite_reply`.
//!
//! Every case starts from the native reply-bearing fixture.  The host's
//! reply writer is copy-on-write and stamps a fresh date and random storage
//! UUID on a new reply, so the write cases are differential (Tier B): both
//! outputs reopen, the host and leaf thread reads agree, and the leaf
//! inverse restores the exact source.

use litchi_iwa::pages;
use litchi_iwa_common::table::coordinate::At;
use litchi_pages::Package;
use litchi_pages::table::comments::{Edit, Error};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::pages_table_cases::{
    COMMENT_COLUMN, COMMENT_ROW, COMMENTS, TABLE, comment_position, host_comment_thread,
    host_model, leaf_comment_thread, masked_new_pages_comment_storage,
};

fn source() -> TestResult<Fixture> {
    fixture(COMMENTS)
}

/// The host storage identifier of the source reply at `index`.
fn host_reply_id(editor: &pages::Draft, model: u64, index: usize) -> TestResult<u64> {
    Ok(editor
        .table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?
        .get(index)
        .ok_or("the source thread has no reply at the requested ordinal")?
        .storage_id
        .get())
}

/// Run one staged leaf edit, then its inverse.
fn leaf_run(
    package: &Package,
    position: At,
    stage: impl for<'a> FnOnce(Edit<'a>) -> Result<Edit<'a>, Error>,
) -> TestResult<LeafRun> {
    let commit = stage(package.edit_body_table_cell_comment(TABLE, position)?)?.commit()?;
    let restored = commit
        .package()
        .apply_body_table_cell_comment(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// `pages_table_cases::comment_case` with the new-storage date and UUID
/// masked: the host stamps the clock and a random UUID on a new reply.
fn thread_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut pages::Draft, u64) -> TestResult + 'static,
    stage: impl for<'a> FnOnce(Edit<'a>) -> Result<Edit<'a>, Error> + 'static,
) -> TestResult {
    let input = source()?;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(move |bytes| {
            let mut editor = pages::Draft::from_bytes(bytes)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |bytes| {
            let package = Package::from_bytes(bytes)?;
            leaf_run(&package, comment_position()?, stage)
        })
        .leaf_structural_metadata()
        .read("host table comment thread", host_comment_thread)
        .read("leaf table comment thread", leaf_comment_thread)
        .known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (new-object date and storage UUID masked)",
                masked_new_pages_comment_storage,
            )
        })
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_comment_set_with_replies_parity() -> TestResult {
    thread_case(
        "pages_body_table_comment_set_with_replies_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W4-T root")?;
            Ok(())
        },
        |edit| edit.set("W4-T root"),
    )
}

#[test]
fn pages_body_table_comment_clear_with_replies_parity() -> TestResult {
    thread_case(
        "pages_body_table_comment_clear_with_replies_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.clear_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |edit| edit.clear(),
    )
}

#[test]
fn pages_body_table_comment_reply_add_parity() -> TestResult {
    thread_case(
        "pages_body_table_comment_reply_add_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.add_table_cell_comment_reply(
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W4-T reply",
            )?;
            Ok(())
        },
        |edit| edit.add_reply("W4-T reply"),
    )
}

#[test]
fn pages_body_table_comment_reply_set_parity() -> TestResult {
    thread_case(
        "pages_body_table_comment_reply_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.set_table_cell_comment_reply(
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                reply,
                "W4-T edited reply",
            )?;
            Ok(())
        },
        |edit| edit.set_reply(0, "W4-T edited reply"),
    )
}

#[test]
fn pages_body_table_comment_reply_remove_parity() -> TestResult {
    thread_case(
        "pages_body_table_comment_reply_remove_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.remove_table_cell_comment_reply(model, COMMENT_ROW, COMMENT_COLUMN, reply)?;
            Ok(())
        },
        |edit| edit.remove_reply(0),
    )
}

/// A reply ordinal past the source thread is a typed refusal, and an
/// unchanged reply text is a no-op commit.
#[test]
fn pages_body_table_comment_reply_refusals_are_typed() -> TestResult {
    let input = source()?;
    let package = Package::from_bytes(&input.bytes)?;
    let position = comment_position()?;
    let replies = package
        .body_table_cell_comment(TABLE, position)?
        .ok_or("the source has no root comment")?
        .replies()
        .map_or(0, <[_]>::len);
    let refused = package
        .edit_body_table_cell_comment(TABLE, position)?
        .remove_reply(replies)?
        .commit();
    assert!(
        matches!(refused, Err(Error::ReplyNotFound { index }) if index == replies),
        "an out-of-range reply ordinal must be refused: {refused:?}"
    );
    let text = package
        .body_table_cell_comment(TABLE, position)?
        .and_then(|comment| {
            comment
                .replies()
                .and_then(|replies| replies.first().cloned())
        })
        .ok_or("the source has no reply")?;
    let unchanged = package
        .edit_body_table_cell_comment(TABLE, position)?
        .set_reply(0, text.text())?
        .commit()?;
    assert!(
        unchanged.patch().is_noop(),
        "rewriting a reply with its own text must be a no-op"
    );
    Ok(())
}

fn thread_probe(
    host: impl Fn(&mut pages::Draft, u64) -> TestResult,
    stage: impl for<'a> Fn(Edit<'a>) -> Result<Edit<'a>, Error>,
) -> TestResult {
    let input = source()?;
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |bytes| {
            let mut editor = pages::Draft::from_bytes(bytes)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            Ok(editor.to_bytes()?)
        },
        |bytes| {
            let package = Package::from_bytes(bytes)?;
            let commit = stage(package.edit_body_table_cell_comment(TABLE, position)?)?.commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_reply_add_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            editor.add_table_cell_comment_reply(model, COMMENT_ROW, COMMENT_COLUMN, "W4-T")?;
            Ok(())
        },
        |edit| edit.add_reply("W4-T"),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_reply_set_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.set_table_cell_comment_reply(
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                reply,
                "W4-T",
            )?;
            Ok(())
        },
        |edit| edit.set_reply(0, "W4-T"),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_reply_remove_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.remove_table_cell_comment_reply(model, COMMENT_ROW, COMMENT_COLUMN, reply)?;
            Ok(())
        },
        |edit| edit.remove_reply(0),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_thread_set_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W4-T")?;
            Ok(())
        },
        |edit| edit.set("W4-T"),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_thread_clear_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            editor.clear_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |edit| edit.clear(),
    )
}
