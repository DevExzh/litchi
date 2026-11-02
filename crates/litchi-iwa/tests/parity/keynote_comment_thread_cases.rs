//! Keynote slide-table comment threads (record 0412c): set and clear of a
//! reply-bearing root, and the direct-reply lifecycle (append, replace,
//! remove) through `litchi_iwa_table::comments_native::rewrite_reply`.
//!
//! Every case starts from a host-authored two-reply thread.  The host's
//! reply writer is copy-on-write and stamps a fresh date and random storage
//! UUID on a new reply, so the write cases are differential (Tier B): both
//! outputs reopen, the host and leaf thread reads agree, and the leaf
//! inverse restores the exact source.

use litchi_iwa::keynote;
use litchi_iwa_common::table::coordinate::At;
use litchi_keynote::slide::table::comments::{Edit, Error};
use litchi_keynote::{Package, SlideSelector, TableSelector};

use crate::harness::{Evidence, Fixture, LeafRun, TestResult, Tier};
use crate::keynote_table_cases::{
    COMMENT_COLUMN, COMMENT_ROW, SLIDE, comment_case, host_model, populated_comment_fixture,
};
use crate::leaf_bytes;

/// The host-authored root with two replies.
fn source() -> TestResult<Fixture> {
    let base = populated_comment_fixture(true)?;
    let mut editor = keynote::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&editor)?;
    editor.add_slide_table_cell_comment_reply(
        SLIDE,
        model,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "W4-T second reply",
    )?;
    Ok(Fixture {
        path: "generated:keynote-table-comment-root-and-two-replies",
        bytes: editor.to_bytes()?,
    })
}

fn comment_position() -> TestResult<At> {
    Ok(At::new(
        u32::try_from(COMMENT_ROW)?,
        u32::try_from(COMMENT_COLUMN)?,
    ))
}

fn edit(package: &Package, position: At) -> Result<Edit<'_>, Error> {
    package.edit_slide_table_cell_comment(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        position,
    )
}

fn read(
    package: &Package,
    position: At,
) -> Result<Option<litchi_keynote::slide::table::comments::Comment>, Error> {
    package.slide_table_cell_comment(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        position,
    )
}

/// The host storage identifier of the source reply at `index`.
fn host_reply_id(editor: &keynote::Draft, model: u64, index: usize) -> TestResult<u64> {
    Ok(editor
        .slide_table_cell_comment_replies(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?
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
    let commit = stage(edit(package, position)?)?.commit()?;
    let restored = commit
        .package()
        .apply_slide_table_cell_comment(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn thread_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft, u64) -> TestResult + 'static,
    stage: impl for<'a> FnOnce(Edit<'a>) -> Result<Edit<'a>, Error> + 'static,
) -> TestResult {
    let input = source()?;
    comment_case(
        name,
        tier,
        evidence,
        &input,
        true,
        host,
        move |package, position| leaf_run(&package, position, stage),
    )
}

#[test]
fn keynote_slide_table_comment_set_with_replies_parity() -> TestResult {
    thread_case(
        "keynote_slide_table_comment_set_with_replies_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.set_slide_table_cell_comment(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W4-T root",
            )?;
            Ok(())
        },
        |edit| edit.set("W4-T root"),
    )
}

#[test]
fn keynote_slide_table_comment_clear_with_replies_parity() -> TestResult {
    thread_case(
        "keynote_slide_table_comment_clear_with_replies_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.clear_slide_table_cell_comment(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |edit| edit.clear(),
    )
}

#[test]
fn keynote_slide_table_comment_reply_add_parity() -> TestResult {
    thread_case(
        "keynote_slide_table_comment_reply_add_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            editor.add_slide_table_cell_comment_reply(
                SLIDE,
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
fn keynote_slide_table_comment_reply_set_parity() -> TestResult {
    thread_case(
        "keynote_slide_table_comment_reply_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.set_slide_table_cell_comment_reply(
                SLIDE,
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
fn keynote_slide_table_comment_reply_remove_parity() -> TestResult {
    thread_case(
        "keynote_slide_table_comment_reply_remove_parity",
        Tier::B,
        Evidence::Queued,
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.remove_slide_table_cell_comment_reply(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                reply,
            )?;
            Ok(())
        },
        |edit| edit.remove_reply(0),
    )
}

/// A reply ordinal past the source thread is a typed refusal, and an
/// unchanged reply text is a no-op commit.
#[test]
fn keynote_slide_table_comment_reply_refusals_are_typed() -> TestResult {
    let input = source()?;
    let package = Package::from_bytes(&input.bytes)?;
    let position = comment_position()?;
    let replies = read(&package, position)?
        .ok_or("the source has no root comment")?
        .replies()
        .map_or(0, <[_]>::len);
    let refused = edit(&package, position)?.remove_reply(replies)?.commit();
    assert!(
        matches!(refused, Err(Error::ReplyNotFound { index }) if index == replies),
        "an out-of-range reply ordinal must be refused: {refused:?}"
    );
    let text = read(&package, position)?
        .and_then(|comment| {
            comment
                .replies()
                .and_then(|replies| replies.first().cloned())
        })
        .ok_or("the source has no reply")?;
    let unchanged = edit(&package, position)?
        .set_reply(0, text.text())?
        .commit()?;
    assert!(
        unchanged.patch().is_noop(),
        "rewriting a reply with its own text must be a no-op"
    );
    Ok(())
}

fn thread_probe(
    host: impl Fn(&mut keynote::Draft, u64) -> TestResult,
    stage: impl for<'a> Fn(Edit<'a>) -> Result<Edit<'a>, Error>,
) -> TestResult {
    let input = source()?;
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |bytes| {
            let mut editor = keynote::Draft::from_bytes(bytes)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            Ok(editor.to_bytes()?)
        },
        |bytes| {
            let package = Package::from_bytes(bytes)?;
            let commit = stage(edit(&package, position)?)?.commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_reply_add_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            editor.add_slide_table_cell_comment_reply(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W4-T",
            )?;
            Ok(())
        },
        |edit| edit.add_reply("W4-T"),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_reply_set_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.set_slide_table_cell_comment_reply(
                SLIDE,
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
fn keynote_slide_table_comment_reply_remove_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            let reply = host_reply_id(editor, model, 0)?;
            editor.remove_slide_table_cell_comment_reply(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                reply,
            )?;
            Ok(())
        },
        |edit| edit.remove_reply(0),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_thread_set_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            editor.set_slide_table_cell_comment(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W4-T",
            )?;
            Ok(())
        },
        |edit| edit.set("W4-T"),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_thread_clear_cost_abba() -> TestResult {
    thread_probe(
        |editor, model| {
            editor.clear_slide_table_cell_comment(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |edit| edit.clear(),
    )
}
