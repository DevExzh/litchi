//! Numbers table-cell comment threads (record 0412c): replacement of a
//! reply-bearing root and the direct-reply lifecycle (append, replace,
//! remove).
//!
//! The host's `Draft` reply methods delegate to the focused leaf owner, so
//! the oracle is the legacy package writer behind the hidden
//! `__parity_*_legacy` seams.  The host stamps a fresh date and random
//! storage UUID on a new reply, so the cases are differential (Tier B) with
//! those two new-object fields masked.

#![cfg(feature = "parity-host-oracle")]

use litchi_iwa::numbers;
use litchi_numbers::cell::comment::CommentReplyIndex;
use litchi_numbers::cell::comment::transaction::Commit;
use litchi_numbers::{Package, SheetSelector, TableSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;
use crate::numbers_table_comment_cases::{
    host_model, masked_new_comment_storage, native_fixture, position,
};

/// The native commented fixture with one reply added by the legacy host
/// writer: the builder-made `populated_comment_fixture` is not an exact
/// source the leaf transactions admit.
fn populated_comment_fixture() -> TestResult<Fixture> {
    let base = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&base.bytes)?;
    editor.__parity_add_table_cell_comment_reply_legacy(
        TableSelector::index(0),
        position().row() as usize,
        position().column() as usize,
        "W4-T source reply",
    )?;
    Ok(Fixture {
        path: "generated:numbers-native-comment-root-and-reply",
        bytes: editor.to_bytes()?,
    })
}

fn host_thread(bytes: &[u8]) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let root = editor.__parity_read_table_cell_comment(TableSelector::index(0), position())?;
    let model = host_model(&editor)?;
    let replies = if root.is_some() {
        editor.cell_comment_replies(
            model,
            position().row() as usize,
            position().column() as usize,
        )?
    } else {
        Vec::new()
    };
    Ok(format!(
        "root={:?};replies={:?}",
        root.as_ref().map(|value| value.comment.text.as_str()),
        replies
            .iter()
            .map(|reply| reply.comment.text.as_str())
            .collect::<Vec<_>>()
    ))
}

fn leaf_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let root =
        package.table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?;
    let replies = if root.is_some() {
        package
            .table_cell_comment_replies(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
            )?
            .to_vec()
    } else {
        Vec::new()
    };
    Ok(format!(
        "root={:?};replies={:?}",
        root.as_ref().map(|value| value.text()),
        replies.iter().map(|reply| reply.text()).collect::<Vec<_>>()
    ))
}

fn host_reply_id(editor: &numbers::Draft, index: usize) -> TestResult<u64> {
    let model = host_model(editor)?;
    Ok(editor
        .cell_comment_replies(
            model,
            position().row() as usize,
            position().column() as usize,
        )?
        .get(index)
        .ok_or("the source thread has no reply at the requested ordinal")?
        .storage_id
        .get())
}

fn reply_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut numbers::Draft) -> TestResult + 'static,
    leaf: impl FnOnce(&Package) -> TestResult<Commit> + 'static,
) -> TestResult {
    let input = populated_comment_fixture()?;
    Case::new(name, Format::Numbers, tier, evidence, &input)
        .host(move |bytes| {
            let mut editor = numbers::Draft::from_bytes(bytes)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |bytes| {
            let package = Package::from_bytes(bytes)?;
            let commit = leaf(&package)?;
            let restored = commit
                .package()
                .apply_table_cell_comment_reply(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table-cell comment thread", host_thread)
        .read("leaf table-cell comment thread", leaf_thread)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (new-object date and storage UUID masked)",
                masked_new_comment_storage,
            )
        })
        .run()?;
    Ok(())
}

#[test]
fn numbers_table_cell_comment_set_with_replies_parity() -> TestResult {
    let input = populated_comment_fixture()?;
    Case::new(
        "numbers_table_cell_comment_set_with_replies_parity",
        Format::Numbers,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|bytes| {
        let mut editor = numbers::Draft::from_bytes(bytes)?;
        editor.__parity_set_table_cell_comment_legacy(
            TableSelector::index(0),
            position().row() as usize,
            position().column() as usize,
            "W4-T root",
        )?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|bytes| {
        let package = Package::from_bytes(bytes)?;
        let commit = package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W4-T root",
        )?;
        let restored = commit
            .package()
            .apply_table_cell_comment(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host table-cell comment thread", host_thread)
    .read("leaf table-cell comment thread", leaf_thread)
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .run()?;
    Ok(())
}

#[test]
fn numbers_table_cell_comment_reply_add_parity() -> TestResult {
    reply_case(
        "numbers_table_cell_comment_reply_add_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.__parity_add_table_cell_comment_reply_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                "W4-T reply",
            )?;
            Ok(())
        },
        |package| {
            Ok(package.add_table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                "W4-T reply",
            )?)
        },
    )
}

#[test]
fn numbers_table_cell_comment_reply_set_parity() -> TestResult {
    reply_case(
        "numbers_table_cell_comment_reply_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let reply = host_reply_id(editor, 0)?;
            editor.__parity_set_table_cell_comment_reply_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                reply,
                "W4-T edited reply",
            )?;
            Ok(())
        },
        |package| {
            Ok(package.set_table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                CommentReplyIndex::new(0),
                "W4-T edited reply",
            )?)
        },
    )
}

#[test]
fn numbers_table_cell_comment_reply_remove_parity() -> TestResult {
    reply_case(
        "numbers_table_cell_comment_reply_remove_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let reply = host_reply_id(editor, 0)?;
            editor.__parity_remove_table_cell_comment_reply_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                reply,
            )?;
            Ok(())
        },
        |package| {
            Ok(package.remove_table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                CommentReplyIndex::new(0),
            )?)
        },
    )
}

fn thread_probe(
    host: impl Fn(&mut numbers::Draft) -> TestResult,
    leaf: impl Fn(&Package) -> TestResult<Vec<u8>>,
) -> TestResult {
    let input = populated_comment_fixture()?;
    crate::pages_cost_probes::abba(
        &input,
        |bytes| {
            let mut editor = numbers::Draft::from_bytes(bytes)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        },
        |bytes| leaf(&Package::from_bytes(bytes)?),
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_reply_add_cost_abba() -> TestResult {
    thread_probe(
        |editor| {
            editor.__parity_add_table_cell_comment_reply_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                "W4-T",
            )?;
            Ok(())
        },
        |package| {
            let commit = package.add_table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                "W4-T",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_reply_set_cost_abba() -> TestResult {
    thread_probe(
        |editor| {
            let reply = host_reply_id(editor, 0)?;
            editor.__parity_set_table_cell_comment_reply_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                reply,
                "W4-T",
            )?;
            Ok(())
        },
        |package| {
            let commit = package.set_table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                CommentReplyIndex::new(0),
                "W4-T",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_reply_remove_cost_abba() -> TestResult {
    thread_probe(
        |editor| {
            let reply = host_reply_id(editor, 0)?;
            editor.__parity_remove_table_cell_comment_reply_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                reply,
            )?;
            Ok(())
        },
        |package| {
            let commit = package.remove_table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                CommentReplyIndex::new(0),
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_thread_set_cost_abba() -> TestResult {
    thread_probe(
        |editor| {
            editor.__parity_set_table_cell_comment_legacy(
                TableSelector::index(0),
                position().row() as usize,
                position().column() as usize,
                "W4-T",
            )?;
            Ok(())
        },
        |package| {
            let commit = package.set_table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                "W4-T",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
