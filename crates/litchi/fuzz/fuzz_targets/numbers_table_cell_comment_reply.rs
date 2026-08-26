#![no_main]

//! Bounded selector-first Numbers direct-comment-reply lifecycle fuzzing.
//!
//! The input is offered to the checked package reader first. Every package
//! that survives bounded ingress is then probed with the same finite command
//! stream: reply reads by ordinal and A1 address, collection append/set/remove
//! edits, the direct add/set/remove conveniences, exact patch application,
//! conflict rejection, inverse restoration, and source-atomic failures. A
//! command corpus is intentionally not treated as a package fixture; a valid
//! reply-bearing package supplied by a fuzz input reaches the lifecycle path
//! without embedding native identifiers or an unverified native artifact.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::comment::{CommentReply, CommentReplyIndex, transaction::Error},
};

const MAX_INPUT_BYTES: u64 = 512 * 1024;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_OBJECTS: usize = 4 * 1024;
const MAX_SHEETS: usize = 128;
const MAX_TABLES: usize = 512;
const MAX_REFERENCES: usize = 8 * 1024;
const MAX_MATERIALIZED_CELLS: usize = 64 * 1024;
const MAX_TEXT_BYTES: usize = 512 * 1024;
const MAX_REPLY_TEXT_BYTES: usize = 4 * 1024;
const MAX_SCAN_ROWS: u32 = 8;
const MAX_SCAN_COLUMNS: u32 = 8;
const MAX_REPLY_ORDINALS: usize = 4;
const OVERSIZED_INPUT_BYTES: usize = 512 * 1024 + 1;
const PRIVATE_SHEET: &str = "__litchi_private_reply_sheet_93__";
const PRIVATE_TABLE: &str = "__litchi_private_reply_table_93__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_reply_input_93__";
const ZIP_LOCAL_HEADER: &[u8] = b"PK\x03\x04";

fuzz_target!(|data: &[u8]| {
    // A command byte keeps the checked-in corpus useful without embedding a
    // private native artifact.  Accept both `command || package` inputs and
    // an unprefixed ZIP supplied by a caller who has a valid reply-bearing
    // package.  The package reader always sees only the selected source
    // bytes, while the command stream chooses bounded selectors/text.
    let (command, package_input) = command_prefix(data);
    match Package::from_bytes_with_options(package_input, options()) {
        Ok(package) => exercise_package(&package, package_input, command),
        Err(error) => observe_error(error),
    }
    exercise_input_limit();
});

fn command_prefix(data: &[u8]) -> (u8, &[u8]) {
    if data.starts_with(ZIP_LOCAL_HEADER) {
        // Plain package input remains admissible; bytes after the ZIP magic
        // are only command entropy and never alter the source bytes.
        (data.get(4).copied().unwrap_or_default(), data)
    } else {
        data.split_first()
            .map(|(&command, input)| (command, input))
            .unwrap_or((0, data))
    }
}

fn options() -> PackageReadOptions {
    static OPTIONS: OnceLock<PackageReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = PackageLimits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid reply archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid reply semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid reply projection limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn exercise_package(package: &Package, data: &[u8], command: u8) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let position = position_from_bytes(data, command);
    let source_bytes = package_bytes(package);

    // Read-one and A1 calls are exercised even when no rooted comment exists;
    // stale ordinals must be typed errors and must never mutate the source.
    let stale = CommentReplyIndex::new(u32::MAX);
    observe_result(package.table_cell_comment_reply(sheet, table, position, stale));
    observe_result(package.table_cell_comment_reply_a1(sheet, table, "A1", stale));
    observe_result(package.table_cell_comment_replies(sheet, table, position));
    observe_result(package.table_cell_comment_replies_a1(sheet, table, "A1"));
    assert_eq!(package_bytes(package), source_bytes);

    exercise_selector_errors(package, position, stale);
    exercise_ingress_limits(&source_bytes);

    let Some((thread_position, replies)) = find_reply_thread(package, data, command) else {
        // A package can be valid Numbers while carrying no comment root. A
        // collection edit and every direct convenience must still fail
        // closed without changing its immutable source snapshot.
        exercise_missing_thread(package, position, stale, &source_bytes);
        return;
    };
    exercise_thread(package, thread_position, &replies, data);
}

fn find_reply_thread(
    package: &Package,
    data: &[u8],
    command: u8,
) -> Option<(CellPosition, Box<[CommentReply]>)> {
    let mut candidates = Vec::new();
    candidates.push(position_from_bytes(data, command));
    for row in 0..MAX_SCAN_ROWS {
        for column in 0..MAX_SCAN_COLUMNS {
            candidates.push(CellPosition::new(row, column));
        }
    }
    candidates.push(
        CellPosition::from_a1("A1")
            .unwrap_or_else(|error| unreachable!("A1 is a valid reply probe address: {error}")),
    );

    for position in candidates {
        match package.table_cell_comment_replies(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        ) {
            Ok(replies) => return Some((position, replies)),
            Err(error) => {
                black_box(error.to_string());
            },
        }
    }
    None
}

fn exercise_missing_thread(
    package: &Package,
    position: CellPosition,
    stale: CommentReplyIndex,
    source_bytes: &[u8],
) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    if let Ok(edit) = package.edit_table_cell_comment_replies(sheet, table, position) {
        let result = edit.append("reply without a rooted comment").commit();
        observe_atomic_failure(package, source_bytes, result);
    }

    observe_atomic_failure(
        package,
        source_bytes,
        package.add_table_cell_comment_reply(sheet, table, position, "reply without a root"),
    );
    observe_atomic_failure(
        package,
        source_bytes,
        package.add_table_cell_comment_reply_a1(sheet, table, "A1", "reply without a root"),
    );
    observe_atomic_failure(
        package,
        source_bytes,
        package.set_table_cell_comment_reply(sheet, table, position, stale, "stale reply ordinal"),
    );
    observe_atomic_failure(
        package,
        source_bytes,
        package.set_table_cell_comment_reply_a1(sheet, table, "A1", stale, "stale reply ordinal"),
    );
    observe_atomic_failure(
        package,
        source_bytes,
        package.remove_table_cell_comment_reply(sheet, table, position, stale),
    );
    observe_atomic_failure(
        package,
        source_bytes,
        package.remove_table_cell_comment_reply_a1(sheet, table, "A1", stale),
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_thread(
    package: &Package,
    position: CellPosition,
    replies: &[CommentReply],
    data: &[u8],
) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let source_bytes = package_bytes(package);
    let ordinal = reply_ordinal(data, replies.len());
    let index = CommentReplyIndex::new(ordinal);

    // Read every selected ordinal independently. Equal text in adjacent
    // replies is intentional: the selector is an ordinal, never text-based.
    for ordinal in 0..replies.len().min(MAX_REPLY_ORDINALS) {
        let index = CommentReplyIndex::new(
            u32::try_from(ordinal)
                .unwrap_or_else(|error| unreachable!("bounded reply ordinal: {error}")),
        );
        observe_result(package.table_cell_comment_reply(sheet, table, position, index));
    }
    observe_result(package.table_cell_comment_reply_a1(sheet, table, "A1", index));

    let duplicate_text = replies
        .first()
        .map(|reply| reply.text().to_owned())
        .unwrap_or_else(|| reply_text(data, 0));
    let replacement = if replies.len() > 1 && data.get(1).copied().unwrap_or_default() & 1 != 0 {
        duplicate_text.clone()
    } else {
        reply_text(data, 1)
    };

    // Collection edit: append, ordinal set (including duplicate-text ordinal),
    // and remove are all dispatched from the same source snapshot.
    if let Ok(edit) = package.edit_table_cell_comment_replies(sheet, table, position) {
        match data.first().copied().unwrap_or_default() % 3 {
            0 => verify_edit_commit(
                package,
                &source_bytes,
                edit.append(replacement.clone()).commit(),
            ),
            1 => verify_edit_commit(
                package,
                &source_bytes,
                edit.set(index, replacement.clone()).commit(),
            ),
            _ => verify_edit_commit(package, &source_bytes, edit.remove(index).commit()),
        }
    }
    if let Ok(edit) = package.edit_table_cell_comment_replies_a1(sheet, table, "A1") {
        verify_edit_commit(
            package,
            &source_bytes,
            edit.set(index, replacement.clone()).commit(),
        );
    }

    // Direct convenience methods exercise the same transaction engine while
    // also covering their A1 selector variants.
    verify_edit_commit(
        package,
        &source_bytes,
        package.add_table_cell_comment_reply(sheet, table, position, replacement.clone()),
    );
    verify_edit_commit(
        package,
        &source_bytes,
        package.add_table_cell_comment_reply_a1(sheet, table, "A1", replacement.clone()),
    );
    verify_edit_commit(
        package,
        &source_bytes,
        package.set_table_cell_comment_reply(sheet, table, position, index, replacement.clone()),
    );
    verify_edit_commit(
        package,
        &source_bytes,
        package.set_table_cell_comment_reply_a1(sheet, table, "A1", index, replacement),
    );
    verify_edit_commit(
        package,
        &source_bytes,
        package.remove_table_cell_comment_reply(sheet, table, position, index),
    );
    verify_edit_commit(
        package,
        &source_bytes,
        package.remove_table_cell_comment_reply_a1(sheet, table, "A1", index),
    );

    // A deliberately invalid ordinal is a bounded atomic failure even when
    // the package has a valid thread.
    exercise_missing_thread(
        package,
        position,
        CommentReplyIndex::new(u32::MAX),
        &source_bytes,
    );
}

fn verify_edit_commit(
    package: &Package,
    source_bytes: &[u8],
    result: Result<litchi::numbers::cell::comment::transaction::Commit, Error>,
) {
    match result {
        Ok(commit) => {
            let patch = commit.patch().clone();
            let target_bytes = package_bytes(commit.package());
            assert_eq!(patch.is_noop(), target_bytes == source_bytes);
            black_box((
                patch.path(),
                patch.source_fingerprint(),
                patch.target_fingerprint(),
                commit.diagnostics(),
            ));

            let applied = package
                .apply_table_cell_comment_reply(&patch)
                .unwrap_or_else(|error| panic!("reply patch must apply: {error}"));
            assert_eq!(package_bytes(applied.package()), target_bytes);
            if !patch.is_noop() {
                let conflict = package.apply_table_cell_comment_reply(&patch);
                assert!(conflict.is_err(), "reply patch conflict was accepted");
            }

            let inverse = patch.inverse();
            assert_eq!(inverse.inverse(), patch);
            let restored = commit
                .package()
                .apply_table_cell_comment_reply(&inverse)
                .unwrap_or_else(|error| panic!("reply inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);

            // The serialized candidate is independently reopened before the
            // exact forward/inverse replay, so malformed publication cannot
            // hide behind an in-memory snapshot.
            let reopened_source = Package::from_bytes_with_options(source_bytes, options())
                .unwrap_or_else(|error| panic!("reply source reopen failed: {error}"));
            let reopened_target = Package::from_bytes_with_options(&target_bytes, options())
                .unwrap_or_else(|error| panic!("reply candidate reopen failed: {error}"));
            let replayed = reopened_source
                .apply_table_cell_comment_reply(&patch)
                .unwrap_or_else(|error| panic!("reopened reply patch failed: {error}"));
            assert_eq!(package_bytes(replayed.package()), target_bytes);
            let reopened_restored = reopened_target
                .apply_table_cell_comment_reply(&inverse)
                .unwrap_or_else(|error| panic!("reopened reply inverse failed: {error}"));
            assert_eq!(package_bytes(reopened_restored.package()), source_bytes);
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
        },
    }
}

fn observe_atomic_failure(
    package: &Package,
    source_bytes: &[u8],
    result: Result<litchi::numbers::cell::comment::transaction::Commit, Error>,
) {
    match result {
        Ok(commit) => verify_edit_commit(package, source_bytes, Ok(commit)),
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
        },
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition, index: CommentReplyIndex) {
    if let Err(error) = package.table_cell_comment_reply(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
        index,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
        index,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) = package.table_cell_comment_reply_a1(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        "A0",
        index,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_INPUT, options()) {
        observe_redacted(
            error,
            std::str::from_utf8(PRIVATE_INPUT).unwrap_or("reply-input"),
        );
    }
}

fn exercise_ingress_limits(source_bytes: &[u8]) {
    let Ok(source_len) = u64::try_from(source_bytes.len()) else {
        return;
    };
    if source_len == 0 || source_len > MAX_INPUT_BYTES {
        return;
    }
    let exact_archive = PackageLimits::new(
        source_len,
        MAX_ENTRIES,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    )
    .unwrap_or_else(|error| panic!("exact reply limits invalid: {error}"));
    let exact = Package::from_bytes_with_options(
        source_bytes,
        PackageReadOptions::new(exact_archive, PackageSemanticLimits::default()),
    );
    if let Err(error) = exact {
        observe_error(error);
        return;
    }
    if source_len > 1 {
        let tight_archive = PackageLimits::new(
            source_len - 1,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("tight reply limits invalid: {error}"));
        let tight = Package::from_bytes_with_options(
            source_bytes,
            PackageReadOptions::new(tight_archive, PackageSemanticLimits::default()),
        );
        assert!(tight.is_err(), "reply input-minus-one was admitted");
        black_box(tight.err().map(|error| error.to_string()));
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    let result = Package::from_bytes_with_options(bytes, options());
    assert!(result.is_err(), "oversized reply input was admitted");
    black_box(result.err().map(|error| error.to_string()));
}

fn position_from_bytes(data: &[u8], command: u8) -> CellPosition {
    CellPosition::new(
        u32::from(command) % MAX_SCAN_ROWS,
        u32::from(data.get(1).copied().unwrap_or_default()) % MAX_SCAN_COLUMNS,
    )
}

fn reply_ordinal(data: &[u8], reply_count: usize) -> u32 {
    if reply_count == 0 {
        return 0;
    }
    u32::try_from(usize::from(data.get(2).copied().unwrap_or_default()) % reply_count)
        .unwrap_or_else(|error| unreachable!("bounded reply ordinal: {error}"))
}

fn reply_text(data: &[u8], offset: usize) -> String {
    let seed = data.get(offset).copied().unwrap_or_default();
    let width = (usize::from(data.get(offset + 1).copied().unwrap_or_default()) % 24) + 1;
    let mut text = String::new();
    text.try_reserve(width + 24)
        .unwrap_or_else(|error| unreachable!("bounded reply text allocation: {error}"));
    text.push_str("Wave93 reply ");
    for index in 0..width {
        let byte = data
            .get(offset.saturating_add(2).saturating_add(index))
            .copied()
            .unwrap_or(seed);
        text.push(char::from(b'a' + byte % 26));
    }
    debug_assert!(text.len() <= MAX_REPLY_TEXT_BYTES);
    text
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("reply package write failed: {error}"));
    bytes
}

fn observe_result<T, E>(result: Result<T, E>)
where
    T: Debug,
    E: Debug + Display,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => {
            observe_error(error);
        },
    }
}

fn observe_error<E>(error: E)
where
    E: Debug + Display,
{
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted<E>(error: E, private: &str)
where
    E: Debug + Display,
{
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
