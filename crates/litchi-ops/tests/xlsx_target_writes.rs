//! Unified XLSX range and selector writes through the public target API.
//!
//! The target adapter resolves a range or selector against the input
//! snapshot, then commits the complete target set in one native operation.
//! These tests keep the original owner available so source and eager routes
//! exercise the same immutable contract.
#![cfg(feature = "xlsx")]

use std::io::{Cursor, Read, Write};

use litchi_core::FileFormat;
use litchi_ops::capabilities::Capabilities;
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, Props, Result, SourceMode, props,
};

const A1: &str = "/sheet[1]/A1";
const B1: &str = "/sheet[1]/B1";
const C1: &str = "/sheet[1]/C1";
const A2: &str = "/sheet[1]/A2";
const B2: &str = "/sheet[1]/B2";
const C2: &str = "/sheet[1]/C2";

fn fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("look up native worksheet")
            .expect("minimal workbook has Sheet1");
        sheet.set("A1", "seed").expect("seed text cell");
        sheet.set("A2", "seed").expect("seed second text cell");
        sheet.set("B1", 1_i32).expect("seed first number");
        sheet.set("C1", 2_i32).expect("seed second number");
        sheet.set("B2", 3_i32).expect("seed third number");
        sheet.set("C2", 4_i32).expect("seed fourth number");
    }
    edit.commit()
        .expect("commit native fixture")
        .into_package()
        .to_bytes()
        .expect("serialize native fixture")
}

fn multi_sheet_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("look up first worksheet")
            .expect("minimal workbook has Sheet1");
        sheet.set("A1", 1_i32).expect("seed first sheet");
    }
    {
        let mut sheet = edit.add("Sheet2").expect("add second worksheet");
        sheet.set("A1", 2_i32).expect("seed second sheet");
    }
    edit.commit()
        .expect("commit multi-sheet fixture")
        .into_package()
        .to_bytes()
        .expect("serialize multi-sheet fixture")
}

fn corrupt_deferred_worksheet_fixture() -> Vec<u8> {
    let mut archive =
        zip::ZipArchive::new(Cursor::new(fixture())).expect("open XLSX fixture archive");
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let (name, contents) = {
            let mut entry = archive.by_index(index).expect("read XLSX fixture member");
            let name = entry.name().to_owned();
            let mut contents = Vec::new();
            entry
                .read_to_end(&mut contents)
                .expect("read XLSX fixture member bytes");
            (name, contents)
        };
        writer
            .start_file(name.as_str(), options)
            .expect("write rewritten XLSX member");
        if name == "xl/worksheets/sheet1.xml" {
            writer
                .write_all(b"<worksheet")
                .expect("write deferred malformed worksheet");
        } else {
            writer
                .write_all(&contents)
                .expect("copy XLSX fixture member");
        }
    }
    writer
        .finish()
        .expect("finish rewritten XLSX archive")
        .into_inner()
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    let mode = if eager {
        SourceMode::Eager
    } else {
        SourceMode::Auto
    };
    litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode))
        .expect("open XLSX owner")
}

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .text()
        .expect("read cell text")
        .map(std::borrow::Cow::into_owned)
}

fn cell_formula(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .format()
        .expect("read cell properties")
        .find(|property| property.key() == "formula")
        .map(|property| property.value().to_string())
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

fn assert_unsupported_target(error: litchi_ops::Error) {
    assert!(
        matches!(
            error.code(),
            Code::InvalidSelector | Code::UnsupportedElement | Code::UnsupportedCapability
        ),
        "expected a typed target refusal, got {error:?}"
    );
}

#[test]
fn range_writes_cover_existing_and_missing_cells_for_source_and_eager_owners() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        let valued = original
            .set("Sheet1!A1:C2", props! { "value" => "range" })
            .expect("set a two-row range");
        for path in [A1, B1, C1, A2, B2, C2] {
            assert_eq!(
                cell_text(&valued, path).as_deref(),
                Some("range"),
                "range write missed {path}"
            );
        }

        let formula = valued
            .set("Sheet1!A1:C2", props! { "formula" => "1+1" })
            .expect("set a formula across the range");
        for path in [A1, B1, C1, A2, B2, C2] {
            assert_eq!(
                cell_formula(&formula, path).as_deref(),
                Some("1+1"),
                "range formula write missed {path}"
            );
        }

        let cleared = formula
            .set("Sheet1!A1:C2", props! { "clear" => true })
            .expect("clear the complete range");
        for path in [A1, B1, C1, A2, B2, C2] {
            assert_eq!(cell_text(&cleared, path).as_deref(), Some(""));
            assert_eq!(cell_formula(&cleared, path), None);
        }

        assert_eq!(
            original.to_bytes().expect("publish original owner again"),
            original_bytes,
            "range writes mutated the detached {eager:?} owner"
        );
    }
}

#[test]
fn selector_writes_use_the_original_match_set_and_one_transaction_commit() {
    let bytes = fixture();
    let selector = "Sheet1!cell[value>0]";

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let old_matches = original
            .query(selector)
            .expect("compile scoped cell selector")
            .collect::<Result<Vec<_>>>()
            .expect("walk the original selector");
        assert_eq!(old_matches.len(), 4, "the fixture has four positive cells");

        let commit = original
            .edit(|transaction| {
                transaction.set(selector, props! { "value" => 0 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("set all original selector matches");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.diagnostics().commits(), 1);

        for path in [B1, C1, B2, C2] {
            assert_eq!(cell_text(commit.file(), path).as_deref(), Some("0"));
        }
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("seed"));
        let remaining = commit
            .file()
            .query(selector)
            .expect("query the committed selector")
            .collect::<Result<Vec<_>>>()
            .expect("walk the committed selector");
        assert!(remaining.is_empty(), "the selector must use old matches");

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay selector patch on its owner");
        assert_eq!(
            replayed.to_bytes().expect("publish replayed selector"),
            commit
                .file()
                .to_bytes()
                .expect("publish committed selector")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse selector patch");
        assert_eq!(
            restored.to_bytes().expect("publish restored selector"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("publish detached original"),
            original_bytes
        );
    }
}

#[test]
fn atomic_range_batch_rolls_back_before_an_unsupported_mixed_target() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let error = original
            .batch(
                [
                    batch_set("Sheet1!A1:A2", Props::new().with("value", "staged")),
                    batch_set(
                        "Sheet1!cell[value>0], row[1]",
                        Props::new().with("value", "must-not-apply"),
                    ),
                    batch_set("Sheet1!C1", Props::new().with("value", "suffix")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("mixed cell/row targets must refuse atomically");
        assert_eq!(error.item(), Some(1));
        assert_unsupported_target(error);
        assert_eq!(
            original.to_bytes().expect("publish atomic input"),
            original_bytes,
            "atomic target refusal published a partial range"
        );
        assert_eq!(cell_text(&original, A1).as_deref(), Some("seed"));
    }
}

#[test]
fn exact_range_noops_reuse_the_input_bytes_and_zero_matches_are_typed() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        let exact = original
            .set("Sheet1!A1:A2", props! { "value" => "seed" })
            .expect("repeat an exact two-cell range value");
        assert_eq!(
            exact.to_bytes().expect("publish exact target no-op"),
            original_bytes
        );
        let blank_clear = original
            .set("Sheet1!D4:E5", props! { "clear" => true })
            .expect("clear an entirely blank rectangle");
        assert_eq!(
            blank_clear.to_bytes().expect("publish blank clear no-op"),
            original_bytes
        );

        assert_eq!(
            original
                .set("Sheet1!cell[value>999]", props! { "value" => "no-match" })
                .expect_err("zero-match selector must be refused")
                .code(),
            Code::ZeroMatches
        );
        assert_eq!(
            original
                .to_bytes()
                .expect("publish after zero-match refusal"),
            original_bytes
        );
    }
}

#[test]
fn scoped_multi_sheet_selector_keeps_same_addresses_distinct_in_one_commit() {
    let bytes = multi_sheet_fixture();
    let selector = litchi_core::selector::Query::parse("cell[value>0]")
        .unwrap()
        .with_scope(litchi_core::selector::Scope::Path(Box::new(
            litchi_core::selector::Path::parse("/").unwrap(),
        )));

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let matches = original
            .query(selector.clone())
            .expect("compile workbook-scoped selector")
            .collect::<Result<Vec<_>>>()
            .expect("walk both worksheets");
        assert_eq!(matches.len(), 2, "both sheets own a positive A1");

        let commit = original
            .edit(|transaction| {
                transaction.set(selector.clone(), props! { "value" => 0 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("mutate both scoped matches in one transaction");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            cell_text(commit.file(), "/sheet[1]/A1").as_deref(),
            Some("0")
        );
        assert_eq!(
            cell_text(commit.file(), "/sheet[2]/A1").as_deref(),
            Some("0")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay multi-sheet patch");
        assert_eq!(
            replayed
                .to_bytes()
                .expect("publish replayed multi-sheet patch"),
            commit
                .file()
                .to_bytes()
                .expect("publish committed multi-sheet file")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse multi-sheet patch");
        assert_eq!(
            restored
                .to_bytes()
                .expect("publish restored multi-sheet file"),
            original_bytes
        );
        assert_eq!(
            original
                .to_bytes()
                .expect("publish detached multi-sheet owner"),
            original_bytes
        );
    }
}

#[test]
fn range_expansion_limit_and_unknown_predicates_fail_before_mutation() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        // The leaf expands finite ranges before opening the native editor.
        // This is one cell beyond the 65,536-cell admission ceiling.
        let limit_error = original
            .set("Sheet1!A1:A65537", props! { "value" => "too-many" })
            .expect_err("the bounded range target must refuse oversized expansion");
        assert_eq!(limit_error.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("publish after limit refusal"),
            original_bytes
        );

        let invalid = original
            .set(
                "Sheet1!cell[unknownPrivatePredicate=1]",
                props! { "value" => "must-not-apply" },
            )
            .expect_err("unknown mutation predicates must be refused");
        assert_eq!(invalid.code(), Code::InvalidSelector);
        assert_eq!(
            original.to_bytes().expect("publish after selector refusal"),
            original_bytes
        );
    }
}

#[test]
fn custom_expanded_target_limits_reject_before_mutation_and_allow_the_bound() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let strict = SemanticLimits::builder()
            .max_expanded_targets(2)
            .expect("two targets is a valid limit")
            .build()
            .expect("build target limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(strict),
                |transaction| transaction.set("Sheet1!A1:A3", props! { "value" => "bounded" }),
            )
            .expect_err("three targets must exceed a two-target limit");
        assert_eq!(error.item(), Some(0));
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: 3,
                maximum: 2,
            })
        ));
        assert_eq!(
            original.to_bytes().expect("publish after target refusal"),
            original_bytes
        );

        let raised = SemanticLimits::builder()
            .max_expanded_targets(3)
            .expect("three targets is a valid limit")
            .build()
            .expect("build raised target limit");
        let commit = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(raised),
                |transaction| transaction.set("Sheet1!A1:A3", props! { "value" => "bounded" }),
            )
            .expect("the configured bound admits three targets");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        for path in [A1, A2, "/sheet[1]/A3"] {
            assert_eq!(cell_text(commit.file(), path).as_deref(), Some("bounded"));
        }
        assert_eq!(
            original.to_bytes().expect("publish detached target input"),
            original_bytes
        );
    }
}

#[test]
fn staged_bytes_are_checked_before_a_huge_range_can_reserve_targets() {
    let bytes = fixture();
    let limits = SemanticLimits::builder()
        .max_expanded_targets(isize::MAX as usize)
        .expect("the largest supported target bound is valid")
        .max_staged_bytes(64)
        .expect("64 staged bytes is a valid limit")
        .build()
        .expect("build huge-range preflight limits");

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set("Sheet1!A1:XFD1048576", Props::new().with("value", "x"))
                },
            )
            .expect_err("staged bytes must fail before huge target reservation");
        assert_eq!(error.item(), Some(0));
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed,
                maximum: 64,
            }) if *observed > 64
        ));
        assert_eq!(
            original
                .to_bytes()
                .expect("publish unchanged huge-range owner"),
            original_bytes
        );
    }
}

#[test]
fn default_target_limit_precedes_a_deferred_corrupt_worksheet_read() {
    let bytes = corrupt_deferred_worksheet_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let error = original
            .set("Sheet1!A1:XFD1048576", props! { "value" => "x" })
            .expect_err("the target limit must fire before worksheet parsing");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed,
                maximum: 65_536,
            }) if *observed > 65_536
        ));
    }
}

#[test]
fn duplicate_selector_matches_are_deduplicated_before_the_target_limit() {
    let bytes = fixture();
    let selector = "Sheet1!cell[value>0], cell[value>0]";

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let limits = SemanticLimits::builder()
            .max_expanded_targets(4)
            .expect("four targets is a valid limit")
            .build()
            .expect("build duplicate-match limit");
        let commit = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| transaction.set(selector, props! { "value" => 0 }),
            )
            .expect("duplicate matches must count once");
        assert_eq!(commit.diagnostics().operations(), 1);
        for path in [B1, C1, B2, C2] {
            assert_eq!(cell_text(commit.file(), path).as_deref(), Some("0"));
        }
        assert_eq!(
            original.to_bytes().expect("publish duplicate-match input"),
            original_bytes
        );
    }
}

#[test]
fn custom_staged_byte_limits_refuse_before_native_staging_and_can_be_raised() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let strict = SemanticLimits::builder()
            .max_staged_bytes(64)
            .expect("64 staged bytes is a valid limit")
            .build()
            .expect("build staged-byte limit");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(strict),
                |transaction| transaction.set(A1, Props::new().with("value", "text")),
            )
            .expect_err("a nonempty scalar must exceed the 64-byte entry charge");
        assert_eq!(error.item(), Some(0));
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed,
                maximum: 64,
            }) if *observed > 64
        ));
        assert_eq!(
            original
                .to_bytes()
                .expect("publish after staged-byte refusal"),
            original_bytes
        );

        let raised = SemanticLimits::builder()
            .max_staged_bytes(1_000)
            .expect("1,000 staged bytes is a valid limit")
            .build()
            .expect("build raised staged-byte limit");
        let commit = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(raised),
                |transaction| transaction.set(A1, Props::new().with("value", "text")),
            )
            .expect("the raised staged-byte limit admits the scalar");
        assert_eq!(cell_text(commit.file(), A1).as_deref(), Some("text"));
    }
}

#[test]
fn large_repeated_text_range_hits_staged_bytes_before_cloning_a_candidate() {
    let bytes = fixture();
    let huge = "x".repeat(32_000);

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let error = original
            .edit(|transaction| {
                transaction.set("Sheet1!A1:AF32", Props::new().with("value", huge.as_str()))
            })
            .expect_err("1,024 repeated 32,000-byte values exceed the default staged budget");
        assert_eq!(error.item(), Some(0));
        assert!(matches!(
            error.leaf(),
            Source::Transaction(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed,
                maximum,
            }) if *observed > *maximum && *maximum == SemanticLimits::default().max_staged_bytes()
        ));
        assert_eq!(
            original
                .to_bytes()
                .expect("publish after large-range refusal"),
            original_bytes,
            "staged-byte admission must leave the original owner unchanged"
        );
    }
}

#[test]
fn range_capability_is_advertised_as_a_single_unified_write_surface() {
    let range = Capabilities::of(FileFormat::Xlsx)
        .element("range")
        .expect("XLSX range capability");
    assert!(range.operations().set);
    for key in ["value", "formula", "clear"] {
        let property = range
            .property(key)
            .unwrap_or_else(|| panic!("missing XLSX range property {key}"));
        assert!(!property.get(), "range {key} must be write-only");
        assert!(property.set(), "range {key} must be settable");
    }
}

#[test]
fn best_effort_keeps_a_successful_range_before_a_failed_selector() {
    let original = open_owner(&fixture(), false);
    let result = original
        .batch(
            [
                batch_set("Sheet1!A1:A2", Props::new().with("value", "kept")),
                batch_set(
                    "Sheet1!cell[value>0], row[1]",
                    Props::new().with("value", "failed"),
                ),
            ],
            BatchMode::BestEffort,
        )
        .expect("best-effort retains a valid range item");
    assert!(matches!(
        &result.value()[0],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &result.value()[1],
        ItemResult::Failed(error) if error.item() == Some(1)
    ));
    assert_eq!(cell_text(result.file(), A1).as_deref(), Some("kept"));
    assert_eq!(cell_text(result.file(), A2).as_deref(), Some("kept"));
}
