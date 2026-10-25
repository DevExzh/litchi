//! Unified XLSX worksheet lifecycle operations.
#![cfg(feature = "xlsx")]

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::path::Position;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, NewNode, OpenOptions, Placement, PropValue,
    Props, SourceMode, props,
};

const SHEET1: &str = "/sheet[1]";
const SHEET2: &str = "/sheet[2]";
const SHEET3: &str = "/sheet[3]";
const A1: &str = "/sheet[1]/A1";

fn fixture() -> Vec<u8> {
    litchi_xlsx::Package::new()
        .expect("create native workbook")
        .to_bytes()
        .expect("serialize native workbook")
}

fn formula_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native edit");
    {
        let mut data = edit.add("Data").expect("add Data worksheet");
        data.set("A1", 7_i32).expect("seed Data");
    }
    edit.sheet("Sheet1")
        .expect("find Sheet1")
        .expect("Sheet1 worksheet")
        .set(
            "A1",
            litchi_xlsx::Formula::new("Data!A1").expect("formula reference"),
        )
        .expect("seed formula");
    edit.commit()
        .expect("commit formula fixture")
        .into_package()
        .to_bytes()
        .expect("serialize formula fixture")
}

fn multi_sheet_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    edit.add("Data").expect("add Data worksheet");
    edit.add("Archive").expect("add Archive worksheet");
    edit.commit()
        .expect("commit multi-sheet fixture")
        .into_package()
        .to_bytes()
        .expect("serialize multi-sheet fixture")
}

fn a1_sheet_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    let mut sheet = edit.add("A1").expect("add A1-named worksheet");
    sheet.set("A1", "seed").expect("seed A1-named worksheet");
    edit.commit()
        .expect("commit A1-named worksheet fixture")
        .into_package()
        .to_bytes()
        .expect("serialize A1-named worksheet fixture")
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

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("resolve sheet")
        .view()
        .format()
        .expect("read sheet properties")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn cell_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .format()
        .expect("read cell properties")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .text()
        .expect("read cell text")
        .map(std::borrow::Cow::into_owned)
}

fn scoped_sheet_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse sheet selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn batch_add(
    parent: &'static str,
    kind: &'static str,
    props: Props<'static>,
) -> BatchItem<'static> {
    let mut item = BatchItem::new("add");
    item.parent = Some(parent.into());
    item.kind = Some(kind.into());
    item.props = Some(props);
    item
}

fn batch_set(path: &'static str, props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

fn batch_remove(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

fn batch_move(path: &'static str, index: i32) -> BatchItem<'static> {
    let mut item = BatchItem::new("move");
    item.path = Some(path.into());
    item.index = Some(index);
    item
}

fn batch_swap(left: &'static str, right: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("swap");
    item.path = Some(left.into());
    item.path2 = Some(right.into());
    item
}

#[test]
fn worksheet_lifecycle_verbs_are_immutable_for_source_and_eager_owners() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");

        let added = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "Data")
                    .prop("hidden", true),
            )
            .expect("add worksheet");
        assert_eq!(added.path.to_string(), SHEET2);
        assert_eq!(
            sheet_property(&added.file, SHEET2, "name").as_deref(),
            Some("Data")
        );
        assert_eq!(
            sheet_property(&added.file, SHEET2, "hidden").as_deref(),
            Some("true")
        );

        let renamed = added
            .file
            .set(SHEET2, props! { "name" => "Renamed" })
            .expect("rename worksheet");
        assert_eq!(
            sheet_property(&renamed, SHEET2, "name").as_deref(),
            Some("Renamed")
        );

        let moved = renamed
            .move_to(SHEET2, Placement::Start)
            .expect("move worksheet to start");
        assert_eq!(
            sheet_property(&moved, SHEET1, "name").as_deref(),
            Some("Renamed")
        );
        assert_eq!(
            sheet_property(&moved, SHEET2, "name").as_deref(),
            Some("Sheet1")
        );

        let swapped = moved.swap(SHEET1, SHEET2).expect("swap worksheets");
        assert_eq!(
            sheet_property(&swapped, SHEET1, "name").as_deref(),
            Some("Sheet1")
        );
        assert_eq!(
            sheet_property(&swapped, SHEET2, "name").as_deref(),
            Some("Renamed")
        );

        let removed = swapped.remove(SHEET2).expect("remove worksheet");
        assert_eq!(
            sheet_property(&removed, SHEET1, "name").as_deref(),
            Some("Sheet1")
        );
        assert!(
            removed
                .find(SHEET2)
                .expect("find removed worksheet")
                .is_none()
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn worksheet_lifecycle_patch_replays_and_mixed_cell_runs_are_barriers() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");
        let commit = original
            .edit(|transaction| {
                transaction.set(A1, props! { "value" => "before sheet" })?;
                let added = transaction.add("/", NewNode::new("sheet").prop("name", "Data"))?;
                assert_eq!(added.to_string(), SHEET2);
                transaction.set(&added, props! { "hidden" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("mixed cell and worksheet edit");

        assert_eq!(commit.diagnostics().operations(), 3);
        assert_eq!(
            commit.diagnostics().commits(),
            3,
            "cell and structural families must close their native runs"
        );
        assert_eq!(
            sheet_property(commit.file(), SHEET2, "name").as_deref(),
            Some("Data")
        );
        assert_eq!(
            sheet_property(commit.file(), SHEET2, "hidden").as_deref(),
            Some("true")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay worksheet patch");
        assert_eq!(
            replayed.to_bytes().expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse worksheet patch");
        assert_eq!(
            restored.to_bytes().expect("serialize restored"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn worksheet_rename_rewrites_formula_references() {
    let bytes = formula_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");
        let renamed = original
            .set(SHEET2, Props::new().with("name", "Input 2026"))
            .expect("rename referenced worksheet");
        let formula = cell_property(&renamed, A1, "formula").expect("formula remains present");
        assert!(formula.contains("Input 2026"), "renamed formula: {formula}");
        assert!(
            !formula.contains("Data!"),
            "stale formula reference: {formula}"
        );
        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn worksheet_structural_guards_use_exact_codes_and_default_names() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize source");

        let default_name = original
            .add("/", NewNode::new("sheet"))
            .expect("the missing name uses the deterministic default");
        assert_eq!(default_name.path.to_string(), SHEET2);
        assert_eq!(
            sheet_property(&default_name.file, SHEET2, "name").as_deref(),
            Some("Sheet2")
        );

        let explicit_null = original
            .add("/", NewNode::new("sheet").prop("name", PropValue::Null))
            .expect_err("an explicit null sheet name is invalid");
        assert_eq!(explicit_null.code(), Code::InvalidValue);
        assert_eq!(explicit_null.property(), Some("name"));

        let empty_name = original
            .add("/", NewNode::new("sheet").prop("name", ""))
            .expect_err("an empty sheet name is invalid");
        assert_eq!(empty_name.code(), Code::InvalidValue);
        assert_eq!(empty_name.property(), Some("name"));

        let duplicate = original
            .add("/", NewNode::new("sheet").prop("name", "Sheet1"))
            .expect_err("duplicate worksheet names must be refused");
        assert_eq!(duplicate.code(), Code::DuplicateName);
        assert_eq!(duplicate.property(), Some("name"));

        let invalid = original
            .add("/", NewNode::new("sheet").prop("name", "bad/name"))
            .expect_err("invalid worksheet names must be refused");
        assert_eq!(invalid.code(), Code::InvalidValue);
        assert_eq!(invalid.property(), Some("name"));

        let last_visible = original
            .set(SHEET1, props! { "hidden" => true })
            .expect_err("the last visible worksheet cannot be hidden");
        assert_eq!(last_visible.code(), Code::InvalidArgument);
        assert_eq!(last_visible.property(), Some("hidden"));

        let last_visible_remove = original
            .remove(SHEET1)
            .expect_err("the last visible worksheet cannot be removed");
        assert_eq!(last_visible_remove.code(), Code::InvalidArgument);
        assert_eq!(last_visible_remove.property(), None);

        let with_hidden = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "Hidden")
                    .prop("hidden", true),
            )
            .expect("add a hidden worksheet");
        let sole_visible_remove = with_hidden
            .file
            .remove(SHEET1)
            .expect_err("a hidden remainder cannot replace the visible tab");
        assert_eq!(sole_visible_remove.code(), Code::InvalidArgument);
        assert_eq!(sole_visible_remove.property(), None);

        assert_eq!(
            original.to_bytes().expect("source remains unchanged"),
            original_bytes
        );
    }
}

#[test]
fn worksheet_add_and_move_accept_all_checked_placements() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);

        let at_position = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "AtPosition")
                    .at(Placement::Position(Position::new(1))),
            )
            .expect("position one appends after the first sheet");
        assert_eq!(at_position.path.to_string(), SHEET2);
        assert_eq!(
            sheet_property(&at_position.file, SHEET2, "name").as_deref(),
            Some("AtPosition")
        );

        let before = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "Before")
                    .at(Placement::before(SHEET1).expect("before placement")),
            )
            .expect("before placement");
        assert_eq!(before.path.to_string(), SHEET1);
        assert_eq!(
            sheet_property(&before.file, SHEET1, "name").as_deref(),
            Some("Before")
        );

        let after = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "After")
                    .at(Placement::after(SHEET1).expect("after placement")),
            )
            .expect("after placement");
        assert_eq!(after.path.to_string(), SHEET2);
        assert_eq!(
            sheet_property(&after.file, SHEET2, "name").as_deref(),
            Some("After")
        );

        let into_root = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "IntoRoot")
                    .at(Placement::inside("/").expect("into placement")),
            )
            .expect("into root placement");
        assert_eq!(into_root.path.to_string(), SHEET2);
        assert_eq!(
            sheet_property(&into_root.file, SHEET2, "name").as_deref(),
            Some("IntoRoot")
        );

        let bad_position = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "BadPosition")
                    .at(Placement::Position(Position::new(2))),
            )
            .expect_err("an insertion position beyond the catalog is refused");
        assert_eq!(bad_position.code(), Code::OutOfRange);

        let bad_anchor = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "BadAnchor")
                    .at(Placement::before("/sheet[99]").expect("parse bad anchor")),
            )
            .expect_err("an unknown placement anchor is refused");
        assert_eq!(bad_anchor.code(), Code::PathNotFound);

        let bad_parent = original
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "BadParent")
                    .at(Placement::inside(SHEET1).expect("parse sheet parent")),
            )
            .expect_err("a worksheet cannot contain another worksheet");
        assert_eq!(bad_parent.code(), Code::UnsupportedCapability);

        let moved_position = after
            .file
            .move_to(SHEET2, Position::new(0))
            .expect("IntoPlacement converts a checked position");
        assert_eq!(
            sheet_property(&moved_position, SHEET1, "name").as_deref(),
            Some("After")
        );
        assert_eq!(
            sheet_property(&moved_position, SHEET2, "name").as_deref(),
            Some("Sheet1")
        );

        let bad_move_position = after
            .file
            .move_to(SHEET2, Position::new(2))
            .expect_err("a move position beyond the catalog is refused");
        assert_eq!(bad_move_position.code(), Code::OutOfRange);
        let bad_move_parent = after
            .file
            .move_to(
                SHEET2,
                Placement::inside(SHEET1).expect("parse invalid move parent"),
            )
            .expect_err("a worksheet cannot move inside another worksheet");
        assert_eq!(bad_move_parent.code(), Code::UnsupportedCapability);
    }
}

#[test]
fn worksheet_batches_cover_structural_atomic_and_best_effort_routes() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");

        let atomic_failure = original
            .batch(
                [
                    batch_add("/", "sheet", Props::new().with("name", "Data")),
                    batch_add("/", "sheet", Props::new().with("name", "Sheet1")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic structural failure discards earlier additions");
        assert_eq!(atomic_failure.item(), Some(1));
        assert_eq!(atomic_failure.code(), Code::DuplicateName);
        assert_eq!(
            original.to_bytes().expect("serialize after atomic refusal"),
            original_bytes
        );

        let atomic = original
            .batch(
                [
                    batch_add("/", "sheet", Props::new().with("name", "Data")),
                    batch_set("/sheet[2]", Props::new().with("hidden", true)),
                    batch_move("/sheet[2]", 0),
                    batch_swap("/sheet[1]", "/sheet[2]"),
                    batch_remove("/sheet[2]"),
                ],
                BatchMode::Atomic,
            )
            .expect("atomic structural commands");
        assert_eq!(atomic.value().len(), 5);
        assert!(
            atomic
                .value()
                .iter()
                .all(|result| matches!(result, ItemResult::Applied(applied) if applied.changed))
        );
        assert_eq!(
            sheet_property(atomic.file(), SHEET1, "name").as_deref(),
            Some("Sheet1")
        );
        assert!(
            atomic
                .file()
                .find(SHEET2)
                .expect("find removed sheet")
                .is_none()
        );

        let best_effort = original
            .batch(
                [
                    batch_add("/", "sheet", Props::new().with("name", "Data")),
                    batch_add("/", "sheet", Props::new().with("name", "Sheet1")),
                    batch_set("/sheet[2]", Props::new().with("hidden", true)),
                    batch_move("/sheet[2]", 0),
                    batch_swap("/sheet[1]", "/sheet[2]"),
                    batch_remove("/sheet[2]"),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort structural commands");
        assert_eq!(best_effort.value().len(), 6);
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::DuplicateName
        ));
        assert!(
            best_effort
                .value()
                .iter()
                .skip(2)
                .all(|result| matches!(result, ItemResult::Applied(applied) if applied.changed))
        );
        assert!(
            best_effort
                .file()
                .find(SHEET2)
                .expect("find removed sheet")
                .is_none()
        );
    }
}

#[test]
fn worksheet_scoped_queries_remove_and_set_bounded_targets_atomically() {
    let bytes = multi_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");
        let data = scoped_sheet_query("sheet[name=Data]");
        let hidden = original
            .set(data.clone(), props! { "hidden" => true })
            .expect("scoped sheet set");
        assert_eq!(
            sheet_property(&hidden, "/sheet[2]", "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(
            sheet_property(&hidden, "/sheet[1]", "hidden").as_deref(),
            None
        );

        let remove_two = scoped_sheet_query("sheet[name!=Sheet1]");
        let commit = original
            .edit(|transaction| {
                transaction.remove(remove_two.clone())?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("remove two independent sheets in one native run");
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            sheet_property(commit.file(), SHEET1, "name").as_deref(),
            Some("Sheet1")
        );
        assert!(
            commit
                .file()
                .find("/sheet[2]")
                .expect("find removed tabs")
                .is_none()
        );

        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("expanded target limit")
            .build()
            .expect("valid semantic limits");
        let limited = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.remove(scoped_sheet_query("sheet[name!=Sheet1]"))?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("the scoped removal exceeds its target bound");
        assert_eq!(limited.code(), Code::LimitExceeded);
        assert_eq!(
            original
                .to_bytes()
                .expect("source remains unchanged after limit"),
            original_bytes
        );

        let no_matches = original
            .remove(scoped_sheet_query("sheet[name=Missing]"))
            .expect_err("zero-match structural query");
        assert_eq!(no_matches.code(), Code::ZeroMatches);

        let unknown = original
            .remove(scoped_sheet_query("sheet[unknownPredicate=true]"))
            .expect_err("unsupported predicates cannot widen removals");
        assert_eq!(unknown.code(), Code::InvalidSelector);

        let staged_limits = SemanticLimits::builder()
            .max_staged_bytes(64)
            .expect("staged byte limit")
            .build()
            .expect("valid limits");
        let staged_remove = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(staged_limits),
                |transaction| {
                    transaction.remove(scoped_sheet_query("sheet[name!=Sheet1]"))?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("two removals exceed the staged charge");
        assert_eq!(staged_remove.code(), Code::LimitExceeded);
        let staged_set = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(staged_limits),
                |transaction| {
                    transaction.set("/sheet[2]", props! { "hidden" => true })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("sheet assignment includes its payload charge");
        assert_eq!(staged_set.code(), Code::LimitExceeded);
        let expanded_set = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(
                        scoped_sheet_query("sheet[name!=Sheet1]"),
                        props! { "hidden" => true },
                    )?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("sheet set enforces expansion limits");
        assert_eq!(expanded_set.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("unchanged after refusals"),
            original_bytes
        );
    }
}

#[test]
fn worksheet_name_that_looks_like_a_cell_keeps_root_and_cell_paths_distinct() {
    let bytes = a1_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        assert_eq!(
            sheet_property(&original, SHEET2, "name").as_deref(),
            Some("A1")
        );
        assert_eq!(cell_text(&original, "/A1/A1").as_deref(), Some("seed"));

        let direct = original
            .set("/A1", props! { "name" => "Renamed A1", "hidden" => true })
            .expect("root A1 resolves to the worksheet tab");
        assert_eq!(
            sheet_property(&direct, SHEET2, "name").as_deref(),
            Some("Renamed A1")
        );
        assert_eq!(
            sheet_property(&direct, SHEET2, "hidden").as_deref(),
            Some("true")
        );

        let transaction = original
            .edit(|edit| {
                edit.set("/A1", props! { "hidden" => true })?;
                edit.set("/A1/A1", props! { "value" => "transaction cell" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("transaction routes sheet and cell paths independently");
        assert_eq!(
            sheet_property(transaction.file(), SHEET2, "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_text(transaction.file(), "/A1/A1").as_deref(),
            Some("transaction cell")
        );

        let batched = original
            .batch(
                [
                    batch_set("/A1", Props::new().with("hidden", true)),
                    batch_set("/A1/A1", Props::new().with("value", "batch cell")),
                ],
                BatchMode::Atomic,
            )
            .expect("batch routes sheet and cell paths independently");
        assert_eq!(
            sheet_property(batched.file(), SHEET2, "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_text(batched.file(), "/A1/A1").as_deref(),
            Some("batch cell")
        );
    }
}

#[test]
fn worksheet_swaps_support_nonadjacent_and_reversed_three_tab_pairs() {
    let bytes = multi_sheet_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (left, right) in [(SHEET1, SHEET3), (SHEET3, SHEET1)] {
            let swapped = original
                .swap(left, right)
                .expect("swap nonadjacent worksheet tabs");
            assert_eq!(
                sheet_property(&swapped, SHEET1, "name").as_deref(),
                Some("Archive")
            );
            assert_eq!(
                sheet_property(&swapped, SHEET2, "name").as_deref(),
                Some("Data")
            );
            assert_eq!(
                sheet_property(&swapped, SHEET3, "name").as_deref(),
                Some("Sheet1")
            );
        }
    }
}
