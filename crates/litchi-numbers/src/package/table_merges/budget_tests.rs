//! Independent resource-accounting tests for the focused merge reader.
//!
//! These tests intentionally exercise the private package boundary with an
//! Apple-authored model.  The production API does not expose a way to inject
//! a budget, so the test first performs the real rooted preflight, then runs
//! the same private wire/result gates with exact and one-below ceilings.

use super::*;

use litchi_iwa_protos::{tsd, tsp, tst};
use prost::Message as _;

const NATIVE_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-merges-native.numbers"
));
const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "shared-model";

struct Prepared<'source> {
    model: &'source [u8],
    budget: Budget,
    expected: merge_wire::MergeRead,
}

fn prepare(package: &Package) -> Prepared<'_> {
    let mut budget = Budget::new(package).expect("native fixture admits the default profile");
    let (sheet_position, table_position) = select_semantic_positions(
        package,
        SheetSelector::name(SHEET_NAME),
        TableSelector::name(TABLE_NAME),
        &mut budget,
    )
    .expect("native fixture has the expected rooted selectors");
    let model = resolve_model_payload(package, sheet_position, table_position, &mut budget)
        .expect("native fixture has a selected table model");
    let expected = merge_wire::read_table_merges(model, merge_wire::ReadLimits::default())
        .expect("native fixture has a valid merge owner");
    Prepared {
        model,
        budget,
        expected,
    }
}

fn run_preflight(package: &Package, mut budget: Budget) -> Result<(), TableMergesError> {
    let (sheet_position, table_position) = select_semantic_positions(
        package,
        SheetSelector::name(SHEET_NAME),
        TableSelector::name(TABLE_NAME),
        &mut budget,
    )?;
    let _ = resolve_model_payload(package, sheet_position, table_position, &mut budget)?;
    Ok(())
}

fn run_with_budget(
    model: &[u8],
    mut budget: Budget,
) -> (Result<Vec<Region>, TableMergesError>, Budget) {
    let result = (|| {
        let max_regions = merge_region_capacity(
            budget.remaining_allocations()?,
            budget.remaining_scratch()?,
            budget.remaining_retained()?,
            budget.remaining_regions()?,
        );
        let limits = merge_wire_limits(&budget, max_regions)?;
        let read = match merge_wire::read_table_merges(model, limits) {
            Ok(read) => read,
            Err(error) => {
                charge_attempted(&mut budget, &error);
                return Err(map_merge_error(error));
            },
        };
        let region_count = read.regions.len();
        budget.charge_wire_report(
            read.report.input_bytes(),
            read.report.fields(),
            read.report.work(),
        )?;
        budget.charge_regions(region_count)?;
        budget.charge_allocations(
            region_count
                .checked_mul(MERGE_READER_ALLOCATIONS_PER_REGION)
                .ok_or(TableMergesError::InvalidSource)?,
        )?;
        budget.charge_scratch(
            region_count
                .checked_mul(MERGE_READER_SCRATCH_PER_REGION)
                .ok_or(TableMergesError::InvalidSource)?,
        )?;
        budget.charge_retained(
            region_count
                .checked_mul(REGION_BYTES)
                .ok_or(TableMergesError::InvalidSource)?,
        )?;
        Ok(read.regions)
    })();
    (result, budget)
}

fn exact_budget(prepared: &Prepared<'_>) -> Budget {
    let mut budget = prepared.budget;
    let report = prepared.expected.report;
    let regions = prepared.expected.regions.len();
    budget.wire_max_input = budget
        .input
        .checked_add(report.input_bytes())
        .expect("native wire input cost fits");
    budget.wire_max_fields = budget
        .fields
        .checked_add(report.fields())
        .expect("native wire field cost fits");
    budget.wire_max_work = budget
        .work
        .checked_add(report.work())
        .expect("native wire work cost fits");
    budget.max_regions = budget
        .regions
        .checked_add(regions)
        .expect("native region cost fits");
    budget.max_allocations = budget
        .allocations
        .checked_add(
            regions
                .checked_mul(MERGE_READER_ALLOCATIONS_PER_REGION)
                .expect("native allocation cost fits"),
        )
        .expect("native allocation ceiling fits");
    budget.max_scratch = budget
        .scratch
        .checked_add(
            regions
                .checked_mul(MERGE_READER_SCRATCH_PER_REGION)
                .expect("native scratch cost fits"),
        )
        .expect("native scratch ceiling fits");
    budget.max_retained = budget
        .retained
        .checked_add(
            regions
                .checked_mul(REGION_BYTES)
                .expect("native retained cost fits"),
        )
        .expect("native retained ceiling fits");
    budget
}

fn counters(
    budget: &Budget,
) -> (
    usize,
    usize,
    usize,
    usize,
    usize,
    usize,
    usize,
    usize,
    usize,
    usize,
) {
    (
        budget.input,
        budget.fields,
        budget.work,
        budget.objects,
        budget.messages,
        budget.items,
        budget.references,
        budget.allocations,
        budget.retained,
        budget.scratch,
    )
}

fn assert_limit(
    result: Result<Vec<Region>, TableMergesError>,
    expected: TableMergesLimitKind,
    label: &str,
) {
    match result {
        Err(TableMergesError::LimitExceeded { kind, .. }) => {
            assert_eq!(kind, expected, "{label} reported the wrong resource");
        },
        other => panic!("{label} unexpectedly returned {other:?}"),
    }
}

fn assert_preflight_limit(
    result: Result<(), TableMergesError>,
    expected: TableMergesLimitKind,
    label: &str,
) {
    match result {
        Err(TableMergesError::LimitExceeded { kind, .. }) => {
            assert_eq!(kind, expected, "{label} reported the wrong resource");
        },
        other => panic!("{label} unexpectedly returned {other:?}"),
    }
}

fn serialized(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .expect("native package can be serialized");
    bytes
}

#[test]
fn rooted_preflight_charges_selector_metadata_and_wire_resources_cumulatively() {
    let package = Package::from_bytes(NATIVE_SOURCE).expect("native fixture parses");
    let before = serialized(&package);
    let prepared = prepare(&package);
    let budget = prepared.budget;

    assert!(budget.input > 0, "rooted preflight must charge input bytes");
    assert!(
        budget.fields > 0,
        "rooted preflight must charge wire fields"
    );
    assert!(budget.work > 0, "selectors and references must charge work");
    assert!(budget.objects > 0, "rooted object lookup must be charged");
    assert!(budget.messages > 0, "message inspection must be charged");
    assert!(
        budget.items > 0,
        "metadata and selector items must be charged"
    );
    assert!(budget.references > 0, "rooted references must be charged");

    let (result, _) = run_with_budget(prepared.model, exact_budget(&prepared));
    assert_eq!(
        result.expect("exact native budget admits the read"),
        prepared.expected.regions
    );
    assert_eq!(
        serialized(&package),
        before,
        "focused reads must preserve source bytes"
    );
}

#[test]
fn exact_cumulative_wire_and_result_budgets_admit_the_native_region() {
    let package = Package::from_bytes(NATIVE_SOURCE).expect("native fixture parses");
    let prepared = prepare(&package);
    let budget = exact_budget(&prepared);
    let expected_counters = (
        budget.input + prepared.expected.report.input_bytes(),
        budget.fields + prepared.expected.report.fields(),
        budget.work + prepared.expected.report.work(),
        budget.objects,
        budget.messages,
        budget.items,
        budget.references,
        budget.allocations + prepared.expected.regions.len() * MERGE_READER_ALLOCATIONS_PER_REGION,
        budget.retained + prepared.expected.regions.len() * REGION_BYTES,
        budget.scratch + prepared.expected.regions.len() * MERGE_READER_SCRATCH_PER_REGION,
    );

    let (result, after) = run_with_budget(prepared.model, budget);
    assert_eq!(
        result.expect("every exact resource ceiling admits the read"),
        prepared.expected.regions
    );
    assert_eq!(counters(&after), expected_counters);
}

#[test]
fn exact_root_metadata_and_selector_caps_admit_and_one_below_refuses() {
    let package = Package::from_bytes(NATIVE_SOURCE).expect("native fixture parses");
    let baseline = prepare(&package).budget;

    let mut exact = Budget::new(&package).expect("native fixture admits the default profile");
    exact.wire_max_input = baseline.input;
    exact.wire_max_fields = baseline.fields;
    exact.wire_max_work = baseline.work;
    exact.max_objects = baseline.objects;
    exact.max_messages = baseline.messages;
    exact.max_items = baseline.items;
    exact.max_references = baseline.references;
    let exact_result = run_preflight(&package, exact);
    assert!(
        exact_result.is_ok(),
        "a cap equal to the complete rooted preflight cost must admit it: {exact_result:?}"
    );

    let mut input = exact;
    input.wire_max_input -= 1;
    assert_preflight_limit(
        run_preflight(&package, input),
        TableMergesLimitKind::WireBytes,
        "root input",
    );

    let mut fields = exact;
    fields.wire_max_fields -= 1;
    assert_preflight_limit(
        run_preflight(&package, fields),
        TableMergesLimitKind::WireFields,
        "root fields",
    );

    let mut work = exact;
    work.wire_max_work -= 1;
    assert_preflight_limit(
        run_preflight(&package, work),
        TableMergesLimitKind::WireWork,
        "selector work",
    );

    let mut objects = exact;
    objects.max_objects -= 1;
    assert_preflight_limit(
        run_preflight(&package, objects),
        TableMergesLimitKind::PayloadObjects,
        "root objects",
    );

    let mut messages = exact;
    messages.max_messages -= 1;
    assert_preflight_limit(
        run_preflight(&package, messages),
        TableMergesLimitKind::PayloadMessages,
        "metadata messages",
    );

    let mut items = exact;
    items.max_items -= 1;
    assert_preflight_limit(
        run_preflight(&package, items),
        TableMergesLimitKind::PayloadItems,
        "metadata items",
    );

    let mut references = exact;
    references.max_references -= 1;
    assert_preflight_limit(
        run_preflight(&package, references),
        TableMergesLimitKind::PayloadReferences,
        "root references",
    );
}

#[test]
fn one_below_each_wire_and_result_budget_is_reported_as_a_typed_refusal() {
    let package = Package::from_bytes(NATIVE_SOURCE).expect("native fixture parses");
    let prepared = prepare(&package);
    let exact = exact_budget(&prepared);

    let mut input = exact;
    input.wire_max_input -= 1;
    let (result, _) = run_with_budget(prepared.model, input);
    assert_limit(result, TableMergesLimitKind::WireBytes, "input");

    let mut fields = exact;
    fields.wire_max_fields -= 1;
    let (result, _) = run_with_budget(prepared.model, fields);
    assert_limit(result, TableMergesLimitKind::WireFields, "fields");

    let mut work = exact;
    work.wire_max_work -= 1;
    let (result, _) = run_with_budget(prepared.model, work);
    assert_limit(result, TableMergesLimitKind::WireWork, "work");

    let mut regions = exact;
    regions.max_regions -= 1;
    let (result, _) = run_with_budget(prepared.model, regions);
    assert_limit(result, TableMergesLimitKind::Regions, "regions");

    let mut allocations = exact;
    allocations.max_allocations -= 1;
    let (result, _) = run_with_budget(prepared.model, allocations);
    assert_limit(result, TableMergesLimitKind::Regions, "allocations");

    let mut retained = exact;
    retained.max_retained -= 1;
    let (result, _) = run_with_budget(prepared.model, retained);
    assert_limit(result, TableMergesLimitKind::Regions, "retained");

    let mut scratch = exact;
    scratch.max_scratch -= 1;
    let (result, _) = run_with_budget(prepared.model, scratch);
    assert_limit(result, TableMergesLimitKind::Regions, "scratch");
}

#[test]
fn zero_wire_residual_refuses_before_entering_the_selected_codec() {
    let package = Package::from_bytes(NATIVE_SOURCE).expect("native fixture parses");
    let prepared = prepare(&package);

    let mut input = prepared.budget;
    input.wire_max_input = input.input;
    let before = counters(&input);
    let (result, after) = run_with_budget(prepared.model, input);
    assert_limit(
        result,
        TableMergesLimitKind::WireBytes,
        "zero input residual",
    );
    assert_eq!(
        counters(&after),
        before,
        "zero residual must fail before extra work"
    );

    let mut fields = prepared.budget;
    fields.wire_max_fields = fields.fields;
    let before = counters(&fields);
    let (result, after) = run_with_budget(prepared.model, fields);
    assert_limit(
        result,
        TableMergesLimitKind::WireFields,
        "zero field residual",
    );
    assert_eq!(
        counters(&after),
        before,
        "zero residual must fail before extra work"
    );

    let mut work = prepared.budget;
    work.wire_max_work = work.work;
    let before = counters(&work);
    let (result, after) = run_with_budget(prepared.model, work);
    assert_limit(result, TableMergesLimitKind::WireWork, "zero work residual");
    assert_eq!(
        counters(&after),
        before,
        "zero residual must fail before extra work"
    );
}

#[test]
fn absent_drawable_lock_is_valid_during_table_ownership_preflight() {
    for locked in [None, Some(false), Some(true)] {
        let source = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                locked,
                ..Default::default()
            },
            table_model: tsp::Reference {
                identifier: 42,
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec();
        let package = Package::from_bytes(NATIVE_SOURCE).expect("native fixture parses");
        let mut budget = Budget::new(&package).expect("native fixture admits the default profile");
        if locked.is_none() {
            assert!(
                !source.iter().any(|byte| *byte == 0x28),
                "the absent-lock case must omit DrawableArchive.locked"
            );
        }
        assert_eq!(
            table_model_identifier(&source, &mut budget),
            Ok(42),
            "an optional lock must not be required to resolve a table model"
        );
    }
}
