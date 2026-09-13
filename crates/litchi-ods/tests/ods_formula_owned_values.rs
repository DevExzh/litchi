use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Profile, Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure,
        value::{
            Context, Limits as ValueLimits, Mode, OwnedArrayView, OwnedEvaluated,
            OwnedReferenceListView, OwnedValueView, Position, SheetExtent, evaluate,
        },
    },
    expression::Expression,
    reference::Reference,
};
use litchi_ods::worksheet::formula::Resolver;
use litchi_ods::{Cell, CellValue, Row, Sheet};

fn make_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        scope.to_owned(),
        litchi_core::Limits::for_profile(Profile::Server),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
        .expect("minimal execution limits are valid");
    let execution = ExecutionContext::new(budget.clone(), token, limits);
    (budget, cancellation, execution)
}

fn sheets_with_values() -> Vec<Sheet> {
    let mut main = Sheet::new("Main").expect("valid sheet name");
    let mut row = Row::new();
    row.push_cell(Cell::new(
        CellValue::Text("resolver α🌟 & \"quoted\"".to_owned()),
        "display",
    ))
    .expect("cell fits row");
    row.push_cell(Cell::new(CellValue::Number(2.0), "2"))
        .expect("cell fits row");
    main.push_row(row).expect("row fits sheet");

    vec![
        main,
        Sheet::new("Data").expect("valid sheet name"),
        Sheet::new("Archive").expect("valid sheet name"),
    ]
}

fn try_own_formula(
    source: &str,
    mode: Mode,
    copy_execution: &ExecutionContext,
    copy_limits: &ValueLimits,
) -> Result<OwnedEvaluated, EvaluationFailure> {
    let sheets = sheets_with_values();
    let (_preparation_budget, _preparation_cancellation, preparation) =
        make_execution("owned-values-preparation");
    let resolver = Resolver::new(&sheets, SheetExtent::new(8, 8), &preparation)
        .expect("fixture resolver is valid");
    let expression = Expression::parse(source).expect("fixture expression is valid");
    let context = Context::new(&preparation, Position::new("Main", 0, 0)).with_mode(mode);
    let evaluated = evaluate(&expression, &resolver, &context, &ValueLimits::default())?;
    evaluated.to_owned(copy_execution, copy_limits)
}

fn own_formula(
    source: &str,
    mode: Mode,
    copy_execution: &ExecutionContext,
    copy_limits: &ValueLimits,
) -> OwnedEvaluated {
    try_own_formula(source, mode, copy_execution, copy_limits)
        .expect("owned conversion should succeed")
}

fn assert_text(value: OwnedValueView<'_>, expected: &str) {
    match value {
        OwnedValueView::Text(actual) => assert_eq!(actual, expected),
        other => panic!("expected owned text, got {other:?}"),
    }
}

fn assert_text_array(array: OwnedArrayView<'_>, expected: &[&str]) {
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (1, expected.len())
    );
    assert_eq!(array.len(), expected.len());
    for (value, expected) in array.iter().zip(expected) {
        assert_text(value, expected);
    }
}

fn assert_reference_list(list: OwnedReferenceListView<'_>, expected_columns: &[usize]) {
    assert_eq!(list.len(), expected_columns.len());
    for (reference, &column) in list.iter().zip(expected_columns) {
        assert_eq!(reference.len(), 1);
        assert_eq!(reference.areas()[0].starts(), [0, 0, column]);
        assert_eq!(reference.areas()[0].ends(), [1, 1, column + 1]);
        assert_eq!(reference.areas()[0].cell_count(), Some(1));
    }
}

#[test]
fn owned_text_and_array_survive_source_drop_with_exact_unicode_contents() {
    let (_copy_budget, _copy_cancellation, copy_execution) = make_execution("owned-text-copy");

    let owned_text = own_formula(
        "=[.A1]",
        Mode::Scalar,
        &copy_execution,
        &ValueLimits::default(),
    );
    assert_text(owned_text.value(), "resolver α🌟 & \"quoted\"");

    let owned_array = own_formula(
        "={\"α🌟\";\"A\"\"B\"}",
        Mode::Matrix,
        &copy_execution,
        &ValueLimits::default(),
    );
    let array = owned_array.as_array().expect("array result");
    assert_text_array(array, &["α🌟", "A\"B"]);
}

#[test]
fn owned_references_preserve_local_3d_metadata_and_list_order() {
    let (_copy_budget, _copy_cancellation, copy_execution) = make_execution("owned-references");

    let owned_3d = own_formula(
        "=[Main.A1:Archive.B2]",
        Mode::Matrix,
        &copy_execution,
        &ValueLimits::default(),
    );
    let reference = owned_3d.as_reference().expect("3D reference result");
    let expected = Reference::parse("[Main.A1:Archive.B2]").expect("valid expected reference");
    assert_eq!(reference.reference(), Some(&expected));
    assert_eq!(reference.len(), 1);
    assert_eq!(reference.areas()[0].starts(), [0, 0, 0]);
    assert_eq!(reference.areas()[0].ends(), [3, 2, 2]);
    assert_eq!(reference.areas()[0].extent(), [3, 2, 2]);

    let owned_list = own_formula(
        "=([.A1]~[.A1]~[.B1])",
        Mode::Matrix,
        &copy_execution,
        &ValueLimits::default(),
    );
    let list = owned_list
        .as_reference_list()
        .expect("reference-list result");
    assert_reference_list(list, &[0, 0, 1]);
    let expected_first = Reference::parse("[.A1]").expect("valid expected reference");
    assert_eq!(
        list.get(0).expect("first list item").reference(),
        Some(&expected_first)
    );
    assert_eq!(
        list.get(1).expect("duplicate list item").reference(),
        Some(&expected_first)
    );
}

#[test]
fn owned_reference_lexical_markers_and_axis_kinds_survive_copy() {
    let (_budget, _cancellation, execution) = make_execution("owned-reference-lexemes");
    for source in ["[$'Main'.$A$1:$'Archive'.$B$2]", "[.$A:.$B]", "[.$1:.$2]"] {
        let owned = own_formula(
            &format!("={source}"),
            Mode::Matrix,
            &execution,
            &ValueLimits::default(),
        );
        let expected = Reference::parse(source).expect("valid reference lexeme");
        assert_eq!(
            owned.as_reference().expect("owned reference").reference(),
            Some(&expected),
            "{source}",
        );
    }
}

#[test]
fn independently_owned_arrays_and_lists_have_structural_equality() {
    let (_left_budget, _left_cancellation, left_execution) = make_execution("owned-equality-left");
    let (_right_budget, _right_cancellation, right_execution) =
        make_execution("owned-equality-right");

    let left_array = own_formula(
        "={\"α🌟\";\"A\"\"B\"}",
        Mode::Matrix,
        &left_execution,
        &ValueLimits::default(),
    );
    let right_array = own_formula(
        "={\"α🌟\";\"A\"\"B\"}",
        Mode::Matrix,
        &right_execution,
        &ValueLimits::default(),
    );
    assert_eq!(left_array.value(), right_array.value());

    let left_list = own_formula(
        "=([.A1]~[.A1]~[.B1])",
        Mode::Matrix,
        &left_execution,
        &ValueLimits::default(),
    );
    let right_list = own_formula(
        "=([.A1]~[.A1]~[.B1])",
        Mode::Matrix,
        &right_execution,
        &ValueLimits::default(),
    );
    assert_eq!(left_list.value(), right_list.value());
}

#[test]
fn owned_conversion_reservations_release_after_successful_drop() {
    let (copy_budget, _copy_cancellation, copy_execution) = make_execution("owned-drop");
    let owned = own_formula(
        "={\"α🌟\";\"A\"\"B\"}",
        Mode::Matrix,
        &copy_execution,
        &ValueLimits::default(),
    );
    assert!(owned.reserved_storage_bytes() > 0);
    assert!(copy_budget.used(Resource::Memory) > 0);

    drop(owned);
    assert_eq!(copy_budget.used(Resource::Memory), 0);
}

#[test]
fn owned_conversion_enforces_text_and_aggregate_storage_limits_atomically() {
    let (text_budget, _text_cancellation, text_execution) = make_execution("owned-text-limit");
    let text_error = try_own_formula(
        "=\"α🌟\"",
        Mode::Scalar,
        &text_execution,
        &ValueLimits::default().with_max_text_bytes(2),
    )
    .expect_err("decoded text exceeds the local text limit");
    assert!(matches!(
        text_error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(text_budget.used(Resource::Memory), 0);

    let (_single_budget, _single_cancellation, single_execution) =
        make_execution("owned-single-measure");
    let single = own_formula(
        "=\"α\"",
        Mode::Scalar,
        &single_execution,
        &ValueLimits::default(),
    );
    let single_storage = single.reserved_storage_bytes();
    assert!(single_storage > 0);
    drop(single);

    let (array_budget, _array_cancellation, array_execution) =
        make_execution("owned-aggregate-limit");
    let aggregate_error = try_own_formula(
        "={\"α\";\"β\"}",
        Mode::Matrix,
        &array_execution,
        &ValueLimits::default().with_max_storage_bytes(single_storage),
    )
    .expect_err("two owned cells exceed the one-text storage budget");
    assert!(matches!(
        aggregate_error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(array_budget.used(Resource::Memory), 0);
}

#[test]
fn owned_conversion_work_and_cancellation_fail_before_publishing_memory() {
    let (work_budget, _work_cancellation, work_execution) = make_execution("owned-work-limit");
    let work_error = try_own_formula(
        "={\"α🌟\";\"A\"\"B\"}",
        Mode::Matrix,
        &work_execution,
        &ValueLimits::default().with_max_steps(0),
    )
    .expect_err("zero conversion work is insufficient");
    assert!(matches!(
        work_error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(work_budget.used(Resource::Memory), 0);

    let (cancel_budget, cancellation, cancel_execution) = make_execution("owned-cancel");
    cancellation.cancel();
    let cancel_error = try_own_formula(
        "={\"α🌟\";\"A\"\"B\"}",
        Mode::Matrix,
        &cancel_execution,
        &ValueLimits::default(),
    )
    .expect_err("pre-cancelled conversion must refuse");
    assert!(matches!(cancel_error, EvaluationFailure::Cancelled));
    assert_eq!(cancel_budget.used(Resource::Memory), 0);
}
