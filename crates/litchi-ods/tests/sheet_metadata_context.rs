//! Execution ownership regressions for source-bound metadata transactions.

use std::{
    num::{NonZeroU64, NonZeroUsize},
    sync::Arc,
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionLimits, Limits, OwnedSource,
    Profile, Resource,
};
use litchi_ods::sheet_metadata::{self, Options, Snapshot};
use litchi_ods::{Builder, SourceBackedSpreadsheet};

const SOURCE: &str = concat!(
    "<office:document-content ",
    "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">",
    "<office:body><office:spreadsheet><table:table table:name=\"Data\">",
    "<table:table-row><table:table-cell/></table:table-row></table:table>",
    "</office:spreadsheet></office:body></office:document-content>"
);

fn context(budget: Budget) -> (CancellationSource, ExecutionContext) {
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero worker count"),
        NonZeroUsize::new(1).expect("nonzero task count"),
        NonZeroU64::new(1024).expect("nonzero byte count"),
        0,
    )
    .expect("finite execution limits");
    (cancellation, ExecutionContext::new(budget, token, limits))
}

fn budget() -> Budget {
    // Equal diagnostic scope strings must not imply shared ownership.
    Budget::root(
        "metadata-context-regression",
        Limits::for_profile(Profile::Server),
    )
}

fn options() -> Options {
    Options::new("sum", vec!["Data.A1:Data.A1".to_owned()], "Data.B1")
        .expect("valid inert consolidation")
}

#[test]
fn equal_limits_and_usage_do_not_authorize_an_unrelated_budget() {
    let source_budget = budget();
    let (_source_cancel, source_context) = context(source_budget.clone());
    let other_budget = budget();
    let (_other_cancel, other_context) = context(other_budget.clone());
    let source =
        Snapshot::parse_with_context(SOURCE, sheet_metadata::Limits::default(), &source_context)
            .expect("source snapshot");
    let mut edit = source.edit();
    edit.set_consolidation(Some(options())).expect("stage edit");
    let usage = source_budget.used(Resource::Memory);
    let imitation = other_budget
        .reserve(Resource::Memory, usage)
        .expect("match source usage with unrelated allocation");
    assert_eq!(other_budget.used(Resource::Memory), usage);

    assert!(edit.commit(&other_context).is_err());
    assert_eq!(source.source_xml(), SOURCE);
    assert!(!edit.is_no_op(), "failed admission must retain staging");
    assert!(
        edit.commit(&source_context)
            .expect("retry source context")
            .changed()
    );
    drop(imitation);
    assert_eq!(other_budget.used(Resource::Memory), 0);
}

#[test]
fn cloned_context_retains_authority() {
    let (_cancel, original) = context(budget());
    let source = Snapshot::parse_with_context(SOURCE, sheet_metadata::Limits::default(), &original)
        .expect("source snapshot");
    let mut edit = source.edit();
    edit.set_consolidation(Some(options())).expect("stage edit");
    let commit = edit
        .commit(&original.clone())
        .expect("shared context clone");
    assert!(commit.changed());
    assert_eq!(commit.patch().source_xml(), SOURCE);
}

#[test]
fn fresh_token_cannot_bypass_retained_source_cancellation() {
    for changed in [false, true] {
        let shared_budget = budget();
        let (cancel, original) = context(shared_budget.clone());
        let (_fresh_cancel, fresh) = context(shared_budget);
        let source =
            Snapshot::parse_with_context(SOURCE, sheet_metadata::Limits::default(), &original)
                .expect("source snapshot");
        let mut edit = source.edit();
        if changed {
            edit.set_consolidation(Some(options())).expect("stage edit");
        }
        cancel.cancel();
        fresh.check().expect("replacement token remains live");
        let error = edit
            .commit(&fresh)
            .expect_err("source cancellation must prevail");
        assert!(error.to_string().contains("cancelled"));
        assert_eq!(edit.is_no_op(), !changed);
        assert_eq!(source.source_xml(), SOURCE);
    }
}

#[test]
fn patch_application_respects_destination_output_limit() {
    let (_source_cancel, source_context) = context(budget());
    let source =
        Snapshot::parse_with_context(SOURCE, sheet_metadata::Limits::default(), &source_context)
            .expect("source snapshot");
    let mut edit = source.edit();
    let mut value = options();
    // Application-defined function strings are permitted metadata. This one
    // makes the accepted target larger than the destination's output policy.
    value.function = "f".repeat(2048);
    edit.set_consolidation(Some(value))
        .expect("stage larger owner");
    let commit = edit.commit(&source_context).expect("originating commit");

    let (_destination_cancel, destination_context) = context(budget());
    let destination = Snapshot::parse_with_context(
        SOURCE,
        sheet_metadata::Limits::new(4096, 1024, 32, 1024, 1_000_000)
            .expect("smaller destination policy"),
        &destination_context,
    )
    .expect("exact source fits the destination policy");
    assert!(commit.patch().target_xml().len() > destination.limits().max_output_bytes());
    let error = commit
        .patch()
        .apply(&destination)
        .expect_err("destination output ceiling");
    let Error::ResourceLimit(limit) = error else {
        panic!("typed resource limit expected: {error:?}")
    };
    assert_eq!(limit.resource, Resource::OutputBytes);
    assert_eq!(limit.limit, 1024);
    assert!(limit.observed > limit.limit);
    assert_eq!(destination.source_xml(), SOURCE);
}

#[test]
fn source_bound_patch_respects_a_new_profile_on_the_same_owner() {
    let bytes = Builder::new()
        .content_xml(SOURCE)
        .build()
        .expect("ODS package");
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source-backed owner");
    let (_source_cancel, source_context) = context(budget());
    let source = owner
        .sheet_metadata_with(sheet_metadata::Limits::default(), &source_context)
        .expect("originating profile");
    let mut edit = source.edit().expect("source edit");
    let mut value = options();
    value.function = "f".repeat(2048);
    edit.set_consolidation(Some(value))
        .expect("stage larger owner");
    let commit = edit.commit(&source_context).expect("originating commit");

    let (_destination_cancel, destination_context) = context(budget());
    let destination = owner
        .sheet_metadata_with(
            sheet_metadata::Limits::new(4096, 1024, 32, 1024, 1_000_000)
                .expect("smaller destination policy"),
            &destination_context,
        )
        .expect("same owner under smaller profile");
    assert!(commit.patch().target_xml().len() > 1024);
    let error = commit
        .patch()
        .apply(&destination)
        .expect_err("destination output ceiling");
    let Error::ResourceLimit(limit) = error else {
        panic!("typed resource limit expected: {error:?}")
    };
    assert_eq!(limit.resource, Resource::OutputBytes);
    assert_eq!(limit.limit, 1024);
    assert!(limit.observed > limit.limit);
    assert_eq!(destination.source_xml(), source.source_xml());
}

#[test]
fn local_input_limits_have_structured_diagnostics_in_both_scanners() {
    let (_cancel, context) = context(budget());
    let maximum = SOURCE.len() - 1;
    let limits =
        sheet_metadata::Limits::new(maximum, 4096, 32, 1024, 1_000_000).expect("finite limits");
    let metadata = Snapshot::parse_with_context(SOURCE, limits, &context).expect_err("input limit");
    let detective_limits =
        sheet_metadata::detective::Limits::new(maximum, 4096, 32, 1024, 1_000_000)
            .expect("finite detective limits");
    let detective =
        sheet_metadata::detective::Snapshot::parse_with_context(SOURCE, detective_limits, &context)
            .expect_err("detective input limit");
    for error in [metadata, detective] {
        let Error::ResourceLimit(limit) = error else {
            panic!("typed resource limit expected: {error:?}")
        };
        assert_eq!(limit.resource, Resource::InputBytes);
        assert_eq!(limit.observed, SOURCE.len() as u64);
        assert_eq!(limit.limit, maximum as u64);
        assert!(limit.scope.contains("input bytes"));
    }
}

#[test]
fn local_staging_limit_reports_observed_operation_count() {
    let (_cancel, context) = context(budget());
    let limits = sheet_metadata::Limits::default()
        .with_item_limits(1, 1, 1)
        .expect("finite one-operation limit");
    let source = Snapshot::parse_with_context(SOURCE, limits, &context).expect("source");
    let mut edit = source.edit();
    edit.set_consolidation(Some(options()))
        .expect("first operation");
    let error = edit.set_consolidation(None).expect_err("second operation");
    let Error::ResourceLimit(limit) = error else {
        panic!("typed resource limit expected: {error:?}")
    };
    assert_eq!(limit.resource, Resource::Objects);
    assert_eq!(limit.observed, 2);
    assert_eq!(limit.limit, 1);
    assert!(
        edit.commit(&context)
            .expect("first operation retained")
            .changed()
    );
}

#[test]
fn scalar_caps_precede_address_validation_and_leave_staging_unchanged() {
    let (_cancel, context) = context(budget());
    let limits = sheet_metadata::Limits::default()
        .with_scalar_limits(128, 64)
        .expect("finite scalar limit");
    let source = Snapshot::parse_with_context(SOURCE, limits, &context).expect("source");
    let mut edit = source.edit();
    let mut value = options();
    // An invalid, over-limit address must hit the finite cap before the
    // syntax validator scans or copies it into an error message.
    value.target_cell_address = "x".repeat(256);
    let error = edit.set_consolidation(Some(value)).expect_err("scalar cap");
    let Error::ResourceLimit(limit) = error else {
        panic!("typed resource limit expected: {error:?}")
    };
    assert_eq!(limit.resource, Resource::Memory);
    assert_eq!(limit.observed, 256);
    assert_eq!(limit.limit, 128);
    assert!(edit.is_no_op());
}
