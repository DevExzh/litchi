//! Exact-boundary coverage for the Pages hidden-state wire codec.
//!
//! The Pages package transaction adds physical archive and transaction-ledger
//! ceilings around this codec.  These tests keep the codec's own independent
//! limits honest while that higher-level owner is assembled: an observation
//! equal to a configured maximum is accepted, and one unit below it is
//! rejected with the corresponding typed resource.

use litchi_iwa_protos::pages_hidden_state_codec::{
    AxisDirection, DecodeError, DecodeLimit, DecodeOptions, HiddenStateExtentSnapshot,
    HiddenStatesOwnerSnapshot, HiddenStatesSnapshot, RewriteExecutionLimits,
    RowOrColumnStateSnapshot, UuidSnapshot,
};

fn varint(mut value: u64) -> Vec<u8> {
    let mut bytes = Vec::new();
    while value >= 0x80 {
        bytes.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    bytes.push(value as u8);
    bytes
}

fn length_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut bytes = varint((u64::from(number) << 3) | 2);
    bytes.extend(varint(payload.len() as u64));
    bytes.extend_from_slice(payload);
    bytes
}

fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut bytes = varint(u64::from(number) << 3);
    bytes.extend(varint(value));
    bytes
}

fn uuid(lower: u64, upper: u64) -> Vec<u8> {
    let mut bytes = varint_field(1, lower);
    bytes.extend(varint_field(2, upper));
    bytes
}

fn row_state(lower: u64, upper: u64, user_hidden: bool) -> Vec<u8> {
    let mut bytes = length_field(1, &uuid(lower, upper));
    if user_hidden {
        bytes.extend(varint_field(2, 1));
    }
    bytes
}

fn extent(lower: u64, upper: u64, direction: AxisDirection, state: &[u8]) -> Vec<u8> {
    let mut bytes = length_field(1, &uuid(lower, upper));
    bytes.extend(length_field(2, state));
    bytes.extend(varint_field(3, direction.native_value() as u64));
    bytes
}

fn owner_source() -> Vec<u8> {
    let column = extent(12, 1, AxisDirection::Column, &row_state(20, 1, true));
    let row = extent(13, 1, AxisDirection::Row, &row_state(21, 1, false));
    let mut hidden = length_field(1, &uuid(11, 1));
    hidden.extend(length_field(2, &column));
    hidden.extend(length_field(3, &row));
    let mut owner = length_field(1, &uuid(10, 1));
    owner.extend(length_field(2, &hidden));
    owner
}

fn generous_options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(source.len(), source.len(), 16_384, 1_000_000, 64, 128)
        .with_max_allocations(16_384)
        .with_max_retained_bytes(1_000_000)
        .with_max_scratch_bytes(1_000_000)
}

#[derive(Clone, Copy, Debug)]
enum ExpectedLimit {
    InputBytes,
    OutputBytes,
    Fields,
    WorkBytes,
    Nesting,
    States,
    Allocations,
    RetainedBytes,
    ScratchBytes,
}

fn assert_limit(error: DecodeError, expected: ExpectedLimit) {
    let limit = error
        .resource_limit()
        .unwrap_or_else(|| panic!("expected typed resource error, got {error}"));
    let matches = matches!(
        (expected, &limit),
        (ExpectedLimit::InputBytes, DecodeLimit::InputBytes { .. })
            | (ExpectedLimit::OutputBytes, DecodeLimit::OutputBytes { .. })
            | (ExpectedLimit::Fields, DecodeLimit::Fields { .. })
            | (ExpectedLimit::WorkBytes, DecodeLimit::WorkBytes { .. })
            | (ExpectedLimit::Nesting, DecodeLimit::Nesting { .. })
            | (ExpectedLimit::States, DecodeLimit::States { .. })
            | (ExpectedLimit::Allocations, DecodeLimit::Allocations { .. })
            | (
                ExpectedLimit::RetainedBytes,
                DecodeLimit::RetainedBytes { .. }
            )
            | (
                ExpectedLimit::ScratchBytes,
                DecodeLimit::ScratchBytes { .. }
            )
    );
    assert!(matches, "expected {expected:?}, got {limit:?}");
}

fn assert_decode_boundary(
    source: &[u8],
    exact: DecodeOptions,
    one_under: DecodeOptions,
    expected: ExpectedLimit,
) {
    litchi_iwa_protos::pages_hidden_state_codec::decode_hidden_states_owner(source, exact)
        .expect("the inclusive resource maximum must be accepted");
    let error =
        litchi_iwa_protos::pages_hidden_state_codec::decode_hidden_states_owner(source, one_under)
            .expect_err("one below a resource maximum must be rejected");
    assert_limit(error, expected);
}

#[test]
fn owner_decode_accepts_exact_and_rejects_one_under_for_every_exposed_axis() {
    let source = owner_source();
    let (owner, report) =
        litchi_iwa_protos::pages_hidden_state_codec::decode_hidden_states_owner_with_report(
            &source,
            generous_options(&source),
        )
        .expect("fixture owner must decode");
    assert_eq!(owner.hidden_states().len(), 1);
    assert!(report.input_bytes() > 0);
    assert!(report.output_bytes() > 0);
    assert!(report.fields() > 0);
    assert!(report.work_bytes() > 0);
    assert!(report.max_depth() > 1);
    assert!(report.states() > 0);
    assert!(report.allocations() > 0);
    assert!(report.retained_bytes() > 0);
    assert!(report.scratch_bytes() > 0);

    assert_decode_boundary(
        &source,
        generous_options(&source),
        generous_options(&source).with_max_input_bytes(report.input_bytes() - 1),
        ExpectedLimit::InputBytes,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source),
        generous_options(&source).with_max_output_bytes(report.output_bytes() - 1),
        ExpectedLimit::OutputBytes,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_max_fields(report.fields()),
        generous_options(&source).with_max_fields(report.fields() - 1),
        ExpectedLimit::Fields,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_max_work_bytes(report.work_bytes()),
        generous_options(&source).with_max_work_bytes(report.work_bytes() - 1),
        ExpectedLimit::WorkBytes,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_recursion_limit(report.max_depth()),
        generous_options(&source).with_recursion_limit(report.max_depth() - 1),
        ExpectedLimit::Nesting,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_max_states(report.states()),
        generous_options(&source).with_max_states(report.states() - 1),
        ExpectedLimit::States,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_max_allocations(report.allocations()),
        generous_options(&source).with_max_allocations(report.allocations() - 1),
        ExpectedLimit::Allocations,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_max_retained_bytes(report.retained_bytes()),
        generous_options(&source).with_max_retained_bytes(report.retained_bytes() - 1),
        ExpectedLimit::RetainedBytes,
    );
    assert_decode_boundary(
        &source,
        generous_options(&source).with_max_scratch_bytes(report.scratch_bytes()),
        generous_options(&source).with_max_scratch_bytes(report.scratch_bytes() - 1),
        ExpectedLimit::ScratchBytes,
    );
}

#[test]
fn prepared_owner_rewrite_accepts_exact_and_rejects_one_under_execution_limits() {
    let source = owner_source();
    let (owner, _) =
        litchi_iwa_protos::pages_hidden_state_codec::decode_hidden_states_owner_with_report(
            &source,
            generous_options(&source),
        )
        .expect("fixture owner must decode");
    let prepared =
        litchi_iwa_protos::pages_hidden_state_codec::prepare_hidden_states_owner_rewrite(
            &source,
            &owner,
            generous_options(&source),
        )
        .expect("fixture owner must prepare");
    let requirements = prepared.execution_requirements();
    assert!(requirements.output_bytes() > 0);
    assert!(requirements.fields() > 0);
    assert!(requirements.work_bytes() > 0);
    assert!(requirements.max_depth() > 1);
    assert!(requirements.states() > 0);
    assert!(requirements.allocations() > 0);
    assert!(requirements.retained_bytes() > 0);
    assert!(requirements.scratch_bytes() > 0);

    prepared
        .clone()
        .execute(RewriteExecutionLimits::exact(requirements))
        .expect("exact execution limits must be inclusive");

    let cases = [
        (
            RewriteExecutionLimits::exact(requirements)
                .with_output_bytes(requirements.output_bytes() - 1),
            ExpectedLimit::OutputBytes,
        ),
        (
            RewriteExecutionLimits::exact(requirements).with_fields(requirements.fields() - 1),
            ExpectedLimit::Fields,
        ),
        (
            RewriteExecutionLimits::exact(requirements)
                .with_work_bytes(requirements.work_bytes() - 1),
            ExpectedLimit::WorkBytes,
        ),
        (
            RewriteExecutionLimits::exact(requirements)
                .with_max_depth(requirements.max_depth() - 1),
            ExpectedLimit::Nesting,
        ),
        (
            RewriteExecutionLimits::exact(requirements).with_states(requirements.states() - 1),
            ExpectedLimit::States,
        ),
        (
            RewriteExecutionLimits::exact(requirements)
                .with_allocations(requirements.allocations() - 1),
            ExpectedLimit::Allocations,
        ),
        (
            RewriteExecutionLimits::exact(requirements)
                .with_retained_bytes(requirements.retained_bytes() - 1),
            ExpectedLimit::RetainedBytes,
        ),
        (
            RewriteExecutionLimits::exact(requirements)
                .with_scratch_bytes(requirements.scratch_bytes() - 1),
            ExpectedLimit::ScratchBytes,
        ),
    ];

    for (limits, expected) in cases {
        let error = prepared
            .clone()
            .execute(limits)
            .expect_err("one below a rewrite requirement must be rejected");
        assert_limit(error, expected);
    }
}

#[test]
fn rewrite_reports_are_stable_at_the_inclusive_boundary() {
    let source = owner_source();
    let options = generous_options(&source);
    let owner =
        litchi_iwa_protos::pages_hidden_state_codec::decode_hidden_states_owner(&source, options)
            .expect("fixture owner must decode");
    let output = litchi_iwa_protos::pages_hidden_state_codec::rewrite_hidden_states_owner(
        &source, &owner, options,
    )
    .expect("exact internal rewrite limits must be inclusive");
    assert_eq!(
        litchi_iwa_protos::pages_hidden_state_codec::decode_hidden_states_owner(
            output.bytes(),
            generous_options(output.bytes()),
        )
        .expect("rewritten owner must decode"),
        owner
    );
}

#[allow(dead_code)]
fn _type_surface_smoke(owner: &HiddenStatesOwnerSnapshot) {
    let _ = UuidSnapshot::new(1, 2);
    let state = RowOrColumnStateSnapshot::new(UuidSnapshot::new(3, 4));
    let _ = HiddenStateExtentSnapshot::new(UuidSnapshot::new(5, 6), AxisDirection::Row, [state]);
    let _ = HiddenStatesSnapshot::new(
        UuidSnapshot::new(7, 8),
        HiddenStateExtentSnapshot::new(UuidSnapshot::new(9, 10), AxisDirection::Column, [])
            .expect("extent"),
        HiddenStateExtentSnapshot::new(UuidSnapshot::new(11, 12), AxisDirection::Row, [])
            .expect("extent"),
    );
    let _ = owner;
}
