//! Native public-PPT commit phase attribution.
//!
//! This module deliberately reuses the inventory, semantic oracle, and
//! corruption controls in the parent probe.  The phase routes differ only in
//! where the outer clocks and the opt-in commit observer are placed.

use super::*;

use serde::Serialize;
use std::hint::black_box;
use std::path::PathBuf;
use std::time::Instant;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PhaseRoute {
    OrdinaryOpaque,
    OrdinarySplit,
    ProfiledEmpty,
    ProfiledClock,
}

impl PhaseRoute {
    fn name(self) -> &'static str {
        match self {
            Self::OrdinaryOpaque => "ordinary-opaque",
            Self::OrdinarySplit => "ordinary-split",
            Self::ProfiledEmpty => "profiled-empty",
            Self::ProfiledClock => "profiled-clock",
        }
    }

    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "ordinary-opaque" => Ok(Self::OrdinaryOpaque),
            "ordinary-split" => Ok(Self::OrdinarySplit),
            "profiled-empty" => Ok(Self::ProfiledEmpty),
            "profiled-clock" => Ok(Self::ProfiledClock),
            other => failure(format!("unknown --route {other:?}")),
        }
    }
}

#[derive(Clone)]
struct PhaseArgs {
    input: PathBuf,
    route: PhaseRoute,
    warmups: usize,
    samples: usize,
    text: String,
}

fn parse_phase_args() -> Result<PhaseArgs, BoxError> {
    let mut input = None;
    let mut route = PhaseRoute::OrdinaryOpaque;
    let mut warmups = 1usize;
    let mut samples = 1usize;
    let mut text = String::from("litchi copy-through baseline replacement text");

    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut next_value = || {
            arguments.next().ok_or_else(|| {
                Box::new(ProbeError(format!("missing value for {flag}"))) as BoxError
            })
        };
        match flag.as_str() {
            "--input" => input = Some(PathBuf::from(next_value()?)),
            "--route" => route = PhaseRoute::parse(&next_value()?)?,
            "--warmups" => warmups = parse_count("--warmups", &next_value()?)?,
            "--samples" => samples = parse_count("--samples", &next_value()?)?,
            "--text" => text = next_value()?,
            "--case" => {
                let case = next_value()?;
                if case != "ppt45543" {
                    return failure(format!(
                        "the 0732 phase probe only measures fixed case ppt45543, got {case:?}"
                    ));
                }
            },
            "--oracle-only" => {
                warmups = 0;
                samples = 0;
            },
            other => return failure(format!("unknown flag {other:?}")),
        }
    }

    Ok(PhaseArgs {
        input: input.ok_or_else(|| Box::new(ProbeError("missing --input".into())) as BoxError)?,
        route,
        warmups,
        samples,
        text,
    })
}

fn phase_public_args(args: &PhaseArgs) -> Args {
    Args {
        case: Case::Ppt45543,
        input: args.input.clone(),
        operation: Operation::Format,
        policy: Policy::Reuse,
        warmups: args.warmups,
        samples: args.samples,
        text: args.text.clone(),
    }
}

#[derive(Clone, Debug, Serialize)]
struct SplitTiming {
    open_ns: u128,
    edit_ns: u128,
    remove_ns: u128,
    commit_ns: u128,
    output_copy_ns: u128,
    split_sum_ns: u128,
    whole_residual_ns: u128,
    /// Offset from the outer lifecycle start to the commit window start.
    commit_start_ns: u128,
    /// Offset from the outer lifecycle start to the commit window end.
    commit_end_ns: u128,
}

fn split_timing(
    whole_ns: u128,
    open_ns: u128,
    edit_ns: u128,
    remove_ns: u128,
    commit_ns: u128,
    output_copy_ns: u128,
    commit_window_ns: (u128, u128),
) -> SplitTiming {
    let (commit_start_ns, commit_end_ns) = commit_window_ns;
    let split_sum_ns = open_ns
        .saturating_add(edit_ns)
        .saturating_add(remove_ns)
        .saturating_add(commit_ns)
        .saturating_add(output_copy_ns);
    SplitTiming {
        open_ns,
        edit_ns,
        remove_ns,
        commit_ns,
        output_copy_ns,
        split_sum_ns,
        whole_residual_ns: whole_ns.saturating_sub(split_sum_ns),
        commit_start_ns,
        commit_end_ns,
    }
}

struct TimedPhaseRoute {
    output: Vec<u8>,
    whole_ns: u128,
    split: Option<SplitTiming>,
    diagnostics: Option<DiagnosticTrace>,
    observer_clock_control_ns: Option<u128>,
}

fn ordinary_opaque_phase(source: &[u8], args: &PhaseArgs) -> Result<TimedPhaseRoute, BoxError> {
    let public_args = phase_public_args(args);
    let whole_start = Instant::now();
    // Keep this route on the original public-format function.  The only
    // timed operation is its complete open/edit/remove/commit lifecycle.
    let output = measured_public_format(source, &public_args)?;
    black_box(output.len());
    let whole_ns = whole_start.elapsed().as_nanos();
    Ok(TimedPhaseRoute {
        output,
        whole_ns,
        split: None,
        diagnostics: None,
        observer_clock_control_ns: None,
    })
}

fn ordinary_split_phase(source: &[u8], _args: &PhaseArgs) -> Result<TimedPhaseRoute, BoxError> {
    let whole_start = Instant::now();
    let (output, split) = {
        let open_start = Instant::now();
        let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
        let open_ns = open_start.elapsed().as_nanos();

        let (output, timings) = {
            let edit_start = Instant::now();
            let mut edit = snapshot.edit()?;
            let edit_ns = edit_start.elapsed().as_nanos();

            let remove_start = Instant::now();
            edit.remove_slide(Position::new(1))?;
            let remove_ns = remove_start.elapsed().as_nanos();

            let commit_start = Instant::now();
            let commit = edit.commit()?;
            let commit_end = Instant::now();
            let commit_ns = commit_end.duration_since(commit_start).as_nanos();
            let commit_start_ns = commit_start.duration_since(whole_start).as_nanos();
            let commit_end_ns = commit_end.duration_since(whole_start).as_nanos();

            let output_start = Instant::now();
            let output = commit.snapshot().bytes().to_vec();
            let output_copy_ns = output_start.elapsed().as_nanos();
            (
                output,
                (
                    edit_ns,
                    remove_ns,
                    commit_ns,
                    output_copy_ns,
                    commit_start_ns,
                    commit_end_ns,
                ),
            )
        };
        (
            output,
            split_timing(
                0,
                open_ns,
                timings.0,
                timings.1,
                timings.2,
                timings.3,
                (timings.4, timings.5),
            ),
        )
    };
    let whole_ns = whole_start.elapsed().as_nanos();
    let mut split = split;
    split.whole_residual_ns = whole_ns.saturating_sub(split.split_sum_ns);
    Ok(TimedPhaseRoute {
        output,
        whole_ns,
        split: Some(split),
        diagnostics: None,
        observer_clock_control_ns: None,
    })
}

fn profiled_empty_phase(source: &[u8], _args: &PhaseArgs) -> Result<TimedPhaseRoute, BoxError> {
    let whole_start = Instant::now();
    let (output, split) = {
        let open_start = Instant::now();
        let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
        let open_ns = open_start.elapsed().as_nanos();

        let (output, timings) = {
            let edit_start = Instant::now();
            let mut edit = snapshot.edit()?;
            let edit_ns = edit_start.elapsed().as_nanos();

            let remove_start = Instant::now();
            edit.remove_slide(Position::new(1))?;
            let remove_ns = remove_start.elapsed().as_nanos();

            let commit_start = Instant::now();
            let commit = edit.commit_profiled(|_| {})?;
            let commit_end = Instant::now();
            let commit_ns = commit_end.duration_since(commit_start).as_nanos();
            let commit_start_ns = commit_start.duration_since(whole_start).as_nanos();
            let commit_end_ns = commit_end.duration_since(whole_start).as_nanos();

            let output_start = Instant::now();
            let output = commit.snapshot().bytes().to_vec();
            let output_copy_ns = output_start.elapsed().as_nanos();
            (
                output,
                (
                    edit_ns,
                    remove_ns,
                    commit_ns,
                    output_copy_ns,
                    commit_start_ns,
                    commit_end_ns,
                ),
            )
        };
        (
            output,
            split_timing(
                0,
                open_ns,
                timings.0,
                timings.1,
                timings.2,
                timings.3,
                (timings.4, timings.5),
            ),
        )
    };
    let whole_ns = whole_start.elapsed().as_nanos();
    let mut split = split;
    split.whole_residual_ns = whole_ns.saturating_sub(split.split_sum_ns);
    Ok(TimedPhaseRoute {
        output,
        whole_ns,
        split: Some(split),
        diagnostics: None,
        observer_clock_control_ns: None,
    })
}

#[derive(Clone, Copy)]
struct RawPhaseEvent {
    started: bool,
    phase: u8,
    outcome: u8,
    stamp_ns: u128,
}

struct ClockTrace {
    base: Instant,
    events: [Option<RawPhaseEvent>; 32],
    len: usize,
    overflow: bool,
}

impl ClockTrace {
    fn new(base: Instant) -> Self {
        Self {
            base,
            events: [None; 32],
            len: 0,
            overflow: false,
        }
    }

    fn record(&mut self, event: litchi_ppt::slide_order::DiagnosticEvent) {
        let (started, phase, outcome) = match event {
            litchi_ppt::slide_order::DiagnosticEvent::Started { phase } => {
                (true, diagnostic_phase_code(phase), 0)
            },
            litchi_ppt::slide_order::DiagnosticEvent::Finished { phase, outcome } => (
                false,
                diagnostic_phase_code(phase),
                match outcome {
                    litchi_ppt::slide_order::DiagnosticOutcome::Success => 1,
                    litchi_ppt::slide_order::DiagnosticOutcome::Error => 2,
                },
            ),
        };
        let event = RawPhaseEvent {
            started,
            phase,
            outcome,
            stamp_ns: self.base.elapsed().as_nanos(),
        };
        if let Some(slot) = self.events.get_mut(self.len) {
            *slot = Some(event);
        } else {
            self.overflow = true;
        }
        self.len = self.len.saturating_add(1);
    }
}

fn diagnostic_phase_code(phase: litchi_ppt::slide_order::DiagnosticPhase) -> u8 {
    use litchi_ppt::slide_order::DiagnosticPhase;
    match phase {
        DiagnosticPhase::DocumentCommit => 1,
        DiagnosticPhase::BeforePayloadCapture => 2,
        DiagnosticPhase::EmbeddedOpen => 3,
        DiagnosticPhase::LiveDocumentRead => 4,
        DiagnosticPhase::EmbeddedFinish => 5,
        DiagnosticPhase::UnrelatedStreamValidation => 6,
        DiagnosticPhase::PublicReopen => 7,
        DiagnosticPhase::AfterPayloadCapture => 8,
        DiagnosticPhase::ArtifactHashBefore => 9,
        DiagnosticPhase::ArtifactHashAfter => 10,
        DiagnosticPhase::StructuralNoOp => 11,
    }
}

fn diagnostic_phase_name(code: u8) -> &'static str {
    match code {
        1 => "DocumentCommit",
        2 => "BeforePayloadCapture",
        3 => "EmbeddedOpen",
        4 => "LiveDocumentRead",
        5 => "EmbeddedFinish",
        6 => "UnrelatedStreamValidation",
        7 => "PublicReopen",
        8 => "AfterPayloadCapture",
        9 => "ArtifactHashBefore",
        10 => "ArtifactHashAfter",
        11 => "StructuralNoOp",
        _ => "unknown",
    }
}

fn diagnostic_outcome_name(outcome: u8) -> &'static str {
    match outcome {
        0 => "started",
        1 => "success",
        2 => "error",
        _ => "unknown",
    }
}

#[derive(Clone, Debug, Serialize)]
struct DiagnosticEventReport {
    kind: String,
    phase: String,
    outcome: String,
    t_ns: u128,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct DiagnosticSpanReport {
    phase: String,
    outcome: String,
    start_ns: u128,
    finish_ns: u128,
    duration_ns: u128,
}

#[derive(Clone, Debug, Serialize)]
struct DiagnosticTraceReport {
    event_count: usize,
    overflow: bool,
    balanced: bool,
    sequence_ok: bool,
    outcomes_ok: bool,
    timestamps_monotonic: bool,
    expected_phases: Vec<String>,
    events: Vec<DiagnosticEventReport>,
    spans: Vec<DiagnosticSpanReport>,
}

#[derive(Clone, Debug, Serialize)]
struct DiagnosticTrace {
    commit: DiagnosticTraceReport,
}

fn commit_phase_codes() -> [u8; 10] {
    [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
}

fn trace_report(trace: &ClockTrace, expected: &[u8]) -> DiagnosticTraceReport {
    let stored = trace.len.min(trace.events.len());
    let mut events = Vec::with_capacity(stored);
    let mut timestamps_monotonic = !trace.overflow;
    let mut previous = 0u128;
    for raw in trace.events.iter().take(stored).flatten() {
        timestamps_monotonic &= raw.stamp_ns >= previous;
        previous = raw.stamp_ns;
        events.push(DiagnosticEventReport {
            kind: if raw.started { "started" } else { "finished" }.into(),
            phase: diagnostic_phase_name(raw.phase).into(),
            outcome: diagnostic_outcome_name(raw.outcome).into(),
            t_ns: raw.stamp_ns,
        });
    }

    let mut stack = [0u8; 32];
    let mut depth = 0usize;
    let mut balanced = !trace.overflow;
    for raw in trace.events.iter().take(stored).flatten() {
        if raw.started {
            if let Some(slot) = stack.get_mut(depth) {
                *slot = raw.phase;
                depth += 1;
            } else {
                balanced = false;
            }
        } else if depth == 0 {
            balanced = false;
        } else {
            depth -= 1;
            if stack[depth] != raw.phase {
                balanced = false;
            }
        }
    }
    balanced &= depth == 0 && trace.len == expected.len().saturating_mul(2);

    let outcomes_ok = !trace.overflow
        && trace.len == expected.len().saturating_mul(2)
        && expected.iter().enumerate().all(|(index, _)| {
            let Some(start) = trace.events.get(index * 2).copied().flatten() else {
                return false;
            };
            let Some(finish) = trace.events.get(index * 2 + 1).copied().flatten() else {
                return false;
            };
            start.started && start.outcome == 0 && !finish.started && finish.outcome == 1
        });

    let sequence_ok = !trace.overflow
        && trace.len == expected.len().saturating_mul(2)
        && expected.iter().enumerate().all(|(index, phase)| {
            let Some(start) = trace.events.get(index * 2).copied().flatten() else {
                return false;
            };
            let Some(finish) = trace.events.get(index * 2 + 1).copied().flatten() else {
                return false;
            };
            start.started
                && start.outcome == 0
                && !finish.started
                && start.phase == *phase
                && finish.phase == *phase
                && finish.outcome == 1
        });

    let mut spans = Vec::with_capacity(expected.len());
    for index in 0..expected.len() {
        let Some(start) = trace.events.get(index * 2).copied().flatten() else {
            continue;
        };
        let Some(finish) = trace.events.get(index * 2 + 1).copied().flatten() else {
            continue;
        };
        if start.started && !finish.started && start.phase == finish.phase {
            spans.push(DiagnosticSpanReport {
                phase: diagnostic_phase_name(start.phase).into(),
                outcome: diagnostic_outcome_name(finish.outcome).into(),
                start_ns: start.stamp_ns,
                finish_ns: finish.stamp_ns,
                duration_ns: finish.stamp_ns.saturating_sub(start.stamp_ns),
            });
        }
    }

    DiagnosticTraceReport {
        event_count: trace.len,
        overflow: trace.overflow,
        balanced,
        sequence_ok,
        outcomes_ok,
        timestamps_monotonic,
        expected_phases: expected
            .iter()
            .map(|phase| diagnostic_phase_name(*phase).into())
            .collect(),
        events,
        spans,
    }
}

fn diagnostic_phase_from_code(code: u8) -> litchi_ppt::slide_order::DiagnosticPhase {
    use litchi_ppt::slide_order::DiagnosticPhase;
    match code {
        1 => DiagnosticPhase::DocumentCommit,
        2 => DiagnosticPhase::BeforePayloadCapture,
        3 => DiagnosticPhase::EmbeddedOpen,
        4 => DiagnosticPhase::LiveDocumentRead,
        5 => DiagnosticPhase::EmbeddedFinish,
        6 => DiagnosticPhase::UnrelatedStreamValidation,
        7 => DiagnosticPhase::PublicReopen,
        8 => DiagnosticPhase::AfterPayloadCapture,
        9 => DiagnosticPhase::ArtifactHashBefore,
        10 => DiagnosticPhase::ArtifactHashAfter,
        _ => DiagnosticPhase::StructuralNoOp,
    }
}

fn observer_clock_control_ns() -> u128 {
    let base = Instant::now();
    let start = Instant::now();
    let mut trace = ClockTrace::new(base);
    for phase in commit_phase_codes() {
        trace.record(litchi_ppt::slide_order::DiagnosticEvent::Started {
            phase: diagnostic_phase_from_code(phase),
        });
        trace.record(litchi_ppt::slide_order::DiagnosticEvent::Finished {
            phase: diagnostic_phase_from_code(phase),
            outcome: litchi_ppt::slide_order::DiagnosticOutcome::Success,
        });
    }
    black_box(&trace);
    start.elapsed().as_nanos()
}

fn profiled_clock_phase(source: &[u8], _args: &PhaseArgs) -> Result<TimedPhaseRoute, BoxError> {
    let whole_start = Instant::now();
    let mut commit_trace = ClockTrace::new(whole_start);
    let (output, split) = {
        let open_start = Instant::now();
        let snapshot = litchi_ppt::slide_order::Snapshot::from_bytes(source.to_vec())?;
        let open_ns = open_start.elapsed().as_nanos();

        let (output, timings) = {
            let edit_start = Instant::now();
            let mut edit = snapshot.edit()?;
            let edit_ns = edit_start.elapsed().as_nanos();

            let remove_start = Instant::now();
            edit.remove_slide(Position::new(1))?;
            let remove_ns = remove_start.elapsed().as_nanos();

            let commit_start = Instant::now();
            let commit = edit.commit_profiled(|event| commit_trace.record(event))?;
            let commit_end = Instant::now();
            let commit_ns = commit_end.duration_since(commit_start).as_nanos();
            let commit_start_ns = commit_start.duration_since(whole_start).as_nanos();
            let commit_end_ns = commit_end.duration_since(whole_start).as_nanos();

            let output_start = Instant::now();
            let output = commit.snapshot().bytes().to_vec();
            let output_copy_ns = output_start.elapsed().as_nanos();
            (
                output,
                (
                    edit_ns,
                    remove_ns,
                    commit_ns,
                    output_copy_ns,
                    commit_start_ns,
                    commit_end_ns,
                ),
            )
        };
        (
            output,
            split_timing(
                0,
                open_ns,
                timings.0,
                timings.1,
                timings.2,
                timings.3,
                (timings.4, timings.5),
            ),
        )
    };
    let whole_ns = whole_start.elapsed().as_nanos();
    let mut split = split;
    split.whole_residual_ns = whole_ns.saturating_sub(split.split_sum_ns);
    let diagnostics = DiagnosticTrace {
        commit: trace_report(&commit_trace, &commit_phase_codes()),
    };
    Ok(TimedPhaseRoute {
        output,
        whole_ns,
        split: Some(split),
        diagnostics: Some(diagnostics),
        observer_clock_control_ns: Some(observer_clock_control_ns()),
    })
}

fn execute_phase_route(source: &[u8], args: &PhaseArgs) -> Result<TimedPhaseRoute, BoxError> {
    match args.route {
        PhaseRoute::OrdinaryOpaque => ordinary_opaque_phase(source, args),
        PhaseRoute::OrdinarySplit => ordinary_split_phase(source, args),
        PhaseRoute::ProfiledEmpty => profiled_empty_phase(source, args),
        PhaseRoute::ProfiledClock => profiled_clock_phase(source, args),
    }
}

#[derive(Clone, Debug, Serialize)]
struct PhaseRouteSample {
    index: usize,
    route: String,
    whole_ns: u128,
    split: Option<SplitTiming>,
    output_sha256: String,
    output_inventory: InventorySummary,
    oracle: Oracle,
    diagnostics: Option<DiagnosticTrace>,
    observer_clock_control_ns: Option<u128>,
}

#[derive(Debug, Serialize)]
struct PhaseProbeOutput {
    schema_version: u32,
    mode: String,
    case: String,
    format: String,
    operation: String,
    scope: String,
    route: String,
    policy: String,
    policy_applied: bool,
    policy_application_scope: String,
    policy_argument_effect: String,
    policy_contract: String,
    phase_contract: String,
    diagnostic_contract: String,
    observer_contract: String,
    input: String,
    text_utf16_units: usize,
    timing_claim: bool,
    allocator_instrumented: bool,
    allocation_ownership_contract: String,
    directory_metadata_fields: Vec<String>,
    warmups: usize,
    samples_requested: usize,
    source_sha256: String,
    expected_output_sha256: String,
    replacements_sha256: String,
    source_inventory: InventorySummary,
    expected_output_inventory: InventorySummary,
    replacements: Vec<ReplacementSummary>,
    changed_length_proof: ChangedLengthProof,
    expected_oracle: Oracle,
    oracle_controls: Vec<OracleControl>,
    samples: Vec<PhaseRouteSample>,
}

pub fn run_phase() -> Result<(), BoxError> {
    let args = parse_phase_args()?;
    let public_args = phase_public_args(&args);
    let source_bytes = std::fs::read(&args.input)?;
    let source = inventory(&source_bytes)?;
    let expected_bytes = public_format_edit(&source_bytes, &public_args)?;
    let expected = inventory(&expected_bytes)?;
    let replacements = derive_replacements(&source, &expected, Case::Ppt45543)?;
    let replacement_paths = replacements
        .iter()
        .map(|replacement| replacement.path.clone())
        .collect::<std::collections::BTreeSet<_>>();
    let mut length_proof = changed_length_proof(&source, &expected);
    let expected_oracle = oracle_for_output(
        &public_args,
        &source,
        &expected,
        &expected,
        &source_bytes,
        &expected_bytes,
        &expected_bytes,
        &replacement_paths,
        false,
    );
    finalize_length_proof(Case::Ppt45543, &expected_oracle, &mut length_proof);
    let oracle_controls = oracle_controls(
        &public_args,
        &source,
        &expected,
        &source_bytes,
        &expected_bytes,
        &replacement_paths,
    );
    if !expected_oracle.oracle_ok || !length_proof.logical_stream_length_change_proven {
        return failure(format!(
            "public PPT oracle failed before measurement: {:?}; length_proof={length_proof:?}",
            expected_oracle.failure_reasons
        ));
    }

    for _ in 0..args.warmups {
        let result = execute_phase_route(&source_bytes, &args)?;
        drop(result.output);
    }

    let mut samples = Vec::with_capacity(args.samples);
    for index in 0..args.samples {
        let result = execute_phase_route(&source_bytes, &args)?;
        if result.whole_ns == 0 {
            return failure("PPT phase route reported a zero outer lifecycle");
        }
        if let Some(split) = &result.split
            && (split.commit_end_ns < split.commit_start_ns
                || split.commit_end_ns > result.whole_ns)
        {
            return failure("PPT phase route reported an invalid commit window");
        }
        if let Some(diagnostics) = &result.diagnostics {
            let report = &diagnostics.commit;
            if !report.balanced
                || !report.sequence_ok
                || !report.outcomes_ok
                || !report.timestamps_monotonic
            {
                return failure("PPT profiled-clock trace failed closed validation");
            }
            let split = result.split.as_ref().ok_or_else(|| {
                Box::new(ProbeError("profiled clock has no split".into())) as BoxError
            })?;
            if report
                .events
                .iter()
                .any(|event| event.t_ns < split.commit_start_ns || event.t_ns > split.commit_end_ns)
            {
                return failure("PPT diagnostic event escaped the commit window");
            }
            if report.spans.iter().any(|span| {
                span.start_ns < split.commit_start_ns || span.finish_ns > split.commit_end_ns
            }) {
                return failure("PPT diagnostic event escaped the commit window");
            }
        }
        let output_sha256 = sha256_hex(&result.output);
        let output_inventory = inventory(&result.output)?;
        let mut oracle = oracle_for_output(
            &public_args,
            &source,
            &expected,
            &output_inventory,
            &source_bytes,
            &expected_bytes,
            &result.output,
            &replacement_paths,
            true,
        );
        if !length_proof.logical_stream_length_change_proven {
            oracle
                .failure_reasons
                .push("public edit did not prove a logical length-changing stream edit".into());
            oracle.oracle_ok = false;
        }
        samples.push(PhaseRouteSample {
            index,
            route: args.route.name().into(),
            whole_ns: result.whole_ns,
            split: result.split,
            output_sha256,
            output_inventory: output_inventory.summary,
            oracle,
            diagnostics: result.diagnostics,
            observer_clock_control_ns: result.observer_clock_control_ns,
        });
    }

    let result = PhaseProbeOutput {
        schema_version: 1,
        mode: "ppt_public_phase_attribution".into(),
        case: Case::Ppt45543.name().into(),
        format: "ppt".into(),
        operation: "format".into(),
        scope: "public_ppt_open_edit_remove_commit_output_copy".into(),
        route: args.route.name().into(),
        policy: "reuse".into(),
        policy_applied: false,
        policy_application_scope: "not_applied_public_format_route".into(),
        policy_argument_effect: "ignored_public_format_default_route".into(),
        policy_contract: "Reuse preserves the normalized raw directory image outside planner-owned allocation fields; Rewrite may normalize physical directory layout and timestamps. Both policies must preserve logical streams, semantic directory metadata, root/storage CLSIDs, and public edit meaning; raw differences are reported in raw_directory and directory_metadata_differences".into(),
        phase_contract: "whole_ns encloses the complete public PPT lifecycle and all local owner drops; ordinary-opaque is the original public_format_edit lifecycle, while split routes clock source open, edit construction, slide removal, commit, and output byte extraction. split_sum_ns is the sum of those windows; whole_residual_ns is checked nonnegative. commit_start_ns and commit_end_ns are offsets into whole_ns and bound profiled diagnostic events".into(),
        diagnostic_contract: "profiled-empty calls the feature-gated PPT commit_profiled API with a no-op observer and keeps ordinary Snapshot::from_bytes open semantics. profiled-clock calls the same API and timestamps only content-free DiagnosticEvent callbacks in a fixed-capacity stack; sequence, balance, outcome, and timestamp validation happen after the timed lifecycle".into(),
        observer_contract: "observer_clock_control_ns measures the fixed-capacity timestamp callback recorder against the expected changed-commit event count outside the workflow. It is an observer-cost control and is not subtracted from workflow timings".into(),
        input: args.input.display().to_string(),
        text_utf16_units: args.text.encode_utf16().count(),
        timing_claim: true,
        allocator_instrumented: alloc_metrics::instrumented(),
        allocation_ownership_contract: "This phase binary is uninstrumented. Snapshot, edit, commit, and local owner drops remain inside each whole lifecycle; the returned output Vec is retained only for untimed oracle validation. No allocation, peak-memory, or RSS claim is made.".into(),
        directory_metadata_fields: vec![
            "entry_type".into(),
            "name_utf16".into(),
            "clsid".into(),
            "bytes".into(),
            "start_sector".into(),
            "is_minifat".into(),
            "raw_left_sibling".into(),
            "raw_right_sibling".into(),
            "raw_child".into(),
            "raw_color".into(),
            "raw_state_bits".into(),
            "raw_creation_time".into(),
            "raw_modification_time".into(),
        ],
        warmups: args.warmups,
        samples_requested: args.samples,
        source_sha256: sha256_hex(&source_bytes),
        expected_output_sha256: sha256_hex(&expected_bytes),
        replacements_sha256: replacement_digest(&replacements),
        replacements: replacement_summaries(&source, &expected, &replacements),
        source_inventory: source.summary,
        expected_output_inventory: expected.summary,
        changed_length_proof: length_proof,
        expected_oracle,
        oracle_controls,
        samples,
    };
    println!("{}", serde_json::to_string(&result)?);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_ppt::slide_order::{DiagnosticEvent, DiagnosticOutcome};

    fn record_success(trace: &mut ClockTrace, phases: &[u8]) {
        for phase in phases {
            trace.record(DiagnosticEvent::Started {
                phase: diagnostic_phase_from_code(*phase),
            });
            trace.record(DiagnosticEvent::Finished {
                phase: diagnostic_phase_from_code(*phase),
                outcome: DiagnosticOutcome::Success,
            });
        }
    }

    #[test]
    fn ppt_trace_accepts_the_complete_changed_commit_sequence() {
        let mut trace = ClockTrace::new(Instant::now());
        let expected = commit_phase_codes();
        record_success(&mut trace, &expected);
        let report = trace_report(&trace, &expected);
        assert!(report.balanced);
        assert!(report.sequence_ok);
        assert!(report.outcomes_ok);
        assert!(report.timestamps_monotonic);
        assert_eq!(report.event_count, expected.len() * 2);
        assert_eq!(report.spans.len(), expected.len());
    }

    #[test]
    fn ppt_trace_rejects_a_balanced_but_reordered_event_sequence() {
        let mut trace = ClockTrace::new(Instant::now());
        let expected = commit_phase_codes();
        let mut reversed = expected;
        reversed.reverse();
        record_success(&mut trace, &reversed);
        let report = trace_report(&trace, &expected);
        assert!(report.balanced);
        assert!(!report.sequence_ok);
        assert!(report.outcomes_ok);
    }

    #[test]
    fn ppt_trace_rejects_a_missing_event() {
        let mut trace = ClockTrace::new(Instant::now());
        let expected = commit_phase_codes();
        record_success(&mut trace, &expected[..expected.len() - 1]);
        let report = trace_report(&trace, &expected);
        assert!(!report.balanced);
        assert!(!report.sequence_ok);
        assert!(!report.outcomes_ok);
    }

    #[test]
    fn ppt_trace_rejects_an_error_outcome() {
        let mut trace = ClockTrace::new(Instant::now());
        let expected = commit_phase_codes();
        for (index, phase) in expected.iter().enumerate() {
            trace.record(DiagnosticEvent::Started {
                phase: diagnostic_phase_from_code(*phase),
            });
            trace.record(DiagnosticEvent::Finished {
                phase: diagnostic_phase_from_code(*phase),
                outcome: if index == 4 {
                    DiagnosticOutcome::Error
                } else {
                    DiagnosticOutcome::Success
                },
            });
        }
        let report = trace_report(&trace, &expected);
        assert!(report.balanced);
        assert!(!report.sequence_ok);
        assert!(!report.outcomes_ok);
    }

    #[test]
    fn ppt_split_residual_is_saturating_and_never_negative() {
        let split = split_timing(100, 10, 20, 30, 10, 10, (40, 50));
        assert_eq!(split.split_sum_ns, 80);
        assert_eq!(split.whole_residual_ns, 20);
        let over = split_timing(50, 20, 20, 20, 20, 20, (20, 40));
        assert_eq!(over.whole_residual_ns, 0);
    }
}
