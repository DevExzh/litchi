//! Candidate-only profile for the public `Evaluated::to_owned` boundary.
//!
//! The `own` phase parses and evaluates once before starting its timer.  Its
//! timed body only converts that prepared borrowed result to an
//! `OwnedEvaluated`, inspects a structural checksum and the retained
//! reservation, then drops the owned result.  `setup` is a separate repeated
//! parse/evaluate preparation control.  `parse-evaluate-own` intentionally
//! includes all three operations in every timed call and is not an apples to
//! apples replacement for the isolated conversion phase.
//!
//! Allocator counters describe the timed batch.  `owned_reserved_bytes` is
//! the exact memory reservation reported while one owned result is live;
//! `peak_live_delta` is the allocator's live-byte high-water mark relative to
//! the pre-timer baseline, and `max_rss_kib` is supplied by the runner's
//! `/usr/bin/time -v` wrapper.  None of these measurements is a general
//! zero-copy or speedup claim.  Each invocation validates the known scalar,
//! array, and reference outputs, including reference lexical metadata, with
//! an untimed ownership copy before it begins sampling.  Warmups, iterations,
//! and repeats have finite command-line ceilings so a malformed invocation
//! cannot create an unbounded profile.

use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    error::Error,
    hint::black_box,
    num::{NonZeroU64, NonZeroUsize},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value, evaluate,
    },
    evaluation::{EvaluationFailure, ScalarError},
    expression::Expression,
    reference::{Address, Cell, Column, EndpointValue, Reference, SheetSelector},
};

type AnyResult<T> = Result<T, Box<dyn Error>>;

const MAX_WARMUPS: usize = 100;
const MAX_ITERATIONS: usize = 1_000;
const MAX_REPEAT: usize = 4_096;

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let live =
            LIVE_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed) + layout.size() as u64;
        record_peak(live);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(size as u64, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let old = layout.size() as u64;
        let new = size as u64;
        let live = if new >= old {
            LIVE_BYTES.fetch_add(new - old, Ordering::Relaxed) + (new - old)
        } else {
            LIVE_BYTES.fetch_sub(old - new, Ordering::Relaxed) - (old - new)
        };
        record_peak(live);
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn record_peak(candidate: u64) {
    let mut current = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while candidate > current {
        match PEAK_LIVE_BYTES.compare_exchange_weak(
            current,
            candidate,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => current = observed,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Phase {
    Setup,
    Own,
    ParseEvaluateOwn,
}

impl Phase {
    fn parse(value: &str) -> Self {
        match value {
            "setup" => Self::Setup,
            "own" => Self::Own,
            "parse-evaluate-own" => Self::ParseEvaluateOwn,
            other => panic!("unknown phase {other:?}"),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Setup => "setup",
            Self::Own => "own",
            Self::ParseEvaluateOwn => "parse-evaluate-own",
        }
    }
}

#[derive(Debug)]
struct Config {
    revision: String,
    workload: String,
    group: String,
    phase: Phase,
    case: String,
    warmups: usize,
    iterations: usize,
    repeat: usize,
}

impl Config {
    fn from_args() -> Self {
        let mut revision = None;
        let mut workload = None;
        let mut group = "all".to_owned();
        let mut phase = None;
        let mut case = None;
        let mut warmups = 3;
        let mut iterations = 15;
        let mut repeat = None;
        let mut args = env::args().skip(1);
        while let Some(argument) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {argument}"));
            match argument.as_str() {
                "--revision" => revision = Some(value),
                "--workload" => workload = Some(value),
                "--group" => group = value,
                "--phase" => phase = Some(Phase::parse(&value)),
                "--case" => case = Some(value),
                "--warmups" => {
                    warmups = parse_count("warmups", &value, 0, MAX_WARMUPS);
                },
                "--iterations" => {
                    iterations = parse_count("iterations", &value, 1, MAX_ITERATIONS);
                },
                "--repeat" => {
                    repeat = Some(parse_count("repeat", &value, 1, MAX_REPEAT));
                },
                other => panic!("unknown option {other}"),
            }
        }
        let case = case.unwrap_or_else(|| panic!("--case is required"));
        let repeat = repeat.unwrap_or_else(|| {
            let value = default_repeat(&case);
            if value == 0 || value > MAX_REPEAT {
                panic!("repeat default for {case:?} must be in 1..={MAX_REPEAT}");
            }
            value
        });
        Self {
            revision: revision.unwrap_or_else(|| panic!("--revision is required")),
            workload: workload.unwrap_or_else(|| panic!("--workload is required")),
            group,
            phase: phase.unwrap_or_else(|| panic!("--phase is required")),
            case,
            warmups,
            iterations,
            repeat,
        }
    }
}

fn parse_count(name: &str, value: &str, minimum: usize, maximum: usize) -> usize {
    let parsed = value
        .parse::<usize>()
        .unwrap_or_else(|_| panic!("{name} must be an unsigned integer"));
    if parsed < minimum || parsed > maximum {
        panic!("{name} must be in {minimum}..={maximum}");
    }
    parsed
}

#[derive(Clone, Copy, Debug)]
enum FixtureCell {
    Empty,
}

#[derive(Debug)]
struct FixtureResolver {
    extent: SheetExtent,
    cell: FixtureCell,
}

impl FixtureResolver {
    fn new() -> Self {
        Self {
            extent: SheetExtent::new(4096, 4096),
            cell: FixtureCell::Empty,
        }
    }
}

impl Resolver for FixtureResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        execution.check()?;
        Ok(match sheet {
            "Main" | "Data" | "Archive" => Some(self.extent),
            _ => None,
        })
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        _row: usize,
        _column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        execution.check()?;
        if matches!(sheet, "Main" | "Data" | "Archive") {
            return Ok(match self.cell {
                FixtureCell::Empty => CellRead::Empty,
            });
        }
        Ok(CellRead::Empty)
    }

    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        execution.check()?;
        Ok(match sheet {
            "Main" => Some(0),
            "Data" => Some(1),
            "Archive" => Some(2),
            _ => None,
        })
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        execution.check()?;
        Ok(match index {
            0 => Some("Main"),
            1 => Some("Data"),
            2 => Some("Archive"),
            _ => None,
        })
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        execution.check()?;
        Ok(3)
    }

    fn source_version(
        &self,
        execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        execution.check()?;
        Ok(None)
    }
}

fn scale(case: &str) -> Option<usize> {
    case.rsplit_once('-')?.1.parse().ok()
}

fn square_side(cells: usize) -> usize {
    let mut side = 1usize;
    while side.saturating_mul(side) < cells {
        side = side.saturating_add(1);
    }
    assert_eq!(side.saturating_mul(side), cells, "case size must be square");
    side
}

fn array_literal(size: usize) -> String {
    let side = square_side(size);
    let mut rows = Vec::with_capacity(side);
    for row in 0..side {
        let mut cells = Vec::with_capacity(side);
        for column in 0..side {
            cells.push((row * side + column + 1).to_string());
        }
        rows.push(cells.join(";"));
    }
    format!("={{{}}}", rows.join("|"))
}

fn reference_list(size: usize, three_dimensional: bool) -> String {
    if size == 1 {
        return if three_dimensional {
            "=[Main.A1:Archive.A1]".to_owned()
        } else {
            "=[.A1]".to_owned()
        };
    }
    let item = if three_dimensional {
        "[Main.A1:Archive.A1]"
    } else {
        "[.A1]"
    };
    format!(
        "=({})",
        std::iter::repeat_n(item, size)
            .collect::<Vec<_>>()
            .join("~")
    )
}

fn source_for(case: &str) -> String {
    match case {
        "scalar-number" => "=42".to_owned(),
        "unicode-text" | "unicode-limit" => "=\"α🌟 & \"\"quoted\"\"\"".to_owned(),
        "array-limit-memory" | "array-limit-work" | "array-cancelled" => array_literal(16),
        _ if case.starts_with("array-") => array_literal(scale(case).expect("array size")),
        _ if case.starts_with("duplicate-list-") => {
            reference_list(scale(case).expect("duplicate-list size"), false)
        },
        _ if case.starts_with("three-d-list-") => {
            reference_list(scale(case).expect("three-d-list size"), true)
        },
        other => panic!("unknown owned-harness case {other:?}"),
    }
}

fn is_refusal(case: &str) -> bool {
    matches!(
        case,
        "unicode-limit" | "array-limit-memory" | "array-limit-work" | "array-cancelled"
    )
}

fn expected_failure(case: &str, phase: Phase) -> &'static str {
    if phase == Phase::Setup {
        return "none";
    }
    match case {
        "unicode-limit" | "array-limit-memory" => "resource-memory",
        "array-limit-work" => "resource-work",
        "array-cancelled" => "cancelled",
        _ => "none",
    }
}

fn expected_success(case: &str, phase: Phase) -> bool {
    phase == Phase::Setup || !is_refusal(case)
}

fn mode_for(case: &str) -> Mode {
    if case.starts_with("array-")
        || case.starts_with("duplicate-list-")
        || case.starts_with("three-d-list-")
    {
        Mode::Matrix
    } else {
        Mode::Scalar
    }
}

fn metadata(case: &str) -> (usize, usize, usize, &'static str) {
    if matches!(
        case,
        "array-limit-memory" | "array-limit-work" | "array-cancelled"
    ) {
        return (4, 4, 16, "array");
    }
    if let Some(size) = scale(case) {
        if case.starts_with("array-") {
            let side = square_side(size);
            return (side, side, size, "array");
        }
        if case.starts_with("duplicate-list-") || case.starts_with("three-d-list-") {
            return (
                0,
                0,
                size,
                if size == 1 {
                    "reference"
                } else {
                    "reference-list"
                },
            );
        }
    }
    if case == "scalar-number" {
        (0, 0, 0, "number")
    } else {
        (0, 0, 0, "text")
    }
}

fn default_repeat(case: &str) -> usize {
    if is_refusal(case) {
        return 1;
    }
    match scale(case) {
        Some(4096) => 1,
        Some(256) => 8,
        Some(16) => 32,
        Some(1) => 128,
        _ => 128,
    }
}

fn copy_limits(case: &str) -> Limits {
    match case {
        "unicode-limit" => Limits::default().with_max_text_bytes(2),
        "array-limit-memory" => Limits::default().with_max_storage_bytes(0),
        "array-limit-work" => Limits::default().with_max_steps(0),
        _ => Limits::default(),
    }
}

fn make_execution(scope: &str) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(1_u64 << 40).expect("finite in-flight bytes"),
        0,
    )
    .expect("valid execution limits");
    (source, ExecutionContext::new(budget, token, limits))
}

fn reset_observer() -> u64 {
    ALLOC_CALLS.store(0, Ordering::Relaxed);
    DEALLOC_CALLS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
    let live = LIVE_BYTES.load(Ordering::Acquire);
    PEAK_LIVE_BYTES.store(live, Ordering::Relaxed);
    live
}

#[derive(Clone, Copy, Debug)]
struct Sample {
    elapsed_ns: u64,
    alloc_calls: u64,
    dealloc_calls: u64,
    requested_bytes: u64,
    released_bytes: u64,
    live_before: u64,
    live_after: u64,
    peak_live_delta: u64,
    work_used: u64,
    memory_retained_used: u64,
    owned_reserved_bytes: u64,
    successes: u64,
    refusals: u64,
    checksum: u64,
    failure: &'static str,
}

fn finish_sample(
    start: Instant,
    repeat: usize,
    live_before: u64,
    execution: &ExecutionContext,
    owned_reserved_bytes: u64,
    memory_retained_used: u64,
    successes: u64,
    refusals: u64,
    checksum: u64,
    failure: &'static str,
) -> Sample {
    let elapsed_total = start.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
    let live_after = LIVE_BYTES.load(Ordering::Acquire);
    let peak = PEAK_LIVE_BYTES.load(Ordering::Acquire);
    Sample {
        elapsed_ns: elapsed_total / repeat as u64,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after,
        peak_live_delta: peak.saturating_sub(live_before),
        work_used: execution.budget().used(Resource::Work),
        memory_retained_used: memory_retained_used.max(execution.budget().used(Resource::Memory)),
        owned_reserved_bytes,
        successes,
        refusals,
        checksum,
        failure,
    }
}

fn merge_failure(current: &mut &'static str, next: &'static str) {
    if next == "none" {
        return;
    }
    if *current == "none" || *current == next {
        *current = next;
    } else {
        *current = "mixed";
    }
}

fn failure_code(error: &EvaluationFailure) -> &'static str {
    match error {
        EvaluationFailure::Cancelled => "cancelled",
        EvaluationFailure::ResourceLimit(limit) => match limit.resource {
            Resource::Memory => "resource-memory",
            Resource::Work => "resource-work",
            Resource::Objects => "resource-objects",
            Resource::InputBytes => "resource-input",
            Resource::OutputBytes => "resource-output",
            Resource::Depth => "resource-depth",
            _ => "resource-other",
        },
        EvaluationFailure::Allocation { .. } => "allocation",
        EvaluationFailure::Unsupported(_) => "unsupported",
        EvaluationFailure::InvalidExpression(_) => "invalid-expression",
        EvaluationFailure::Execution(_) => "execution",
        EvaluationFailure::SourceChanged { .. } => "source-changed",
        EvaluationFailure::SourceVersionAvailabilityChanged => "source-version",
        _ => "other",
    }
}

fn scalar_error_code(error: ScalarError) -> u64 {
    match error {
        ScalarError::NotAvailable => 1,
        ScalarError::Name => 2,
        ScalarError::Value => 3,
        ScalarError::DivisionByZero => 4,
        ScalarError::Reference => 5,
        ScalarError::Number => 6,
        ScalarError::Null => 7,
        _ => 255,
    }
}

fn hash_bytes(mut hash: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn hash_mix(hash: u64, value: u64) -> u64 {
    hash.wrapping_mul(0x1000_0000_01b3) ^ value
}

fn hash_string(hash: u64, value: &str) -> u64 {
    hash_mix(hash_bytes(hash, value.as_bytes()), value.len() as u64)
}

fn hash_bool(hash: u64, value: bool) -> u64 {
    hash_mix(hash, u64::from(value))
}

fn checksum_reference_ast(hash: u64, reference: &Reference) -> u64 {
    match reference {
        Reference::Error => hash_mix(hash, 1),
        Reference::Local(address) => checksum_address(hash_mix(hash, 2), address),
        Reference::Source { source, address } => {
            checksum_address(hash_string(hash_mix(hash, 3), source.as_str()), address)
        },
    }
}

fn checksum_address(hash: u64, address: &Address) -> u64 {
    match address {
        Address::Cell(endpoint) => checksum_endpoint(hash_mix(hash, 1), endpoint),
        Address::Cells(first, second) => {
            checksum_endpoint(checksum_endpoint(hash_mix(hash, 2), first), second)
        },
        Address::Columns(first, second) => {
            checksum_endpoint(checksum_endpoint(hash_mix(hash, 3), first), second)
        },
        Address::Rows(first, second) => {
            checksum_endpoint(checksum_endpoint(hash_mix(hash, 4), first), second)
        },
    }
}

fn checksum_endpoint(hash: u64, endpoint: &litchi_ods::codec::formula::reference::Endpoint) -> u64 {
    let hash = checksum_selector(hash, &endpoint.sheet);
    match &endpoint.value {
        EndpointValue::Cell(cell) => checksum_cell(hash_mix(hash, 1), cell),
        EndpointValue::Column(column) => checksum_column(hash_mix(hash, 2), column),
        EndpointValue::Row(row) => hash_mix(
            hash_mix(hash, 3),
            u64::from(row.number) ^ (u64::from(row.absolute) << 32),
        ),
    }
}

fn checksum_column(hash: u64, column: &Column) -> u64 {
    hash_bool(hash_string(hash, &column.label), column.absolute)
}

fn checksum_cell(hash: u64, cell: &Cell) -> u64 {
    let hash = checksum_column(hash, &cell.column);
    hash_mix(
        hash_mix(hash, u64::from(cell.row.number)),
        u64::from(cell.row.absolute),
    )
}

fn checksum_selector(hash: u64, selector: &SheetSelector) -> u64 {
    match selector {
        SheetSelector::Current => hash_mix(hash, 1),
        SheetSelector::Inherited => hash_mix(hash, 2),
        SheetSelector::Explicit(locator) => {
            let mut hash = hash_mix(hash, 3);
            hash = hash_string(hash, locator.sheet_name().as_str());
            hash = hash_bool(hash, locator.sheet_name().absolute);
            hash = hash_bool(hash, locator.sheet_name().quoted);
            hash = hash_mix(hash, locator.subtables().len() as u64);
            for subtable in locator.subtables() {
                hash = match subtable {
                    litchi_ods::codec::formula::reference::Subtable::Cell(cell) => {
                        checksum_cell(hash_mix(hash, 1), cell)
                    },
                    litchi_ods::codec::formula::reference::Subtable::Name(name) => {
                        let hash = hash_mix(hash, 2);
                        let hash = hash_string(hash, name.as_str());
                        let hash = hash_bool(hash, name.absolute);
                        hash_bool(hash, name.quoted)
                    },
                };
            }
            hash
        },
    }
}

fn checksum_owned(value: OwnedValueView<'_>) -> u64 {
    match value {
        OwnedValueView::Empty => 1,
        OwnedValueView::Number(number) => number.to_bits(),
        OwnedValueView::Logical(logical) => u64::from(logical).wrapping_add(3),
        OwnedValueView::Text(text) => hash_bytes(4, text.as_bytes()),
        OwnedValueView::Error(error) => scalar_error_code(error),
        OwnedValueView::Array(array) => {
            let mut checksum =
                ((array.shape().rows() as u64) << 32) ^ array.shape().columns() as u64 ^ 0xa2;
            for cell in array.iter() {
                checksum = checksum.rotate_left(5) ^ checksum_owned(cell);
            }
            checksum
        },
        OwnedValueView::Reference(reference) => checksum_reference(reference),
        OwnedValueView::ReferenceList(list) => {
            list.iter().fold(list.len() as u64, |hash, reference| {
                hash.rotate_left(5) ^ checksum_reference(reference)
            })
        },
        _ => 255,
    }
}

fn checksum_reference(
    reference: litchi_ods::codec::formula::evaluation::value::ReferenceView<'_>,
) -> u64 {
    let mut hash = 0x8422_2325_cbf2_9ce4;
    hash = hash_mix(hash, u64::from(reference.reference().is_some()));
    if let Some(reference_ast) = reference.reference() {
        hash = checksum_reference_ast(hash, reference_ast);
    }
    hash = hash_mix(hash, reference.len() as u64);
    for area in reference.areas() {
        for bound in area.starts().into_iter().chain(area.ends()) {
            hash = hash_mix(hash, bound as u64);
        }
    }
    hash
}

fn validation_error(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(std::io::Error::new(
        std::io::ErrorKind::Other,
        message.into(),
    ))
}

fn expected_array_size(case: &str) -> Option<usize> {
    if matches!(
        case,
        "array-limit-memory" | "array-limit-work" | "array-cancelled"
    ) {
        Some(16)
    } else if case.starts_with("array-") {
        scale(case)
    } else {
        None
    }
}

fn expected_reference_size(case: &str) -> Option<usize> {
    if case.starts_with("duplicate-list-") || case.starts_with("three-d-list-") {
        scale(case)
    } else {
        None
    }
}

fn validate_borrowed_array(
    case: &str,
    array: litchi_ods::codec::formula::evaluation::value::ArrayView<'_>,
    size: usize,
) -> AnyResult<()> {
    let side = square_side(size);
    if array.shape().rows() != side || array.shape().columns() != side || array.len() != size {
        return Err(validation_error(format!(
            "{case} borrowed array is {}x{} with {} cells, expected {side}x{side} with {size}",
            array.shape().rows(),
            array.shape().columns(),
            array.len()
        )));
    }
    for index in 0..size {
        if !matches!(array.get(index), Some(Value::Number(value)) if value == (index + 1) as f64) {
            return Err(validation_error(format!(
                "{case} borrowed cell {index} is {:?}, expected Number({})",
                array.get(index),
                index + 1
            )));
        }
    }
    Ok(())
}

fn validate_owned_array(
    case: &str,
    array: litchi_ods::codec::formula::evaluation::value::OwnedArrayView<'_>,
    size: usize,
) -> AnyResult<()> {
    let side = square_side(size);
    if array.shape().rows() != side || array.shape().columns() != side || array.len() != size {
        return Err(validation_error(format!(
            "{case} owned array is {}x{} with {} cells, expected {side}x{side} with {size}",
            array.shape().rows(),
            array.shape().columns(),
            array.len()
        )));
    }
    for index in 0..size {
        if !matches!(array.get(index), Some(OwnedValueView::Number(value)) if value == (index + 1) as f64)
        {
            return Err(validation_error(format!(
                "{case} owned cell {index} is {:?}, expected Number({})",
                array.get(index),
                index + 1
            )));
        }
    }
    Ok(())
}

fn validate_cell_coordinates(
    case: &str,
    endpoint: &litchi_ods::codec::formula::reference::Endpoint,
) -> AnyResult<()> {
    let cell = match &endpoint.value {
        EndpointValue::Cell(cell) => cell,
        other => {
            return Err(validation_error(format!(
                "{case} reference endpoint has {other:?}, expected a cell"
            )));
        },
    };
    if cell.column.label != "A" || cell.column.absolute || cell.row.number != 1 || cell.row.absolute
    {
        return Err(validation_error(format!(
            "{case} reference endpoint has unexpected A1 metadata: {cell:?}"
        )));
    }
    Ok(())
}

fn validate_current_cell_reference(
    case: &str,
    reference: litchi_ods::codec::formula::evaluation::value::ReferenceView<'_>,
) -> AnyResult<()> {
    let ast = reference
        .reference()
        .ok_or_else(|| validation_error(format!("{case} reference lost its lexical owner")))?;
    match ast {
        Reference::Local(Address::Cell(endpoint)) => {
            if !matches!(endpoint.sheet, SheetSelector::Current) {
                return Err(validation_error(format!(
                    "{case} local reference changed its current-sheet selector: {:?}",
                    endpoint.sheet
                )));
            }
            validate_cell_coordinates(case, endpoint)?;
        },
        other => {
            return Err(validation_error(format!(
                "{case} retained unexpected local reference AST: {other:?}"
            )));
        },
    }
    if reference.areas().len() != 1
        || reference.areas()[0].starts() != [0, 0, 0]
        || reference.areas()[0].ends() != [1, 1, 1]
    {
        return Err(validation_error(format!(
            "{case} retained unexpected local reference areas: {:?}",
            reference.areas()
        )));
    }
    Ok(())
}

fn validate_named_cell_reference_endpoint(
    case: &str,
    endpoint: &litchi_ods::codec::formula::reference::Endpoint,
    expected_sheet: &str,
) -> AnyResult<()> {
    let locator = match &endpoint.sheet {
        SheetSelector::Explicit(locator) => locator,
        other => {
            return Err(validation_error(format!(
                "{case} 3-D endpoint uses {other:?}, expected explicit {expected_sheet}"
            )));
        },
    };
    if locator.sheet_name().as_str() != expected_sheet
        || locator.sheet_name().absolute
        || locator.sheet_name().quoted
        || !locator.subtables().is_empty()
    {
        return Err(validation_error(format!(
            "{case} 3-D endpoint has unexpected {expected_sheet} locator: {:?}",
            locator
        )));
    }
    validate_cell_coordinates(case, endpoint)
}

fn validate_three_dimensional_reference(
    case: &str,
    reference: litchi_ods::codec::formula::evaluation::value::ReferenceView<'_>,
) -> AnyResult<()> {
    let ast = reference
        .reference()
        .ok_or_else(|| validation_error(format!("{case} reference lost its lexical owner")))?;
    match ast {
        Reference::Local(Address::Cells(first, second)) => {
            validate_named_cell_reference_endpoint(case, first, "Main")?;
            validate_named_cell_reference_endpoint(case, second, "Archive")?;
        },
        other => {
            return Err(validation_error(format!(
                "{case} retained unexpected 3-D reference AST: {other:?}"
            )));
        },
    }
    if reference.areas().len() != 1
        || reference.areas()[0].starts() != [0, 0, 0]
        || reference.areas()[0].ends() != [3, 1, 1]
    {
        return Err(validation_error(format!(
            "{case} retained unexpected 3-D reference areas: {:?}",
            reference.areas()
        )));
    }
    Ok(())
}

fn validate_borrowed_references(case: &str, value: Value<'_>, size: usize) -> AnyResult<()> {
    if size == 1 {
        return match value {
            Value::Reference(reference) if case.starts_with("duplicate-list-") => {
                validate_current_cell_reference(case, reference)
            },
            Value::Reference(reference) => validate_three_dimensional_reference(case, reference),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected one first-class reference"
            ))),
        };
    }
    let list = match value {
        Value::ReferenceList(list) => list,
        other => {
            return Err(validation_error(format!(
                "{case} returned {other:?}, expected an ordered reference list"
            )));
        },
    };
    if list.len() != size {
        return Err(validation_error(format!(
            "{case} returned {} references, expected {size}",
            list.len()
        )));
    }
    for reference in list.iter() {
        if case.starts_with("duplicate-list-") {
            validate_current_cell_reference(case, reference)?;
        } else {
            validate_three_dimensional_reference(case, reference)?;
        }
    }
    Ok(())
}

fn validate_owned_references(case: &str, value: OwnedValueView<'_>, size: usize) -> AnyResult<()> {
    if size == 1 {
        return match value {
            OwnedValueView::Reference(reference) if case.starts_with("duplicate-list-") => {
                validate_current_cell_reference(case, reference)
            },
            OwnedValueView::Reference(reference) => {
                validate_three_dimensional_reference(case, reference)
            },
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected one owned reference"
            ))),
        };
    }
    let list = match value {
        OwnedValueView::ReferenceList(list) => list,
        other => {
            return Err(validation_error(format!(
                "{case} returned {other:?}, expected an owned reference list"
            )));
        },
    };
    if list.len() != size {
        return Err(validation_error(format!(
            "{case} returned {} owned references, expected {size}",
            list.len()
        )));
    }
    for reference in list.iter() {
        if case.starts_with("duplicate-list-") {
            validate_current_cell_reference(case, reference)?;
        } else {
            validate_three_dimensional_reference(case, reference)?;
        }
    }
    Ok(())
}

fn validate_borrowed_value(case: &str, value: Value<'_>) -> AnyResult<()> {
    if case == "scalar-number" {
        return match value {
            Value::Number(number) if number == 42.0 => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected Number(42)"
            ))),
        };
    }
    if matches!(case, "unicode-text" | "unicode-limit") {
        return match value {
            Value::Text(text) if text == "α🌟 & \"quoted\"" => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected the exact Unicode text"
            ))),
        };
    }
    if let Some(size) = expected_array_size(case) {
        return match value {
            Value::Array(array) => validate_borrowed_array(case, array, size),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected a square numeric array"
            ))),
        };
    }
    if let Some(size) = expected_reference_size(case) {
        return validate_borrowed_references(case, value, size);
    }
    Err(validation_error(format!(
        "{case} has no expected borrowed value"
    )))
}

fn validate_owned_value(case: &str, value: OwnedValueView<'_>) -> AnyResult<()> {
    if case == "scalar-number" {
        return match value {
            OwnedValueView::Number(number) if number == 42.0 => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected owned Number(42)"
            ))),
        };
    }
    if matches!(case, "unicode-text" | "unicode-limit") {
        return match value {
            OwnedValueView::Text(text) if text == "α🌟 & \"quoted\"" => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected the exact owned Unicode text"
            ))),
        };
    }
    if let Some(size) = expected_array_size(case) {
        return match value {
            OwnedValueView::Array(array) => validate_owned_array(case, array, size),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected an owned square numeric array"
            ))),
        };
    }
    if let Some(size) = expected_reference_size(case) {
        return validate_owned_references(case, value, size);
    }
    Err(validation_error(format!(
        "{case} has no expected owned value"
    )))
}

fn validate_prepared_result(case: &str, evaluated: &Evaluated<'_>) -> AnyResult<()> {
    validate_borrowed_value(case, evaluated.value())?;

    // This copy is deliberately outside every measured phase. It validates
    // ownership semantics with an independent default context and is dropped
    // before the first timed sample, so it cannot affect timing or allocator
    // counters for the reported samples.
    let (_cancellation, ownership_execution) = make_execution("owned-validation");
    let owned = evaluated
        .to_owned(&ownership_execution, &Limits::default())
        .map_err(|error| validation_error(format!("{case} ownership preflight failed: {error}")))?;
    validate_owned_value(case, owned.value())?;

    let borrowed_checksum = checksum_borrowed(evaluated.value());
    let owned_checksum = checksum_owned(owned.value());
    if borrowed_checksum != owned_checksum {
        return Err(validation_error(format!(
            "{case} borrowed/owned checksums differ before timing: {borrowed_checksum} != {owned_checksum}"
        )));
    }
    Ok(())
}

fn prepare_limits() -> Limits {
    Limits::default()
}

fn measure_setup(source: &str, resolver: &FixtureResolver, case: &str, repeat: usize) -> Sample {
    let (_cancellation, execution) = make_execution("owned-setup");
    let live_before = reset_observer();
    let start = Instant::now();
    let mut successes: u64 = 0;
    let mut refusals: u64 = 0;
    let mut checksum: u64 = 0;
    let mut retained: u64 = 0;
    let mut failure = "none";
    for _ in 0..repeat {
        let expression = match Expression::parse(black_box(source)) {
            Ok(expression) => expression,
            Err(_) => {
                refusals += 1;
                merge_failure(&mut failure, "parse");
                continue;
            },
        };
        let context =
            Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode_for(case));
        match evaluate(&expression, resolver, &context, &prepare_limits()) {
            Ok(evaluated) => {
                checksum = checksum.wrapping_add(checksum_borrowed(evaluated.value()));
                retained = retained.max(execution.budget().used(Resource::Memory));
                black_box(evaluated.value());
                successes += 1;
            },
            Err(error) => {
                refusals += 1;
                merge_failure(&mut failure, failure_code(&error));
            },
        }
    }
    finish_sample(
        start,
        repeat,
        live_before,
        &execution,
        0,
        retained,
        successes,
        refusals,
        checksum,
        failure,
    )
}

fn checksum_borrowed(value: litchi_ods::codec::formula::evaluation::value::Value<'_>) -> u64 {
    match value {
        litchi_ods::codec::formula::evaluation::value::Value::Empty => 1,
        litchi_ods::codec::formula::evaluation::value::Value::Number(number) => number.to_bits(),
        litchi_ods::codec::formula::evaluation::value::Value::Logical(logical) => {
            u64::from(logical).wrapping_add(3)
        },
        litchi_ods::codec::formula::evaluation::value::Value::Text(text) => {
            hash_bytes(4, text.as_bytes())
        },
        litchi_ods::codec::formula::evaluation::value::Value::Error(error) => {
            scalar_error_code(error)
        },
        litchi_ods::codec::formula::evaluation::value::Value::Array(array) => {
            let mut checksum =
                ((array.shape().rows() as u64) << 32) ^ array.shape().columns() as u64 ^ 0xa2;
            for index in 0..array.len() {
                if let Some(cell) = array.get(index) {
                    checksum = checksum.rotate_left(5) ^ checksum_borrowed(cell);
                }
            }
            checksum
        },
        litchi_ods::codec::formula::evaluation::value::Value::Reference(reference) => {
            checksum_reference(reference)
        },
        litchi_ods::codec::formula::evaluation::value::Value::ReferenceList(list) => {
            list.iter().fold(list.len() as u64, |hash, reference| {
                hash.rotate_left(5) ^ checksum_reference(reference)
            })
        },
        _ => 255,
    }
}

fn measure_own(evaluated: &Evaluated<'_>, case: &str, repeat: usize) -> Sample {
    let (cancellation, execution) = make_execution("owned-copy");
    if case == "array-cancelled" {
        cancellation.cancel();
    }
    let limits = copy_limits(case);
    let live_before = reset_observer();
    let start = Instant::now();
    let mut successes: u64 = 0;
    let mut refusals: u64 = 0;
    let mut checksum: u64 = 0;
    let mut owned_reserved: u64 = 0;
    let mut retained: u64 = 0;
    let mut failure = "none";
    for _ in 0..repeat {
        match evaluated.to_owned(&execution, &limits) {
            Ok(owned) => {
                let reservation = owned.reserved_storage_bytes() as u64;
                let value_checksum = checksum_owned(owned.value());
                let memory = execution.budget().used(Resource::Memory);
                owned_reserved = owned_reserved.max(reservation);
                retained = retained.max(memory);
                checksum = checksum.wrapping_add(value_checksum);
                black_box(value_checksum);
                drop(owned);
                successes += 1;
            },
            Err(error) => {
                refusals += 1;
                merge_failure(&mut failure, failure_code(&error));
                retained = retained.max(execution.budget().used(Resource::Memory));
            },
        }
    }
    finish_sample(
        start,
        repeat,
        live_before,
        &execution,
        owned_reserved,
        retained,
        successes,
        refusals,
        checksum,
        failure,
    )
}

fn measure_end_to_end(
    source: &str,
    resolver: &FixtureResolver,
    case: &str,
    repeat: usize,
) -> Sample {
    let (cancellation, execution) = make_execution("owned-end-to-end");
    let limits = copy_limits(case);
    let live_before = reset_observer();
    let start = Instant::now();
    let mut successes: u64 = 0;
    let mut refusals: u64 = 0;
    let mut checksum: u64 = 0;
    let mut owned_reserved: u64 = 0;
    let mut retained: u64 = 0;
    let mut failure = "none";
    for _ in 0..repeat {
        let expression = match Expression::parse(black_box(source)) {
            Ok(expression) => expression,
            Err(_) => {
                refusals += 1;
                merge_failure(&mut failure, "parse");
                continue;
            },
        };
        let context =
            Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode_for(case));
        let evaluated = match evaluate(&expression, resolver, &context, &prepare_limits()) {
            Ok(evaluated) => evaluated,
            Err(error) => {
                refusals += 1;
                merge_failure(&mut failure, failure_code(&error));
                continue;
            },
        };
        if case == "array-cancelled" {
            cancellation.cancel();
        }
        match evaluated.to_owned(&execution, &limits) {
            Ok(owned) => {
                let reservation = owned.reserved_storage_bytes() as u64;
                let value_checksum = checksum_owned(owned.value());
                let memory = execution.budget().used(Resource::Memory);
                owned_reserved = owned_reserved.max(reservation);
                retained = retained.max(memory);
                checksum = checksum.wrapping_add(value_checksum);
                black_box(value_checksum);
                drop(owned);
                successes += 1;
            },
            Err(error) => {
                refusals += 1;
                merge_failure(&mut failure, failure_code(&error));
                retained = retained.max(execution.budget().used(Resource::Memory));
            },
        }
    }
    finish_sample(
        start,
        repeat,
        live_before,
        &execution,
        owned_reserved,
        retained,
        successes,
        refusals,
        checksum,
        failure,
    )
}

fn percentile(samples: &[Sample], metric: impl Fn(&Sample) -> u64, percent: usize) -> u64 {
    let mut values: Vec<u64> = samples.iter().map(metric).collect();
    values.sort_unstable();
    values[(values.len().saturating_sub(1) * percent) / 100]
}

fn mean_ns(samples: &[Sample]) -> u64 {
    let total: u128 = samples
        .iter()
        .map(|sample| u128::from(sample.elapsed_ns))
        .sum();
    (total / samples.len() as u128).min(u128::from(u64::MAX)) as u64
}

fn maximum(samples: &[Sample], metric: impl Fn(&Sample) -> u64) -> u64 {
    samples.iter().map(metric).max().unwrap_or(0)
}

fn failure_summary(samples: &[Sample]) -> &'static str {
    let first = samples.first().map_or("none", |sample| sample.failure);
    if samples.iter().all(|sample| sample.failure == first) {
        first
    } else {
        "mixed"
    }
}

fn main() -> AnyResult<()> {
    let config = Config::from_args();
    assert_eq!(config.workload, "owned-evaluation", "unexpected workload");
    assert_eq!(
        config.revision, "candidate",
        "owned harness is candidate-only"
    );
    let source = source_for(&config.case);
    let (rows, columns, elements, value_kind) = metadata(&config.case);
    let expected_success = expected_success(&config.case, config.phase);
    println!(
        "config revision={} workload={} group={} phase={} case={} input_bytes={} repeat={} warmups={} iterations={} expected_success={} expected_failure={} rows={} columns={} elements={} mode={} value_kind={}",
        config.revision,
        config.workload,
        config.group,
        config.phase.label(),
        config.case,
        source.len(),
        config.repeat,
        config.warmups,
        config.iterations,
        expected_success,
        expected_failure(&config.case, config.phase),
        rows,
        columns,
        elements,
        match mode_for(&config.case) {
            Mode::Matrix => "matrix",
            Mode::Scalar => "scalar",
            _ => "other",
        },
        value_kind,
    );

    let resolver = FixtureResolver::new();
    let (_preparation_cancellation, preparation) = make_execution("owned-preparation");
    let expression = Expression::parse(&source)
        .map_err(|error| format!("{} preparation parse failed: {error}", config.case))?;
    let context =
        Context::new(&preparation, Position::new("Main", 0, 0)).with_mode(mode_for(&config.case));
    let evaluated = evaluate(&expression, &resolver, &context, &prepare_limits())
        .map_err(|error| format!("{} preparation evaluation failed: {error}", config.case))?;
    black_box(evaluated.value());
    validate_prepared_result(&config.case, &evaluated)?;

    let measure = |phase: Phase| match phase {
        Phase::Setup => measure_setup(&source, &resolver, &config.case, config.repeat),
        Phase::Own => measure_own(&evaluated, &config.case, config.repeat),
        Phase::ParseEvaluateOwn => {
            measure_end_to_end(&source, &resolver, &config.case, config.repeat)
        },
    };
    for _ in 0..config.warmups {
        black_box(measure(config.phase));
    }
    let samples: Vec<Sample> = (0..config.iterations)
        .map(|_| measure(config.phase))
        .collect();
    let first = samples.first().expect("at least one sample");
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} alloc_calls_p50={} alloc_calls_max={} dealloc_calls_p50={} dealloc_calls_max={} requested_bytes_p50={} requested_bytes_max={} released_bytes_p50={} released_bytes_max={} live_before_p50={} live_after_p50={} live_after_max={} peak_live_delta_p50={} peak_live_delta_max={} work_used_p50={} work_used_max={} memory_retained_used_p50={} memory_retained_used_max={} owned_reserved_bytes_p50={} owned_reserved_bytes_max={} successes_p50={} successes_max={} refusals_p50={} refusals_max={} checksum_p50={} checksum_max={} failure={}",
        mean_ns(&samples),
        percentile(&samples, |sample| sample.elapsed_ns, 50),
        percentile(&samples, |sample| sample.elapsed_ns, 95),
        percentile(&samples, |sample| sample.elapsed_ns, 99),
        percentile(&samples, |sample| sample.alloc_calls, 50),
        maximum(&samples, |sample| sample.alloc_calls),
        percentile(&samples, |sample| sample.dealloc_calls, 50),
        maximum(&samples, |sample| sample.dealloc_calls),
        percentile(&samples, |sample| sample.requested_bytes, 50),
        maximum(&samples, |sample| sample.requested_bytes),
        percentile(&samples, |sample| sample.released_bytes, 50),
        maximum(&samples, |sample| sample.released_bytes),
        percentile(&samples, |sample| sample.live_before, 50),
        percentile(&samples, |sample| sample.live_after, 50),
        maximum(&samples, |sample| sample.live_after),
        percentile(&samples, |sample| sample.peak_live_delta, 50),
        maximum(&samples, |sample| sample.peak_live_delta),
        percentile(&samples, |sample| sample.work_used, 50),
        maximum(&samples, |sample| sample.work_used),
        percentile(&samples, |sample| sample.memory_retained_used, 50),
        maximum(&samples, |sample| sample.memory_retained_used),
        percentile(&samples, |sample| sample.owned_reserved_bytes, 50),
        maximum(&samples, |sample| sample.owned_reserved_bytes),
        percentile(&samples, |sample| sample.successes, 50),
        maximum(&samples, |sample| sample.successes),
        percentile(&samples, |sample| sample.refusals, 50),
        maximum(&samples, |sample| sample.refusals),
        percentile(&samples, |sample| sample.checksum, 50),
        maximum(&samples, |sample| sample.checksum),
        failure_summary(&samples),
    );
    // Keep a use of the first sample visible to the optimizer while retaining
    // the complete aggregate values above.
    black_box(first.elapsed_ns);
    Ok(())
}
