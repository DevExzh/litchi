//! Process-isolated allocator and runtime evidence for detached ink-action CRUD.
//!
//! Fixture construction and semantic source/patch checks stay outside each
//! timed operation. This binary is an opt-in evidence harness; it is not a
//! production dependency or benchmark API.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in profile owns a process-local allocator observer and emits JSON"
)]

use litchi_drawingml::Error as DrawingError;
use litchi_drawingml::ink::{
    ACTION_NAMESPACE, INKML_NAMESPACE,
    actions::{
        self, ActionDataDraft, ActionDraft, ActionParent, ActionSelector, ActionType,
        ChildSelector, Commit, Draft, Edit, LengthUnit, Limits, OpaquePayload, Prepared, Profile,
        PropertyDraft, TimeUnit,
    },
};
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::hint::black_box;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

type BoxError = Box<dyn Error + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;
const SMALL_ACTIONS: usize = 8;
const SCALED_ACTIONS: usize = 128;
const NEAR_ACTIONS: usize = 1_024;
const OPAQUE_DRAFT_ACTIONS: usize = 64;
const CAP_VALUE: &str =
    "cap-expanded-value-012345678901234567890123456789012345678901234567890123456789";

const LANES: &[&str] = &[
    "draft_small_8",
    "draft_scaled_128",
    "draft_near_1024",
    "draft_opaque_64",
    "scalar_edit_small_8",
    "scalar_edit_scaled_128",
    "scalar_edit_near_1024",
    "scalar_batch_scaled_128",
    "scalar_batch_near_1024",
    "scalar_coalesce_scaled_128",
    "scalar_coalesce_near_1024",
    "no_op_small_8",
    "no_op_scaled_128",
    "no_op_near_1024",
    "add_small_8",
    "add_scaled_128",
    "add_near_1024",
    "insert_batch_scaled_128",
    "insert_batch_near_1024",
    "remove_small_8",
    "remove_scaled_128",
    "remove_near_1024",
    "remove_batch_scaled_128",
    "remove_batch_near_1024",
    "clear_batch_scaled_128",
    "clear_batch_near_1024",
    "move_small_8",
    "move_scaled_128",
    "move_near_1024",
    "move_batch_scaled_128",
    "move_batch_near_1024",
    "cap_refusal_small_8",
    "cap_refusal_scaled_128",
    "cap_refusal_near_1024",
];

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: every method forwards the allocation contract to System; atomics
// observe successful operations without changing pointer semantics.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair belongs to the caller.
        unsafe { System.dealloc(pointer, layout) };
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            REALLOC_OLD.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            REALLOC_NEW.fetch_add(as_u64(new_size), Ordering::Relaxed);
            if new_size >= layout.size() {
                observe_growth(new_size - layout.size());
            } else {
                subtract_live(layout.size() - new_size);
            }
        }
        result
    }
}

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn observe_alloc(size: usize) {
    ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
    DIRECT_BYTES.fetch_add(as_u64(size), Ordering::Relaxed);
    observe_growth(size);
}

fn observe_growth(size: usize) {
    let size = as_u64(size);
    let live = LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    let mut old = PEAK_BYTES.load(Ordering::Relaxed);
    while live > old {
        match PEAK_BYTES.compare_exchange_weak(old, live, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => break,
            Err(observed) => old = observed,
        }
    }
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        INVALID.store(true, Ordering::Release);
    }
}

#[derive(Clone, Copy)]
struct AllocSnapshot {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live: u64,
    peak: u64,
    failed: u64,
    invalid: bool,
}

impl AllocSnapshot {
    fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: REALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
            direct: DIRECT_BYTES.load(Ordering::Acquire),
            realloc_old: REALLOC_OLD.load(Ordering::Acquire),
            realloc_new: REALLOC_NEW.load(Ordering::Acquire),
            deallocated: DEALLOC_BYTES.load(Ordering::Acquire),
            live: LIVE_BYTES.load(Ordering::Acquire),
            peak: PEAK_BYTES.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: INVALID.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> AllocDelta {
        AllocDelta {
            calls: after.calls.saturating_sub(self.calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            direct: after.direct.saturating_sub(self.direct),
            realloc_old: after.realloc_old.saturating_sub(self.realloc_old),
            realloc_new: after.realloc_new.saturating_sub(self.realloc_new),
            deallocated: after.deallocated.saturating_sub(self.deallocated),
            live_before: self.live,
            live_after: after.live,
            peak_delta: after.peak.saturating_sub(self.peak),
            failed: after.failed.saturating_sub(self.failed),
            invalid: self.invalid || after.invalid,
        }
    }
}

#[derive(Clone, Copy)]
struct AllocDelta {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live_before: u64,
    live_after: u64,
    peak_delta: u64,
    failed: u64,
    invalid: bool,
}

impl AllocDelta {
    fn requested(self) -> u64 {
        self.direct.saturating_add(self.realloc_new)
    }

    fn balanced(self) -> bool {
        self.live_before
            .checked_add(self.direct)
            .and_then(|value| value.checked_add(self.realloc_new))
            .and_then(|value| value.checked_sub(self.realloc_old))
            .and_then(|value| value.checked_sub(self.deallocated))
            == Some(self.live_after)
    }
}

fn reset_counters() {
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
    ALLOC_CALLS.store(0, Ordering::Release);
    REALLOC_CALLS.store(0, Ordering::Release);
    DEALLOC_CALLS.store(0, Ordering::Release);
    DIRECT_BYTES.store(0, Ordering::Release);
    REALLOC_OLD.store(0, Ordering::Release);
    REALLOC_NEW.store(0, Ordering::Release);
    DEALLOC_BYTES.store(0, Ordering::Release);
    ALLOC_FAILED.store(0, Ordering::Release);
    INVALID.store(false, Ordering::Release);
}

fn allocator_counter_self_test() -> Result<()> {
    reset_counters();
    let before = AllocSnapshot::now();
    let layout = Layout::from_size_align(8, std::mem::align_of::<usize>())?;
    // SAFETY: this deliberately exercises the observer with a valid layout.
    let pointer = unsafe { std::alloc::alloc(layout) };
    if pointer.is_null() {
        reset_counters();
        return Err("allocator self-test allocation failed".into());
    }
    // SAFETY: pointer was returned with layout, and the new size is valid.
    let resized = unsafe { std::alloc::realloc(pointer, layout, 32) };
    if resized.is_null() {
        // SAFETY: a failed reallocation leaves the original allocation valid.
        unsafe { std::alloc::dealloc(pointer, layout) };
        reset_counters();
        return Err("allocator self-test reallocation failed".into());
    }
    let resized_layout = Layout::from_size_align(32, layout.align())?;
    // SAFETY: resized uses the same alignment and requested new size.
    unsafe { std::alloc::dealloc(resized, resized_layout) };
    let delta = before.delta(AllocSnapshot::now());
    let valid = delta.calls == 1
        && delta.realloc_calls == 1
        && delta.dealloc_calls == 1
        && delta.direct == 8
        && delta.realloc_old == 8
        && delta.realloc_new == 32
        && delta.deallocated == 32
        && delta.requested() == 40
        && delta.balanced()
        && !delta.invalid
        && delta.failed == 0;
    reset_counters();
    if valid {
        Ok(())
    } else {
        Err("allocator self-test counters did not balance".into())
    }
}

struct Fixture {
    action_count: usize,
    profile: Profile,
}

struct Inputs {
    small: Fixture,
    scaled: Fixture,
    near: Fixture,
    definitions_payload: Vec<u8>,
    trace_payload: Vec<u8>,
}

impl Inputs {
    fn load() -> Result<Self> {
        let definitions_payload = definitions_payload();
        let trace_payload = trace_payload();
        let small = fixture(SMALL_ACTIONS)?;
        let scaled = fixture(SCALED_ACTIONS)?;
        let near = fixture(NEAR_ACTIONS)?;
        for fixture in [&small, &scaled, &near] {
            if fixture.profile.actions().count() != fixture.action_count {
                return Err("fixture action count readback mismatch".into());
            }
            if !contains_opaque(fixture.profile.source()) {
                return Err("fixture lost opaque namespace payload".into());
            }
        }
        Ok(Self {
            small,
            scaled,
            near,
            definitions_payload,
            trace_payload,
        })
    }
}

fn definitions_payload() -> Vec<u8> {
    format!(
        r#"<inkml:definitions xmlns:inkml="{INKML_NAMESPACE}" xmlns:v="urn:vendor:ink-action"><v:future v:flag="opaque"/></inkml:definitions>"#
    )
    .into_bytes()
}

fn trace_payload() -> Vec<u8> {
    format!(
        r#"<inkml:trace xmlns:inkml="{INKML_NAMESPACE}" xmlns:v="urn:vendor:ink-action"><v:future v:payload="opaque"/></inkml:trace>"#
    )
    .into_bytes()
}

fn fixture(action_count: usize) -> Result<Fixture> {
    let definitions = String::from_utf8(definitions_payload())?;
    let trace = String::from_utf8(trace_payload())?;
    let mut xml = String::new();
    write!(
        &mut xml,
        r#"<?xml version="1.0" encoding="UTF-8"?><!--leading-->
<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:inkml="{INKML_NAMESPACE}" lengthUnit="cm" timeUnit="ms" xml:id="root">
  <!--root comment-->
  {definitions}
"#
    )?;
    for index in 0..action_count {
        write!(
            &mut xml,
            "  <iact:action xml:id=\"a{index}\" type=\"add\" startTime=\"{index}\">\n    <iact:property name=\"kind\" value=\"v\"/>\n"
        )?;
        if index == 0 {
            write!(
                &mut xml,
                "    <iact:actionData xml:id=\"d0\" name=\"stroke\">{trace}</iact:actionData>\n"
            )?;
        }
        xml.push_str("  </iact:action>\n");
    }
    xml.push_str("</iact:actions><!--trailing-->");
    let source = xml.into_bytes();
    let profile = actions::read_profile(&source)?;
    Ok(Fixture {
        action_count,
        profile,
    })
}

fn contains_opaque(source: &[u8]) -> bool {
    source
        .windows(b"<v:future v:flag=\"opaque\"/>".len())
        .any(|window| window == b"<v:future v:flag=\"opaque\"/>")
        && source
            .windows(b"<v:future v:payload=\"opaque\"/>".len())
            .any(|window| window == b"<v:future v:payload=\"opaque\"/>")
}

#[derive(Clone, Copy)]
enum LaneKind {
    Draft { opaque: bool },
    Scalar,
    ScalarBatch,
    ScalarCoalesce,
    Noop,
    Add,
    InsertBatch,
    Remove,
    RemoveBatch,
    ClearBatch,
    Move,
    MoveBatch,
    CapRefusal,
}

fn lane_kind(lane: &str) -> Option<LaneKind> {
    if lane.starts_with("draft_opaque_") {
        Some(LaneKind::Draft { opaque: true })
    } else if lane.starts_with("draft_") {
        Some(LaneKind::Draft { opaque: false })
    } else if lane.starts_with("scalar_edit_") {
        Some(LaneKind::Scalar)
    } else if lane.starts_with("scalar_batch_") {
        Some(LaneKind::ScalarBatch)
    } else if lane.starts_with("scalar_coalesce_") {
        Some(LaneKind::ScalarCoalesce)
    } else if lane.starts_with("no_op_") {
        Some(LaneKind::Noop)
    } else if lane.starts_with("add_") {
        Some(LaneKind::Add)
    } else if lane.starts_with("insert_batch_") {
        Some(LaneKind::InsertBatch)
    } else if lane.starts_with("remove_batch_") {
        Some(LaneKind::RemoveBatch)
    } else if lane.starts_with("remove_") {
        Some(LaneKind::Remove)
    } else if lane.starts_with("clear_batch_") {
        Some(LaneKind::ClearBatch)
    } else if lane.starts_with("move_batch_") {
        Some(LaneKind::MoveBatch)
    } else if lane.starts_with("move_") {
        Some(LaneKind::Move)
    } else if lane.starts_with("cap_refusal_") {
        Some(LaneKind::CapRefusal)
    } else {
        None
    }
}

fn is_refusal(lane: &str) -> bool {
    matches!(lane_kind(lane), Some(LaneKind::CapRefusal))
}

fn action_count_for_lane(lane: &str) -> usize {
    if lane.ends_with("_small_8") {
        SMALL_ACTIONS
    } else if lane.ends_with("_scaled_128") {
        SCALED_ACTIONS
    } else if lane.ends_with("_near_1024") {
        NEAR_ACTIONS
    } else if lane.ends_with("_opaque_64") {
        OPAQUE_DRAFT_ACTIONS
    } else {
        0
    }
}

fn result_action_count_for_lane(lane: &str) -> usize {
    let count = action_count_for_lane(lane);
    match lane_kind(lane) {
        Some(LaneKind::Add) => count + 1,
        Some(LaneKind::InsertBatch) => count * 2,
        Some(LaneKind::Remove) => count - 1,
        Some(LaneKind::RemoveBatch) => count / 2,
        _ => count,
    }
}

fn operation_count_for_lane(lane: &str) -> usize {
    let count = action_count_for_lane(lane);
    match lane_kind(lane) {
        Some(LaneKind::ScalarBatch)
        | Some(LaneKind::ScalarCoalesce)
        | Some(LaneKind::InsertBatch) => count,
        Some(LaneKind::RemoveBatch) | Some(LaneKind::ClearBatch) | Some(LaneKind::MoveBatch) => {
            count / 2
        },
        Some(LaneKind::Draft { .. }) => count,
        Some(LaneKind::CapRefusal) => 1,
        Some(
            LaneKind::Scalar | LaneKind::Noop | LaneKind::Add | LaneKind::Remove | LaneKind::Move,
        ) => 1,
        None => 0,
    }
}

fn fixture_for<'a>(inputs: &'a Inputs, lane: &str) -> Result<&'a Fixture> {
    let fixture = if lane.ends_with("_small_8") {
        &inputs.small
    } else if lane.ends_with("_scaled_128") {
        &inputs.scaled
    } else if lane.ends_with("_near_1024") {
        &inputs.near
    } else {
        return Err(format!("lane {lane} has no source fixture").into());
    };
    Ok(fixture)
}

fn action_draft(index: usize, trace: Option<&[u8]>) -> Result<ActionDraft> {
    let mut action = ActionDraft::new(ActionType::Add, index.to_string())?
        .with_xml_id(format!("a{index}"))?
        .property(PropertyDraft::new("kind", "v")?)?;
    if let Some(trace) = trace {
        let data = ActionDataDraft::new()
            .with_xml_id("d0")?
            .trace(OpaquePayload::new(trace)?)?;
        action = action.data(data)?;
    }
    Ok(action)
}

fn draft_operation(inputs: &Inputs, count: usize, opaque: bool) -> Result<Prepared> {
    let mut draft = Draft::with_units(LengthUnit::Centimeter, TimeUnit::Millisecond)?;
    if opaque {
        draft = draft.definitions(OpaquePayload::new(&inputs.definitions_payload)?)?;
    }
    for index in 0..count {
        let trace = if opaque && index == 0 {
            Some(inputs.trace_payload.as_slice())
        } else {
            None
        };
        draft = draft.action(action_draft(index, trace)?)?;
    }
    Ok(draft.finish()?)
}

fn added_action(index: usize) -> Result<ActionDraft> {
    Ok(ActionDraft::new(ActionType::Add, index.to_string())?
        .with_xml_id(format!("added-{index}"))?
        .property(PropertyDraft::new("kind", "added")?)?)
}

enum Work {
    Prepared(Prepared),
    Commit(Commit),
    Rejected {
        resource: &'static str,
        limit: usize,
        pre_source: Vec<u8>,
        pre_source_hash: u64,
        pre_state: ProfileState,
        retained_profile: Profile,
    },
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct ProfileState {
    action_ids: Vec<String>,
    action_types: Vec<String>,
    start_times: Vec<String>,
    properties: Vec<Vec<(String, String)>>,
    child_counts: Vec<usize>,
}

fn profile_state(profile: &Profile) -> ProfileState {
    let mut state = ProfileState {
        action_ids: Vec::new(),
        action_types: Vec::new(),
        start_times: Vec::new(),
        properties: Vec::new(),
        child_counts: Vec::new(),
    };
    for action in profile.actions() {
        state
            .action_ids
            .push(action.xml_id().unwrap_or_default().to_owned());
        state
            .action_types
            .push(action.action_type().as_str().to_owned());
        state.start_times.push(action.start_time().to_owned());
        state.properties.push(
            action
                .properties()
                .iter()
                .map(|property| (property.name().to_owned(), property.value().to_owned()))
                .collect(),
        );
        state.child_counts.push(action.children().len());
    }
    state
}

fn source_hash(source: &[u8]) -> u64 {
    let mut hash = 14_695_981_039_346_656_037_u64;
    for byte in source {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(1_099_511_628_211_u64);
    }
    hash
}

struct RejectionContext {
    pre_source: Vec<u8>,
    pre_source_hash: u64,
    pre_state: ProfileState,
    edit_profile: Profile,
    retained_profile: Profile,
}

fn rejection_context(inputs: &Inputs, lane: &str) -> Result<RejectionContext> {
    let fixture = source_fixture(inputs, lane)?;
    let profile = fixture.profile.clone();
    let pre_source = profile.source().to_vec();
    Ok(RejectionContext {
        pre_source_hash: source_hash(&pre_source),
        pre_state: profile_state(&profile),
        edit_profile: profile.clone(),
        retained_profile: profile,
        pre_source,
    })
}

fn source_fixture<'a>(inputs: &'a Inputs, lane: &str) -> Result<&'a Fixture> {
    fixture_for(inputs, lane)
}

fn operation(inputs: &Inputs, lane: &str, rejection: Option<RejectionContext>) -> Result<Work> {
    let kind = lane_kind(lane).ok_or_else(|| format!("unknown lane {lane}"))?;
    let count = action_count_for_lane(lane);
    match kind {
        LaneKind::Draft { opaque } => Ok(Work::Prepared(draft_operation(inputs, count, opaque)?)),
        LaneKind::Scalar => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            edit.set_action_type(
                ActionSelector::ordinal(count - 1),
                ActionType::Custom("edited".into()),
            )?;
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::ScalarBatch => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            for index in 0..count {
                edit.set_property_value(
                    ChildSelector::Property {
                        action: ActionSelector::direct(index),
                        index: 0,
                    },
                    format!("batch-{index}"),
                )?;
            }
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::ScalarCoalesce => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            let target = ChildSelector::Property {
                action: ActionSelector::direct(count - 1),
                index: 0,
            };
            for index in 0..count {
                edit.set_property_value(target, format!("coalesced-{index}"))?;
            }
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::Noop => {
            let fixture = source_fixture(inputs, lane)?;
            let edit = Edit::from_profile(&fixture.profile)?;
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::Add => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            edit.add_action(ActionParent::Root, added_action(count)?)?;
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::InsertBatch => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            for index in 0..count {
                edit.add_action(ActionParent::Root, added_action(index)?)?;
            }
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::Remove => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            edit.remove_action(ActionSelector::ordinal(count - 1))?;
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::RemoveBatch => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            for index in (1..count).step_by(2) {
                edit.remove_action(ActionSelector::direct(index))?;
            }
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::ClearBatch => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            for index in (1..count).step_by(2) {
                edit.clear(ActionSelector::direct(index))?;
            }
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::Move => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            edit.move_before(
                ActionSelector::ordinal(count - 1),
                ActionSelector::ordinal(0),
            )?;
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::MoveBatch => {
            let fixture = source_fixture(inputs, lane)?;
            let mut edit = Edit::from_profile(&fixture.profile)?;
            for index in (1..count).step_by(2) {
                edit.move_before(
                    ActionSelector::direct(index),
                    ActionSelector::direct(index - 1),
                )?;
            }
            Ok(Work::Commit(edit.finish()?))
        },
        LaneKind::CapRefusal => {
            let context = rejection.ok_or("caller-cap rejection context is missing")?;
            let mut limits = Limits::default();
            limits.max_output_bytes = context.pre_source.len();
            let RejectionContext {
                pre_source,
                pre_source_hash,
                pre_state,
                edit_profile,
                retained_profile,
            } = context;
            let mut edit = Edit::with_limits(edit_profile, limits)?;
            edit.set_property_value(
                ChildSelector::Property {
                    action: ActionSelector::ordinal(0),
                    index: 0,
                },
                CAP_VALUE,
            )?;
            match edit.finish() {
                Err(DrawingError::Limit { resource, limit })
                    if resource == "ink action output bytes" =>
                {
                    Ok(Work::Rejected {
                        resource,
                        limit,
                        pre_source,
                        pre_source_hash,
                        pre_state,
                        retained_profile,
                    })
                },
                Err(error) => Err(error.into()),
                Ok(_) => Err("caller output cap unexpectedly accepted the edit".into()),
            }
        },
    }
}

#[derive(Clone, Copy)]
struct Sample {
    elapsed_ns: u64,
    allocation: AllocDelta,
    expected_success: bool,
    actual_success: bool,
    semantic_ok: Option<bool>,
    source_exact: Option<bool>,
    source_shared: Option<bool>,
    inverse_ok: Option<bool>,
    opaque_preserved: Option<bool>,
    output_exact: Option<bool>,
    rejection_ok: Option<bool>,
    rejection_source_unchanged: Option<bool>,
    rejection_state_unchanged: Option<bool>,
    rejection_resource: Option<&'static str>,
    rejection_limit: Option<usize>,
    rejection_pre_source_hash: Option<u64>,
    rejection_post_source_hash: Option<u64>,
}

#[derive(Clone, Copy)]
struct Validation {
    semantic_ok: Option<bool>,
    source_exact: Option<bool>,
    source_shared: Option<bool>,
    inverse_ok: Option<bool>,
    opaque_preserved: Option<bool>,
    output_exact: Option<bool>,
    rejection_ok: Option<bool>,
    rejection_source_unchanged: Option<bool>,
    rejection_state_unchanged: Option<bool>,
    rejection_resource: Option<&'static str>,
    rejection_limit: Option<usize>,
    rejection_pre_source_hash: Option<u64>,
    rejection_post_source_hash: Option<u64>,
}

fn action_ids(profile: &Profile) -> Vec<String> {
    profile
        .actions()
        .map(|action| action.xml_id().unwrap_or_default().to_owned())
        .collect()
}

fn original_action_ids(count: usize) -> Vec<String> {
    (0..count).map(|index| format!("a{index}")).collect()
}

fn ids_match(profile: &Profile, expected: &[String]) -> bool {
    action_ids(profile) == expected
}

fn property_is(action: &litchi_drawingml::ink::actions::Action, expected: &str) -> bool {
    action
        .properties()
        .first()
        .is_some_and(|property| property.name() == "kind" && property.value() == expected)
}

fn validate_work(inputs: &Inputs, lane: &str, work: &Work) -> Result<Validation> {
    let kind = lane_kind(lane).ok_or_else(|| format!("unknown lane {lane}"))?;
    if matches!(kind, LaneKind::Draft { .. }) {
        let Work::Prepared(prepared) = work else {
            return Err(format!("lane {lane} returned an unexpected result").into());
        };
        let count = action_count_for_lane(lane);
        let opaque = matches!(kind, LaneKind::Draft { opaque: true });
        let expected_ids = original_action_ids(count);
        let semantic_ok = prepared.profile().actions().count() == count
            && ids_match(prepared.profile(), &expected_ids)
            && (!opaque || contains_opaque(prepared.as_bytes()));
        let opaque_preserved = !opaque || contains_opaque(prepared.as_bytes());
        return Ok(Validation {
            semantic_ok: Some(semantic_ok),
            source_exact: None,
            source_shared: None,
            inverse_ok: None,
            opaque_preserved: Some(opaque_preserved),
            output_exact: Some(prepared.profile().source() == prepared.as_bytes()),
            rejection_ok: None,
            rejection_source_unchanged: None,
            rejection_state_unchanged: None,
            rejection_resource: None,
            rejection_limit: None,
            rejection_pre_source_hash: None,
            rejection_post_source_hash: None,
        });
    }

    if let Work::Rejected {
        resource,
        limit,
        pre_source,
        pre_source_hash,
        pre_state,
        retained_profile,
    } = work
    {
        let post_source = retained_profile.source();
        let post_source_hash = source_hash(post_source);
        let post_profile = actions::read_profile(post_source)?;
        let post_state = profile_state(&post_profile);
        let no_op = Edit::from_profile(&post_profile)?.finish()?;
        let no_op_state = profile_state(no_op.profile());
        let source_unchanged = post_source == pre_source.as_slice()
            && post_source_hash == *pre_source_hash
            && no_op.as_bytes() == pre_source.as_slice();
        let state_unchanged = profile_state(retained_profile) == *pre_state
            && post_state == *pre_state
            && no_op_state == *pre_state;
        return Ok(Validation {
            semantic_ok: None,
            source_exact: None,
            source_shared: None,
            inverse_ok: None,
            opaque_preserved: Some(contains_opaque(post_source)),
            output_exact: None,
            rejection_ok: Some(
                *resource == "ink action output bytes"
                    && *limit == pre_source.len()
                    && source_unchanged
                    && state_unchanged,
            ),
            rejection_source_unchanged: Some(source_unchanged),
            rejection_state_unchanged: Some(state_unchanged),
            rejection_resource: Some(*resource),
            rejection_limit: Some(*limit),
            rejection_pre_source_hash: Some(*pre_source_hash),
            rejection_post_source_hash: Some(post_source_hash),
        });
    }

    let fixture = source_fixture(inputs, lane)?;
    let source = fixture.profile.source();
    let Work::Commit(commit) = work else {
        return Err(format!("lane {lane} returned an unexpected result").into());
    };
    let after = commit.as_bytes();
    let profile = commit.profile();
    let patch = commit.patch();
    let source_exact = patch.before() == source;
    let output_exact = patch.after() == after && profile.source() == after;
    let inverse_ok = patch.apply(source)? == after && patch.inverse().apply(after)? == source;
    let opaque_preserved = contains_opaque(after);
    let count = action_count_for_lane(lane);
    let action_count = profile.actions().count();
    let mut semantic_ok = action_count == result_action_count_for_lane(lane);
    let source_shared = matches!(kind, LaneKind::Noop)
        && after.as_ptr() == source.as_ptr()
        && after.len() == source.len();
    let original_ids = original_action_ids(count);

    match kind {
        LaneKind::Scalar => {
            semantic_ok &= ids_match(profile, &original_ids)
                && profile
                    .actions()
                    .nth(count - 1)
                    .is_some_and(|action| action.action_type().as_str() == "edited");
        },
        LaneKind::ScalarBatch => {
            semantic_ok &= ids_match(profile, &original_ids)
                && profile
                    .actions()
                    .enumerate()
                    .all(|(index, action)| property_is(action, &format!("batch-{index}")));
        },
        LaneKind::ScalarCoalesce => {
            let expected = format!("coalesced-{}", count - 1);
            semantic_ok &= ids_match(profile, &original_ids)
                && profile
                    .actions()
                    .nth(count - 1)
                    .is_some_and(|action| property_is(action, &expected));
        },
        LaneKind::Noop => {
            semantic_ok &= ids_match(profile, &original_ids) && after == source;
        },
        LaneKind::Add => {
            let mut expected = original_ids;
            expected.push(format!("added-{count}"));
            semantic_ok &= ids_match(profile, &expected)
                && profile
                    .actions()
                    .last()
                    .is_some_and(|action| property_is(action, "added"));
        },
        LaneKind::InsertBatch => {
            let mut expected = original_ids;
            expected.extend((0..count).map(|index| format!("added-{index}")));
            semantic_ok &= ids_match(profile, &expected)
                && profile
                    .actions()
                    .skip(count)
                    .all(|action| property_is(action, "added"));
        },
        LaneKind::Remove => {
            let expected = original_ids[..count - 1].to_vec();
            semantic_ok &= ids_match(profile, &expected);
        },
        LaneKind::RemoveBatch => {
            let expected = original_ids
                .into_iter()
                .enumerate()
                .filter_map(|(index, id)| (index % 2 == 0).then_some(id))
                .collect::<Vec<_>>();
            semantic_ok &= ids_match(profile, &expected);
        },
        LaneKind::ClearBatch => {
            semantic_ok &= ids_match(profile, &original_ids)
                && profile.actions().enumerate().all(|(index, action)| {
                    if index % 2 == 1 {
                        action.properties().is_empty() && action.children().is_empty()
                    } else {
                        property_is(action, "v")
                    }
                })
                && profile
                    .actions()
                    .next()
                    .is_some_and(|action| !action.children().is_empty());
        },
        LaneKind::Move => {
            let mut expected = vec![format!("a{}", count - 1)];
            expected.extend((0..count - 1).map(|index| format!("a{index}")));
            semantic_ok &= ids_match(profile, &expected);
        },
        LaneKind::MoveBatch => {
            let expected = (0..count)
                .map(|index| {
                    let expected_index = if index % 2 == 0 { index + 1 } else { index - 1 };
                    format!("a{expected_index}")
                })
                .collect::<Vec<_>>();
            semantic_ok &= ids_match(profile, &expected);
        },
        LaneKind::CapRefusal | LaneKind::Draft { .. } => {},
    }
    Ok(Validation {
        semantic_ok: Some(semantic_ok),
        source_exact: Some(source_exact),
        source_shared: Some(source_shared),
        inverse_ok: Some(inverse_ok),
        opaque_preserved: Some(opaque_preserved),
        output_exact: Some(output_exact),
        rejection_ok: None,
        rejection_source_unchanged: None,
        rejection_state_unchanged: None,
        rejection_resource: None,
        rejection_limit: None,
        rejection_pre_source_hash: None,
        rejection_post_source_hash: None,
    })
}

fn run_sample(inputs: &Inputs, lane: &str) -> Result<Sample> {
    reset_counters();
    let rejection = if is_refusal(lane) {
        Some(rejection_context(inputs, lane)?)
    } else {
        None
    };
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let work = operation(inputs, lane, rejection)?;
    black_box(&work);
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
    let allocation = before.delta(AllocSnapshot::now());
    let actual_success = !matches!(&work, Work::Rejected { .. });
    let expected = !is_refusal(lane);
    if actual_success != expected {
        return Err(
            format!("lane {lane} returned success={actual_success}, expected={expected}").into(),
        );
    }
    if allocation.invalid || allocation.failed != 0 || !allocation.balanced() {
        return Err(format!("allocator accounting failed for {lane}").into());
    }
    let validation = validate_work(inputs, lane, &work)?;
    if is_refusal(lane) {
        if validation.rejection_ok != Some(true)
            || validation.rejection_source_unchanged != Some(true)
            || validation.rejection_state_unchanged != Some(true)
            || validation.opaque_preserved != Some(true)
            || validation.semantic_ok.is_some()
            || validation.source_exact.is_some()
            || validation.source_shared.is_some()
            || validation.inverse_ok.is_some()
            || validation.output_exact.is_some()
            || validation.rejection_pre_source_hash.is_none()
            || validation.rejection_post_source_hash.is_none()
        {
            return Err(format!("caller-cap rejection gate failed for {lane}").into());
        }
    } else if matches!(lane_kind(lane), Some(LaneKind::Draft { .. })) {
        if validation.semantic_ok != Some(true)
            || validation.opaque_preserved != Some(true)
            || validation.output_exact != Some(true)
            || validation.source_exact.is_some()
            || validation.source_shared.is_some()
            || validation.inverse_ok.is_some()
            || validation.rejection_ok.is_some()
            || validation.rejection_source_unchanged.is_some()
            || validation.rejection_state_unchanged.is_some()
            || validation.rejection_resource.is_some()
            || validation.rejection_limit.is_some()
            || validation.rejection_pre_source_hash.is_some()
            || validation.rejection_post_source_hash.is_some()
        {
            return Err(format!("post-timer draft semantic gate failed for {lane}").into());
        }
    } else if validation.semantic_ok != Some(true)
        || validation.source_exact != Some(true)
        || validation.inverse_ok != Some(true)
        || validation.opaque_preserved != Some(true)
        || validation.output_exact != Some(true)
        || validation.rejection_ok.is_some()
        || validation.rejection_source_unchanged.is_some()
        || validation.rejection_state_unchanged.is_some()
        || validation.rejection_resource.is_some()
        || validation.rejection_limit.is_some()
        || validation.rejection_pre_source_hash.is_some()
        || validation.rejection_post_source_hash.is_some()
    {
        return Err(format!("post-timer semantic gate failed for {lane}").into());
    }
    Ok(Sample {
        elapsed_ns,
        allocation,
        expected_success: expected,
        actual_success,
        semantic_ok: validation.semantic_ok,
        source_exact: validation.source_exact,
        source_shared: validation.source_shared,
        inverse_ok: validation.inverse_ok,
        opaque_preserved: validation.opaque_preserved,
        output_exact: validation.output_exact,
        rejection_ok: validation.rejection_ok,
        rejection_source_unchanged: validation.rejection_source_unchanged,
        rejection_state_unchanged: validation.rejection_state_unchanged,
        rejection_resource: validation.rejection_resource,
        rejection_limit: validation.rejection_limit,
        rejection_pre_source_hash: validation.rejection_pre_source_hash,
        rejection_post_source_hash: validation.rejection_post_source_hash,
    })
}

fn run_lane(inputs: &Inputs, lane: &str, warmup: usize, samples: usize) -> Result<String> {
    for _ in 0..warmup {
        let _ = run_sample(inputs, lane)?;
    }
    let mut values = Vec::with_capacity(samples);
    for _ in 0..samples {
        values.push(run_sample(inputs, lane)?);
    }
    let input_bytes =
        if is_refusal(lane) || !matches!(lane_kind(lane), Some(LaneKind::Draft { .. })) {
            source_fixture(inputs, lane)?.profile.source().len()
        } else {
            0
        };
    let mut output = String::new();
    write!(
        &mut output,
        "{{\"schema\":\"ink-action-edit-profile-v2\",\"lane\":\"{lane}\",\"pid\":{},\"warmup\":{warmup},\"sample_count\":{},\"expected_success\":{},\"action_count\":{},\"result_action_count\":{},\"operation_count\":{},\"input_bytes\":{input_bytes},\"samples\":[",
        std::process::id(),
        values.len(),
        !is_refusal(lane),
        action_count_for_lane(lane),
        result_action_count_for_lane(lane),
        operation_count_for_lane(lane),
    )?;
    for (index, sample) in values.iter().enumerate() {
        if index != 0 {
            output.push(',');
        }
        write!(
            &mut output,
            "{{\"elapsed_ns\":{},\"requested_alloc_bytes\":{},\"direct_allocated_bytes\":{},\"realloc_old_bytes\":{},\"realloc_new_bytes\":{},\"deallocated_bytes\":{},\"alloc_calls\":{},\"realloc_calls\":{},\"dealloc_calls\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"alloc_balance_ok\":{},\"alloc_invalid\":{},\"alloc_failed\":{},\"expected_success\":{},\"actual_success\":{},\"semantic_ok\":{},\"source_exact\":{},\"source_shared\":{},\"inverse_ok\":{},\"opaque_preserved\":{},\"output_exact\":{},\"rejection_ok\":{},\"rejection_source_unchanged\":{},\"rejection_state_unchanged\":{},\"rejection_resource\":{},\"rejection_limit\":{},\"rejection_pre_source_hash\":{},\"rejection_post_source_hash\":{}}}",
            sample.elapsed_ns,
            sample.allocation.requested(),
            sample.allocation.direct,
            sample.allocation.realloc_old,
            sample.allocation.realloc_new,
            sample.allocation.deallocated,
            sample.allocation.calls,
            sample.allocation.realloc_calls,
            sample.allocation.dealloc_calls,
            sample.allocation.live_before,
            sample.allocation.live_after,
            sample.allocation.peak_delta,
            sample.allocation.balanced(),
            sample.allocation.invalid,
            sample.allocation.failed,
            sample.expected_success,
            sample.actual_success,
            json_bool(sample.semantic_ok),
            json_bool(sample.source_exact),
            json_bool(sample.source_shared),
            json_bool(sample.inverse_ok),
            json_bool(sample.opaque_preserved),
            json_bool(sample.output_exact),
            json_bool(sample.rejection_ok),
            json_bool(sample.rejection_source_unchanged),
            json_bool(sample.rejection_state_unchanged),
            json_opt_string(sample.rejection_resource),
            json_opt_usize(sample.rejection_limit),
            json_opt_u64(sample.rejection_pre_source_hash),
            json_opt_u64(sample.rejection_post_source_hash),
        )?;
    }
    output.push_str("]}\n");
    Ok(output)
}

fn json_bool(value: Option<bool>) -> &'static str {
    match value {
        Some(true) => "true",
        Some(false) => "false",
        None => "null",
    }
}

fn json_opt_string(value: Option<&str>) -> String {
    value.map_or_else(|| "null".to_owned(), |item| format!("\"{item}\""))
}

fn json_opt_usize(value: Option<usize>) -> String {
    value.map_or_else(|| "null".to_owned(), |item| item.to_string())
}

fn json_opt_u64(value: Option<u64>) -> String {
    value.map_or_else(|| "null".to_owned(), |item| item.to_string())
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let value = value.parse::<usize>()?;
    if value == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(value)
}

fn main() -> Result<()> {
    let mut lane = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--lane" => lane = Some(args.next().ok_or("missing --lane value")?),
            "--warmup" => {
                warmup = parse_positive(&args.next().ok_or("missing --warmup value")?, "--warmup")?
            },
            "--samples" => {
                samples =
                    parse_positive(&args.next().ok_or("missing --samples value")?, "--samples")?
            },
            "--help" | "-h" => {
                println!(
                    "--lane <draft_*|scalar_edit_*|scalar_batch_*|scalar_coalesce_*|no_op_*|add_*|insert_batch_*|remove_*|remove_batch_*|clear_batch_*|move_*|move_batch_*|cap_refusal_*> [--warmup N] [--samples N]"
                );
                return Ok(());
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    let lane = lane.ok_or("--lane is required")?;
    if !LANES.contains(&lane.as_str()) {
        return Err(format!("unknown lane {lane}").into());
    }
    allocator_counter_self_test()?;
    let inputs = Inputs::load()?;
    print!("{}", run_lane(&inputs, &lane, warmup, samples)?);
    Ok(())
}
