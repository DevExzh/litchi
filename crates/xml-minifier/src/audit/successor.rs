//! Successor proofs: the source verdict on a replacement, proved from a
//! verified original's recorded content boundaries without scanning the
//! original again (records 0313 and 0317).
//!
//! [`super::VerifiedSource::verify_indexed`] audits a payload like
//! [`super::verify_source`] and, through an [`Observer`] that only reads the
//! auditor's state, records every *content boundary* down to an
//! [`IndexDepth`]: each offset between two accepted tokens at which at most
//! that many elements are open (two by default: the document element's
//! content and its children's). A boundary keeps the element whose content it
//! lies in (its *parent*), the aggregate counters the audit had charged up to
//! it, and, shared per parent, the namespace bindings in scope there.
//!
//! Two boundaries of the same parent enclose a run of complete sibling
//! tokens. A replacement that repeats the original outside such a run is
//! proved by auditing only the bytes that took the run's place, as
//! [`super::audit_window`] audits an element window: from the parent's depth
//! and bindings, balanced, beginning with `<` and ending with markup, with the
//! aggregate budgets re-totalled for the whole replacement. The invariant at
//! [`super::scan`] is what makes that sound, exactly as for element windows;
//! a boundary carries every piece of state the invariant lets a source-policy
//! check read. The run may be empty, which is a pure insertion.
//!
//! # Derived indexes (record 0317)
//!
//! A successor proved this way carries a *derived* index, so that the next
//! edit of its bytes is proved from its own boundaries too. The derived index
//! is exactly the index a recording audit of the replacement would make:
//!
//! * a boundary before the run is the original's: the replacement repeats
//!   every byte before it, and the scan up to it is the same scan;
//! * a boundary inside the window is recorded while the window is audited,
//!   at the window's absolute offset and with the run start's counters added:
//!   by the invariant at [`super::scan`], the window scan agrees with the
//!   complete scan on every token of the window, its depth and its bindings;
//! * a boundary at or after the run's end is the original's, moved by the
//!   length change and by the change of the counters the window was charged,
//!   since the window is balanced and closes every binding it opens.
//!
//! Debug builds compare every derived index with the index a recording audit
//! of the replacement makes.
//!
//! # Bounds
//!
//! One boundary per 32 input bytes plus a constant, at 20 bytes each; a
//! parent costs 4 bytes and every parent of an index is referenced by one of
//! its boundaries; the namespace snapshots are charged against the
//! multi-window snapshot budget (8 MiB). Exceeding any bound, or an input
//! longer than `u32` offsets address, records or derives no index, and a
//! successor of such a proof is audited completely.

use std::sync::Arc;

use super::{
    Counters, Limits, MAX_REPLACEMENT_NAMESPACE_CLONE_BYTES, Namespaces, Observer, Policy,
    ReplacementProof, Report, State, Token, add_counters, audit_window_observed, common_prefix_len,
    common_suffix_len, verify_observed,
};

/// The deepest content an index may record boundaries in.
pub(super) const MAX_INDEXED_DEPTH: usize = 3;
/// Boundaries recorded per input byte, as a divisor.
const BYTES_PER_BOUNDARY: usize = 32;
/// Boundaries recorded beyond the per-byte allowance.
const EXTRA_BOUNDARIES: usize = 64;

/// How deep a recording audit records content boundaries.
///
/// A boundary is recorded at every offset between two tokens at which at
/// most this many elements are open. Two, the default, covers the document
/// element's content and its children's: WordprocessingML body blocks and
/// SpreadsheetML rows. Three also covers the grandchildren's content, such as
/// the shapes of a PresentationML slide's shape tree
/// (`p:sld/p:cSld/p:spTree`). The bounds of the index are the same at every
/// depth; a deeper index only reaches them sooner.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct IndexDepth(u8);

impl IndexDepth {
    /// The document element's content and its children's.
    pub const CHILDREN: Self = Self(2);
    /// [`Self::CHILDREN`] and the grandchildren's content too.
    pub const GRANDCHILDREN: Self = Self(3);

    /// The number of open elements down to which boundaries are recorded.
    #[must_use]
    pub const fn get(self) -> usize {
        self.0 as usize
    }
}

impl Default for IndexDepth {
    fn default() -> Self {
        Self::CHILDREN
    }
}

/// One recorded content boundary. Every field fits `u32` because the index
/// is recorded only for inputs shorter than `u32::MAX` bytes, and no counter
/// exceeds the input's length plus one.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Boundary {
    offset: u32,
    /// Index into [`BoundaryIndex::parents`].
    parent: u32,
    attributes: u32,
    events: u32,
    text_bytes: u32,
}

impl Boundary {
    fn counters(self) -> Counters {
        Counters {
            attributes: self.attributes as usize,
            events: self.events as usize,
            text_bytes: self.text_bytes as usize,
        }
    }

    fn at(offset: usize, parent: u32, counters: Counters) -> Option<Self> {
        Some(Self {
            offset: u32::try_from(offset).ok()?,
            parent,
            attributes: u32::try_from(counters.attributes).ok()?,
            events: u32::try_from(counters.events).ok()?,
            text_bytes: u32::try_from(counters.text_bytes).ok()?,
        })
    }
}

/// How many boundaries an index of an `input_len`-byte payload may hold.
const fn boundary_allowance(input_len: usize) -> usize {
    (input_len / BYTES_PER_BOUNDARY).saturating_add(EXTRA_BOUNDARIES)
}

/// The charge of keeping one namespace snapshot behind an `Arc`.
fn scope_charge(namespaces: &Namespaces) -> Option<usize> {
    namespaces
        .clone_memory_charge()?
        .checked_add(2 * size_of::<usize>())
}

/// The content boundaries one source audit recorded, with the audit's
/// totals.
pub(super) struct BoundaryIndex {
    /// How deep boundaries were recorded, at most [`MAX_INDEXED_DEPTH`].
    depth: usize,
    /// The original's aggregate totals, as its [`Report`] gives them.
    totals: Counters,
    /// `levels[d - 1]`: the boundaries at which `d` elements are open, in
    /// offset order. Levels deeper than `depth` are empty.
    levels: [Vec<Boundary>; MAX_INDEXED_DEPTH],
    /// Each parent's index into `scopes`.
    parents: Vec<u32>,
    /// The parent whose content is the document element's.
    root: u32,
    /// The namespace bindings in scope in a parent's content.
    scopes: Vec<Arc<Namespaces>>,
    /// What `scopes` is charged against the snapshot budget.
    scope_bytes: usize,
}

impl BoundaryIndex {
    /// How many boundaries the index holds.
    pub(super) fn len(&self) -> usize {
        self.levels.iter().map(Vec::len).sum()
    }

    /// How deep the index records boundaries.
    pub(super) fn index_depth(&self) -> IndexDepth {
        IndexDepth(u8::try_from(self.depth).unwrap_or(u8::MAX))
    }

    /// Prove `replacement`'s source-policy verdict from that of `original`,
    /// the bytes this index was recorded from, or return `None` so the caller
    /// audits the replacement completely. A proof comes with the index of the
    /// replacement derived from this one, when it fits its bounds.
    ///
    /// Only success is decided here, as in `Window::prove`.
    pub(super) fn prove(
        &self,
        original: &[u8],
        replacement: &[u8],
        limits: Limits,
    ) -> Option<(ReplacementProof, Option<Self>)> {
        if replacement.len() > limits.bytes || original == replacement {
            return None;
        }
        let prefix = common_prefix_len(original, replacement);
        let suffix = common_suffix_len(&original[prefix..], &replacement[prefix..]);
        // `original[..prefix]` and `original[tail..]` are repeated by the
        // replacement, and `prefix <= tail`.
        let tail = original.len() - suffix;
        // An insertion usually repeats the start of the sibling after it, so
        // the common prefix runs into that sibling. Measuring the common tail
        // again from the run's start finds the insertion itself; when the
        // bytes that took that smaller run's place do not begin and end with
        // markup, the run the first measure gives is tried too.
        let mut tried = None;
        for refine in [true, false] {
            let Some((depth, start, end)) =
                self.enclosing(original, replacement, prefix, tail, refine)
            else {
                continue;
            };
            if tried == Some((start.offset, end.offset)) {
                continue;
            }
            tried = Some((start.offset, end.offset));
            if let Some(proof) =
                self.prove_run(original.len(), replacement, depth, start, end, limits)
            {
                return Some(proof);
            }
        }
        None
    }

    /// Prove the replacement in which the bytes from `start` up to the
    /// matching place of `end` took the place of the original's run
    /// `start..end`, and derive the replacement's index.
    fn prove_run(
        &self,
        original_len: usize,
        replacement: &[u8],
        depth: usize,
        start: Boundary,
        end: Boundary,
        limits: Limits,
    ) -> Option<(ReplacementProof, Option<Self>)> {
        let run_start = start.offset as usize;
        let run_end = end.offset as usize;
        let window_end = run_end
            .checked_add(replacement.len())?
            .checked_sub(original_len)?;
        // A window of more than half the payload saves nothing.
        if window_end < run_start || (window_end - run_start).checked_mul(2)? > replacement.len() {
            return None;
        }
        let removed = end.counters().since(start.counters())?;
        let scope = self
            .parents
            .get(start.parent as usize)
            .and_then(|scope| self.scopes.get(*scope as usize))?;
        let bytes = replacement.get(run_start..window_end)?;
        let mut recorder =
            BoundaryRecorder::resume(self, depth, start, boundary_allowance(replacement.len()));
        let added = audit_window_observed(bytes, limits, depth, scope.try_clone()?, &mut recorder)?;
        let totals = self.totals.replaced(removed, added)?;
        if !totals.within(limits) {
            return None;
        }
        let proof = ReplacementProof::Splice {
            original: run_start..run_end,
            replacement: run_start..window_end,
        };
        let derived = self.spliced(
            replacement.len(),
            start,
            end,
            window_end,
            add_counters(start.counters(), added)?,
            totals,
            recorder,
        );
        Some((proof, derived))
    }

    /// The index of the replacement in which the window ending at
    /// `window_end` took the place of the run `start..end`: this index's
    /// boundaries before the run, the window's, which `recorder` recorded,
    /// and this index's boundaries from the run's end on, moved past the
    /// window. `after_window` holds the counters charged up to the window's
    /// end. `None` when the result exceeds a bound.
    #[allow(
        clippy::too_many_arguments,
        reason = "one splice's geometry, each piece named"
    )]
    fn spliced(
        &self,
        replacement_len: usize,
        start: Boundary,
        end: Boundary,
        window_end: usize,
        after_window: Counters,
        totals: Counters,
        recorder: BoundaryRecorder,
    ) -> Option<Self> {
        if recorder.failed || u32::try_from(replacement_len).map_or(true, |len| len == u32::MAX) {
            return None;
        }
        let allowance = boundary_allowance(replacement_len);
        let run_start = start.offset as usize;
        let run_end = end.offset as usize;
        let BoundaryRecorder {
            levels: window_levels,
            parents,
            scopes,
            ..
        } = recorder;
        // Every boundary from the run's end on moves by these.
        let shift = |from: u32, to: usize| -> Option<i32> {
            i32::try_from(i64::try_from(to).ok()? - i64::from(from)).ok()
        };
        let after_counters = end.counters();
        let offset_shift = shift(end.offset, window_end)?;
        let attributes_shift = shift(
            u32::try_from(after_counters.attributes).ok()?,
            after_window.attributes,
        )?;
        let events_shift = shift(
            u32::try_from(after_counters.events).ok()?,
            after_window.events,
        )?;
        let text_shift = shift(
            u32::try_from(after_counters.text_bytes).ok()?,
            after_window.text_bytes,
        )?;
        let mut levels: [Vec<Boundary>; MAX_INDEXED_DEPTH] = Default::default();
        let mut count = 0usize;
        // Whether a dropped boundary was the only reference to a parent other
        // than the run's own, which the start boundary keeps referenced.
        let mut orphaned = false;
        for (level, (old, window)) in levels
            .iter_mut()
            .zip(self.levels.iter().zip(window_levels.iter()))
        {
            let before = old.partition_point(|boundary| boundary.offset as usize <= run_start);
            let after = old.partition_point(|boundary| (boundary.offset as usize) < run_end);
            let inside = window.partition_point(|boundary| (boundary.offset as usize) < window_end);
            let total = before
                .checked_add(inside)?
                .checked_add(old.len().checked_sub(after)?)?;
            count = count.checked_add(total)?;
            if count > allowance {
                return None;
            }
            // The run's start boundary is kept, so for an empty run `before`
            // is one past `after` and nothing is dropped.
            orphaned |= old
                .get(before.min(after)..after)?
                .iter()
                .any(|boundary| boundary.parent != start.parent);
            level.try_reserve_exact(total).ok()?;
            level.extend_from_slice(&old[..before]);
            level.extend_from_slice(&window[..inside]);
            for boundary in &old[after..] {
                level.push(Boundary {
                    offset: boundary.offset.checked_add_signed(offset_shift)?,
                    parent: boundary.parent,
                    attributes: boundary.attributes.checked_add_signed(attributes_shift)?,
                    events: boundary.events.checked_add_signed(events_shift)?,
                    text_bytes: boundary.text_bytes.checked_add_signed(text_shift)?,
                });
            }
        }
        if !orphaned && parents.len() == self.parents.len() {
            // No parent was opened in the window and none lost its last
            // boundary: the parents and scopes are this index's.
            return Some(Self {
                depth: self.depth,
                totals,
                levels,
                parents,
                root: self.root,
                scopes,
                scope_bytes: self.scope_bytes,
            });
        }
        compact(self.depth, totals, levels, &parents, &scopes, self.root)
    }

    /// The deepest pair of boundaries of one parent that enclose every
    /// difference: the last one at or before `prefix`, and the first one at
    /// or after the common tail, with their depth. With `refine`, the common
    /// tail is measured again from the first boundary, which can only move
    /// it earlier; `original[..prefix]` is repeated, so any earlier start is
    /// too.
    fn enclosing(
        &self,
        original: &[u8],
        replacement: &[u8],
        prefix: usize,
        tail: usize,
        refine: bool,
    ) -> Option<(usize, Boundary, Boundary)> {
        (1..=self.depth).rev().find_map(|depth| {
            let level = &self.levels[depth - 1];
            let before = level.partition_point(|boundary| boundary.offset as usize <= prefix);
            let start = *level.get(before.checked_sub(1)?)?;
            let tail = if refine {
                let from = start.offset as usize;
                original.len() - common_suffix_len(original.get(from..)?, replacement.get(from..)?)
            } else {
                tail
            };
            let after = level.partition_point(|boundary| (boundary.offset as usize) < tail);
            let end = *level.get(after)?;
            (start.parent == end.parent).then_some((depth, start, end))
        })
    }

    /// The namespace bindings in scope in `parent`'s content.
    #[cfg(debug_assertions)]
    fn scope_of(&self, parent: usize) -> Option<&Namespaces> {
        let scope = *self.parents.get(parent)? as usize;
        self.scopes.get(scope).map(|scope| &**scope)
    }

    /// Whether `other` records the same boundaries: the same offsets and
    /// counters at every level, parents that correspond one to one, and the
    /// same bindings in each corresponding parent's scope. Parent and scope
    /// numbering may differ.
    #[cfg(debug_assertions)]
    pub(super) fn same_boundaries(&self, other: &Self) -> bool {
        let mut forward = vec![u32::MAX; self.parents.len()];
        let mut backward = vec![u32::MAX; other.parents.len()];
        let mut corresponds = |left: u32, right: u32| {
            let (Some(forward), Some(backward)) = (
                forward.get_mut(left as usize),
                backward.get_mut(right as usize),
            ) else {
                return false;
            };
            if *forward == u32::MAX && *backward == u32::MAX {
                *forward = right;
                *backward = left;
            }
            *forward == right && *backward == left
        };
        let levels_match = self.depth == other.depth
            && self.totals == other.totals
            && self.levels.iter().zip(&other.levels).all(|(left, right)| {
                left.len() == right.len()
                    && left.iter().zip(right).all(|(left, right)| {
                        left.offset == right.offset
                            && left.counters() == right.counters()
                            && corresponds(left.parent, right.parent)
                    })
            })
            && corresponds(self.root, other.root);
        levels_match
            && forward.iter().enumerate().all(|(left, right)| {
                *right == u32::MAX || {
                    match (self.scope_of(left), other.scope_of(*right as usize)) {
                        (Some(left), Some(right)) => left.same_bindings(right),
                        _ => false,
                    }
                }
            })
    }
}

/// Build an index from spliced `levels`, keeping only the parents its
/// boundaries reference (and `root`), and only the scopes those parents use.
fn compact(
    depth: usize,
    totals: Counters,
    mut levels: [Vec<Boundary>; MAX_INDEXED_DEPTH],
    parents: &[u32],
    scopes: &[Arc<Namespaces>],
    root: u32,
) -> Option<BoundaryIndex> {
    const UNUSED: u32 = u32::MAX;
    let mut parent_map = Vec::new();
    parent_map.try_reserve_exact(parents.len()).ok()?;
    parent_map.resize(parents.len(), UNUSED);
    *parent_map.get_mut(root as usize)? = 0;
    for boundary in levels.iter().flatten() {
        *parent_map.get_mut(boundary.parent as usize)? = 0;
    }
    let mut scope_map = Vec::new();
    scope_map.try_reserve_exact(scopes.len()).ok()?;
    scope_map.resize(scopes.len(), UNUSED);
    let mut kept_parents = Vec::new();
    let mut kept_scopes = Vec::new();
    let mut scope_bytes = 0usize;
    for (old, slot) in parent_map.iter_mut().enumerate() {
        if *slot == UNUSED {
            continue;
        }
        let scope = *parents.get(old)? as usize;
        let mapped = scope_map.get_mut(scope)?;
        if *mapped == UNUSED {
            let namespaces = scopes.get(scope)?;
            scope_bytes = scope_bytes.checked_add(scope_charge(namespaces)?)?;
            if scope_bytes > MAX_REPLACEMENT_NAMESPACE_CLONE_BYTES {
                return None;
            }
            kept_scopes.try_reserve(1).ok()?;
            kept_scopes.push(Arc::clone(namespaces));
            *mapped = u32::try_from(kept_scopes.len() - 1).ok()?;
        }
        kept_parents.try_reserve(1).ok()?;
        kept_parents.push(*mapped);
        *slot = u32::try_from(kept_parents.len() - 1).ok()?;
    }
    for boundary in levels.iter_mut().flatten() {
        boundary.parent = parent_map[boundary.parent as usize];
    }
    Some(BoundaryIndex {
        depth,
        totals,
        levels,
        parents: kept_parents,
        root: parent_map[root as usize],
        scopes: kept_scopes,
        scope_bytes,
    })
}

/// Records the [`BoundaryIndex`] of one source audit, or the boundaries of
/// one window audited in place of a run of an index's original.
///
/// It only reads the auditor's state, so the audit's verdict is unchanged.
/// Any bound it cannot keep stops it, and then no index is recorded.
pub(super) struct BoundaryRecorder {
    depth: usize,
    levels: [Vec<Boundary>; MAX_INDEXED_DEPTH],
    parents: Vec<u32>,
    scopes: Vec<Arc<Namespaces>>,
    /// The parent whose content is open at each depth.
    current: [u32; MAX_INDEXED_DEPTH],
    /// Bindings in scope in the document element's content.
    root_bindings: usize,
    boundaries_left: usize,
    scope_bytes: usize,
    /// Added to every recorded offset and counter: zero for a complete
    /// audit, the run's start for a window.
    offset_base: usize,
    counter_base: Counters,
    failed: bool,
}

impl BoundaryRecorder {
    pub(super) fn new(input_len: usize, depth: IndexDepth) -> Self {
        Self {
            depth: depth.get().min(MAX_INDEXED_DEPTH),
            levels: Default::default(),
            parents: Vec::new(),
            scopes: Vec::new(),
            current: [0; MAX_INDEXED_DEPTH],
            root_bindings: 0,
            boundaries_left: boundary_allowance(input_len),
            scope_bytes: 0,
            offset_base: 0,
            counter_base: Counters::default(),
            failed: u32::try_from(input_len).map_or(true, |len| len == u32::MAX),
        }
    }

    /// A recorder for the window audited in place of the run that starts at
    /// `start`, with `depth` elements open, in `index`'s original. It knows
    /// `index`'s parents, so the window's boundaries name them.
    fn resume(index: &BoundaryIndex, depth: usize, start: Boundary, allowance: usize) -> Self {
        let mut recorder = Self {
            depth: index.depth,
            levels: Default::default(),
            parents: Vec::new(),
            scopes: Vec::new(),
            current: [index.root; MAX_INDEXED_DEPTH],
            root_bindings: index
                .parents
                .get(index.root as usize)
                .and_then(|scope| index.scopes.get(*scope as usize))
                .map_or(0, |scope| scope.bindings()),
            boundaries_left: allowance,
            scope_bytes: index.scope_bytes,
            offset_base: start.offset as usize,
            counter_base: start.counters(),
            failed: false,
        };
        // The window never closes an element it did not open, so only the
        // run's own parent and the document element's are named from it.
        if let Some(slot) = depth
            .checked_sub(1)
            .and_then(|at| recorder.current.get_mut(at))
        {
            *slot = start.parent;
        }
        let copied = recorder
            .parents
            .try_reserve_exact(index.parents.len())
            .is_ok()
            && recorder
                .scopes
                .try_reserve_exact(index.scopes.len())
                .is_ok();
        if copied {
            recorder.parents.extend_from_slice(&index.parents);
            recorder.scopes.extend(index.scopes.iter().cloned());
        } else {
            recorder.failed = true;
        }
        recorder
    }

    /// The recorded index, for an audit that passed with `report`.
    pub(super) fn finish(self, report: Report) -> Option<BoundaryIndex> {
        if self.failed {
            return None;
        }
        Some(BoundaryIndex {
            depth: self.depth,
            totals: Counters {
                attributes: report.attributes,
                events: report.events,
                text_bytes: report.text_bytes,
            },
            levels: self.levels,
            parents: self.parents,
            // The document element's start tag opens the first parent.
            root: 0,
            scopes: self.scopes,
            scope_bytes: self.scope_bytes,
        })
    }

    /// Start recording the content of the element whose start tag left
    /// `depth` elements open.
    fn open_parent(&mut self, depth: usize, state: &State) -> Option<()> {
        let bindings = state.namespaces.as_ref().map_or(0, Namespaces::bindings);
        let shared = (depth > 1 && bindings == self.root_bindings)
            .then(|| self.parents.get(self.current[0] as usize).copied())
            .flatten();
        let scope = match shared {
            Some(scope) => scope,
            None => {
                let namespaces = match state.namespaces.as_ref() {
                    Some(namespaces) => namespaces.try_clone()?,
                    None => Namespaces::new(),
                };
                self.scope_bytes = self.scope_bytes.checked_add(scope_charge(&namespaces)?)?;
                if self.scope_bytes > MAX_REPLACEMENT_NAMESPACE_CLONE_BYTES {
                    return None;
                }
                self.scopes.try_reserve(1).ok()?;
                self.scopes.push(Arc::new(namespaces));
                u32::try_from(self.scopes.len() - 1).ok()?
            },
        };
        if depth == 1 {
            self.root_bindings = bindings;
        }
        self.parents.try_reserve(1).ok()?;
        self.parents.push(scope);
        self.current[depth - 1] = u32::try_from(self.parents.len() - 1).ok()?;
        Some(())
    }

    /// Record the boundary at `offset`, in the content of the parent open at
    /// `depth`.
    fn record(&mut self, depth: usize, offset: usize, state: &State) -> Option<()> {
        self.boundaries_left = self.boundaries_left.checked_sub(1)?;
        let boundary = Boundary::at(
            self.offset_base.checked_add(offset)?,
            self.current[depth - 1],
            add_counters(self.counter_base, Counters::of(state))?,
        )?;
        let level = &mut self.levels[depth - 1];
        level.try_reserve(1).ok()?;
        level.push(boundary);
        Some(())
    }
}

impl Observer for BoundaryRecorder {
    #[inline]
    fn listening(&self) -> bool {
        !self.failed
    }

    #[inline]
    fn accepted(
        &mut self,
        token: Token,
        _start: usize,
        end: usize,
        _before: Counters,
        state: &State,
    ) {
        let depth = state.depth;
        if depth == 0 || depth > self.depth {
            return;
        }
        let recorded = if token == Token::Start {
            self.open_parent(depth, state)
        } else {
            Some(())
        }
        .and_then(|()| self.record(depth, end, state));
        if recorded.is_none() {
            self.failed = true;
        }
    }
}

/// Audit `input` under the source policy and record its index down to
/// `depth`: the verdict and error of [`super::verify_source`].
pub(super) fn verify_recorded(
    input: &[u8],
    limits: Limits,
    depth: IndexDepth,
) -> Result<Option<BoundaryIndex>, super::Error> {
    let mut recorder = BoundaryRecorder::new(input.len(), depth);
    let report = verify_observed(input, limits, Policy::SOURCE, &mut recorder)?;
    Ok(recorder.finish(report))
}

/// Two observers of one scan; each is asked only while it listens.
pub(super) struct Both<'a, A, B>(pub(super) &'a mut A, pub(super) &'a mut B);

impl<A: Observer, B: Observer> Observer for Both<'_, A, B> {
    #[inline]
    fn listening(&self) -> bool {
        self.0.listening() || self.1.listening()
    }

    #[inline]
    fn accepted(
        &mut self,
        token: Token,
        start: usize,
        end: usize,
        before: Counters,
        state: &State,
    ) {
        if self.0.listening() {
            self.0.accepted(token, start, end, before, state);
        }
        if self.1.listening() {
            self.1.accepted(token, start, end, before, state);
        }
    }
}
