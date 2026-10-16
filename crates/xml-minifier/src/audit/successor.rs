//! Successor proofs: the source verdict on a replacement, proved from a
//! verified original's recorded content boundaries without scanning the
//! original again (record 0313).
//!
//! [`super::VerifiedSource::verify_indexed`] audits a payload like
//! [`super::verify_source`] and, through an [`Observer`] that only reads the
//! auditor's state, records every *content boundary* of the document element
//! and of its children: each offset between two accepted tokens at which one
//! or two elements are open. A boundary keeps the element whose content it
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
//! The index is bounded: one boundary per 32 input bytes plus a constant, at
//! 20 bytes each, and the namespace snapshots by the multi-window snapshot
//! budget. Exceeding either, or an input longer than `u32` offsets address,
//! records no index, and a successor of such a proof is audited completely.

use super::{
    Counters, Limits, MAX_REPLACEMENT_NAMESPACE_CLONE_BYTES, Namespaces, Observer,
    ReplacementProof, Report, State, Token, audit_window, common_prefix_len, common_suffix_len,
};

/// Content boundaries are recorded while at most this many elements are
/// open: the document element's content and its children's.
const INDEXED_DEPTH: usize = 2;
/// Boundaries recorded per input byte, as a divisor.
const BYTES_PER_BOUNDARY: usize = 32;
/// Boundaries recorded beyond the per-byte allowance.
const EXTRA_BOUNDARIES: usize = 64;

/// One recorded content boundary. Every field fits `u32` because the index
/// is recorded only for inputs shorter than `u32::MAX` bytes, and no counter
/// exceeds the input's length plus one.
#[derive(Clone, Copy, Debug)]
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
}

/// The content boundaries one source audit recorded, with the audit's
/// totals.
pub(super) struct BoundaryIndex {
    /// The original's aggregate totals, as its [`Report`] gives them.
    totals: Counters,
    /// `levels[d - 1]`: the boundaries at which `d` elements are open, in
    /// offset order.
    levels: [Vec<Boundary>; INDEXED_DEPTH],
    /// Each parent's index into `scopes`.
    parents: Vec<u32>,
    /// The namespace bindings in scope in a parent's content.
    scopes: Vec<Namespaces>,
}

impl BoundaryIndex {
    /// How many boundaries the index holds.
    pub(super) fn len(&self) -> usize {
        self.levels.iter().map(Vec::len).sum()
    }

    /// Prove `replacement`'s source-policy verdict from that of `original`,
    /// the bytes this index was recorded from, or return `None` so the caller
    /// audits the replacement completely.
    ///
    /// Only success is decided here, as in `Window::prove`.
    pub(super) fn prove(
        &self,
        original: &[u8],
        replacement: &[u8],
        limits: Limits,
    ) -> Option<ReplacementProof> {
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
    /// `start..end`.
    fn prove_run(
        &self,
        original_len: usize,
        replacement: &[u8],
        depth: usize,
        start: Boundary,
        end: Boundary,
        limits: Limits,
    ) -> Option<ReplacementProof> {
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
        let added = audit_window(bytes, limits, depth, scope.try_clone()?)?;
        let totals = self.totals.replaced(removed, added)?;
        totals.within(limits).then_some(ReplacementProof::Splice {
            original: run_start..run_end,
            replacement: run_start..window_end,
        })
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
        (1..=INDEXED_DEPTH).rev().find_map(|depth| {
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
}

/// Records the [`BoundaryIndex`] of one source audit.
///
/// It only reads the auditor's state, so the audit's verdict is unchanged.
/// Any bound it cannot keep stops it, and then no index is recorded.
pub(super) struct BoundaryRecorder {
    levels: [Vec<Boundary>; INDEXED_DEPTH],
    parents: Vec<u32>,
    scopes: Vec<Namespaces>,
    /// The parent whose content is open at each depth.
    current: [u32; INDEXED_DEPTH],
    /// Bindings in scope in the document element's content.
    root_bindings: usize,
    boundaries_left: usize,
    scope_bytes_left: usize,
    failed: bool,
}

impl BoundaryRecorder {
    pub(super) fn new(input_len: usize) -> Self {
        Self {
            levels: [Vec::new(), Vec::new()],
            parents: Vec::new(),
            scopes: Vec::new(),
            current: [0; INDEXED_DEPTH],
            root_bindings: 0,
            boundaries_left: (input_len / BYTES_PER_BOUNDARY).saturating_add(EXTRA_BOUNDARIES),
            scope_bytes_left: MAX_REPLACEMENT_NAMESPACE_CLONE_BYTES,
            failed: u32::try_from(input_len).map_or(true, |len| len == u32::MAX),
        }
    }

    /// The recorded index, for an audit that passed with `report`.
    pub(super) fn finish(self, report: Report) -> Option<BoundaryIndex> {
        if self.failed {
            return None;
        }
        Some(BoundaryIndex {
            totals: Counters {
                attributes: report.attributes,
                events: report.events,
                text_bytes: report.text_bytes,
            },
            levels: self.levels,
            parents: self.parents,
            scopes: self.scopes,
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
                    Some(namespaces) => {
                        let charge = namespaces.clone_memory_charge()?;
                        self.scope_bytes_left = self.scope_bytes_left.checked_sub(charge)?;
                        namespaces.try_clone()?
                    },
                    None => Namespaces::new(),
                };
                self.scopes.try_reserve(1).ok()?;
                self.scopes.push(namespaces);
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
        let boundary = Boundary {
            offset: u32::try_from(offset).ok()?,
            parent: self.current[depth - 1],
            attributes: u32::try_from(state.attributes).ok()?,
            events: u32::try_from(state.events).ok()?,
            text_bytes: u32::try_from(state.text_bytes).ok()?,
        };
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
        if depth == 0 || depth > INDEXED_DEPTH {
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
