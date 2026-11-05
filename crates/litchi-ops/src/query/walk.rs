//! Query execution (record 0010 §5.3): one iterative depth-first walk in
//! document order.
//!
//! The walk keeps an explicit stack of frames, one per container it has
//! entered: the container's node, its kind, its position among the parent's
//! children of that kind, the lazy iterator over its children, and its
//! per-kind child counters. No recursion follows the document, so nesting
//! cannot grow the call stack; what recursion there is follows the selector
//! (chain links, filter expressions, nested pseudo-classes), which the
//! grammar's limits bound.
//!
//! - **Pruning.** A child is entered only if the spec's child tables say a
//!   wanted kind can occur below it, and the adapter is asked only for the
//!   kinds that are wanted or lead to one.
//! - **Incremental paths.** The counters give every visited node its
//!   position as it is visited, so a match's path is written from the stack
//!   ([`Hit::path`]) without rescanning siblings.
//! - **Cheapest test first.** Kind, then structure (position, number, cell),
//!   then ancestors (memoized per frame), then properties (one property read
//!   per candidate), then text, then `:not` and `:has`.
//! - **Streaming.** Matches are produced one at a time; nothing is visited
//!   beyond the match the caller last pulled, `limit` stops the walk, and an
//!   adapter error ends it (the walk is fused).
//!
//! The stack, the counters and the ancestor memo are reused for the whole
//! walk, so visiting a node allocates nothing once they have grown to the
//! document's depth.

use std::borrow::Cow;
use std::fmt::{self, Write as _};
use std::iter::FusedIterator;
use std::ops::Deref;
use std::sync::Arc;

use litchi_core::selector::{Combinator, Path, Position, Step};

use super::inspect::{HeaderBinding, HeaderKey, Inspect};
use super::plan::{
    Chain, ExprNode, FastPlan, Plan, Positional, ScopePlan, Shortcut, Structural, Target,
};
use crate::error::{Code, Error, Result};
use crate::node::{Format, Value};
use crate::path::PathArg;
use crate::path::fence::Scope;
use crate::path::navigate::{ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::resolve;
use crate::value::Scalar;
use crate::value::text::contains_ignore_case;

/// The plan a walk runs: compiled for it, or borrowed from the caller.
#[derive(Debug)]
pub(crate) enum PlanRef<'q> {
    /// Compiled for this walk from a selector; boxed, so a walk stays small
    /// whichever way it got its plan.
    Owned(Box<Plan<'q>>),
    /// Prepared once by the caller and reused.
    Borrowed(&'q Plan<'q>),
    /// Shared with this thread's plan cache (record 0291).
    Shared(Arc<Plan<'static>>),
}

impl<'q> From<super::cache::Compiled<'q>> for PlanRef<'q> {
    fn from(plan: super::cache::Compiled<'q>) -> Self {
        match plan {
            super::cache::Compiled::Shared(plan) => Self::Shared(plan),
            super::cache::Compiled::Fresh(plan) => Self::Owned(plan),
        }
    }
}

impl<'q> Deref for PlanRef<'q> {
    type Target = Plan<'q>;

    fn deref(&self) -> &Plan<'q> {
        match self {
            Self::Owned(plan) => plan,
            Self::Borrowed(plan) => plan,
            Self::Shared(plan) => plan,
        }
    }
}

/// A running query: an iterator over the matching nodes, in document order.
///
/// Each item is fallible because a lazily read part can fail mid-walk; after
/// the first error, and after `limit` matches, the iterator is fused.
pub(crate) struct Walk<'q, 'f, N>
where
    N: Inspect<'f>,
{
    nav: N,
    plan: PlanRef<'q>,
    state: State<'f, N>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Phase {
    Start,
    Walk,
    Done,
}

/// The per-alternative state of subject positional predicates.
#[derive(Debug, Clone, Copy, Default)]
struct Alt {
    /// Matches of the alternative so far.
    count: usize,
    /// The alternative's `[N]` was found: it selects nothing more.
    done: bool,
}

struct State<'f, N>
where
    N: Inspect<'f>,
{
    phase: Phase,
    frames: Vec<Frame<'f, N>>,
    /// `(kind ordinal, children so far)` per frame, each frame's entries
    /// starting at its `counters` index.
    counters: Vec<(u8, u32)>,
    /// Two halves of `words` bits per frame: which ancestor-memo slots are
    /// known, and their values.
    memo: Vec<u64>,
    words: usize,
    alts: Vec<Alt>,
    /// The subject kinds still wanted.
    interest: u128,
    /// `(interest, kinds worth asking the adapter for)`, the two most recent.
    relevant: [(u128, u128); 2],
    /// The latest match, from the step that found it until the iterator or
    /// the hit takes it (record 0287: the node moves once into this slot and
    /// once out, instead of through every layer of the step's result).
    out: Option<N::Node>,
    /// The latest match of a `[last()]` subject, with its ancestors.
    pending: Option<(N::Node, N::Kind, u32)>,
    pending_path: Vec<(N::Node, N::Kind, u32)>,
    /// One header-row binding per header plan.
    bindings: Vec<HeaderBinding>,
    /// The last sibling count: `(frame serial, kind ordinal, count)`.
    siblings: Option<(u64, u8, u32)>,
    serial: u64,
    limit: usize,
    yielded: usize,
    visited: u64,
    /// The walk is likely to stop early: it has a `limit`, or every
    /// alternative selects its `[N]`th match. Its containers are then read
    /// through [`Inspect::early_children`] (record 0267).
    early: bool,
    /// The kinds of the query's one alternative when it tests the kind
    /// only and selects no position (`cell`, `p`, `shape`): a visited node
    /// then matches exactly when it is of one of them, and the visit skips
    /// the per-alternative bookkeeping (record 0291).
    kind_only: Option<u128>,
    /// Tests only: walk on the general step even when the plan has a
    /// [`FastPlan`] (record 0315).
    #[cfg(test)]
    general: bool,
}

struct Frame<'f, N>
where
    N: Inspect<'f>,
{
    node: N::Node,
    kind: N::Kind,
    /// The position among the parent's children of the same kind.
    ordinal: u32,
    /// The kinds to ask the adapter for.
    kinds: u128,
    source: Source<'f, N>,
    /// Where this frame's child counters start.
    counters: usize,
    /// A number unique to this push, for caches keyed by frame.
    serial: u64,
    /// The frame of the node's parent. It is the frame below, except for
    /// the anchor of a `:has` search started from an ancestor frame, which
    /// sits on top of the stack while its parent is further down.
    parent: Option<usize>,
}

enum Source<'f, N>
where
    N: Inspect<'f>,
{
    /// The adapter's iterator, created when the first child is needed.
    Lazy,
    Open(N::Children),
    /// Worksheet rows by number, records or not (header-row queries).
    Rows {
        next: u32,
        end: u32,
        kind: N::Kind,
    },
    Done,
}

/// A match found by the walk; the node itself waits in the state's `out`
/// slot.
struct Found<'f, N: Navigate<'f>> {
    kind: N::Kind,
    ordinal: u32,
    ancestors: Ancestors,
}

/// A match with its node, as a [`Hit`] holds it.
struct Matched<'f, N: Navigate<'f>> {
    node: N::Node,
    found: Found<'f, N>,
}

/// Where a match's ancestors are, for its path.
#[derive(Debug, Clone, Copy)]
enum Ancestors {
    /// `frames[1..n]`.
    Frames(usize),
    /// The snapshot taken for a `[last()]` match.
    Pending,
}

/// One node a compound is tested against.
struct Candidate<'c, 'f, N: Navigate<'f>> {
    node: &'c N::Node,
    kind: N::Kind,
    ordinal: u32,
    /// The parent's frame, if the node has one on the stack.
    parent: Option<usize>,
}

/// The result of testing a visited node against the top-level alternatives.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Verdict {
    No,
    Yield,
    Pending,
}

/// What one candidate's tests have read, so each part is read once.
struct Probe<'c> {
    text: Option<Option<Cow<'c, str>>>,
    style: Option<Option<Cow<'c, str>>>,
    format: Option<Format<'c>>,
}

impl<'c> Probe<'c> {
    const fn new() -> Self {
        Self {
            text: None,
            style: None,
            format: None,
        }
    }

    fn text<'f, N: Inspect<'f>>(&mut self, nav: &'c N, node: &'c N::Node) -> Result<Option<&str>> {
        if self.text.is_none() {
            self.text = Some(nav.text(node)?);
        }
        Ok(self.text.as_ref().and_then(|text| text.as_deref()))
    }

    fn style<'f, N: Inspect<'f>>(&mut self, nav: &'c N, node: &'c N::Node) -> Result<Option<&str>> {
        if self.style.is_none() {
            self.style = Some(nav.style(node)?);
        }
        Ok(self.style.as_ref().and_then(|style| style.as_deref()))
    }

    fn format<'f, N: Inspect<'f>>(&mut self, nav: &'c N, node: &'c N::Node) -> Result<&Format<'c>> {
        if self.format.is_none() {
            self.format = Some(nav.format(node)?);
        }
        self.format
            .as_ref()
            .ok_or_else(|| Error::new(Code::InternalError, "a property read was lost"))
    }

    fn property<'f, N: Inspect<'f>>(
        &mut self,
        nav: &'c N,
        node: &'c N::Node,
        index: usize,
    ) -> Result<Option<Value<'c>>> {
        if N::INDIVIDUAL_PROPERTIES {
            nav.prop(node, index)
        } else {
            Ok(self.format(nav, node)?.get(index))
        }
    }
}

/// The bit of the kind with `ordinal`; none for an ordinal past 127, which
/// a well-formed adapter never reports, so such a node matches nothing.
const fn bit_of(ordinal: u8) -> u128 {
    match 1_u128.checked_shl(ordinal as u32) {
        Some(bit) => bit,
        None => 0,
    }
}

fn bit<K: ElementKind>(kind: K) -> u128 {
    bit_of(kind.ordinal())
}

/// Whether a property value is absent text.
fn is_empty(value: &Value<'_>) -> bool {
    value.as_str().is_some_and(str::is_empty)
}

impl<'q, 'f, N> Walk<'q, 'f, N>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
    /// A walk of `plan` over the document `nav` reads. Nothing is read until
    /// the first match is pulled.
    pub(crate) fn new(nav: N, plan: PlanRef<'q>) -> Result<Self> {
        if plan.format != N::FORMAT {
            return Err(Error::new(
                Code::InvalidArgument,
                format!(
                    "the query plan was compiled for {:?} documents, not {:?}",
                    plan.format,
                    N::FORMAT
                ),
            )
            .with_format(N::FORMAT)
            .with_suggestion("compile the selector again for this document's format"));
        }
        let words = (plan.memo_slots as usize).div_ceil(64);
        let alternatives = plan.top.len();
        let interest = plan.wanted;
        let early = !plan.top.is_empty()
            && plan.top.iter().all(|&chain| {
                matches!(
                    plan.chains[chain as usize].position,
                    Some(Positional::Nth(_))
                )
            });
        let kind_only = match plan.top.as_slice() {
            [chain] => {
                let chain = plan.chains[*chain as usize];
                (chain.shortcut == Shortcut::KindOnly && chain.position.is_none())
                    .then_some(chain.kinds)
            },
            _ => None,
        };
        Ok(Self {
            nav,
            plan,
            state: State {
                phase: Phase::Start,
                frames: Vec::with_capacity(16),
                counters: Vec::with_capacity(16),
                // Sixteen frames' worth; deeper documents grow it once more.
                memo: Vec::with_capacity(16 * 2 * words),
                words,
                alts: vec![Alt::default(); alternatives],
                interest,
                relevant: [(0, 0); 2],
                out: None,
                pending: None,
                pending_path: Vec::new(),
                bindings: Vec::new(),
                siblings: None,
                serial: 0,
                limit: usize::MAX,
                yielded: 0,
                visited: 0,
                early,
                kind_only,
                #[cfg(test)]
                general: false,
            },
        })
    }

    /// The same walk, stopping after `limit` matches.
    #[must_use]
    pub(crate) fn limit(mut self, limit: usize) -> Self {
        self.state.limit = limit;
        self.state.early |= limit != usize::MAX;
        self
    }

    /// The first match, or `None`; the walk stops there.
    pub(crate) fn first(self) -> Result<Option<N::Node>> {
        self.limit(1).next().transpose()
    }

    /// The plan the walk runs.
    pub(crate) fn plan(&self) -> &Plan<'q> {
        &self.plan
    }

    /// The same walk on the general step, without its [`FastPlan`] path: the
    /// reference the fast path is tested against (record 0315).
    #[cfg(test)]
    pub(crate) fn general(mut self) -> Self {
        self.state.general = true;
        self
    }

    /// Whether the walk takes the [`FastPlan`] path.
    pub(crate) fn is_fast(&self) -> bool {
        #[cfg(test)]
        if self.state.general {
            return false;
        }
        self.plan.fast.is_some()
    }

    /// The nodes the adapter has produced for this walk so far: children
    /// visited, `:has` descendants searched, siblings counted, rows probed.
    pub(crate) fn visited(&self) -> u64 {
        self.state.visited
    }

    /// The next match, with access to its path.
    pub(crate) fn next_hit(&mut self) -> Option<Result<Hit<'_, 'q, 'f, N>>> {
        match self.advance() {
            Ok(Some(found)) => {
                let node = self.state.out.take()?;
                Some(Ok(Hit {
                    walk: self,
                    found: Matched { node, found },
                }))
            },
            Ok(None) => None,
            Err(error) => Some(Err(error)),
        }
    }

    fn engine(&mut self) -> Engine<'_, 'q, 'f, N> {
        Engine {
            plan: &self.plan,
            nav: &self.nav,
            state: &mut self.state,
        }
    }

    #[inline]
    fn advance(&mut self) -> Result<Option<Found<'f, N>>> {
        if matches!(self.state.phase, Phase::Done) || self.state.yielded >= self.state.limit {
            self.finish();
            return Ok(None);
        }
        // One check after the step fences every read of it: the adapter's
        // nested reads skip their own checks while the scope is active, and a
        // source that changed before or during the step fails this check, so
        // its error wins over whatever the step read, as it would if checked
        // before as well (record 0262).
        let result = {
            let _scope = Scope::enter(self.nav.fence_key());
            self.advance_checked()
        };
        let result = match self.nav.check_source() {
            Ok(()) => result,
            Err(error) => Err(error),
        };
        if result.is_err() {
            self.finish();
            self.state.out = None;
            self.state.pending_path.clear();
        }
        result
    }

    #[inline]
    fn advance_checked(&mut self) -> Result<Option<Found<'f, N>>> {
        match self.state.phase {
            Phase::Done => return Ok(None),
            Phase::Start => {
                // An error while starting ends the walk.
                self.state.phase = Phase::Done;
                self.engine().start()?;
                self.state.phase = Phase::Walk;
            },
            Phase::Walk => {},
        }
        if self.state.yielded >= self.state.limit {
            self.finish();
            return Ok(None);
        }
        match self.engine().next_found() {
            Ok(Some(found)) => {
                self.state.yielded += 1;
                Ok(Some(found))
            },
            Ok(None) => {
                let pending = self.state.pending.take();
                self.finish();
                Ok(pending.map(|(node, kind, ordinal)| {
                    self.state.yielded += 1;
                    self.state.out = Some(node);
                    Found {
                        kind,
                        ordinal,
                        ancestors: Ancestors::Pending,
                    }
                }))
            },
            Err(error) => {
                self.finish();
                self.state.pending_path.clear();
                Err(error)
            },
        }
    }

    /// Ends the walk, dropping the stack.
    fn finish(&mut self) {
        self.state.phase = Phase::Done;
        self.state.frames.clear();
        self.state.counters.clear();
        self.state.pending = None;
    }
}

impl<'f, N> Iterator for Walk<'_, 'f, N>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
    type Item = Result<N::Node>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        // The match moves out of the walk's slot once, without a `Hit`.
        self.next_map(|node| node)
    }
}

impl<'f, N> Walk<'_, 'f, N>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
    /// The next match, wrapped by `wrap` as it leaves the walk's slot, so
    /// the node moves once into the caller's item (record 0287).
    #[inline]
    pub(crate) fn next_map<T>(&mut self, wrap: impl FnOnce(N::Node) -> T) -> Option<Result<T>> {
        match self.advance() {
            Ok(Some(_)) => self.state.out.take().map(|node| Ok(wrap(node))),
            Ok(None) => None,
            Err(error) => Some(Err(error)),
        }
    }
}

impl<'f, N> Walk<'_, 'f, N>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
    /// [`Self::next_map`] on the [`FastPlan`] path: [`Self::advance`] with the
    /// match handed straight to `wrap` instead of through the walk's slot
    /// (record 0315). The bracket, the limit, the fusing and the error
    /// handling are [`Self::advance`]'s.
    ///
    /// The caller takes this route only for a walk whose
    /// [`Self::is_fast`] holds, decided once when the walk starts
    /// ([`super::Matches`]), so the general route's code is not touched.
    #[inline(never)]
    pub(crate) fn next_fast_map<T>(
        &mut self,
        wrap: impl FnOnce(N::Node) -> T,
    ) -> Option<Result<T>> {
        // Only a walk with a fast plan comes here; calling the general route
        // instead would give it a second call site and change how it is
        // compiled.
        let fast = self.plan.fast?;
        if matches!(self.state.phase, Phase::Done) || self.state.yielded >= self.state.limit {
            self.finish();
            return None;
        }
        let result = {
            let _scope = Scope::enter(self.nav.fence_key());
            self.fast_checked(fast)
        };
        let result = match self.nav.check_source() {
            Ok(()) => result,
            Err(error) => Err(error),
        };
        match result {
            Ok(Some(node)) => Some(Ok(wrap(node))),
            Ok(None) => None,
            Err(error) => {
                self.finish();
                self.state.out = None;
                self.state.pending_path.clear();
                Some(Err(error))
            },
        }
    }

    /// [`Self::advance_checked`] on the [`FastPlan`] path, which has no pending
    /// `[last()]` match.
    #[inline(always)]
    fn fast_checked(&mut self, fast: FastPlan) -> Result<Option<N::Node>> {
        match self.state.phase {
            Phase::Done => return Ok(None),
            Phase::Start => {
                self.state.phase = Phase::Done;
                self.engine().start()?;
                self.state.phase = Phase::Walk;
            },
            Phase::Walk => {},
        }
        if self.state.yielded >= self.state.limit {
            self.finish();
            return Ok(None);
        }
        match self.engine().next_fast(fast) {
            Ok(Some(matched)) => {
                self.state.yielded += 1;
                Ok(Some(matched.node))
            },
            Ok(None) => {
                self.finish();
                Ok(None)
            },
            Err(error) => {
                self.finish();
                self.state.pending_path.clear();
                Err(error)
            },
        }
    }
}

impl<'f, N> FusedIterator for Walk<'_, 'f, N>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
}

/// One match, borrowed from its walk until the next is pulled.
pub(crate) struct Hit<'w, 'q, 'f, N>
where
    N: Inspect<'f>,
{
    walk: &'w Walk<'q, 'f, N>,
    found: Matched<'f, N>,
}

impl<'w, 'q, 'f, N> Hit<'w, 'q, 'f, N>
where
    N: Inspect<'f>,
{
    /// The matching node.
    pub(crate) fn node(&self) -> &N::Node {
        &self.found.node
    }

    /// The matching node's kind.
    pub(crate) fn kind(&self) -> N::Kind {
        self.found.found.kind
    }

    /// The match's path, written from the walk's stack: the scope as the
    /// query wrote it, then one step per ancestor and the match itself, each
    /// written by the adapter ([`Inspect::write_step`]).
    pub(crate) fn path(&self) -> HitPath<'_, 'w, 'q, 'f, N> {
        HitPath { hit: self }
    }
}

/// The path of a [`Hit`], written on demand without allocating.
pub(crate) struct HitPath<'h, 'w, 'q, 'f, N>
where
    N: Inspect<'f>,
{
    hit: &'h Hit<'w, 'q, 'f, N>,
}

impl<'f, N> fmt::Display for HitPath<'_, '_, '_, 'f, N>
where
    N: Inspect<'f>,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let walk = self.hit.walk;
        let nav = &walk.nav;
        match &walk.plan.scope {
            ScopePlan::Root => {},
            ScopePlan::Path(path) => {
                if !path.steps().is_empty() {
                    fmt::Display::fmt(path, f)?;
                }
            },
            ScopePlan::Sheet(name) => {
                f.write_char('/')?;
                fmt::Display::fmt(&Step::new(name.as_ref()), f)?;
            },
        }
        let mut step = |node: &N::Node, kind: N::Kind, ordinal: u32| {
            f.write_char('/')?;
            nav.write_step(node, kind, Position::new(ordinal as usize), f)
        };
        match self.hit.found.found.ancestors {
            Ancestors::Frames(depth) => {
                for frame in walk.state.frames.iter().take(depth).skip(1) {
                    step(&frame.node, frame.kind, frame.ordinal)?;
                }
            },
            Ancestors::Pending => {
                for (node, kind, ordinal) in &walk.state.pending_path {
                    step(node, *kind, *ordinal)?;
                }
            },
        }
        let found = &self.hit.found;
        step(&found.node, found.found.kind, found.found.ordinal)
    }
}

/// The plan, the adapter and the walk state, borrowed separately so the
/// evaluation can read the plan while it updates the stack.
struct Engine<'e, 'q, 'f, N>
where
    N: Inspect<'f>,
{
    plan: &'e Plan<'q>,
    nav: &'e N,
    state: &'e mut State<'f, N>,
}

impl<'e, 'q, 'f, N> Engine<'e, 'q, 'f, N>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
    fn kind_at(&self, ordinal: u8) -> Result<N::Kind> {
        N::SPEC
            .elements
            .get(usize::from(ordinal))
            .map(|element| element.kind)
            .ok_or_else(|| Error::new(Code::InternalError, "a query plan names an unknown kind"))
    }

    /// Resolves the scope, binds header-row predicates, and pushes the root
    /// frame.
    fn start(&mut self) -> Result<()> {
        let (plan, nav) = (self.plan, self.nav);
        let (root, kind) = match &plan.scope {
            ScopePlan::Root => (nav.root()?, N::Kind::ROOT),
            ScopePlan::Path(path) => {
                let mut arg = PathArg::from(&**path);
                arg.raw = plan.raw;
                let node = resolve::get(nav, &arg)?;
                let kind = nav.kind(&node);
                (node, kind)
            },
            ScopePlan::Sheet(name) => {
                let path = Path::from_steps([Step::new(name.as_ref())]);
                let node = resolve::get(nav, &PathArg::from(path))?;
                let kind = nav.kind(&node);
                (node, kind)
            },
        };
        for key in &plan.collisions {
            if nav.names_column(&root, key)? {
                return Err(Error::new(
                    Code::InvalidSelector,
                    "a query key names both a row property and a table column",
                )
                .with_format(N::FORMAT)
                .with_suggestion("use col. for a table column or @ for a row property"));
            }
        }
        let mut bindings = Vec::with_capacity(plan.headers.len());
        for header in &plan.headers {
            let keys: Vec<HeaderKey<'_>> = header
                .keys
                .iter()
                .map(|(name, forced)| HeaderKey {
                    name,
                    forced: *forced,
                })
                .collect();
            let binding = nav.bind_header(&root, &keys)?;
            let (first, last) = binding.rows;
            if binding.columns.len() != keys.len() || first == 0 || first > last.saturating_add(1) {
                return Err(Error::new(
                    Code::InternalError,
                    "the format adapter returned an inconsistent header-row binding",
                ));
            }
            bindings.push(binding);
        }
        self.state.bindings = bindings;
        let kinds = self.relevant(plan.wanted);
        let Some(header) = plan.rows_mode else {
            self.push(root, kind, 0, kinds, Source::Lazy)?;
            return Ok(());
        };
        let header_plan = &plan.headers[header as usize];
        let (sheet_kind, row_kind) = (
            self.kind_at(header_plan.sheet_kind)?,
            self.kind_at(header_plan.row_kind)?,
        );
        let binding = &self.state.bindings[header as usize];
        let (sheet_position, (first, last)) = (binding.sheet, binding.rows);
        let rows = Source::Rows {
            next: first,
            end: last,
            kind: row_kind,
        };
        if kind == sheet_kind {
            self.push(root, kind, 0, kinds, rows)?;
            return Ok(());
        }
        let lookup = nav.child(
            &root,
            &Resolved {
                kind: sheet_kind,
                select: Select::Ordinal(sheet_position),
            },
        )?;
        let Lookup::Found(sheet) = lookup else {
            return Err(Error::new(
                Code::InternalError,
                "the format adapter bound a header row to a worksheet it cannot find",
            ));
        };
        self.push(root, kind, 0, kinds, Source::Done)?;
        let ordinal = u32::try_from(sheet_position.get()).unwrap_or(u32::MAX);
        self.push(sheet, sheet_kind, ordinal, kinds, rows)?;
        Ok(())
    }

    /// Walks to the next match of the top-level alternatives.
    fn next_found(&mut self) -> Result<Option<Found<'f, N>>> {
        loop {
            if self.state.interest == 0 || self.state.frames.is_empty() {
                return Ok(None);
            }
            let Some(item) = self.next_child() else {
                self.pop();
                continue;
            };
            let node = item?;
            self.charge_visit()?;
            let kind = self.nav.kind(&node);
            let ordinal = self.bump(kind);
            let depth = self.state.frames.len();
            let verdict = self.visit(&node, kind, ordinal, depth)?;
            let descend = self.descend(kind) && {
                let kinds = self.relevant(self.state.interest);
                !self.nav.childless(&node, KindSet::from_bits(kinds))
            };
            match verdict {
                Verdict::No => {
                    if descend {
                        let kinds = self.relevant(self.state.interest);
                        self.push(node, kind, ordinal, kinds, Source::Lazy)?;
                    }
                },
                Verdict::Pending => {
                    self.set_pending(&node, kind, ordinal, depth);
                    if descend {
                        let kinds = self.relevant(self.state.interest);
                        self.push(node, kind, ordinal, kinds, Source::Lazy)?;
                    }
                },
                Verdict::Yield => {
                    self.state.out = Some(if descend {
                        let kinds = self.relevant(self.state.interest);
                        let yielded = node.clone();
                        self.push(node, kind, ordinal, kinds, Source::Lazy)?;
                        yielded
                    } else {
                        node
                    });
                    return Ok(Some(Found {
                        kind,
                        ordinal,
                        ancestors: Ancestors::Frames(depth),
                    }));
                },
            }
        }
    }

    /// [`Self::next_found`] for a plan with a [`FastPlan`], on the route
    /// [`Walk::next_fast_map`] takes (`Matches`; a [`Hit`] walk keeps the
    /// general step): the same visits, reads and matches without the
    /// per-alternative bookkeeping, whose outcome the plan fixes (record
    /// 0315). The match is returned rather than left in the walk's slot.
    #[inline(always)]
    fn next_fast(&mut self, fast: FastPlan) -> Result<Option<Matched<'f, N>>> {
        loop {
            if self.state.frames.is_empty() {
                return Ok(None);
            }
            let Some(item) = self.next_child_inline() else {
                self.pop();
                continue;
            };
            let node = item?;
            self.charge_visit()?;
            let kind = self.nav.kind(&node);
            let ordinal = self.bump(kind);
            let depth = self.state.frames.len();
            let kind_bit = bit(kind);
            let matched = fast.kinds & kind_bit != 0 && {
                let candidate = Candidate {
                    node: &node,
                    kind,
                    ordinal,
                    parent: depth.checked_sub(1),
                };
                let mut probe = Probe::new();
                match fast.test {
                    Shortcut::KindOnly => true,
                    Shortcut::Expr(expr) => self.expr(expr, &candidate, &mut probe)?,
                    Shortcut::Content(compound) => {
                        self.content_matches(compound, &candidate, &mut probe)?
                    },
                    Shortcut::None => {
                        return Err(Error::new(
                            Code::InternalError,
                            "a query fast path lost its shortcut",
                        ));
                    },
                }
            };
            let descend = fast.enter & kind_bit != 0
                && !self.nav.childless(&node, KindSet::from_bits(fast.ask));
            if matched {
                let node = if descend {
                    let yielded = node.clone();
                    self.push(node, kind, ordinal, fast.ask, Source::Lazy)?;
                    yielded
                } else {
                    node
                };
                return Ok(Some(Matched {
                    node,
                    found: Found {
                        kind,
                        ordinal,
                        ancestors: Ancestors::Frames(depth),
                    },
                }));
            }
            if descend {
                self.push(node, kind, ordinal, fast.ask, Source::Lazy)?;
            }
        }
    }

    /// Whether the main walk enters a visited node of `kind`: a wanted kind
    /// can occur below it, and, for a lone `*`, it may lead to a top-level
    /// container (`*` never matches deeper).
    fn descend(&self, kind: N::Kind) -> bool {
        let plan = self.plan;
        let reach = plan
            .reach
            .get(usize::from(kind.ordinal()))
            .copied()
            .unwrap_or(0);
        if reach & self.state.interest == 0 {
            return false;
        }
        if !plan.shallow {
            return true;
        }
        match plan.scope {
            // `*` lists the scope's direct children only.
            ScopePlan::Path(_) | ScopePlan::Sheet(_) => false,
            ScopePlan::Root => (bit(kind) | reach) & plan.top_parents != 0,
        }
    }

    /// The next child of the top frame, or `None` when it has no more.
    fn next_child(&mut self) -> Option<Result<N::Node>> {
        self.next_child_inline()
    }

    /// [`Self::next_child`], inlined into the fast step (record 0315): the
    /// general step keeps its out-of-line call.
    #[inline(always)]
    fn next_child_inline(&mut self) -> Option<Result<N::Node>> {
        let nav = self.nav;
        let early = self.state.early;
        loop {
            let frame = self.state.frames.last_mut()?;
            match &mut frame.source {
                Source::Lazy => match if early {
                    nav.early_children(&frame.node, KindSet::from_bits(frame.kinds))
                } else {
                    nav.children(&frame.node, KindSet::from_bits(frame.kinds))
                } {
                    Ok(children) => frame.source = Source::Open(children),
                    Err(error) => {
                        frame.source = Source::Done;
                        return Some(Err(error));
                    },
                },
                Source::Open(children) => return children.next(),
                Source::Rows { next, end, kind } => {
                    if *next > *end {
                        return None;
                    }
                    let number = *next as usize;
                    let wanted = Resolved {
                        kind: *kind,
                        select: Select::Number(number),
                    };
                    if *next == *end {
                        frame.source = Source::Done;
                    } else {
                        *next += 1;
                    }
                    if self.state.visited >= 1_000_000 {
                        return Some(Err(Error::new(
                            Code::LimitExceeded,
                            "query traversal exceeds the visit limit",
                        )));
                    }
                    self.state.visited += 1;
                    match nav.child(&frame.node, &wanted) {
                        Ok(Lookup::Found(row)) => return Some(Ok(row)),
                        Ok(Lookup::Missing) => {},
                        Ok(_) => {
                            return Some(Err(Error::new(
                                Code::InternalError,
                                "the format adapter cannot read a worksheet row by number",
                            )));
                        },
                        Err(error) => return Some(Err(error)),
                    }
                },
                Source::Done => return None,
            }
        }
    }

    /// Tests a visited node against the top-level alternatives.
    fn visit(
        &mut self,
        node: &N::Node,
        kind: N::Kind,
        ordinal: u32,
        depth: usize,
    ) -> Result<Verdict> {
        let plan = self.plan;
        let kind_bit = bit(kind);
        if self.state.interest & kind_bit == 0 {
            return Ok(Verdict::No);
        }
        // The loop below, for one kind-only alternative without a position:
        // its shortcut matches every node of its kinds, which yields.
        if let Some(kinds) = self.state.kind_only {
            return Ok(if kinds & kind_bit == 0 {
                Verdict::No
            } else {
                Verdict::Yield
            });
        }
        let candidate = Candidate {
            node,
            kind,
            ordinal,
            parent: depth.checked_sub(1),
        };
        let mut probe = Probe::new();
        let mut verdict = Verdict::No;
        for (index, &chain_id) in plan.top.iter().enumerate() {
            if self.state.alts[index].done {
                continue;
            }
            let chain = plan.chains[chain_id as usize];
            if chain.kinds & kind_bit == 0
                || (verdict == Verdict::Yield && chain.position.is_none())
            {
                continue;
            }
            if !self.shortcut_matches(chain, chain_id, &candidate, &mut probe, 0)? {
                continue;
            }
            match chain.position {
                None => verdict = Verdict::Yield,
                Some(Positional::Nth(number)) => {
                    let alt = &mut self.state.alts[index];
                    alt.count += 1;
                    if alt.count == number {
                        alt.done = true;
                        verdict = Verdict::Yield;
                        self.refresh_interest();
                    }
                },
                Some(Positional::Last) => {
                    if verdict == Verdict::No {
                        verdict = Verdict::Pending;
                    }
                },
            }
        }
        Ok(verdict)
    }

    fn refresh_interest(&mut self) {
        let plan = self.plan;
        self.state.interest = plan
            .top
            .iter()
            .zip(&self.state.alts)
            .filter(|(_, alt)| !alt.done)
            .fold(0, |bits, (&chain, _)| {
                bits | plan.chains[chain as usize].kinds
            });
    }

    /// Records the latest match of a `[last()]` subject and its ancestors.
    fn set_pending(&mut self, node: &N::Node, kind: N::Kind, ordinal: u32, depth: usize) {
        let state = &mut *self.state;
        state.pending = Some((node.clone(), kind, ordinal));
        state.pending_path.clear();
        for frame in state.frames.iter().take(depth).skip(1) {
            state
                .pending_path
                .push((frame.node.clone(), frame.kind, frame.ordinal));
        }
    }

    /// Whether a candidate of one of `chain`'s kinds matches it: through its
    /// shortcut when it has one (a kind-only compound matches; a compound
    /// with only a filter is its filter; a compound that tests nothing
    /// about the node's place is its content tests), else in full.
    #[inline]
    fn shortcut_matches<'c>(
        &mut self,
        chain: Chain,
        chain_id: u32,
        candidate: &Candidate<'c, 'f, N>,
        probe: &mut Probe<'c>,
        floor: usize,
    ) -> Result<bool>
    where
        'e: 'c,
    {
        match chain.shortcut {
            Shortcut::KindOnly => Ok(true),
            Shortcut::Expr(expr) => self.expr(expr, candidate, probe),
            Shortcut::Content(compound) => self.content_matches(compound, candidate, probe),
            Shortcut::None => self.chain_matches(chain_id, candidate, probe, floor),
        }
    }

    /// Whether a candidate matches a chain: its subject compound, then its
    /// ancestor links against the frames from `candidate.parent` down to
    /// `floor`.
    fn chain_matches<'c>(
        &mut self,
        chain_id: u32,
        candidate: &Candidate<'c, 'f, N>,
        probe: &mut Probe<'c>,
        floor: usize,
    ) -> Result<bool>
    where
        'e: 'c,
    {
        let plan = self.plan;
        let chain = plan.chains[chain_id as usize];
        let subject = plan.links[chain.first as usize].compound;
        // Cheapest first: the subject's structure, then its ancestors (read
        // from the stack and memoized per frame), then its content.
        Ok(self.structure_matches(subject, candidate)?
            && (chain.len == 1 || self.ancestors(chain, 1, candidate.parent, floor)?)
            && self.content_matches(subject, candidate, probe)?)
    }

    /// Whether link `k` of `chain` and every link left of it match `from`
    /// and its ancestors, following the frames' parent links and stopping
    /// below `floor` (a `:has` search does not look above its anchor). The
    /// result for each frame is memoized in the frame: it depends only on
    /// the frame's ancestors, which do not change while it is on the stack.
    /// Ancestors are scanned in a loop; only the chain's links recurse, as
    /// deep as the selector, never as deep as the document.
    fn ancestors(
        &mut self,
        chain: Chain,
        k: u32,
        from: Option<usize>,
        floor: usize,
    ) -> Result<bool> {
        if k == chain.len {
            return Ok(true);
        }
        let link = self.plan.links[(chain.first + k) as usize];
        let Some(start) = from.filter(|&at| at >= floor) else {
            return Ok(false);
        };
        if link.combinator == Combinator::Child {
            if let Some(known) = self.memo_get(start, link.memo) {
                return Ok(known);
            }
            let parent = self.state.frames[start].parent;
            let result = self.frame_matches(link.compound, start)?
                && self.ancestors(chain, k + 1, parent, floor)?;
            self.memo_set(start, link.memo, result);
            return Ok(result);
        }
        // Descendant: the nearest ancestor that matches with the rest of the
        // chain above it. Every frame scanned before the answer has the same
        // answer, so it is memoized along the way.
        let mut cursor = Some(start);
        let (result, end) = loop {
            let Some(at) = cursor.filter(|&at| at >= floor) else {
                break (false, None);
            };
            if let Some(known) = self.memo_get(at, link.memo) {
                break (known, Some(at));
            }
            let parent = self.state.frames[at].parent;
            if self.frame_matches(link.compound, at)?
                && self.ancestors(chain, k + 1, parent, floor)?
            {
                break (true, parent);
            }
            cursor = parent;
        };
        let mut cursor = Some(start);
        while let Some(at) = cursor.filter(|&at| at >= floor && Some(at) != end) {
            self.memo_set(at, link.memo, result);
            cursor = self.state.frames[at].parent;
        }
        Ok(result)
    }

    /// Tests a compound against the node of frame `index`.
    fn frame_matches(&mut self, compound: u32, index: usize) -> Result<bool> {
        let (node, kind, ordinal, parent) = {
            let frame = &self.state.frames[index];
            (frame.node.clone(), frame.kind, frame.ordinal, frame.parent)
        };
        let candidate = Candidate {
            node: &node,
            kind,
            ordinal,
            parent,
        };
        let mut probe = Probe::new();
        self.compound_matches(compound, &candidate, &mut probe)
    }

    /// Tests one compound against a candidate, cheapest test first.
    fn compound_matches<'c>(
        &mut self,
        id: u32,
        candidate: &Candidate<'c, 'f, N>,
        probe: &mut Probe<'c>,
    ) -> Result<bool>
    where
        'e: 'c,
    {
        Ok(self.structure_matches(id, candidate)? && self.content_matches(id, candidate, probe)?)
    }

    /// The tests of a compound that read no node content: kind, `*`,
    /// positions, numbers, cells and the header-row gate.
    fn structure_matches(&mut self, id: u32, candidate: &Candidate<'_, 'f, N>) -> Result<bool> {
        let plan = self.plan;
        let compound = &plan.compounds[id as usize];
        if compound.kinds & bit(candidate.kind) == 0 {
            return Ok(false);
        }
        if compound.any && !self.top_level(candidate) {
            return Ok(false);
        }
        for &test in &compound.tests {
            if !self.structural(test, candidate)? {
                return Ok(false);
            }
        }
        Ok(compound
            .header
            .is_none_or(|header| self.header_gate(header, candidate)))
    }

    /// The tests of a compound that read the node: properties (one read per
    /// candidate), text, alternative text, then `:not` and `:has`.
    fn content_matches<'c>(
        &mut self,
        id: u32,
        candidate: &Candidate<'c, 'f, N>,
        probe: &mut Probe<'c>,
    ) -> Result<bool>
    where
        'e: 'c,
    {
        let plan = self.plan;
        let nav: &'c N = self.nav;
        let compound = &plan.compounds[id as usize];
        if let Some(expr) = compound.expr {
            if !self.expr(expr, candidate, probe)? {
                return Ok(false);
            }
        }
        if !compound.contains.is_empty() || compound.empty {
            let text = probe.text(nav, candidate.node)?;
            for needle in &compound.contains {
                if !text.is_some_and(|text| contains_ignore_case(text, needle)) {
                    return Ok(false);
                }
            }
            if compound.empty && text.is_some_and(|text| !text.trim().is_empty()) {
                return Ok(false);
            }
        }
        if let Some(alts) = &compound.no_alt {
            let Some(index) = alts
                .get(usize::from(candidate.kind.ordinal()))
                .copied()
                .flatten()
            else {
                return Ok(false);
            };
            let alt = probe.property(nav, candidate.node, usize::from(index))?;
            if alt.is_some_and(|alt| !is_empty(&alt)) {
                return Ok(false);
            }
        }
        for &(start, len) in &compound.not {
            for &chain in &plan.lists[start as usize..(start + len) as usize] {
                let plan_chain = plan.chains[chain as usize];
                if plan_chain.kinds & bit(candidate.kind) != 0
                    && self.shortcut_matches(plan_chain, chain, candidate, probe, 0)?
                {
                    return Ok(false);
                }
            }
        }
        for &(start, len) in &compound.has {
            if !self.has(start, len, candidate)? {
                return Ok(false);
            }
        }
        Ok(true)
    }

    /// `*`: a direct child of the scope, or of a top-level container when
    /// the query names no scope.
    fn top_level(&self, candidate: &Candidate<'_, 'f, N>) -> bool {
        let Some(parent) = candidate.parent else {
            return false;
        };
        match self.plan.scope {
            ScopePlan::Root => self
                .state
                .frames
                .get(parent)
                .is_some_and(|frame| bit(frame.kind) & self.plan.top_parents != 0),
            _ => parent == 0,
        }
    }

    fn structural(&mut self, test: Structural, candidate: &Candidate<'_, 'f, N>) -> Result<bool> {
        Ok(match test {
            Structural::Sibling(position) => candidate.ordinal as usize == position.get(),
            Structural::LastSibling => match candidate.parent {
                Some(parent) => {
                    candidate.ordinal + 1 == self.sibling_count(parent, candidate.kind)?
                },
                None => false,
            },
            Structural::Number(number) => self.nav.number(candidate.node) == Some(number),
            Structural::Placeholder => self.nav.is_placeholder(candidate.node),
            Structural::Column(column) => self
                .nav
                .cell_at(candidate.node)
                .is_some_and(|(cell_column, _)| cell_column == column),
            Structural::Range { first, last } => {
                self.nav
                    .cell_at(candidate.node)
                    .is_some_and(|(column, row)| {
                        (first.0..=last.0).contains(&column) && (first.1..=last.1).contains(&row)
                    })
            },
        })
    }

    /// The number of children of kind `kind` of frame `parent`, read once
    /// per frame and kind.
    fn sibling_count(&mut self, parent: usize, kind: N::Kind) -> Result<u32> {
        let serial = self.state.frames[parent].serial;
        if let Some((cached, cached_kind, count)) = self.state.siblings {
            if cached == serial && cached_kind == kind.ordinal() {
                return Ok(count);
            }
        }
        let mut count = 0_u32;
        for child in self
            .nav
            .children(&self.state.frames[parent].node, KindSet::of(kind))?
        {
            child?;
            self.charge_visit()?;
            count += 1;
        }
        self.state.siblings = Some((serial, kind.ordinal(), count));
        Ok(count)
    }

    /// A worksheet row a header-row compound may match: in the bound table's
    /// worksheet and among its data rows.
    fn header_gate(&self, header: u32, candidate: &Candidate<'_, 'f, N>) -> bool {
        let sheet_kind = self.plan.headers[header as usize].sheet_kind;
        let binding = &self.state.bindings[header as usize];
        let mut cursor = candidate.parent;
        let mut sheet = None;
        while let Some(index) = cursor {
            let frame = &self.state.frames[index];
            if frame.kind.ordinal() == sheet_kind {
                sheet = Some(index);
                break;
            }
            cursor = frame.parent;
        }
        let in_sheet = match sheet {
            // The scope itself is the worksheet the binding was made in.
            Some(0) => !matches!(self.plan.scope, ScopePlan::Root),
            Some(index) => self.state.frames[index].ordinal as usize == binding.sheet.get(),
            None => false,
        };
        let (first, last) = binding.rows;
        in_sheet
            && self
                .nav
                .number(candidate.node)
                .is_some_and(|row| (first as usize..=last as usize).contains(&row))
    }

    fn expr<'c>(
        &mut self,
        id: u32,
        candidate: &Candidate<'c, 'f, N>,
        probe: &mut Probe<'c>,
    ) -> Result<bool>
    where
        'e: 'c,
    {
        let plan = self.plan;
        match plan.exprs[id as usize] {
            ExprNode::Cond(cond) => self.cond(cond, candidate, probe),
            ExprNode::And(start, len) => {
                for &arg in &plan.expr_args[start as usize..(start + len) as usize] {
                    if !self.expr(arg, candidate, probe)? {
                        return Ok(false);
                    }
                }
                Ok(true)
            },
            ExprNode::Or(start, len) => {
                for &arg in &plan.expr_args[start as usize..(start + len) as usize] {
                    if self.expr(arg, candidate, probe)? {
                        return Ok(true);
                    }
                }
                Ok(false)
            },
            ExprNode::Not(inner) => Ok(!self.expr(inner, candidate, probe)?),
        }
    }

    fn cond<'c>(
        &mut self,
        id: u32,
        candidate: &Candidate<'c, 'f, N>,
        probe: &mut Probe<'c>,
    ) -> Result<bool>
    where
        'e: 'c,
    {
        let plan = self.plan;
        let nav: &'c N = self.nav;
        let cond = &plan.conds[id as usize];
        let target = cond
            .by_kind
            .get(usize::from(candidate.kind.ordinal()))
            .and_then(|&index| cond.targets.get(usize::from(index)))
            .or_else(|| cond.targets.first())
            .copied()
            .ok_or_else(|| Error::new(Code::InternalError, "a query condition has no target"))?;
        let literal = cond.literal.as_ref();
        Ok(match target {
            Target::Absent(test) => test.eval(literal, None),
            Target::Prop { index, test } => {
                let value = if cond.direct || N::INDIVIDUAL_PROPERTIES {
                    nav.prop(candidate.node, usize::from(index))?
                } else {
                    probe.format(nav, candidate.node)?.get(usize::from(index))
                };
                test.eval(literal, value.as_ref().map(Value::as_scalar))
            },
            Target::Text(test) => {
                test.eval(literal, probe.text(nav, candidate.node)?.map(Scalar::Text))
            },
            Target::Type(test) => {
                test.eval(literal, Some(Scalar::Text(nav.type_name(candidate.node))))
            },
            Target::Style { prop, name, test } => {
                let declared = prop
                    .map(|index| probe.property(nav, candidate.node, usize::from(index)))
                    .transpose()?
                    .flatten();
                let display = name
                    .map(|index| probe.property(nav, candidate.node, usize::from(index)))
                    .transpose()?
                    .flatten();
                let member = probe.style(nav, candidate.node)?;
                let values = [
                    declared.as_ref().map(Value::as_scalar),
                    member.map(Scalar::Text),
                    display.as_ref().map(Value::as_scalar),
                ];
                match test.polarity() {
                    // OfficeCLI's dual style key: `=` holds when any form
                    // equals, `!=` when none does.
                    Some(true) => values
                        .into_iter()
                        .flatten()
                        .any(|value| test.eval(literal, Some(value))),
                    Some(false) => values
                        .into_iter()
                        .flatten()
                        .all(|value| test.eval(literal, Some(value))),
                    None => test.eval(literal, values[0].or(values[1])),
                }
            },
            Target::Column { header, key, test } => {
                let column = self.state.bindings[usize::from(header)].columns[usize::from(key)];
                let raw = plan.headers[usize::from(header)].raw[usize::from(key)];
                let value = nav.row_cell(candidate.node, column, raw)?;
                // A row without a cell there compares as empty text, as
                // OfficeCLI's probe row does.
                let scalar = value.as_ref().map_or(Scalar::Text(""), Value::as_scalar);
                test.eval(literal, Some(scalar))
            },
        })
    }

    /// `:has(list)`: some descendant of the candidate matches a chain of the
    /// list, with its ancestor links matched below the candidate. Stops at
    /// the first hit.
    fn has(&mut self, start: u32, len: u32, candidate: &Candidate<'_, 'f, N>) -> Result<bool> {
        let plan = self.plan;
        let chains = &plan.lists[start as usize..(start + len) as usize];
        let interest = chains
            .iter()
            .fold(0, |bits, &chain| bits | plan.chains[chain as usize].kinds);
        let reachable = plan
            .reach
            .get(usize::from(candidate.kind.ordinal()))
            .is_some_and(|reach| reach & interest != 0);
        if !reachable {
            return Ok(false);
        }
        let barrier = self.state.frames.len();
        let kinds = self.relevant(interest);
        self.push_with_parent(
            candidate.node.clone(),
            candidate.kind,
            candidate.ordinal,
            kinds,
            Source::Lazy,
            candidate.parent,
        )?;
        let result = self.has_walk(barrier, chains, interest);
        self.truncate(barrier);
        result
    }

    fn has_walk(&mut self, barrier: usize, chains: &[u32], interest: u128) -> Result<bool> {
        let plan = self.plan;
        while self.state.frames.len() > barrier {
            let Some(item) = self.next_child() else {
                self.pop();
                continue;
            };
            let node = item?;
            self.charge_visit()?;
            let kind = self.nav.kind(&node);
            let ordinal = self.bump(kind);
            let candidate = Candidate {
                node: &node,
                kind,
                ordinal,
                parent: Some(self.state.frames.len() - 1),
            };
            let mut probe = Probe::new();
            for &chain in chains {
                let plan_chain = plan.chains[chain as usize];
                if plan_chain.kinds & bit(kind) != 0
                    && self.shortcut_matches(
                        plan_chain,
                        chain,
                        &candidate,
                        &mut probe,
                        barrier + 1,
                    )?
                {
                    return Ok(true);
                }
            }
            let descend = plan
                .reach
                .get(usize::from(kind.ordinal()))
                .is_some_and(|reach| reach & interest != 0);
            if descend {
                let kinds = self.relevant(interest);
                self.push(node, kind, ordinal, kinds, Source::Lazy)?;
            }
        }
        Ok(false)
    }

    /// The kinds worth asking the adapter for under `interest`: the wanted
    /// kinds and the kinds that can contain one.
    fn relevant(&mut self, interest: u128) -> u128 {
        let cache = &mut self.state.relevant;
        if let Some(&(_, kinds)) = cache.iter().find(|(key, _)| *key == interest && *key != 0) {
            return kinds;
        }
        let mut kinds = interest;
        for (ordinal, reach) in self.plan.reach.iter().enumerate() {
            if reach & interest != 0 {
                kinds |= 1 << ordinal;
            }
        }
        cache[1] = cache[0];
        cache[0] = (interest, kinds);
        kinds
    }

    /// Pushes a child of the top frame.
    fn push(
        &mut self,
        node: N::Node,
        kind: N::Kind,
        ordinal: u32,
        kinds: u128,
        source: Source<'f, N>,
    ) -> Result<()> {
        let parent = self.state.frames.len().checked_sub(1);
        self.push_with_parent(node, kind, ordinal, kinds, source, parent)
    }

    fn push_with_parent(
        &mut self,
        node: N::Node,
        kind: N::Kind,
        ordinal: u32,
        kinds: u128,
        source: Source<'f, N>,
        parent: Option<usize>,
    ) -> Result<()> {
        let state = &mut *self.state;
        let index = state.frames.len();
        if index >= 256 {
            return Err(Error::new(
                Code::MaxDepthExceeded,
                "query traversal exceeds the frame limit",
            ));
        }
        state.serial += 1;
        state.frames.push(Frame {
            node,
            kind,
            ordinal,
            kinds,
            source,
            counters: state.counters.len(),
            serial: state.serial,
            parent,
        });
        if state.words > 0 {
            let start = index * 2 * state.words;
            let end = start + 2 * state.words;
            if state.memo.len() < end {
                state.memo.resize(end, 0);
            }
            state.memo[start..end].fill(0);
        }
        Ok(())
    }

    fn charge_visit(&mut self) -> Result<()> {
        if self.state.visited >= 1_000_000 {
            return Err(Error::new(
                Code::LimitExceeded,
                "query traversal exceeds the visit limit",
            ));
        }
        self.state.visited += 1;
        Ok(())
    }

    fn pop(&mut self) {
        if let Some(frame) = self.state.frames.pop() {
            self.state.counters.truncate(frame.counters);
        }
    }

    fn truncate(&mut self, len: usize) {
        while self.state.frames.len() > len {
            self.pop();
        }
    }

    /// Counts a child of the top frame; returns its position among the
    /// frame's children of the same kind.
    fn bump(&mut self, kind: N::Kind) -> u32 {
        let state = &mut *self.state;
        let start = state.frames.last().map_or(0, |frame| frame.counters);
        let key = kind.ordinal();
        for entry in &mut state.counters[start..] {
            if entry.0 == key {
                let ordinal = entry.1;
                entry.1 += 1;
                return ordinal;
            }
        }
        state.counters.push((key, 1));
        0
    }

    fn memo_get(&self, frame: usize, slot: u32) -> Option<bool> {
        let words = self.state.words;
        if words == 0 {
            return None;
        }
        let (word, mask) = (slot as usize / 64, 1_u64 << (slot % 64));
        let base = frame * 2 * words;
        let known = self.state.memo.get(base + word)?;
        (known & mask != 0).then(|| self.state.memo[base + words + word] & mask != 0)
    }

    fn memo_set(&mut self, frame: usize, slot: u32, value: bool) {
        let words = self.state.words;
        if words == 0 {
            return;
        }
        let (word, mask) = (slot as usize / 64, 1_u64 << (slot % 64));
        let base = frame * 2 * words;
        if let Some(known) = self.state.memo.get_mut(base + word) {
            *known |= mask;
            let stored = &mut self.state.memo[base + words + word];
            if value {
                *stored |= mask;
            } else {
                *stored &= !mask;
            }
        }
    }
}
