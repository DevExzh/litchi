//! Hierarchical, thread-safe resource budgets.

use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use thiserror::Error;

const RESOURCE_COUNT: usize = 9;

/// Resource dimensions charged by parsing, editing, and serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Resource {
    Memory,
    InputBytes,
    OutputBytes,
    Objects,
    Depth,
    Work,
    /// Peak concurrency permits for worker threads or executor slots held by
    /// an operation.
    ///
    /// Reserved before the first task starts and released when the workers are
    /// joined, so a parent budget bounds the sum of every child session's
    /// width rather than each session's own policy value.
    Workers,
    /// Peak permits for concurrent positional reads.
    ///
    /// A read session reserves these permits before the first source read and
    /// releases them when its batch or operation completes. A shared parent
    /// therefore bounds read concurrency across ZIP, CFB and OPC sessions.
    IoConcurrency,
    /// Cumulative count of scheduled CPU work units.
    ///
    /// Counted in tasks, not bytes: one deflate, decompression, parse or
    /// validation unit is one task. Consumed like [`Resource::Work`] and never
    /// released.
    CpuTasks,
}

impl Resource {
    const fn index(self) -> usize {
        match self {
            Self::Memory => 0,
            Self::InputBytes => 1,
            Self::OutputBytes => 2,
            Self::Objects => 3,
            Self::Depth => 4,
            Self::Work => 5,
            Self::Workers => 6,
            Self::IoConcurrency => 7,
            Self::CpuTasks => 8,
        }
    }
}

/// Named finite production profiles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Profile {
    Server,
    Desktop,
    TrustedBatch,
}

/// Finite limits for every resource dimension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    values: [u64; RESOURCE_COUNT],
}

impl Limits {
    /// Creates a fully explicit finite limit set.
    ///
    /// [`Resource::Workers`], [`Resource::IoConcurrency`] and
    /// [`Resource::CpuTasks`] are left unbounded; the execution builders set
    /// them. Every caller that predates those dimensions therefore keeps
    /// exactly its current meaning.
    #[must_use]
    pub const fn new(
        memory: u64,
        input_bytes: u64,
        output_bytes: u64,
        objects: u64,
        depth: u64,
        work: u64,
    ) -> Self {
        Self {
            values: [
                memory,
                input_bytes,
                output_bytes,
                objects,
                depth,
                work,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ],
        }
    }

    /// Bounds worker and CPU-task execution dimensions while retaining an
    /// unbounded positional-read budget.
    ///
    /// `workers` is the peak number of worker threads or executor slots every
    /// operation charged against this budget may hold at once; `cpu_tasks` is
    /// the cumulative number of CPU work units they may schedule.
    #[must_use]
    pub const fn with_execution(mut self, workers: u64, cpu_tasks: u64) -> Self {
        self.values[Resource::Workers.index()] = workers;
        self.values[Resource::CpuTasks.index()] = cpu_tasks;
        self
    }

    /// Bounds workers, positional-read concurrency and cumulative CPU tasks.
    ///
    /// This additive builder keeps the two-argument [`Self::with_execution`]
    /// constructor source-compatible for callers that do not opt into a
    /// shared read budget.
    #[must_use]
    pub const fn with_execution_io(
        mut self,
        workers: u64,
        io_concurrency: u64,
        cpu_tasks: u64,
    ) -> Self {
        self.values[Resource::Workers.index()] = workers;
        self.values[Resource::IoConcurrency.index()] = io_concurrency;
        self.values[Resource::CpuTasks.index()] = cpu_tasks;
        self
    }

    /// Conservative named defaults. Workload-specific limits remain explicit.
    #[must_use]
    pub const fn for_profile(profile: Profile) -> Self {
        const MIB: u64 = 1024 * 1024;
        const GIB: u64 = 1024 * MIB;
        match profile {
            Profile::Server => {
                Self::new(256 * MIB, 2 * GIB, 4 * GIB, 10_000_000, 256, 1_000_000_000)
                    .with_execution_io(64, 64, 1_000_000)
            },
            Profile::Desktop => Self::new(GIB, 8 * GIB, 16 * GIB, 50_000_000, 512, 5_000_000_000)
                .with_execution_io(256, 256, 5_000_000),
            Profile::TrustedBatch => Self::new(
                4 * GIB,
                64 * GIB,
                128 * GIB,
                250_000_000,
                1024,
                50_000_000_000,
            )
            .with_execution_io(1024, 1024, 50_000_000),
        }
    }

    /// Returns the limit for one dimension.
    #[must_use]
    pub const fn get(self, resource: Resource) -> u64 {
        self.values[resource.index()]
    }
}

#[derive(Debug)]
struct Node {
    scope: Arc<str>,
    limits: Limits,
    used: [AtomicU64; RESOURCE_COUNT],
    parent: Option<Arc<Node>>,
}

impl Node {
    fn new(scope: Arc<str>, limits: Limits, parent: Option<Arc<Node>>) -> Self {
        Self {
            scope,
            limits,
            used: std::array::from_fn(|_| AtomicU64::new(0)),
            parent,
        }
    }
}

/// Clone-cheap handle to a hierarchical resource budget.
#[derive(Debug, Clone)]
pub struct Budget {
    node: Arc<Node>,
}

impl Budget {
    /// Creates a root budget.
    pub fn root(scope: impl Into<Arc<str>>, limits: Limits) -> Self {
        Self {
            node: Arc::new(Node::new(scope.into(), limits, None)),
        }
    }

    /// Creates a child charged both locally and against every ancestor.
    #[must_use]
    pub fn child(&self, scope: impl Into<Arc<str>>, limits: Limits) -> Self {
        Self {
            node: Arc::new(Node::new(scope.into(), limits, Some(self.node.clone()))),
        }
    }

    /// Reserves outstanding capacity and releases it when the token is dropped.
    ///
    /// # Errors
    ///
    /// Returns `ResourceLimit` if charging `amount` would exceed the limit of
    /// this budget or any ancestor.
    pub fn reserve(&self, resource: Resource, amount: u64) -> Result<Reservation, ResourceLimit> {
        charge_chain(&self.node, resource, amount)?;
        // One handle to the charged node is enough: it keeps every ancestor
        // alive through its parent links, and the release walks that same
        // immutable chain.
        Ok(Reservation {
            node: Arc::clone(&self.node),
            resource,
            amount,
        })
    }

    /// Reserves outstanding capacity for as long as the returned token borrows
    /// this budget.
    ///
    /// The charge, its refusal, its release on drop and
    /// [`ScopedReservation::commit`] are exactly those of [`Self::reserve`]
    /// and [`Reservation::commit`]. The token borrows this handle instead of
    /// sharing it, so neither making nor dropping it touches a reference
    /// count; prefer it when the reservation does not outlive the operation
    /// that holds the budget, such as one sequential write.
    ///
    /// # Errors
    ///
    /// Returns `ResourceLimit` if charging `amount` would exceed the limit of
    /// this budget or any ancestor.
    pub fn reserve_scoped(
        &self,
        resource: Resource,
        amount: u64,
    ) -> Result<ScopedReservation<'_>, ResourceLimit> {
        charge_chain(&self.node, resource, amount)?;
        Ok(ScopedReservation {
            node: &self.node,
            resource,
            amount,
        })
    }

    /// Charges cumulative work that is not released during this budget's life.
    ///
    /// # Errors
    ///
    /// Returns `ResourceLimit` if charging `amount` would exceed the limit of
    /// this budget or any ancestor.
    pub fn consume(&self, resource: Resource, amount: u64) -> Result<(), ResourceLimit> {
        charge_chain(&self.node, resource, amount)
    }

    /// Current local usage for one resource.
    #[must_use]
    pub fn used(&self, resource: Resource) -> u64 {
        self.node.used[resource.index()].load(Ordering::Acquire)
    }

    /// Current local limit for one resource.
    #[must_use]
    pub fn limit(&self, resource: Resource) -> u64 {
        self.node.limits.get(resource)
    }
}

/// Charges `amount` against `leaf` and then each ancestor in turn.
///
/// Every level is one atomic check-and-add, so no level is ever observed above
/// its limit. When a level refuses, the levels already charged are released
/// before the refusal is returned, leaving the hierarchy as it was found. The
/// walk borrows the immutable parent chain; it takes no reference counts.
fn charge_chain(leaf: &Node, resource: Resource, amount: u64) -> Result<(), ResourceLimit> {
    let mut charged = 0usize;
    let mut current = Some(leaf);
    while let Some(node) = current {
        let counter = &node.used[resource.index()];
        let limit = node.limits.get(resource);
        let result = counter.fetch_update(Ordering::AcqRel, Ordering::Acquire, |used| {
            used.checked_add(amount).filter(|next| *next <= limit)
        });
        match result {
            Ok(_) => charged = charged.saturating_add(1),
            Err(used) => {
                release_ancestor_prefix(leaf, charged, resource, amount);
                return Err(ResourceLimit {
                    resource,
                    observed: used.saturating_add(amount),
                    limit,
                    scope: node.scope.clone(),
                });
            },
        }
        current = node.parent.as_deref();
    }
    Ok(())
}

/// RAII token for outstanding budget usage.
///
/// The token holds one handle to the budget node it charged. That node owns
/// its ancestors through its parent links, so releasing walks the same chain
/// the charge walked, whatever the depth, without per-level handles.
#[derive(Debug)]
pub struct Reservation {
    node: Arc<Node>,
    resource: Resource,
    amount: u64,
}

impl Reservation {
    /// Reserved amount.
    #[must_use]
    pub const fn amount(&self) -> u64 {
        self.amount
    }

    /// Reserved resource kind.
    #[must_use]
    pub const fn resource(&self) -> Resource {
        self.resource
    }

    /// Merges another already-charged reservation into this token.
    ///
    /// The reservations can be merged only when they charge the same
    /// resource through the exact same budget-node chain.  A successful merge
    /// keeps one chain and makes the consumed token inert, so the combined
    /// amount is released exactly once when this token is dropped.  The
    /// counters are already charged by both reservations and therefore do not
    /// change during a merge.
    ///
    /// # Errors
    ///
    /// Returns `other` unchanged when either reservation has a different
    /// resource or node chain, or when their amounts cannot be added.  `self`
    /// is unchanged on every error path.
    pub fn try_merge(&mut self, mut other: Reservation) -> Result<(), Reservation> {
        // A token holds only the node it charged, and that node owns its
        // ancestors through immutable parent links, so two tokens charge the
        // exact same chain precisely when they hold the same node.
        if self.resource != other.resource || !Arc::ptr_eq(&self.node, &other.node) {
            return Err(other);
        }

        let Some(amount) = self.amount.checked_add(other.amount) else {
            return Err(other);
        };

        self.amount = amount;
        // A zero amount makes the consumed token inert: its `Drop` releases
        // nothing, so the combined amount is released once, by `self`.
        other.amount = 0;
        Ok(())
    }

    /// Commits at most the reserved amount as cumulative usage.
    ///
    /// A reservation normally releases all of its charge when dropped.  A
    /// sequential writer can instead preflight a maximum write, perform the
    /// sink operation, and commit the exact number of bytes accepted without
    /// releasing the charge into a race window.  Returns `false` when
    /// `amount` exceeds the reservation; in that case the reservation is
    /// released normally and no cumulative usage is retained.
    #[must_use = "check whether the requested amount was committed"]
    pub fn commit(mut self, amount: u64) -> bool {
        if amount > self.amount {
            return false;
        }
        if amount < self.amount {
            release_chain(&self.node, self.resource, self.amount - amount);
        }
        // Nothing remains outstanding: the committed amount stays charged as
        // cumulative usage and `Drop` has nothing left to release.
        self.amount = 0;
        true
    }
}

impl Drop for Reservation {
    fn drop(&mut self) {
        if self.amount != 0 {
            release_chain(&self.node, self.resource, self.amount);
        }
    }
}

/// RAII token for outstanding budget usage that borrows its budget.
///
/// Made by [`Budget::reserve_scoped`]. It charges, commits and releases
/// exactly as a [`Reservation`] does; it differs only in borrowing the budget
/// handle for its lifetime instead of holding a shared handle of its own.
#[derive(Debug)]
pub struct ScopedReservation<'budget> {
    node: &'budget Node,
    resource: Resource,
    amount: u64,
}

impl ScopedReservation<'_> {
    /// Reserved amount.
    #[must_use]
    pub const fn amount(&self) -> u64 {
        self.amount
    }

    /// Reserved resource kind.
    #[must_use]
    pub const fn resource(&self) -> Resource {
        self.resource
    }

    /// Commits at most the reserved amount as cumulative usage.
    ///
    /// Behaves exactly as [`Reservation::commit`]: the unused remainder is
    /// released, and an `amount` larger than the reservation returns `false`
    /// and releases the whole reservation without retaining any usage.
    #[must_use = "check whether the requested amount was committed"]
    pub fn commit(mut self, amount: u64) -> bool {
        if amount > self.amount {
            return false;
        }
        if amount < self.amount {
            release_chain(self.node, self.resource, self.amount - amount);
        }
        self.amount = 0;
        true
    }
}

impl Drop for ScopedReservation<'_> {
    fn drop(&mut self) {
        if self.amount != 0 {
            release_chain(self.node, self.resource, self.amount);
        }
    }
}

/// A structured resource-limit failure.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[error("{resource:?} budget exceeded in {scope}: observed {observed}, limit {limit}")]
pub struct ResourceLimit {
    pub resource: Resource,
    pub observed: u64,
    pub limit: u64,
    pub scope: Arc<str>,
}

/// Releases `amount` from `leaf` and every ancestor.
fn release_chain(leaf: &Node, resource: Resource, amount: u64) {
    let mut current = Some(leaf);
    while let Some(node) = current {
        release_node(node, resource, amount);
        current = node.parent.as_deref();
    }
}

fn release_ancestor_prefix(mut node: &Node, count: usize, resource: Resource, amount: u64) {
    for _ in 0..count {
        release_node(node, resource, amount);
        let Some(parent) = node.parent.as_deref() else {
            break;
        };
        node = parent;
    }
}

fn release_node(node: &Node, resource: Resource, amount: u64) {
    let counter = &node.used[resource.index()];
    let _prev = counter.fetch_update(Ordering::AcqRel, Ordering::Acquire, |used| {
        Some(used.saturating_sub(amount))
    });
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic by design"
    )]

    use super::*;

    fn limits(memory: u64) -> Limits {
        Limits::new(memory, 100, 100, 100, 100, 100)
    }

    #[test]
    fn reservations_release_capacity() {
        let budget = Budget::root("document", limits(10));
        let first = budget.reserve(Resource::Memory, 7).expect("within limit");
        assert_eq!(budget.used(Resource::Memory), 7);
        assert!(budget.reserve(Resource::Memory, 4).is_err());
        drop(first);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert!(budget.reserve(Resource::Memory, 10).is_ok());
    }

    #[test]
    fn reservations_can_commit_an_exact_short_write() {
        let budget = Budget::root("document", limits(10));
        let reservation = budget.reserve(Resource::Memory, 7).expect("reserve");
        assert!(reservation.commit(3));
        assert_eq!(budget.used(Resource::Memory), 3);
        assert!(budget.reserve(Resource::Memory, 7).is_ok());
    }

    #[test]
    fn reservations_merge_for_clone_handles_without_changing_counters() {
        let budget = Budget::root("document", limits(20));
        let mut first = budget.reserve(Resource::Memory, 7).expect("first reserve");
        let second = budget
            .clone()
            .reserve(Resource::Memory, 5)
            .expect("clone reserve");
        assert_eq!(budget.used(Resource::Memory), 12);

        assert!(first.try_merge(second).is_ok());
        assert_eq!(first.amount(), 12);
        assert_eq!(budget.used(Resource::Memory), 12);

        drop(first);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn reservations_merge_preserves_parent_and_child_accounting() {
        let root = Budget::root("document", limits(20));
        let child = root.child("worksheet", limits(20));
        let mut first = child
            .reserve(Resource::OutputBytes, 7)
            .expect("first reserve");
        let second = child
            .clone()
            .reserve(Resource::OutputBytes, 5)
            .expect("clone reserve");
        assert_eq!(root.used(Resource::OutputBytes), 12);
        assert_eq!(child.used(Resource::OutputBytes), 12);

        assert!(first.try_merge(second).is_ok());
        assert_eq!(first.amount(), 12);
        assert_eq!(root.used(Resource::OutputBytes), 12);
        assert_eq!(child.used(Resource::OutputBytes), 12);

        drop(first);
        assert_eq!(root.used(Resource::OutputBytes), 0);
        assert_eq!(child.used(Resource::OutputBytes), 0);
    }

    #[test]
    fn reservations_refuse_different_chains_and_resources_without_accounting_changes() {
        let root = Budget::root("document", limits(20));
        let left = root.child("left", limits(20));
        let right = root.child("right", limits(20));
        let mut sibling = left.reserve(Resource::Memory, 2).expect("left reserve");
        let sibling_other = right.reserve(Resource::Memory, 3).expect("right reserve");
        assert_eq!(root.used(Resource::Memory), 5);
        let sibling_other = sibling
            .try_merge(sibling_other)
            .expect_err("sibling chains must not merge");
        assert_eq!(sibling.amount(), 2);
        assert_eq!(sibling_other.amount(), 3);
        assert_eq!(root.used(Resource::Memory), 5);
        assert_eq!(left.used(Resource::Memory), 2);
        assert_eq!(right.used(Resource::Memory), 3);
        drop(sibling);
        drop(sibling_other);
        assert_eq!(root.used(Resource::Memory), 0);

        let scoped_left = Budget::root("scope-left", limits(20));
        let scoped_right = Budget::root("scope-right", limits(20));
        let mut scoped = scoped_left
            .reserve(Resource::Memory, 2)
            .expect("scoped left reserve");
        let scoped_other = scoped_right
            .reserve(Resource::Memory, 3)
            .expect("scoped right reserve");
        let scoped_other = scoped
            .try_merge(scoped_other)
            .expect_err("different scope chains must not merge");
        assert_eq!(scoped.amount(), 2);
        assert_eq!(scoped_other.amount(), 3);
        assert_eq!(scoped_left.used(Resource::Memory), 2);
        assert_eq!(scoped_right.used(Resource::Memory), 3);
        drop(scoped);
        drop(scoped_other);

        let mut memory = root.reserve(Resource::Memory, 2).expect("memory reserve");
        let input = root
            .reserve(Resource::InputBytes, 3)
            .expect("input reserve");
        let input = memory
            .try_merge(input)
            .expect_err("different resources must not merge");
        assert_eq!(memory.amount(), 2);
        assert_eq!(memory.resource(), Resource::Memory);
        assert_eq!(input.amount(), 3);
        assert_eq!(input.resource(), Resource::InputBytes);
        assert_eq!(root.used(Resource::Memory), 2);
        assert_eq!(root.used(Resource::InputBytes), 3);
        drop(memory);
        drop(input);
        assert_eq!(root.used(Resource::Memory), 0);
        assert_eq!(root.used(Resource::InputBytes), 0);
    }

    #[test]
    fn zero_amount_reservations_merge_without_charging() {
        let budget = Budget::root("document", limits(20));
        let mut first = budget.reserve(Resource::Memory, 0).expect("first reserve");
        let second = budget.reserve(Resource::Memory, 0).expect("second reserve");

        assert_eq!(budget.used(Resource::Memory), 0);
        assert!(first.try_merge(second).is_ok());
        assert_eq!(first.amount(), 0);
        assert_eq!(budget.used(Resource::Memory), 0);

        drop(first);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn merged_reservation_can_commit_part_of_combined_amount() {
        let budget = Budget::root("document", limits(20));
        let mut first = budget.reserve(Resource::Memory, 7).expect("first reserve");
        let second = budget.reserve(Resource::Memory, 5).expect("second reserve");
        assert!(first.try_merge(second).is_ok());
        assert_eq!(budget.used(Resource::Memory), 12);

        assert!(first.commit(4));
        assert_eq!(budget.used(Resource::Memory), 4);
        assert!(budget.reserve(Resource::Memory, 16).is_ok());
    }

    #[test]
    fn merge_refuses_amount_overflow_without_mutating_either_token() {
        let budget = Budget::root("document", limits(u64::MAX));
        // Tokens built directly on one node: the charge-free construction
        // isolates the checked amount addition (releases saturate at zero).
        let mut first = Reservation {
            node: Arc::clone(&budget.node),
            resource: Resource::Memory,
            amount: u64::MAX,
        };
        let second = Reservation {
            node: Arc::clone(&budget.node),
            resource: Resource::Memory,
            amount: 1,
        };

        let second = first
            .try_merge(second)
            .expect_err("amount addition must be checked");
        assert_eq!(first.amount(), u64::MAX);
        assert_eq!(second.amount(), 1);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn committed_reservation_preserves_hierarchical_charge() {
        let root = Budget::root("document", limits(100));
        let child = root.child("worksheet", limits(100));
        let reservation = child.reserve(Resource::OutputBytes, 7).expect("reserve");
        assert_eq!(root.used(Resource::OutputBytes), 7);
        assert_eq!(child.used(Resource::OutputBytes), 7);
        assert!(reservation.commit(3));
        assert_eq!(root.used(Resource::OutputBytes), 3);
        assert_eq!(child.used(Resource::OutputBytes), 3);

        let exact = child
            .reserve(Resource::OutputBytes, 4)
            .expect("reserve exact");
        assert!(exact.commit(4));
        assert_eq!(root.used(Resource::OutputBytes), 7);
        assert_eq!(child.used(Resource::OutputBytes), 7);

        let zero = child
            .reserve(Resource::OutputBytes, 5)
            .expect("reserve zero");
        assert!(zero.commit(0));
        assert_eq!(root.used(Resource::OutputBytes), 7);
        assert_eq!(child.used(Resource::OutputBytes), 7);

        let over = child
            .reserve(Resource::OutputBytes, 5)
            .expect("reserve over");
        assert!(!over.commit(6));
        assert_eq!(root.used(Resource::OutputBytes), 7);
        assert_eq!(child.used(Resource::OutputBytes), 7);
        let released = child
            .reserve(Resource::OutputBytes, 93)
            .expect("over-commit must release its complete reservation");
        drop(released);
    }

    #[test]
    fn child_failure_rolls_back_every_level() {
        let root = Budget::root("document", limits(5));
        let child = root.child("worksheet", limits(10));
        let error = child
            .reserve(Resource::Memory, 6)
            .expect_err("parent must cap child");
        assert_eq!(error.scope.as_ref(), "document");
        assert_eq!(child.used(Resource::Memory), 0);
        assert_eq!(root.used(Resource::Memory), 0);
    }

    #[test]
    fn a_reservation_holds_one_handle_whatever_the_depth() {
        let root = Budget::root("root", limits(100));
        let child = root.child("child", limits(100));
        let grandchild = child.child("grandchild", limits(100));
        let leaf = grandchild.child("leaf", limits(100));
        let handles_before = Arc::strong_count(&leaf.node);
        let ancestor_handles_before = Arc::strong_count(&root.node);
        let reservation = leaf
            .reserve(Resource::Memory, 1)
            .expect("four-level reservation");

        // The charged node is shared once; no ancestor gains a handle.
        assert_eq!(Arc::strong_count(&leaf.node), handles_before + 1);
        assert_eq!(Arc::strong_count(&root.node), ancestor_handles_before);
        for budget in [&root, &child, &grandchild, &leaf] {
            assert_eq!(budget.used(Resource::Memory), 1);
        }
        drop(reservation);
        assert_eq!(Arc::strong_count(&leaf.node), handles_before);
        for budget in [&root, &child, &grandchild, &leaf] {
            assert_eq!(budget.used(Resource::Memory), 0);
        }
    }

    #[test]
    fn a_reservation_outlives_the_budget_handles_that_made_it() {
        let root = Budget::root("root", limits(100));
        let leaf = root.child("leaf", limits(100));
        let reservation = leaf.reserve(Resource::Memory, 7).expect("reserve");
        drop(leaf);
        assert_eq!(root.used(Resource::Memory), 7);
        // The reservation still releases the ancestor it charged.
        assert!(reservation.commit(3));
        assert_eq!(root.used(Resource::Memory), 3);
    }

    #[test]
    fn deep_hierarchies_roll_back_exactly() {
        let root = Budget::root("root", limits(1));
        let first = root.child("first", limits(100));
        let second = first.child("second", limits(100));
        let third = second.child("third", limits(100));
        let fourth = third.child("fourth", limits(100));
        let leaf = fourth.child("leaf", limits(100));

        let reservation = leaf
            .reserve(Resource::Memory, 1)
            .expect("six-level reservation");
        for budget in [&root, &first, &second, &third, &fourth, &leaf] {
            assert_eq!(budget.used(Resource::Memory), 1);
        }
        assert!(reservation.commit(1));

        let error = leaf
            .consume(Resource::Memory, 1)
            .expect_err("root limit must reject the deep charge");
        assert_eq!(error.scope.as_ref(), "root");
        let reservation_error = leaf
            .reserve(Resource::Memory, 1)
            .expect_err("spilled reservation must roll back after parent rejection");
        assert_eq!(reservation_error.scope.as_ref(), "root");
        for budget in [&root, &first, &second, &third, &fourth, &leaf] {
            assert_eq!(budget.used(Resource::Memory), 1);
        }
        leaf.consume(Resource::Work, 0)
            .expect("zero consumption must preserve the hierarchy");
        assert_eq!(root.used(Resource::Work), 0);
    }

    fn three_levels(tight: u64) -> [Budget; 3] {
        // The middle level is the tightest; the leaf and the root are looser.
        let root = Budget::root("root", limits(20));
        let middle = root.child("middle", limits(tight));
        let leaf = middle.child("leaf", limits(15));
        [root, middle, leaf]
    }

    #[test]
    fn limits_refuse_at_exactly_one_over_at_the_innermost_refusing_level() {
        for (charge, refused) in [(8, false), (9, false), (10, true)] {
            let consumed_levels = three_levels(9);
            let reserved_levels = three_levels(9);
            let consumed = consumed_levels[2].consume(Resource::Memory, charge);
            let reserved = reserved_levels[2].reserve(Resource::Memory, charge);
            if refused {
                for error in [
                    consumed.expect_err("consume one over"),
                    reserved.expect_err("reserve one over"),
                ] {
                    assert_eq!(error.resource, Resource::Memory);
                    assert_eq!(error.observed, charge);
                    assert_eq!(error.limit, 9);
                    assert_eq!(error.scope.as_ref(), "middle");
                }
                for budget in consumed_levels.iter().chain(&reserved_levels) {
                    assert_eq!(budget.used(Resource::Memory), 0);
                }
            } else {
                consumed.expect("consume within limit");
                assert!(reserved.expect("reserve within limit").commit(charge));
                for budget in consumed_levels.iter().chain(&reserved_levels) {
                    assert_eq!(budget.used(Resource::Memory), charge);
                }
                // The next unit is refused exactly when the tight level is full.
                for levels in [&consumed_levels, &reserved_levels] {
                    let next = levels[2].consume(Resource::Memory, 1);
                    assert_eq!(next.is_err(), charge == 9);
                }
            }
        }
    }

    #[test]
    fn scoped_reservations_charge_commit_and_release_exactly_as_owned_ones() {
        for (charge, commit) in [(8, 8), (9, 9), (9, 4), (9, 0), (9, 10), (10, 10)] {
            let owned_levels = three_levels(9);
            let scoped_levels = three_levels(9);
            let owned = owned_levels[2].reserve(Resource::Memory, charge);
            let scoped = scoped_levels[2].reserve_scoped(Resource::Memory, charge);
            match (owned, scoped) {
                (Ok(owned), Ok(scoped)) => {
                    assert_eq!(
                        (owned.amount(), owned.resource()),
                        (charge, Resource::Memory)
                    );
                    assert_eq!(
                        (scoped.amount(), scoped.resource()),
                        (charge, Resource::Memory)
                    );
                    for (owned, scoped) in owned_levels.iter().zip(&scoped_levels) {
                        assert_eq!(owned.used(Resource::Memory), scoped.used(Resource::Memory));
                    }
                    assert_eq!(owned.commit(commit), scoped.commit(commit));
                },
                (Err(owned), Err(scoped)) => assert_eq!(owned, scoped),
                (owned, scoped) => panic!("reserve {owned:?} but reserve_scoped {scoped:?}"),
            }
            for (owned, scoped) in owned_levels.iter().zip(&scoped_levels) {
                assert_eq!(owned.used(Resource::Memory), scoped.used(Resource::Memory));
            }
        }

        // Dropping an uncommitted scoped reservation releases every level.
        let [root, middle, leaf] = three_levels(9);
        let scoped = leaf.reserve_scoped(Resource::Memory, 6).expect("reserve");
        assert_eq!(root.used(Resource::Memory), 6);
        assert!(leaf.reserve_scoped(Resource::Memory, 4).is_err());
        drop(scoped);
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Memory), 0);
        }
    }

    #[test]
    fn partial_and_zero_commits_release_every_level() {
        let root = Budget::root("root", limits(100));
        let middle = root.child("middle", limits(100));
        let leaf = middle.child("leaf", limits(100));
        let reservation = leaf.reserve(Resource::Memory, 10).expect("reserve");
        assert!(reservation.commit(4));
        let empty = leaf.reserve(Resource::Memory, 10).expect("reserve");
        assert!(empty.commit(0));
        let zero = leaf.reserve(Resource::Memory, 0).expect("zero reserve");
        drop(zero);
        let refused = leaf.reserve(Resource::Memory, 10).expect("reserve");
        assert!(!refused.commit(11));
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Memory), 4);
        }
    }

    #[test]
    fn concurrent_reservations_never_exceed_limit() {
        let budget = Budget::root("document", limits(1));
        let barrier = Arc::new(std::sync::Barrier::new(2));
        std::thread::scope(|scope| {
            let first = budget.clone();
            let second = budget.clone();
            let first_barrier = barrier.clone();
            let second_barrier = barrier.clone();
            let left = scope.spawn(move || {
                let reservation = first.reserve(Resource::Memory, 1).ok();
                first_barrier.wait();
                reservation.is_some()
            });
            let right = scope.spawn(move || {
                let reservation = second.reserve(Resource::Memory, 1).ok();
                second_barrier.wait();
                reservation.is_some()
            });
            let successes =
                u8::from(left.join().unwrap_or(false)) + u8::from(right.join().unwrap_or(false));
            assert_eq!(successes, 1);
        });
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn concurrent_consumption_never_exceeds_limit() {
        let budget = Budget::root("document", limits(1));
        let barrier = Arc::new(std::sync::Barrier::new(2));
        std::thread::scope(|scope| {
            let first = budget.clone();
            let second = budget.clone();
            let first_barrier = barrier.clone();
            let second_barrier = barrier.clone();
            let left = scope.spawn(move || {
                first_barrier.wait();
                first.consume(Resource::Memory, 1).is_ok()
            });
            let right = scope.spawn(move || {
                second_barrier.wait();
                second.consume(Resource::Memory, 1).is_ok()
            });
            let successes =
                u8::from(left.join().unwrap_or(false)) + u8::from(right.join().unwrap_or(false));
            assert_eq!(successes, 1);
        });
        assert_eq!(budget.used(Resource::Memory), 1);
    }

    /// The concurrency test's hierarchy, by index:
    ///
    /// ```text
    /// root ─┬─ left ─┬─ left_a
    ///       │        └─ left_b
    ///       └─ right ── right_a
    /// ```
    const TREE_NAMES: [&str; 6] = ["root", "left", "left_a", "left_b", "right", "right_a"];
    const TREE_PARENT: [Option<usize>; 6] = [None, Some(0), Some(1), Some(1), Some(0), Some(4)];
    const TREE_MEMORY: [u64; 6] = [48, 32, 20, 20, 32, 20];
    const TREE_WORK: [u64; 6] = [9_000, 6_000, 2_500, 2_500, 4_000, 3_000];
    /// The node each worker charges: two siblings under `left`, the leaf under
    /// `right`, a middle level and the root itself.
    const TREE_WORKER_NODE: [usize; 8] = [2, 2, 3, 3, 5, 5, 1, 0];
    const TREE_STEPS: usize = 4_000;

    /// A node and its ancestors, innermost first.
    fn tree_chain(mut node: usize) -> Vec<usize> {
        let mut chain = vec![node];
        while let Some(parent) = TREE_PARENT[node] {
            chain.push(parent);
            node = parent;
        }
        chain
    }

    #[derive(Debug, Default)]
    struct TreeWorkerOutcome {
        work_granted: u64,
        consume_demand: u64,
        work_refusals: u64,
    }

    /// One worker's fixed pseudo-random sequence of owned and scoped
    /// reservations, commits and releases, and consumptions. Memory is only
    /// ever reserved and then released (dropped, committed at zero or
    /// over-committed); Work is kept exactly where the budget granted it.
    fn tree_worker(budget: &Budget, chain: &[usize], seed: u64) -> TreeWorkerOutcome {
        let refused = |error: &ResourceLimit, resource: Resource| {
            let level = chain
                .iter()
                .copied()
                .find(|&level| TREE_NAMES[level] == &*error.scope)
                .expect("a refusal names a level of the charged chain");
            let limit = if resource == Resource::Memory {
                TREE_MEMORY[level]
            } else {
                TREE_WORK[level]
            };
            assert_eq!(error.resource, resource);
            assert_eq!(error.limit, limit);
            assert!(error.observed > limit);
        };
        let mut outcome = TreeWorkerOutcome::default();
        let mut owned: Vec<Reservation> = Vec::new();
        let mut scoped: Vec<ScopedReservation<'_>> = Vec::new();
        let mut state = seed;
        for _ in 0..TREE_STEPS {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            let amount = 1 + state % 6;
            let release = (state >> 16) % 3;
            match (state >> 8) % 8 {
                0 => match budget.reserve(Resource::Memory, amount) {
                    Ok(reservation) if owned.len() < 3 => owned.push(reservation),
                    Ok(reservation) => drop(reservation),
                    Err(error) => refused(&error, Resource::Memory),
                },
                1 => match budget.reserve_scoped(Resource::Memory, amount) {
                    Ok(reservation) if scoped.len() < 2 => scoped.push(reservation),
                    Ok(reservation) => drop(reservation),
                    Err(error) => refused(&error, Resource::Memory),
                },
                2 => {
                    if let Some(reservation) = owned.pop() {
                        let held = reservation.amount();
                        match release {
                            0 => drop(reservation),
                            1 => assert!(reservation.commit(0)),
                            _ => assert!(!reservation.commit(held + 1)),
                        }
                    }
                },
                3 => {
                    if let Some(reservation) = scoped.pop() {
                        let held = reservation.amount();
                        match release {
                            0 => drop(reservation),
                            1 => assert!(reservation.commit(0)),
                            _ => assert!(!reservation.commit(held + 1)),
                        }
                    }
                },
                4 => match budget.reserve(Resource::Work, amount) {
                    Ok(reservation) => {
                        let keep = (state >> 20) % (amount + 1);
                        assert!(reservation.commit(keep));
                        outcome.work_granted += keep;
                    },
                    Err(error) => {
                        refused(&error, Resource::Work);
                        outcome.work_refusals += 1;
                    },
                },
                5 => match budget.reserve_scoped(Resource::Work, amount) {
                    Ok(reservation) => {
                        let keep = (state >> 20) % (amount + 1);
                        assert!(reservation.commit(keep));
                        outcome.work_granted += keep;
                    },
                    Err(error) => {
                        refused(&error, Resource::Work);
                        outcome.work_refusals += 1;
                    },
                },
                6 => {
                    outcome.consume_demand += amount;
                    match budget.consume(Resource::Work, amount) {
                        Ok(()) => outcome.work_granted += amount,
                        Err(error) => {
                            refused(&error, Resource::Work);
                            outcome.work_refusals += 1;
                        },
                    }
                },
                _ => match budget.reserve_scoped(Resource::Memory, amount) {
                    Ok(reservation) => drop(reservation),
                    Err(error) => refused(&error, Resource::Memory),
                },
            }
        }
        drop(scoped);
        drop(owned);
        outcome
    }

    /// Stops the monitor even when a worker's panic unwinds the test.
    struct StopOnDrop<'flag>(&'flag std::sync::atomic::AtomicBool);

    impl Drop for StopOnDrop<'_> {
        fn drop(&mut self) {
            self.0.store(true, Ordering::Release);
        }
    }

    #[test]
    fn concurrent_owned_and_scoped_reservations_keep_a_shared_hierarchy_exact() {
        let mut budgets: Vec<Budget> = Vec::new();
        for index in 0..TREE_NAMES.len() {
            let limits = Limits::new(TREE_MEMORY[index], 100, 100, 100, 100, TREE_WORK[index]);
            let budget = match TREE_PARENT[index] {
                None => Budget::root(TREE_NAMES[index], limits),
                Some(parent) => budgets[parent].child(TREE_NAMES[index], limits),
            };
            budgets.push(budget);
        }
        let stop = std::sync::atomic::AtomicBool::new(false);
        let outcomes: Vec<TreeWorkerOutcome> = std::thread::scope(|scope| {
            // Every read of every level, while the workers run, is within its
            // limit: no charge is ever visible above a limit, even briefly.
            let monitor = scope.spawn(|| {
                let mut passes = 0_u64;
                loop {
                    let last = stop.load(Ordering::Acquire);
                    for (index, budget) in budgets.iter().enumerate() {
                        let memory = budget.used(Resource::Memory);
                        let work = budget.used(Resource::Work);
                        assert!(memory <= TREE_MEMORY[index], "{}", TREE_NAMES[index]);
                        assert!(work <= TREE_WORK[index], "{}", TREE_NAMES[index]);
                    }
                    passes += 1;
                    if last {
                        break passes;
                    }
                    std::thread::yield_now();
                }
            });
            let stop_monitor = StopOnDrop(&stop);
            let workers: Vec<_> = TREE_WORKER_NODE
                .iter()
                .enumerate()
                .map(|(worker, &node)| {
                    let budget = budgets[node].clone();
                    let seed = 0x9E37_79B9_7F4A_7C15_u64
                        ^ (u64::try_from(worker).expect("worker index") + 1)
                            .wrapping_mul(0x2545_F491_4F6C_DD1D);
                    scope.spawn(move || tree_worker(&budget, &tree_chain(node), seed))
                })
                .collect();
            let outcomes = workers
                .into_iter()
                .map(|worker| worker.join().expect("worker"))
                .collect();
            drop(stop_monitor);
            assert!(monitor.join().expect("monitor") > 0);
            outcomes
        });

        // Every level ends exactly at what its subtree was granted: no
        // reserved Memory survives, and Work is the sum of the kept commits
        // and consumptions of every worker that charges through the level.
        for (index, budget) in budgets.iter().enumerate() {
            assert_eq!(budget.used(Resource::Memory), 0, "{}", TREE_NAMES[index]);
            let granted: u64 = TREE_WORKER_NODE
                .iter()
                .zip(&outcomes)
                .filter(|(node, _)| tree_chain(**node).contains(&index))
                .map(|(_, outcome)| outcome.work_granted)
                .sum();
            assert_eq!(
                budget.used(Resource::Work),
                granted,
                "{}",
                TREE_NAMES[index]
            );
        }
        // The seeds ask for more Work through `consume` alone than the root
        // allows, so some Work charge must have been refused, whatever the
        // interleaving; each refusal named a level of its chain and its limit.
        let demand: u64 = outcomes.iter().map(|outcome| outcome.consume_demand).sum();
        assert!(demand > TREE_WORK[0]);
        assert!(outcomes.iter().any(|outcome| outcome.work_refusals > 0));
    }

    #[test]
    fn execution_io_builder_composes_all_runtime_dimensions() {
        let limits = Limits::new(1, 2, 3, 4, 5, 6).with_execution_io(7, 8, 9);
        assert_eq!(limits.get(Resource::Workers), 7);
        assert_eq!(limits.get(Resource::IoConcurrency), 8);
        assert_eq!(limits.get(Resource::CpuTasks), 9);

        let legacy = Limits::new(1, 2, 3, 4, 5, 6).with_execution(7, 9);
        assert_eq!(legacy.get(Resource::Workers), 7);
        assert_eq!(legacy.get(Resource::IoConcurrency), u64::MAX);
        assert_eq!(legacy.get(Resource::CpuTasks), 9);
    }
}
